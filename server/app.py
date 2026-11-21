"""Docxify: local PDF -> Word service.

A small task-based HTTP API (the web UI in ``web/`` is its only client):

  POST   /v1/start                     -> {"server", "task"}
  POST   /v1/upload                    chunked (2 MB): {"chunks","chunk"} ... then
                                       {"server_filename","scanned","pages",...}
  POST   /v1/process                   convert_to=docx, ocr=0|1, ocr_lang=en|latin|ar, ocr_review=0|1,
                                       files[i][server_filename], files[i][filename]
  GET    /v1/task/{task}               status + progress (polled by the progress bar)
  GET    /v1/download/{task}           .docx (one file) or .zip (several)
  GET    /v1/thumbnail/{task}/{file}   first-page preview
  DELETE /v1/upload/{task}/{file}      remove one uploaded file
  DELETE /v1/task/{task}               delete everything now (otherwise after 2 h)

Run:  .venv/bin/uvicorn server.app:app --port 8000
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
import secrets
import shutil
import threading
import time
import zipfile
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import engine

BASE = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("DOCXIFY_DATA_DIR", BASE / "data"))
TTL_SECONDS = int(os.environ.get("DOCXIFY_TTL_SECONDS", 2 * 3600))
MAX_FILE_MB = int(os.environ.get("DOCXIFY_MAX_FILE_MB", 100))
MAX_FILES = int(os.environ.get("DOCXIFY_MAX_FILES", 10))
MAX_PAGES = int(os.environ.get("DOCXIFY_MAX_PAGES", 300))
WORKERS = int(os.environ.get("DOCXIFY_WORKERS", 2))

log = logging.getLogger("docxify.app")
TASK_RE = re.compile(r"^[0-9a-f]{32}$")
FILE_RE = re.compile(r"^[0-9a-f]{64}\.pdf$")
UID_RE = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


# ------------------------------------------------------------------ task store
@dataclass
class FileRec:
    name: str
    path: Path
    size: int
    pages: int
    scanned_pages: list[int]
    thumb: Path


@dataclass
class Upload:
    path: Path
    name: str
    next_chunk: int = 0
    size: int = 0


@dataclass
class Task:
    id: str
    dir: Path
    created: float = field(default_factory=time.time)
    touched: float = field(default_factory=time.time)
    files: dict[str, FileRec] = field(default_factory=dict)
    uploads: dict[str, Upload] = field(default_factory=dict)
    status: str = "TaskWaiting"           # TaskWaiting | TaskProcessing | TaskSuccess | TaskError
    progress: dict = field(default_factory=dict)
    output: Optional[Path] = None
    output_name: str = ""
    result: dict = field(default_factory=dict)
    error: str = ""
    lock: threading.Lock = field(default_factory=threading.Lock)


TASKS: dict[str, Task] = {}
TASKS_LOCK = threading.Lock()
CONVERT_SLOTS = threading.BoundedSemaphore(WORKERS)


def _err(status: int, kind: str, message: str):
    raise HTTPException(status_code=status, detail={"type": kind, "message": message})


def _task(task_id: str) -> Task:
    if not TASK_RE.match(task_id or ""):
        _err(400, "InvalidTask", "Unknown task.")
    with TASKS_LOCK:
        t = TASKS.get(task_id)
    if t is None:
        _err(404, "InvalidTask", "This task has expired. Please start again.")
    t.touched = time.time()
    return t


def _delete_task(task_id: str) -> None:
    with TASKS_LOCK:
        t = TASKS.pop(task_id, None)
    if t is not None:
        shutil.rmtree(t.dir, ignore_errors=True)


def _janitor() -> None:
    while True:
        time.sleep(300)
        now = time.time()
        with TASKS_LOCK:
            expired = [k for k, t in TASKS.items() if now - t.created > TTL_SECONDS]
        for k in expired:
            _delete_task(k)
        # orphaned directories from a previous run
        for d in DATA_DIR.glob("*"):
            if d.is_dir() and d.name not in TASKS and now - d.stat().st_mtime > TTL_SECONDS:
                shutil.rmtree(d, ignore_errors=True)


def _safe_stem(filename: str) -> str:
    stem = Path(filename.replace("\\", "/")).name
    stem = re.sub(r"\.pdf$", "", stem, flags=re.I)
    stem = re.sub(r'[\x00-\x1f<>:"/\\|?*]+', "_", stem).strip(" .")
    return stem[:120] or "document"


# ------------------------------------------------------------------------ app
@asynccontextmanager
async def _lifespan(_: FastAPI):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    # pdf2docx installs a console handler on the root logger at import; keep the
    # server console readable (progress is captured through our own handler).
    import pdf2docx  # noqa: F401

    root = logging.getLogger()
    for h in list(root.handlers):
        if isinstance(h, logging.StreamHandler):
            root.removeHandler(h)
    threading.Thread(target=_janitor, daemon=True, name="docxify-janitor").start()
    yield


app = FastAPI(title="Docxify", docs_url="/api/docs", redoc_url=None, lifespan=_lifespan)


@app.exception_handler(HTTPException)
async def _http_error(_: Request, exc: HTTPException):
    detail = exc.detail if isinstance(exc.detail, dict) else {"type": "Error", "message": str(exc.detail)}
    return JSONResponse({"error": detail}, status_code=exc.status_code)


@app.middleware("http")
async def _headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    return response


@app.post("/v1/start")
def start(request: Request):
    task_id = secrets.token_hex(16)
    t = Task(id=task_id, dir=DATA_DIR / task_id)
    t.dir.mkdir(parents=True, exist_ok=True)
    with TASKS_LOCK:
        TASKS[task_id] = t
    return {
        "server": str(request.base_url).rstrip("/"),
        "task": task_id,
        "limits": {"files": MAX_FILES, "mb": MAX_FILE_MB, "pages": MAX_PAGES},
        "ocr_languages": list(engine.OCR_LANGS),
    }


@app.post("/v1/upload")
def upload(
    task: str = Form(...),
    file: UploadFile = File(...),
    name: str = Form(""),
    uid: str = Form(""),
    chunk: Optional[int] = Form(None),
    chunks: Optional[int] = Form(None),
):
    t = _task(task)
    if t.status == "TaskProcessing":
        _err(409, "TaskBusy", "This task is already converting.")
    name = name or file.filename or "document.pdf"
    key = uid if UID_RE.match(uid or "") else hashlib.sha1(name.encode()).hexdigest()[:16]
    chunk = chunk or 0
    chunks = max(1, chunks or 1)

    with t.lock:
        if chunk == 0:
            if len(t.files) + len(t.uploads) >= MAX_FILES:
                _err(400, "TooManyFiles", f"You can convert up to {MAX_FILES} files at once.")
            t.uploads[key] = Upload(path=t.dir / f"upload-{key}.part", name=name)
        up = t.uploads.get(key)
        if up is None or chunk != up.next_chunk:
            _err(409, "UploadError", "Upload chunks arrived out of order. Please retry.")

        data = file.file.read()
        if chunk == 0 and not data.lstrip()[:5].startswith(b"%PDF-"):
            t.uploads.pop(key, None)
            _err(415, "InvalidExtension", "Only PDF files can be converted.")
        up.size += len(data)
        if up.size > MAX_FILE_MB * 1024 * 1024:
            t.uploads.pop(key, None)
            up.path.unlink(missing_ok=True)
            _err(413, "FileSizeExceeded", f"Files can be up to {MAX_FILE_MB} MB.")
        with open(up.path, "ab") as fh:
            fh.write(data)
        up.next_chunk += 1
        if up.next_chunk < chunks:
            return {"chunks": str(chunks), "chunk": None}
        t.uploads.pop(key, None)

    # last chunk: 64-hex name (sha256 + .pdf). The upload key is mixed
    # in so the same document added twice gets two independent entries.
    digest = hashlib.sha256()
    with open(up.path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            digest.update(block)
    digest.update(f"|{key}|{secrets.token_hex(4)}".encode())
    server_filename = digest.hexdigest() + ".pdf"
    final = t.dir / server_filename
    up.path.replace(final)

    try:
        info = engine.analyze_pdf(final)
    except engine.ConversionError as exc:
        final.unlink(missing_ok=True)
        _err(422, "DamagedFile", str(exc))
    if info.encrypted:
        final.unlink(missing_ok=True)
        _err(422, "PasswordRequired", "This PDF is password-protected. Unlock it first, then convert.")
    if info.pages > MAX_PAGES:
        final.unlink(missing_ok=True)
        _err(422, "PageLimit", f"PDFs can have up to {MAX_PAGES} pages.")

    thumb = t.dir / f"thumb-{digest.hexdigest()}.png"
    try:
        thumb.write_bytes(engine.render_thumbnail(final))
    except Exception:
        log.warning("thumbnail failed for %s", server_filename)
    with t.lock:
        t.files[server_filename] = FileRec(name=name, path=final, size=up.size, pages=info.pages,
                                           scanned_pages=info.scanned_pages, thumb=thumb)
        if t.status in ("TaskSuccess", "TaskError"):
            t.status = "TaskWaiting"
    return {
        "server_filename": server_filename,
        "scanned": "true" if info.scanned else "false",
        "pages": info.pages,
        "scanned_pages": info.scanned_pages,
        "thumbnail": f"/v1/thumbnail/{t.id}/{server_filename}",
    }


@app.get("/v1/thumbnail/{task}/{server_filename}")
def thumbnail(task: str, server_filename: str):
    t = _task(task)
    rec = t.files.get(server_filename) if FILE_RE.match(server_filename) else None
    if rec is None or not rec.thumb.exists():
        _err(404, "NotFound", "No preview.")
    return FileResponse(rec.thumb, media_type="image/png", headers={"Cache-Control": "private, max-age=3600"})


@app.delete("/v1/upload/{task}/{server_filename}")
def remove_file(task: str, server_filename: str):
    t = _task(task)
    if not FILE_RE.match(server_filename):
        _err(400, "InvalidParams", "Unknown file.")
    with t.lock:
        rec = t.files.pop(server_filename, None)
    if rec:
        rec.path.unlink(missing_ok=True)
        rec.thumb.unlink(missing_ok=True)
    return {"deleted": bool(rec)}


@app.post("/v1/process")
async def process(request: Request):
    form = await request.form()
    t = _task(str(form.get("task", "")))
    if str(form.get("tool", "pdf_to_word")) not in ("pdf_to_word", "pdfoffice"):
        _err(400, "InvalidTask", "Only the pdf_to_word tool is available.")
    if str(form.get("convert_to", "docx")) != "docx":
        _err(400, "InvalidParams", "Only conversion to DOCX is supported.")
    ocr = str(form.get("ocr", "0")) in ("1", "true")
    lang = str(form.get("ocr_lang", "en"))
    review = str(form.get("ocr_review", "0")) in ("1", "true")
    out_template = str(form.get("output_filename", "{filename}")) or "{filename}"
    packaged = _safe_stem(str(form.get("packaged_filename", "docxify_converted")))

    selected: list[tuple[str, str]] = []
    i = 0
    while f"files[{i}][server_filename]" in form:
        sf = str(form[f"files[{i}][server_filename]"])
        fname = str(form.get(f"files[{i}][filename]", "document.pdf"))
        if not FILE_RE.match(sf) or sf not in t.files:
            _err(400, "InvalidParams", "A selected file is no longer available. Please upload it again.")
        selected.append((sf, fname))
        i += 1
    if not selected:
        _err(400, "InvalidParams", "No file selected.")

    from starlette.concurrency import run_in_threadpool

    return await run_in_threadpool(_run_task, t, selected, ocr, lang, out_template, packaged, review)


def _run_task(t: Task, selected, ocr: bool, lang: str, out_template: str, packaged: str,
              review: bool = False):
    with t.lock:
        if t.status == "TaskProcessing":
            _err(409, "TaskBusy", "This task is already converting.")
        t.status, t.error, t.result = "TaskProcessing", "", {}
        t.progress = {"stage": "queued", "file": 0, "files": len(selected), "done": 0, "total": 1, "percent": 0}

    t0 = time.time()
    outputs: list[tuple[Path, str]] = []
    per_file: list[dict] = []
    weights = {"analyze": (0, 3), "ocr": (3, 55), "layout": (55, 85), "write": (85, 97), "enrich": (97, 100)}

    def progress_for(index: int, has_ocr: bool):
        def cb(stage: str, done: int, total: int):
            lo, hi = weights.get(stage, (0, 100))
            if not has_ocr and stage in ("layout", "write"):   # spread over the whole bar
                lo, hi = {"layout": (3, 70), "write": (70, 97)}[stage]
            within = lo + (hi - lo) * (done / max(1, total))
            overall = (index + within / 100) / len(selected) * 100
            t.progress = {"stage": stage, "file": index + 1, "files": len(selected),
                          "done": done, "total": total, "percent": round(overall, 1)}
        return cb

    try:
        with CONVERT_SLOTS:
            for idx, (sf, fname) in enumerate(selected):
                rec = t.files[sf]
                stem = _safe_stem(out_template.replace("{filename}", _safe_stem(fname)))
                dst = t.dir / f"out-{idx}.docx"
                stats = engine.convert(rec.path, dst, ocr=ocr, lang=lang, review=review,
                                       progress=progress_for(idx, ocr and bool(rec.scanned_pages)))
                outputs.append((dst, stem + ".docx"))
                per_file.append({"filename": fname, **stats})
    except engine.ConversionError as exc:
        t.status, t.error = "TaskError", str(exc)
        _err(422, "ProcessingError", str(exc))
    except HTTPException:
        t.status = "TaskError"
        raise
    except Exception:  # unexpected engine failure
        log.exception("conversion failed")
        t.status, t.error = "TaskError", "Something went wrong while converting this PDF."
        _err(500, "ProcessingError", t.error)

    # one file -> the .docx itself; several -> one zip (unique names inside)
    if len(outputs) == 1:
        t.output, t.output_name = outputs[0]
    else:
        zpath = t.dir / "output.zip"
        seen: dict[str, int] = {}
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for path, nm in outputs:
                n = seen.get(nm, 0)
                seen[nm] = n + 1
                arc = nm if n == 0 else f"{Path(nm).stem} ({n}).docx"
                z.write(path, arc)
        t.output, t.output_name = zpath, packaged + ".zip"

    t.result = {
        "download_filename": t.output_name,
        "filesize": sum(t.files[sf].size for sf, _ in selected),
        "output_filesize": t.output.stat().st_size,
        "output_filenumber": len(outputs),
        "output_extensions": ["docx"],
        "timer": f"{time.time() - t0:.3f}",
        "status": "TaskSuccess",
        "files": per_file,
    }
    t.status = "TaskSuccess"
    t.progress = {"stage": "done", "file": len(selected), "files": len(selected),
                  "done": 1, "total": 1, "percent": 100}
    return t.result


@app.get("/v1/task/{task}")
def task_status(task: str):
    t = _task(task)
    return {"task": t.id, "status": t.status, "progress": t.progress, "error": t.error or None,
            "files": len(t.files), "expires_in": max(0, int(TTL_SECONDS - (time.time() - t.created)))}


@app.get("/v1/download/{task}")
def download(task: str):
    t = _task(task)
    if t.status != "TaskSuccess" or t.output is None or not t.output.exists():
        _err(404, "NotReady", "There is nothing to download yet.")
    media = ("application/zip" if t.output.suffix == ".zip" else
             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return FileResponse(t.output, media_type=media, filename=t.output_name)


@app.delete("/v1/task/{task}")
def delete_task(task: str):
    _task(task)
    _delete_task(task)
    return {"deleted": True}


app.mount("/", StaticFiles(directory=BASE / "web", html=True), name="web")

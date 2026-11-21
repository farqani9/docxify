/* Docxify: client.
 * Flow: start task -> chunked upload (server reports `scanned`) -> options ->
 * scanned-PDF OCR prompt -> process -> download.
 */
(() => {
  'use strict';

  const CHUNK = 2 * 1024 * 1024; // upload in 2 MB chunks
  const $ = (id) => document.getElementById(id);

  const state = {
    task: null,
    limits: { files: 10, mb: 100, pages: 300 },
    files: [],            // {uid, file, name, size, status, progress, server_filename, pages, scanned, thumbnail, el}
    ocr: false,
    lang: 'en',
    queue: Promise.resolve(),
  };

  // ------------------------------------------------------------------ helpers
  const fmtSize = (b) => (b < 1024 ? `${b} B` : b < 1048576 ? `${(b / 1024).toFixed(0)} KB` : `${(b / 1048576).toFixed(1)} MB`);
  const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 6000);
  }
  function show(view) {
    document.querySelectorAll('.view').forEach((v) => v.classList.toggle('is-active', v.id === `view-${view}`));
    const heading = document.querySelector(`#view-${view} h1, #view-${view} h2`);
    if (heading && view !== 'select') { heading.setAttribute('tabindex', '-1'); heading.focus({ preventScroll: true }); }
  }
  async function api(path, opts = {}) {
    const res = await fetch(path, opts);
    let json = {};
    try { json = await res.json(); } catch (_) { /* non-JSON */ }
    if (!res.ok) throw new Error((json.error && json.error.message) || `Request failed (${res.status})`);
    return json;
  }
  async function ensureTask() {
    if (state.task) return state.task;
    const r = await api('/v1/start', { method: 'POST' });
    state.task = r.task;
    state.limits = r.limits || state.limits;
    return state.task;
  }

  // ------------------------------------------------------------------ files
  function addFiles(list) {
    const incoming = [...list].filter((f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
    if (!incoming.length) { toast('Only PDF files can be converted.'); return; }
    const room = state.limits.files - state.files.length;
    if (room <= 0) { toast(`You can convert up to ${state.limits.files} files at once.`); return; }
    if (incoming.length > room) toast(`Only the first ${room} file(s) were added (limit ${state.limits.files}).`);
    show('files');
    incoming.slice(0, room).forEach((file) => {
      if (file.size > state.limits.mb * 1048576) { toast(`${file.name} is larger than ${state.limits.mb} MB.`); return; }
      const item = { uid: uid(), file, name: file.name, size: file.size, status: 'queued', progress: 0 };
      state.files.push(item);
      renderCard(item);
      state.queue = state.queue.then(() => uploadOne(item)).catch(() => {});
    });
    renderAddCard();
    syncPanel();
  }

  function renderCard(item) {
    const li = document.createElement('li');
    li.className = 'file-card';
    li.innerHTML = `
      <div class="thumb">
        <svg class="placeholder" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3h7l5 5v13H7zM14 3v5h5"/></svg>
        <span class="pages" hidden></span>
        <div class="upload-bar"><span></span></div>
      </div>
      <div class="file-meta">
        <p class="file-name"></p>
        <div class="file-sub"><span class="file-size"></span><span class="badge" hidden></span></div>
      </div>
      <button class="remove" type="button" aria-label="Remove file">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>
      </button>`;
    li.querySelector('.file-name').textContent = item.name;
    li.querySelector('.file-name').title = item.name;
    li.querySelector('.file-size').textContent = fmtSize(item.size);
    li.querySelector('.remove').addEventListener('click', () => removeFile(item));
    item.el = li;
    const grid = $('file-grid');
    const add = grid.querySelector('.add-slot');
    grid.insertBefore(li, add || null);
  }

  function updateCard(item) {
    const el = item.el;
    if (!el) return;
    el.querySelector('.upload-bar').hidden = item.status !== 'uploading' && item.status !== 'queued';
    el.querySelector('.upload-bar span').style.width = `${Math.round(item.progress * 100)}%`;
    const badge = el.querySelector('.badge');
    if (item.status === 'done') {
      badge.hidden = false;
      badge.className = `badge ${item.scanned ? 'badge-scan' : 'badge-text'}`;
      badge.textContent = item.scanned ? 'Scanned' : 'Text';
      badge.title = item.scanned ? 'No selectable text found: OCR recommended' : 'Has selectable text';
      const pages = el.querySelector('.pages');
      pages.hidden = false;
      pages.textContent = `${item.pages} page${item.pages === 1 ? '' : 's'}`;
      if (item.thumbnail) {
        const img = new Image();
        img.alt = `First page of ${item.name}`;
        img.src = item.thumbnail;
        img.onload = () => { const ph = el.querySelector('.placeholder'); if (ph) ph.replaceWith(img); };
      }
    } else if (item.status === 'error') {
      badge.hidden = false;
      badge.className = 'badge badge-err';
      badge.textContent = 'Failed';
    }
  }

  function renderAddCard() {
    const grid = $('file-grid');
    let slot = grid.querySelector('.add-slot');
    if (state.files.length >= state.limits.files) { if (slot) slot.remove(); return; }
    if (!slot) {
      slot = document.createElement('li');
      slot.className = 'add-slot';
      slot.innerHTML = `<button class="add-card" type="button">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>Add more PDFs</button>`;
      slot.querySelector('button').addEventListener('click', () => $('file-input').click());
      grid.appendChild(slot);
    }
  }

  async function removeFile(item) {
    state.files = state.files.filter((f) => f !== item);
    item.removed = true;
    if (item.xhr) item.xhr.abort();
    item.el.remove();
    if (item.server_filename && state.task) {
      fetch(`/v1/upload/${state.task}/${item.server_filename}`, { method: 'DELETE' }).catch(() => {});
    }
    if (!state.files.length) { resetAll(); return; }
    renderAddCard();
    syncPanel();
  }

  function sendChunk(item, form) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      item.xhr = xhr;
      xhr.open('POST', '/v1/upload');
      xhr.responseType = 'json';
      xhr.upload.onprogress = (e) => { if (e.lengthComputable) item._chunkLoaded = e.loaded; item.onChunkProgress && item.onChunkProgress(); };
      xhr.onload = () => {
        const body = xhr.response || {};
        if (xhr.status >= 200 && xhr.status < 300) resolve(body);
        else reject(new Error((body.error && body.error.message) || `Upload failed (${xhr.status})`));
      };
      xhr.onerror = () => reject(new Error('Network error while uploading.'));
      xhr.onabort = () => reject(new Error('aborted'));
      xhr.send(form);
    });
  }

  async function uploadOne(item) {
    if (item.removed) return;
    try {
      await ensureTask();
      item.status = 'uploading';
      updateCard(item);
      syncPanel();
      const chunks = Math.max(1, Math.ceil(item.size / CHUNK));
      let sent = 0;
      let result;
      for (let i = 0; i < chunks; i++) {
        const blob = item.file.slice(i * CHUNK, Math.min(item.size, (i + 1) * CHUNK));
        const form = new FormData();
        form.append('task', state.task);
        form.append('uid', item.uid);
        form.append('name', item.name);
        form.append('chunk', String(i));
        form.append('chunks', String(chunks));
        form.append('file', blob, item.name);
        item._chunkLoaded = 0;
        item.onChunkProgress = () => { item.progress = Math.min(0.99, (sent + item._chunkLoaded) / item.size); updateCard(item); };
        result = await sendChunk(item, form);
        sent += blob.size;
      }
      Object.assign(item, {
        status: 'done',
        progress: 1,
        server_filename: result.server_filename,
        pages: result.pages,
        scanned: result.scanned === 'true',
        thumbnail: result.thumbnail,
      });
    } catch (err) {
      if (item.removed) return;
      item.status = 'error';
      toast(`${item.name}: ${err.message}`);
    } finally {
      item.xhr = null;
      updateCard(item);
      syncPanel();
    }
  }

  // ------------------------------------------------------------------ options panel
  function syncPanel() {
    const ready = state.files.filter((f) => f.status === 'done');
    const busy = state.files.some((f) => f.status === 'uploading' || f.status === 'queued');
    const scanned = ready.filter((f) => f.scanned);
    $('convert').disabled = busy || ready.length === 0;
    $('panel-foot').textContent = busy ? 'Uploading…' : ready.length === 0 ? 'Add a PDF to get started.'
      : `${ready.length} file${ready.length === 1 ? '' : 's'} · ${ready.reduce((n, f) => n + (f.pages || 0), 0)} pages`;
    $('scan-hint').hidden = scanned.length === 0 || state.ocr;
    $('scan-hint-text').textContent = scanned.length === 1 && ready.length === 1
      ? 'This PDF looks scanned. OCR is recommended.'
      : `${scanned.length} of ${ready.length} PDFs look scanned. OCR is recommended.`;
    $('lang-field').hidden = !state.ocr;
    document.querySelector(`input[name=ocr][value="${state.ocr ? 1 : 0}"]`).checked = true;
  }

  document.querySelectorAll('input[name=ocr]').forEach((r) =>
    r.addEventListener('change', () => { state.ocr = r.value === '1' && r.checked; syncPanel(); }));
  $('lang').addEventListener('change', (e) => { state.lang = e.target.value; });

  // ------------------------------------------------------------------ OCR prompt
  function askOcr() {
    return new Promise((resolve) => {
      const modal = $('ocr-modal');
      const prev = document.activeElement;
      modal.hidden = false;
      $('ocr-yes').focus();
      const done = (v) => {
        modal.hidden = true;
        modal.removeEventListener('keydown', onKey);
        $('ocr-yes').onclick = $('ocr-no').onclick = modal.onclick = null;
        if (prev) prev.focus();
        resolve(v);
      };
      const onKey = (e) => {
        if (e.key === 'Escape') done(null);
        if (e.key === 'Tab') { // keep focus inside the dialog
          const f = [$('ocr-no'), $('ocr-yes')];
          const i = f.indexOf(document.activeElement);
          e.preventDefault();
          f[(i + (e.shiftKey ? -1 : 1) + f.length) % f.length].focus();
        }
      };
      modal.addEventListener('keydown', onKey);
      $('ocr-yes').onclick = () => done(true);
      $('ocr-no').onclick = () => done(false);
      modal.onclick = (e) => { if (e.target === modal) done(null); };
    });
  }

  // ------------------------------------------------------------------ convert
  const STAGES = {
    queued: () => 'Waiting for a free converter…',
    analyze: () => 'Analyzing PDF…',
    ocr: (p) => `Recognizing text: page ${Math.min(p.done + 1, p.total)} of ${p.total}`,
    layout: (p) => `Rebuilding layout: page ${p.done} of ${p.total}`,
    write: (p) => `Writing Word document: page ${p.done} of ${p.total}`,
    enrich: () => 'Adding headings and lists…',
    done: () => 'Done',
  };

  function setProgress(pct, text) {
    $('bar-fill').style.width = `${pct}%`;
    $('bar').setAttribute('aria-valuenow', String(Math.round(pct)));
    if (text) $('status-line').textContent = text;
  }

  async function convert() {
    const ready = state.files.filter((f) => f.status === 'done');
    if (!ready.length) return;
    if (!state.ocr && ready.some((f) => f.scanned)) {
      const choice = await askOcr();
      if (choice === null) return;
      state.ocr = choice;
      syncPanel();
    }
    show('progress');
    setProgress(2, 'Starting…');
    const form = new FormData();
    form.append('task', state.task);
    form.append('tool', 'pdf_to_word');
    form.append('convert_to', 'docx');
    form.append('ocr', state.ocr ? '1' : '0');
    form.append('ocr_lang', state.lang);
    form.append('output_filename', '{filename}');
    form.append('packaged_filename', 'docxify_converted');
    ready.forEach((f, i) => {
      form.append(`files[${i}][server_filename]`, f.server_filename);
      form.append(`files[${i}][filename]`, f.name);
    });

    let polling = true;
    const poll = async () => {
      while (polling) {
        try {
          const s = await api(`/v1/task/${state.task}`);
          const p = s.progress || {};
          if (p.stage) {
            const label = (STAGES[p.stage] || (() => 'Working…'))(p);
            const prefix = p.files > 1 ? `File ${p.file} of ${p.files}: ` : '';
            setProgress(Math.max(2, p.percent || 0), prefix + label);
          }
        } catch (_) { /* transient */ }
        await new Promise((r) => setTimeout(r, 350));
      }
    };
    poll();
    try {
      const result = await api('/v1/process', { method: 'POST', body: form });
      polling = false;
      setProgress(100, 'Done');
      showDone(result, ready);
    } catch (err) {
      polling = false;
      show('files');
      toast(err.message);
    }
  }

  function showDone(result, ready) {
    const many = result.output_filenumber > 1;
    $('h-done').textContent = many ? 'Your Word documents are ready' : 'Your Word document is ready';
    $('download-label').textContent = many ? 'Download ZIP' : 'Download Word';
    const link = $('download');
    link.href = `/v1/download/${state.task}`;
    link.setAttribute('download', result.download_filename);
    const files = result.files || [];
    const ocrPages = files.reduce((n, f) => n + (f.ocr_pages || 0), 0);
    const rows = [
      ['File', result.download_filename],
      ['Size', fmtSize(result.output_filesize)],
      ['Pages', String(files.reduce((n, f) => n + (f.pages || 0), 0))],
      ['OCR', ocrPages ? `${ocrPages} page${ocrPages === 1 ? '' : 's'} recognized` : 'Not used'],
      ['Time', `${Number(result.timer).toFixed(1)} s`],
    ];
    const dl = $('summary');
    dl.replaceChildren(...rows.flatMap(([k, v]) => {
      const dt = document.createElement('dt'); dt.textContent = k;
      const dd = document.createElement('dd'); dd.textContent = v;
      return [dt, dd];
    }));
    show('done');
    void ready;
  }

  // ------------------------------------------------------------------ reset / delete
  function resetAll() {
    state.files.forEach((f) => { f.removed = true; if (f.xhr) f.xhr.abort(); });
    state.files = [];
    state.task = null;
    state.ocr = false;
    state.queue = Promise.resolve();
    $('file-grid').replaceChildren();
    $('file-input').value = '';
    syncPanel();
    show('select');
  }

  $('again').addEventListener('click', resetAll);
  $('restart').addEventListener('click', () => {
    if (state.task) fetch(`/v1/task/${state.task}`, { method: 'DELETE' }).catch(() => {});
    resetAll();
  });
  $('delete-now').addEventListener('click', async () => {
    try {
      if (state.task) await api(`/v1/task/${state.task}`, { method: 'DELETE' });
      toast('Your files were deleted.');
    } catch (err) { toast(err.message); }
    resetAll();
  });
  $('convert').addEventListener('click', convert);

  // ------------------------------------------------------------------ pick / drop
  $('pick').addEventListener('click', () => $('file-input').click());
  $('file-input').addEventListener('change', (e) => { addFiles(e.target.files); e.target.value = ''; });

  let dragDepth = 0;
  const canDrop = () => $('view-select').classList.contains('is-active') || $('view-files').classList.contains('is-active');
  window.addEventListener('dragenter', (e) => {
    if (!e.dataTransfer || ![...e.dataTransfer.types].includes('Files') || !canDrop()) return;
    dragDepth++;
    $('drop-veil').hidden = false;
  });
  window.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) $('drop-veil').hidden = true; });
  window.addEventListener('dragover', (e) => { if (canDrop()) e.preventDefault(); });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    $('drop-veil').hidden = true;
    if (canDrop() && e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });
})();

# Security policy

## Reporting a vulnerability

Please report security issues **privately**, through GitHub's
[private vulnerability reporting](../../security/advisories/new) for this repository,
not in a public issue. Include steps to reproduce and, if relevant, a minimal PDF.
You can expect an acknowledgement within a week.

## Deployment notes

Docxify is designed to run on your own machine and listens on `127.0.0.1` by default.
It has no authentication. If you expose it to a network:

- put it behind a reverse proxy that adds authentication and request-size limits;
- keep `DOCXIFY_MAX_FILE_MB`, `DOCXIFY_MAX_PAGES` and `DOCXIFY_WORKERS` conservative;
  PDF parsing and OCR are CPU-intensive;
- remember that uploaded files are kept on disk (in `DOCXIFY_DATA_DIR`) until they are
  deleted or expire (2 hours by default).

Running Docxify as a network service also brings AGPL-3.0 obligations: users of the
service must be able to get the source code of the version you run.

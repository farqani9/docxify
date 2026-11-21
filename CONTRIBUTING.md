# Contributing to Docxify

Thanks for helping! Bug reports with a sample PDF are the most useful contribution.

## Set up

```bash
uv venv --python 3.12 .venv
uv pip install --python .venv/bin/python -r requirements.txt
.venv/bin/uvicorn server.app:app --reload --port 8000
```

## Before you open a pull request

```bash
.venv/bin/python tests/make_fixtures.py   # synthetic test PDFs
.venv/bin/python tests/check_engine.py    # must print ALL CHECKS PASSED
npm install && npx playwright install chromium
npm run test:e2e                          # with the server running
```

- Keep pull requests focused: one problem per PR, with a sentence on how you verified it.
- Conversion changes should come with a check in `tests/check_engine.py`, or a fixture in
  `tests/make_fixtures.py` that reproduces the problem.
- New dependencies need a good reason. The project runs fully offline after install and
  should stay that way.

## Reporting a conversion problem

Open an issue with:

- whether the PDF is digital or scanned, and whether you used OCR;
- what is wrong in the Word file (page, element, a screenshot if possible);
- a sample PDF **only if it contains nothing confidential**. A synthetic or redacted file that
  shows the same problem is ideal.

For local work, the `feedback/` folder holds private regression cases (it is git-ignored);
`tests/check_feedback.py` re-checks every case in it.

## License

By contributing you agree that your contributions are licensed under the
[GNU AGPL-3.0](LICENSE), the license of this project.

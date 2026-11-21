# Local feedback cases

Private regression cases for conversions that went wrong. Everything in this folder
except this README and `_template/` is git-ignored, so real documents stay on your machine.

1. Create a folder per case, e.g. `feedback/invoice-table/` (copy `_template/`).
2. Put inside it the **original PDF**, the **.docx** Docxify produced, and `notes.md`
   describing what went wrong.
3. Run `.venv/bin/python tests/check_feedback.py` to re-check every case after a change.

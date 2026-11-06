# FactorMaker — Windows

Desktop Qt (PySide6) app for Windows.

## Run

```bash
cd windows
python -m venv .venv
.venv\Scripts\activate
pip install -r req.txt
python main.py
```

Requires the bundled `templates/`, `images/`, and `fonts/` next to `main.py`.
Invoice PDF export uses Microsoft Word via `docx2pdf`.

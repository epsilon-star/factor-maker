# FactorMaker

Persian invoice maker by Shadow-Team & Meta-Noia.

This repo ships **two clients** side by side:

| Folder | Platform | Stack |
|--------|----------|--------|
| [`windows/`](windows/) | Windows desktop | PySide6 + Word/PDF |
| [`webapp/`](webapp/) | Browser (mobile-friendly) | FastAPI + responsive RTL UI |

## Quick start

**Windows desktop**

```bash
cd windows
pip install -r req.txt
python main.py
```

**Web app**

```bash
cd webapp
pip install -r requirements.txt
python run.py
```

Then open http://127.0.0.1:8000 — works on phone and desktop.

## Shared ideas

Both apps support seller profiles, free line items, theme colors, SQLite factor storage with ID occupancy checks, and DOCX generation from the same visual templates.

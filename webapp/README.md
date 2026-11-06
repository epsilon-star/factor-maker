# FactorMaker Web

Responsive, mobile-first Persian invoice maker (RTL).

## Run

```bash
cd webapp
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
python run.py
```

Open http://127.0.0.1:8000

On Windows + Python 3.14, use `python run.py` (not bare `uvicorn --reload`) — reload workers hit a `socket.socketpair` bug. Optional: `python run.py --reload`.

## Notes

- Generates `.docx` invoices (same Word templates as the Windows app).
- PDF conversion is optional and only available when `docx2pdf` + Microsoft Word are present.
- Runtime data lives in `webapp/data/` and `webapp/profiles/` (gitignored).
- Outputs go to `webapp/output/`.

"""FactorMaker web API + SPA shell."""

from __future__ import annotations

import asyncio
import os
from pathlib import Path
from typing import Any

import jdatetime
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field

from . import factor_store, profiles as profile_store
from .engine import (
	DEFAULT_THEMES,
	IMAGES_DIR,
	TEMPLATE_PATH,
	THEME_LABELS,
	PersianFactor,
	image_map,
)

APP_DIR = Path(__file__).resolve().parent
WEBAPP_ROOT = APP_DIR.parent
TEMPLATES_DIR = WEBAPP_ROOT / "templates"
STATIC_DIR = WEBAPP_ROOT / "static"

app = FastAPI(title="FactorMaker Web", version="1.0.0")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

engine = PersianFactor()

THEME_ACCENTS = {
	"blue": ("#1d4ed8", "#dbeafe"),
	"cyan": ("#0e7490", "#cffafe"),
	"red": ("#b91c1c", "#fee2e2"),
	"yellow": ("#a16207", "#fef3c7"),
	"green": ("#15803d", "#dcfce7"),
	"gray": ("#475569", "#e2e8f0"),
}

OUTPUT_FORMATS = ("word", "web")


class ItemRow(BaseModel):
	name: str = ""
	count: int = 1
	price: int = 0
	free: bool = False


class FactorPayload(BaseModel):
	factor_id: int
	factor_date: str = ""
	title: str = ""
	addons: str = ""
	sby: str = ""
	phone: str = ""
	location: str = ""
	off: float = 0
	tax: float = 0
	desc1: str = ""
	desc2: str = ""
	items: list[ItemRow] = Field(default_factory=list)
	themes: list[str] = Field(default_factory=lambda: list(DEFAULT_THEMES))
	formats: list[str] = Field(default_factory=lambda: ["web"])
	status: str | None = None
	overwrite: bool = False


class ProfilePayload(BaseModel):
	id: str | None = None
	name: str = ""
	sby: str = ""
	phone: str = ""
	location: str = ""
	off: float = 0
	tax: float = 0
	desc1: str = ""
	desc2: str = ""
	preferred_themes: list[str] = Field(default_factory=lambda: list(DEFAULT_THEMES))
	is_default: bool = False


class TotalPayload(BaseModel):
	items: list[ItemRow] = Field(default_factory=list)
	off: float = 0
	tax: float = 0


def _bootstrap() -> None:
	os.makedirs(WEBAPP_ROOT / "data", exist_ok=True)
	os.makedirs(WEBAPP_ROOT / "profiles", exist_ok=True)
	os.makedirs(WEBAPP_ROOT / "output", exist_ok=True)
	factor_store.init_db()
	profile_store.load_store()


_bootstrap()


def _items_engine(items: list[ItemRow] | list[dict[str, Any]]) -> dict:
	rows = []
	for item in items:
		if isinstance(item, ItemRow):
			rows.append(item.model_dump())
		else:
			rows.append(item)
	return factor_store.items_to_engine_dict(rows)


def _snapshot_dict(payload: FactorPayload) -> dict[str, Any]:
	formats = _normalize_formats(payload.formats)
	if not formats:
		formats = ["web"]
	return {
		"factor_id": payload.factor_id,
		"factor_date": payload.factor_date,
		"title": payload.title,
		"addons": payload.addons,
		"sby": payload.sby,
		"phone": payload.phone,
		"location": payload.location,
		"off": payload.off,
		"tax": payload.tax,
		"desc1": payload.desc1,
		"desc2": payload.desc2,
		"items": [i.model_dump() for i in payload.items],
		"themes": [t for t in (payload.themes or []) if t in image_map] or list(DEFAULT_THEMES),
		"formats": formats,
	}


def _normalize_formats(raw: list[str] | None) -> list[str]:
	"""Return only valid formats. Empty input stays empty (do not invent defaults)."""
	return [f for f in (raw or []) if f in OUTPUT_FORMATS]


def _invoice_view_context(factor: dict[str, Any]) -> dict[str, Any]:
	items = factor.get("items") or []
	engine_items = factor_store.items_to_engine_dict(items)
	subtotal = 0
	rows = []
	for idx, (name, entry) in enumerate(engine_items.items(), start=1):
		count = int(entry[0])
		price = int(entry[1])
		is_free = bool(entry[2]) if len(entry) > 2 else False
		line = 0 if is_free else count * price
		if not is_free:
			subtotal += line
		rows.append(
			{
				"index_fa": engine.persianNumber(str(idx)),
				"name": name,
				"count_fa": engine.persianNumber(str(count)),
				"price_fa": engine.formatNumber(str(price)),
				"line_fa": engine.formatNumber(str(line)),
				"free": is_free,
			}
		)

	tax = float(factor.get("tax") or 0)
	off = float(factor.get("off") or 0)
	after_tax = subtotal * (100 + tax) / 100
	tax_amount = after_tax - subtotal
	total = round(after_tax * (100 - off) / 100)
	off_amount = round(after_tax - total)

	themes = factor.get("themes") or list(DEFAULT_THEMES)
	accent_key = next((t for t in themes if t in THEME_ACCENTS), "blue")
	accent, accent_soft = THEME_ACCENTS[accent_key]

	return {
		"factor": factor,
		"rows": rows,
		"factor_id_fa": engine.persianNumber(str(factor.get("factor_id") or "")),
		"factor_date_fa": engine.persianNumber(str(factor.get("factor_date") or "")),
		"subtotal_fa": engine.formatNumber(str(subtotal)),
		"tax_fa": engine.persianNumber(f"{tax:g}"),
		"off_fa": engine.persianNumber(f"{off:g}"),
		"tax_amount_fa": engine.formatNumber(str(round(tax_amount))),
		"off_amount_fa": engine.formatNumber(str(off_amount)),
		"total_fa": engine.formatNumber(str(total)),
		"total_words": engine.convertWord(total) + " ریال",
		"accent": accent,
		"accent_soft": accent_soft,
	}



def _default_seed() -> int:
	tt = jdatetime.datetime.now()
	return int(str(tt.year)[1:] + "00" + "000")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
	return templates.TemplateResponse(
		request,
		"index.html",
		{
			"theme_labels": THEME_LABELS,
			"default_themes": DEFAULT_THEMES,
			"image_map": image_map,
		},
	)


@app.get("/api/meta")
async def meta():
	return {
		"themes": [
			{
				"id": key,
				"label": THEME_LABELS.get(key, key),
				"image": f"/static/images/{image_map[key]}.jpg",
				"default": key in DEFAULT_THEMES,
			}
			for key in image_map
		],
		"formats": [
			{
				"id": "web",
				"label": "صفحه وب قابل چاپ",
				"hint": "چاپ مستقیم به PDF از مرورگر",
			},
			{
				"id": "word",
				"label": "قالب تصویری (Word)",
				"hint": "خروجی DOCX با تصویر تم",
			},
		],
		"max_items": 13,
	}


@app.get("/invoice/{factor_id}", response_class=HTMLResponse)
async def invoice_page(request: Request, factor_id: int):
	data = factor_store.get_factor(factor_id)
	if not data:
		raise HTTPException(status_code=404, detail="فاکتور پیدا نشد")
	ctx = _invoice_view_context(data)
	ctx["request"] = request
	return templates.TemplateResponse(request, "invoice.html", ctx)



@app.get("/api/bootstrap")
async def bootstrap():
	seed = _default_seed()
	factor_id = factor_store.next_available_id(seed)
	tt = jdatetime.datetime.now()
	active = profile_store.get_active_profile()
	return {
		"factor_id": factor_id,
		"factor_date": tt.strftime("%Y/%m/%d"),
		"profile": active,
		"profiles": profile_store.list_profiles(),
		"themes": active.get("preferred_themes") or list(DEFAULT_THEMES),
	}


@app.post("/api/total")
async def calc_total(payload: TotalPayload):
	items = _items_engine(payload.items)
	total = engine.calculateTotal(list(items.values()), payload.tax, payload.off)
	return {
		"total": total,
		"formatted": engine.formatNumber(str(total)),
		"words": engine.convertWord(total) + " ریال",
	}


@app.get("/api/factors")
async def list_factors():
	return {"factors": factor_store.list_factors()}


@app.get("/api/factors/{factor_id}")
async def get_factor(factor_id: int):
	data = factor_store.get_factor(factor_id)
	if not data:
		raise HTTPException(status_code=404, detail="فاکتور پیدا نشد")
	return data


@app.post("/api/factors")
async def save_factor(payload: FactorPayload):
	factor_id = int(payload.factor_id)
	occupied = factor_store.id_exists(factor_id)
	if occupied and not payload.overwrite:
		raise HTTPException(
			status_code=409,
			detail={
				"message": f"آیدی {factor_id} قبلا ذخیره شده",
				"suggested_id": factor_store.next_available_id(factor_id),
			},
		)
	existing = factor_store.get_factor(factor_id)
	status = payload.status or (existing or {}).get("status") or "draft"
	saved = factor_store.upsert_factor(_snapshot_dict(payload), status=status)
	return saved


@app.delete("/api/factors/{factor_id}")
async def delete_factor(factor_id: int):
	if not factor_store.delete_factor(factor_id):
		raise HTTPException(status_code=404, detail="فاکتور پیدا نشد")
	return {"ok": True}


@app.get("/api/profiles")
async def list_profiles():
	store = profile_store.load_store()
	return {
		"profiles": store["profiles"],
		"last_profile_id": store.get("last_profile_id"),
	}


@app.post("/api/profiles")
async def upsert_profile(payload: ProfilePayload):
	data = payload.model_dump()
	saved = profile_store.upsert_profile(data, make_default=payload.is_default)
	return saved


@app.post("/api/profiles/{profile_id}/activate")
async def activate_profile(profile_id: str):
	profile = profile_store.get_profile(profile_id)
	if not profile:
		raise HTTPException(status_code=404, detail="پروفایل پیدا نشد")
	profile_store.set_last_profile(profile_id)
	return profile


@app.delete("/api/profiles/{profile_id}")
async def delete_profile(profile_id: str):
	if len(profile_store.list_profiles()) <= 1:
		raise HTTPException(status_code=400, detail="حداقل یک پروفایل باید باقی بماند")
	if not profile_store.delete_profile(profile_id):
		raise HTTPException(status_code=404, detail="پروفایل پیدا نشد")
	return {"ok": True, "profiles": profile_store.list_profiles()}


@app.post("/api/generate")
async def generate_factor(payload: FactorPayload):
	# Trust the client list exactly — never silently re-add "word"
	formats = _normalize_formats(list(payload.formats or []))
	want_word = "word" in formats
	want_web = "web" in formats
	themes = [t for t in (payload.themes or []) if t in image_map]

	if not want_word and not want_web:
		raise HTTPException(status_code=400, detail="حداقل یک نوع خروجی را انتخاب کنید")
	if want_word and not themes:
		raise HTTPException(status_code=400, detail="برای خروجی Word حداقل یک رنگ را انتخاب کنید")
	if want_web and not themes:
		themes = list(DEFAULT_THEMES)
	if not payload.title.strip():
		raise HTTPException(status_code=400, detail="لطفا عنوان را مشخص کنید")
	if not payload.sby.strip():
		raise HTTPException(status_code=400, detail="لطفا از را مشخص کنید")
	if not payload.phone.strip():
		raise HTTPException(status_code=400, detail="لطفا تلفن را مشخص کنید")
	if not payload.location.strip():
		raise HTTPException(status_code=400, detail="لطفا آدرس را مشخص کنید")
	if not payload.factor_date.strip():
		raise HTTPException(status_code=400, detail="لطفا تاریخ را مشخص کنید")

	items = _items_engine(payload.items)
	if not items:
		raise HTTPException(status_code=400, detail="حداقل یک قلم اضافه کنید")

	if want_word:
		if not os.path.isfile(TEMPLATE_PATH):
			raise HTTPException(status_code=500, detail="قالب Word پیدا نشد")
		for theme in themes:
			image_path = os.path.join(IMAGES_DIR, f"{image_map[theme]}.jpg")
			if not os.path.isfile(image_path):
				raise HTTPException(status_code=500, detail=f"تصویر تم {theme} پیدا نشد")

	factor_id = int(payload.factor_id)
	if factor_store.id_exists(factor_id) and not payload.overwrite:
		raise HTTPException(
			status_code=409,
			detail={
				"message": f"آیدی {factor_id} قبلا استفاده شده",
				"suggested_id": factor_store.next_available_id(factor_id),
			},
		)

	snap = _snapshot_dict(payload)
	snap["themes"] = themes or list(DEFAULT_THEMES)
	snap["formats"] = formats
	factor_store.upsert_factor(snap, status="draft")

	files: list[dict[str, Any]] = []
	invoice_url = f"/invoice/{factor_id}"
	print_url = f"/invoice/{factor_id}?print=1"

	if want_word:
		datas = {
			"factor_id": engine.persianNumber(str(factor_id)),
			"factor_date": engine.persianNumber(payload.factor_date),
			"factor_addons": engine.persianNumber(payload.addons),
			"title": payload.title,
			"sby": payload.sby,
			"phone": payload.phone,
			"location": payload.location,
			"items": items,
			"off": float(payload.off),
			"tax": float(payload.tax),
			"factor_disc_1": engine.persianNumber(payload.desc1),
			"factor_disc_2": engine.persianNumber(payload.desc2),
		}
		ok = await asyncio.to_thread(
			lambda: engine.makeFactor(
				factor_id,
				TEMPLATE_PATH,
				datas,
				themes=themes,
				make_pdf=False,
			)
		)
		if not ok:
			raise HTTPException(status_code=500, detail="خطا در تولید فاکتور Word")

		out_dir = WEBAPP_ROOT / "output" / str(factor_id)
		if out_dir.is_dir():
			for path in sorted(out_dir.iterdir()):
				if path.suffix.lower() in {".docx", ".pdf"}:
					files.append(
						{
							"name": path.name,
							"url": f"/api/files/{factor_id}/{path.name}",
							"kind": path.suffix.lower().lstrip("."),
						}
					)

	if want_web:
		files.insert(
			0,
			{
				"name": f"صفحه وب فاکتور {factor_id}",
				"url": invoice_url,
				"kind": "web",
				"print_url": print_url,
			},
		)

	factor_store.upsert_factor(snap, status="generated")

	return {
		"ok": True,
		"factor_id": factor_id,
		"themes": themes,
		"formats": formats,
		"files": files,
		"invoice_url": invoice_url if want_web else None,
		"print_url": print_url if want_web else None,
		"open_print": bool(want_web),
	}



@app.get("/api/files/{factor_id}/{filename}")
async def download_file(factor_id: int, filename: str):
	safe = Path(filename).name
	path = WEBAPP_ROOT / "output" / str(factor_id) / safe
	if not path.is_file():
		raise HTTPException(status_code=404, detail="فایل پیدا نشد")
	media = (
		"application/pdf"
		if path.suffix.lower() == ".pdf"
		else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	)
	return FileResponse(path, media_type=media, filename=safe)


@app.get("/api/next-id")
async def next_id(preferred: int | None = None):
	seed = preferred if preferred is not None else _default_seed()
	return {"factor_id": factor_store.next_available_id(seed)}

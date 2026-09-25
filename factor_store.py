"""SQLite store for FactorMaker invoices (draft + generated)."""

from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime, timezone
from typing import Any

from engine import APP_ROOT, DEFAULT_THEMES

DATA_DIR = os.path.join(APP_ROOT, "data")
DB_PATH = os.path.join(DATA_DIR, "factors.db")


def _utc_now() -> str:
	return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _connect() -> sqlite3.Connection:
	os.makedirs(DATA_DIR, exist_ok=True)
	conn = sqlite3.connect(DB_PATH)
	conn.row_factory = sqlite3.Row
	return conn


def init_db() -> None:
	with _connect() as conn:
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS factors (
				factor_id INTEGER PRIMARY KEY,
				factor_date TEXT NOT NULL DEFAULT '',
				title TEXT NOT NULL DEFAULT '',
				addons TEXT NOT NULL DEFAULT '',
				sby TEXT NOT NULL DEFAULT '',
				phone TEXT NOT NULL DEFAULT '',
				location TEXT NOT NULL DEFAULT '',
				off REAL NOT NULL DEFAULT 0,
				tax REAL NOT NULL DEFAULT 0,
				desc1 TEXT NOT NULL DEFAULT '',
				desc2 TEXT NOT NULL DEFAULT '',
				items_json TEXT NOT NULL DEFAULT '[]',
				themes_json TEXT NOT NULL DEFAULT '[]',
				status TEXT NOT NULL DEFAULT 'draft',
				created_at TEXT NOT NULL,
				updated_at TEXT NOT NULL
			)
			"""
		)
		conn.commit()


def id_exists(factor_id: int) -> bool:
	init_db()
	with _connect() as conn:
		row = conn.execute(
			"SELECT 1 FROM factors WHERE factor_id = ? LIMIT 1",
			(int(factor_id),),
		).fetchone()
	return row is not None


def next_available_id(preferred: int | None = None) -> int:
	"""Return preferred if free, otherwise the next free integer at/after preferred."""
	init_db()
	start = int(preferred) if preferred is not None else 1
	candidate = max(start, 1)
	with _connect() as conn:
		while True:
			row = conn.execute(
				"SELECT 1 FROM factors WHERE factor_id = ? LIMIT 1",
				(candidate,),
			).fetchone()
			if row is None:
				return candidate
			candidate += 1


def get_factor(factor_id: int) -> dict[str, Any] | None:
	init_db()
	with _connect() as conn:
		row = conn.execute(
			"SELECT * FROM factors WHERE factor_id = ?",
			(int(factor_id),),
		).fetchone()
	if not row:
		return None
	return _row_to_dict(row)


def list_factors() -> list[dict[str, Any]]:
	init_db()
	with _connect() as conn:
		rows = conn.execute(
			"""
			SELECT factor_id, factor_date, title, status, updated_at
			FROM factors
			ORDER BY factor_id DESC
			"""
		).fetchall()
	return [
		{
			"factor_id": int(r["factor_id"]),
			"factor_date": r["factor_date"] or "",
			"title": r["title"] or "",
			"status": r["status"] or "draft",
			"updated_at": r["updated_at"] or "",
		}
		for r in rows
	]


def upsert_factor(data: dict[str, Any], *, status: str | None = None) -> dict[str, Any]:
	"""Insert or update a factor by factor_id."""
	init_db()
	factor_id = int(data["factor_id"])
	now = _utc_now()
	existing = get_factor(factor_id)
	created_at = (existing or {}).get("created_at") or now
	final_status = status or data.get("status") or (existing or {}).get("status") or "draft"

	items = data.get("items") or []
	themes = data.get("themes") or list(DEFAULT_THEMES)
	payload = (
		factor_id,
		str(data.get("factor_date") or ""),
		str(data.get("title") or ""),
		str(data.get("addons") or ""),
		str(data.get("sby") or ""),
		str(data.get("phone") or ""),
		str(data.get("location") or ""),
		float(data.get("off") or 0),
		float(data.get("tax") or 0),
		str(data.get("desc1") or ""),
		str(data.get("desc2") or ""),
		json.dumps(items, ensure_ascii=False),
		json.dumps(themes, ensure_ascii=False),
		str(final_status),
		created_at,
		now,
	)

	with _connect() as conn:
		conn.execute(
			"""
			INSERT INTO factors (
				factor_id, factor_date, title, addons, sby, phone, location,
				off, tax, desc1, desc2, items_json, themes_json, status,
				created_at, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(factor_id) DO UPDATE SET
				factor_date = excluded.factor_date,
				title = excluded.title,
				addons = excluded.addons,
				sby = excluded.sby,
				phone = excluded.phone,
				location = excluded.location,
				off = excluded.off,
				tax = excluded.tax,
				desc1 = excluded.desc1,
				desc2 = excluded.desc2,
				items_json = excluded.items_json,
				themes_json = excluded.themes_json,
				status = excluded.status,
				updated_at = excluded.updated_at
			""",
			payload,
		)
		conn.commit()

	saved = get_factor(factor_id)
	assert saved is not None
	return saved


def delete_factor(factor_id: int) -> bool:
	init_db()
	with _connect() as conn:
		cur = conn.execute("DELETE FROM factors WHERE factor_id = ?", (int(factor_id),))
		conn.commit()
		return cur.rowcount > 0


def items_from_engine_dict(items: dict) -> list[dict[str, Any]]:
	"""Convert engine items {name: [count, price, free?]} to JSON-friendly list."""
	rows = []
	for name, entry in items.items():
		count = int(entry[0])
		price = int(entry[1])
		is_free = bool(entry[2]) if len(entry) > 2 else False
		rows.append({"name": name, "count": count, "price": price, "free": is_free})
	return rows


def items_to_engine_dict(rows: list[dict[str, Any]] | None) -> dict:
	"""Convert stored item list back to engine items dict."""
	items = {}
	for row in rows or []:
		name = str(row.get("name") or "").strip()
		if not name or name in items:
			continue
		items[name] = [
			int(row.get("count") or 1),
			int(row.get("price") or 0),
			bool(row.get("free")),
		]
	return items


def _row_to_dict(row: sqlite3.Row) -> dict[str, Any]:
	try:
		items = json.loads(row["items_json"] or "[]")
	except json.JSONDecodeError:
		items = []
	try:
		themes = json.loads(row["themes_json"] or "[]")
	except json.JSONDecodeError:
		themes = list(DEFAULT_THEMES)
	if not isinstance(items, list):
		items = []
	if not isinstance(themes, list) or not themes:
		themes = list(DEFAULT_THEMES)
	return {
		"factor_id": int(row["factor_id"]),
		"factor_date": row["factor_date"] or "",
		"title": row["title"] or "",
		"addons": row["addons"] or "",
		"sby": row["sby"] or "",
		"phone": row["phone"] or "",
		"location": row["location"] or "",
		"off": float(row["off"] or 0),
		"tax": float(row["tax"] or 0),
		"desc1": row["desc1"] or "",
		"desc2": row["desc2"] or "",
		"items": items,
		"themes": [str(t) for t in themes],
		"status": row["status"] or "draft",
		"created_at": row["created_at"] or "",
		"updated_at": row["updated_at"] or "",
	}

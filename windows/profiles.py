"""JSON-backed seller profile store for FactorMaker."""

from __future__ import annotations

import json
import os
import uuid
from copy import deepcopy
from typing import Any

from engine import APP_ROOT, DEFAULT_THEMES

PROFILES_DIR = os.path.join(APP_ROOT, "profiles")
PROFILES_PATH = os.path.join(PROFILES_DIR, "profiles.json")

SEED_PROFILE = {
	"id": "metanoia-default",
	"name": "متانویا",
	"sby": "تیم توسعه دهنده متانویا",
	"phone": "09000000000",
	"location": "",
	"off": 0.0,
	"tax": 0.0,
	"desc1": "",
	"desc2": "",
	"preferred_themes": list(DEFAULT_THEMES),
	"is_default": True,
}


def _empty_store() -> dict[str, Any]:
	return {
		"last_profile_id": SEED_PROFILE["id"],
		"profiles": [deepcopy(SEED_PROFILE)],
	}


def _ensure_store() -> dict[str, Any]:
	os.makedirs(PROFILES_DIR, exist_ok=True)
	if not os.path.isfile(PROFILES_PATH):
		store = _empty_store()
		_write_store(store)
		return store
	try:
		with open(PROFILES_PATH, "r", encoding="utf-8") as fh:
			store = json.load(fh)
	except (OSError, json.JSONDecodeError):
		store = _empty_store()
		_write_store(store)
		return store

	if not isinstance(store, dict):
		store = _empty_store()
	profiles = store.get("profiles")
	if not isinstance(profiles, list) or not profiles:
		store = _empty_store()
		_write_store(store)
		return store

	# Normalize entries
	normalized = []
	for raw in profiles:
		if not isinstance(raw, dict):
			continue
		normalized.append(_normalize_profile(raw))
	if not normalized:
		store = _empty_store()
		_write_store(store)
		return store

	# Ensure exactly one default
	defaults = [p for p in normalized if p.get("is_default")]
	if not defaults:
		normalized[0]["is_default"] = True
	elif len(defaults) > 1:
		keep_id = defaults[0]["id"]
		for p in normalized:
			p["is_default"] = p["id"] == keep_id

	last_id = store.get("last_profile_id")
	ids = {p["id"] for p in normalized}
	if last_id not in ids:
		default = next(p for p in normalized if p["is_default"])
		last_id = default["id"]

	store = {"last_profile_id": last_id, "profiles": normalized}
	_write_store(store)
	return store


def _normalize_profile(raw: dict[str, Any]) -> dict[str, Any]:
	from engine import image_map

	themes = raw.get("preferred_themes") or list(DEFAULT_THEMES)
	if not isinstance(themes, list) or not themes:
		themes = list(DEFAULT_THEMES)
	# Migrate removed theme keys
	migrated = []
	for t in themes:
		if not isinstance(t, str):
			continue
		if t == "green":
			t = "cyan"
		if t in image_map and t not in migrated:
			migrated.append(t)
	themes = migrated or list(DEFAULT_THEMES)

	try:
		off = float(raw.get("off", 0) or 0)
	except (TypeError, ValueError):
		off = 0.0
	try:
		tax = float(raw.get("tax", 0) or 0)
	except (TypeError, ValueError):
		tax = 0.0

	return {
		"id": str(raw.get("id") or uuid.uuid4()),
		"name": str(raw.get("name") or "بدون نام"),
		"sby": str(raw.get("sby") or ""),
		"phone": str(raw.get("phone") or ""),
		"location": str(raw.get("location") or ""),
		"off": off,
		"tax": tax,
		"desc1": str(raw.get("desc1") or ""),
		"desc2": str(raw.get("desc2") or ""),
		"preferred_themes": themes,
		"is_default": bool(raw.get("is_default", False)),
	}


def _write_store(store: dict[str, Any]) -> None:
	os.makedirs(PROFILES_DIR, exist_ok=True)
	with open(PROFILES_PATH, "w", encoding="utf-8") as fh:
		json.dump(store, fh, ensure_ascii=False, indent=2)


def load_store() -> dict[str, Any]:
	return _ensure_store()


def list_profiles() -> list[dict[str, Any]]:
	return deepcopy(_ensure_store()["profiles"])


def get_profile(profile_id: str) -> dict[str, Any] | None:
	for profile in _ensure_store()["profiles"]:
		if profile["id"] == profile_id:
			return deepcopy(profile)
	return None


def get_active_profile() -> dict[str, Any]:
	store = _ensure_store()
	last_id = store["last_profile_id"]
	for profile in store["profiles"]:
		if profile["id"] == last_id:
			return deepcopy(profile)
	for profile in store["profiles"]:
		if profile.get("is_default"):
			return deepcopy(profile)
	return deepcopy(store["profiles"][0])


def set_last_profile(profile_id: str) -> None:
	store = _ensure_store()
	ids = {p["id"] for p in store["profiles"]}
	if profile_id not in ids:
		return
	store["last_profile_id"] = profile_id
	_write_store(store)


def upsert_profile(data: dict[str, Any], make_default: bool = False) -> dict[str, Any]:
	store = _ensure_store()
	profile = _normalize_profile(data)
	if not data.get("id"):
		profile["id"] = str(uuid.uuid4())

	replaced = False
	for i, existing in enumerate(store["profiles"]):
		if existing["id"] == profile["id"]:
			# Preserve default flag unless explicitly changing
			if not make_default:
				profile["is_default"] = existing.get("is_default", False)
			store["profiles"][i] = profile
			replaced = True
			break
	if not replaced:
		profile["is_default"] = make_default or not store["profiles"]
		store["profiles"].append(profile)

	if make_default or profile["is_default"]:
		for p in store["profiles"]:
			p["is_default"] = p["id"] == profile["id"]
		profile["is_default"] = True

	# Guarantee one default
	if not any(p.get("is_default") for p in store["profiles"]):
		store["profiles"][0]["is_default"] = True

	store["last_profile_id"] = profile["id"]
	_write_store(store)
	return deepcopy(profile)


def delete_profile(profile_id: str) -> bool:
	store = _ensure_store()
	profiles = store["profiles"]
	if len(profiles) <= 1:
		return False

	remaining = [p for p in profiles if p["id"] != profile_id]
	if len(remaining) == len(profiles):
		return False

	if not any(p.get("is_default") for p in remaining):
		remaining[0]["is_default"] = True

	if store.get("last_profile_id") == profile_id:
		default = next(p for p in remaining if p.get("is_default"))
		store["last_profile_id"] = default["id"]

	store["profiles"] = remaining
	_write_store(store)
	return True


def profile_from_form(
	*,
	profile_id: str | None,
	name: str,
	sby: str,
	phone: str,
	location: str,
	off: float,
	tax: float,
	desc1: str,
	desc2: str,
	preferred_themes: list[str],
	is_default: bool = False,
) -> dict[str, Any]:
	return {
		"id": profile_id or "",
		"name": name,
		"sby": sby,
		"phone": phone,
		"location": location,
		"off": off,
		"tax": tax,
		"desc1": desc1,
		"desc2": desc2,
		"preferred_themes": preferred_themes,
		"is_default": is_default,
	}

(() => {
	"use strict";

	const MAX_ITEMS = 13;
	const state = {
		editingFactorId: null,
		themes: ["blue", "gray"],
		formats: ["web"],
		metaThemes: [],
		metaFormats: [],
		profiles: [],
		profileEditingId: null,
		generateAbort: null,
	};

	const $ = (id) => document.getElementById(id);
	const els = {
		profileSelect: $("profileSelect"),
		profileSelectBtn: $("profileSelectBtn"),
		profileSelectLabel: $("profileSelectLabel"),
		profileSelectMenu: $("profileSelectMenu"),
		profileSelectWrap: $("profileSelectWrap"),
		factorId: $("factorId"),
		factorDate: $("factorDate"),
		title: $("title"),
		addons: $("addons"),
		sby: $("sby"),
		phone: $("phone"),
		location: $("location"),
		off: $("off"),
		tax: $("tax"),
		desc1: $("desc1"),
		desc2: $("desc2"),
		itemsList: $("itemsList"),
		themeGrid: $("themeGrid"),
		formatGrid: $("formatGrid"),
		themeBlock: $("themeBlock"),
		totalLine: $("totalLine"),
		filesPanel: $("filesPanel"),
		filesList: $("filesList"),
		busyOverlay: $("busyOverlay"),
		listDialog: $("listDialog"),
		profileDialog: $("profileDialog"),
		toastDialog: $("toastDialog"),
		toastText: $("toastText"),
		themeToggle: $("themeToggle"),
	};

	async function api(path, options = {}) {
		const timeoutMs = options.timeoutMs ?? 60000;
		const externalSignal = options.signal;
		delete options.timeoutMs;
		delete options.signal;

		const controller = new AbortController();
		const onExternalAbort = () => controller.abort();
		if (externalSignal) {
			if (externalSignal.aborted) controller.abort();
			else externalSignal.addEventListener("abort", onExternalAbort, { once: true });
		}
		const timer = setTimeout(() => controller.abort(), timeoutMs);

		try {
			const res = await fetch(path, {
				headers: { "Content-Type": "application/json", ...(options.headers || {}) },
				...options,
				signal: controller.signal,
			});
			let data = null;
			const text = await res.text();
			if (text) {
				try {
					data = JSON.parse(text);
				} catch {
					data = { detail: text };
				}
			}
			if (!res.ok) {
				let message = "خطا";
				if (typeof data?.detail === "string") message = data.detail;
				else if (data?.detail?.message) message = data.detail.message;
				const err = new Error(message);
				err.status = res.status;
				err.data = data;
				throw err;
			}
			return data;
		} catch (err) {
			if (err?.name === "AbortError") throw new Error("زمان درخواست تمام شد یا لغو شد");
			throw err;
		} finally {
			clearTimeout(timer);
			if (externalSignal) externalSignal.removeEventListener("abort", onExternalAbort);
		}
	}

	function toast(message) {
		els.toastText.textContent = typeof message === "string" ? message : String(message);
		els.toastDialog.showModal();
	}

	function setBusy(on, message) {
		const el = els.busyOverlay;
		if (!el) return;
		el.hidden = !on;
		el.style.display = on ? "grid" : "none";
		const text = $("busyText");
		if (text && message) text.textContent = message;
		if (!on) state.generateAbort = null;
	}

	function openUrl(url) {
		if (!url) return;
		const a = document.createElement("a");
		a.href = url;
		a.target = "_blank";
		a.rel = "noopener";
		document.body.appendChild(a);
		a.click();
		a.remove();
	}

	function normalizeThemes(list) {
		const ids = new Set((state.metaThemes || []).map((t) => t.id));
		let themes = (list || []).map(String).filter((t) => !ids.size || ids.has(t));
		if (themes.includes("green") && ids.has("cyan") && !themes.includes("cyan")) {
			themes = themes.filter((t) => t !== "green").concat("cyan");
		}
		if (!themes.length) themes = ["blue", "gray"].filter((t) => !ids.size || ids.has(t));
		return [...new Set(themes)];
	}

	function normalizeFormats(list) {
		const allowed = new Set(["word", "web"]);
		const formats = (list || []).map(String).filter((f) => allowed.has(f));
		return formats.length ? [...new Set(formats)] : ["web"];
	}

	async function generate() {
		const sameEdit = state.editingFactorId === Number(els.factorId.value);
		let payload = collectPayload(sameEdit);
		if (!Number.isFinite(payload.factor_id)) {
			toast("آیدی باید عدد باشد");
			return;
		}
		if (!payload.formats.length) {
			toast("حداقل یک نوع خروجی را انتخاب کنید");
			return;
		}
		if (payload.formats.includes("word") && !payload.themes.length) {
			toast("برای Word حداقل یک رنگ را انتخاب کنید");
			return;
		}

		const abort = new AbortController();
		state.generateAbort = abort;
		setBusy(true, "در حال ساخت فاکتور…");
		try {
			let result;
			try {
				result = await api("/api/generate", {
					method: "POST",
					body: JSON.stringify(payload),
					signal: abort.signal,
					timeoutMs: 90000,
				});
			} catch (err) {
				if (err.status === 409) {
					const suggested = err.data?.detail?.suggested_id;
					const ok = confirm(
						`${err.data?.detail?.message || "آیدی اشغال است"}\nادامه و بازنویسی؟` +
							(suggested ? `\nپیشنهاد: ${suggested}` : "")
					);
					if (!ok) return;
					payload = collectPayload(true);
					result = await api("/api/generate", {
						method: "POST",
						body: JSON.stringify(payload),
						signal: abort.signal,
						timeoutMs: 90000,
					});
				} else {
					throw err;
				}
			}

			state.editingFactorId = result.factor_id;
			els.filesPanel.hidden = false;
			els.filesList.innerHTML = "";
			(result.files || []).forEach((file) => {
				const wrap = document.createElement("div");
				wrap.className = "file-row";
				const a = document.createElement("a");
				a.href = file.url;
				a.target = "_blank";
				a.rel = "noopener";
				a.textContent = file.kind === "web" ? file.name : `${file.name} (${file.kind})`;
				if (file.kind !== "web") a.download = file.name;
				wrap.appendChild(a);
				if (file.print_url) {
					const printBtn = document.createElement("a");
					printBtn.href = file.print_url;
					printBtn.target = "_blank";
					printBtn.rel = "noopener";
					printBtn.className = "btn tiny primary";
					printBtn.textContent = "چاپ PDF";
					wrap.appendChild(printBtn);
				}
				els.filesList.appendChild(wrap);
			});

			// Always jump to printable page when web output was requested
			if (result.open_print && (result.print_url || result.invoice_url)) {
				openUrl(result.print_url || result.invoice_url);
			}

			toast(
				result.formats?.includes("web")
					? "فاکتور وب آماده شد — صفحه چاپ باز می‌شود"
					: "فاکتور با موفقیت ساخته شد"
			);
		} catch (err) {
			toast(err.message);
		} finally {
			setBusy(false);
		}
	}

	function collectItems() {
		return [...els.itemsList.querySelectorAll(".item-card")].map((card) => ({
			name: card.querySelector(".item-name").value.trim(),
			count: Number(card.querySelector(".item-count").value || 1),
			price: Number(card.querySelector(".item-price").value || 0),
			free: card.querySelector(".item-free").checked,
		}));
	}

	function collectPayload(overwrite = false) {
		return {
			factor_id: Number(els.factorId.value),
			factor_date: els.factorDate.value.trim(),
			title: els.title.value.trim(),
			addons: els.addons.value.trim(),
			sby: els.sby.value.trim(),
			phone: els.phone.value.trim(),
			location: els.location.value.trim(),
			off: Number(els.off.value || 0),
			tax: Number(els.tax.value || 0),
			desc1: els.desc1.value.trim(),
			desc2: els.desc2.value.trim(),
			items: collectItems().filter((i) => i.name),
			themes: normalizeThemes(state.themes),
			formats: normalizeFormats(state.formats),
			overwrite,
		};
	}

	function applyProfile(profile) {
		if (!profile) return;
		els.sby.value = profile.sby || "";
		els.phone.value = profile.phone || "";
		els.location.value = profile.location || "";
		els.off.value = profile.off ?? 0;
		els.tax.value = profile.tax ?? 0;
		els.desc1.value = profile.desc1 || "";
		els.desc2.value = profile.desc2 || "";
		state.themes = normalizeThemes(profile.preferred_themes || []);
		renderThemes();
		updateTotal();
	}

	function applyFactor(data) {
		els.factorId.value = data.factor_id ?? "";
		els.factorDate.value = data.factor_date || "";
		els.title.value = data.title || "";
		els.addons.value = data.addons || "";
		els.sby.value = data.sby || "";
		els.phone.value = data.phone || "";
		els.location.value = data.location || "";
		els.off.value = data.off ?? 0;
		els.tax.value = data.tax ?? 0;
		els.desc1.value = data.desc1 || "";
		els.desc2.value = data.desc2 || "";
		state.themes = normalizeThemes(data.themes || []);
		state.formats = normalizeFormats(data.formats || ["web"]);
		els.itemsList.innerHTML = "";
		const items = data.items || [];
		if (!items.length) addItemRow();
		else items.forEach((item) => addItemRow(item));
		renderFormats();
		renderThemes();
		updateTotal();
		state.editingFactorId = data.factor_id ?? null;
	}

	function addItemRow(item = null) {
		if (els.itemsList.children.length >= MAX_ITEMS) {
			toast(`حداکثر ${MAX_ITEMS} قلم مجاز است`);
			return;
		}
		const card = document.createElement("div");
		card.className = "item-card";
		card.innerHTML = `
			<div class="item-top">
				<input class="control item-name" placeholder="نام قلم" value="${item?.name ? escapeAttr(item.name) : ""}" />
				<button type="button" class="btn tiny danger item-remove" aria-label="حذف">×</button>
			</div>
			<div class="item-meta">
				<label><span>تعداد</span><input class="control item-count" type="number" min="1" value="${item?.count ?? 1}" /></label>
				<label><span>قیمت</span><input class="control item-price" type="number" min="0" value="${item?.price ?? 0}" /></label>
				<label class="check">
					<input class="item-free" type="checkbox" ${item?.free ? "checked" : ""} />
					<span class="check-ui" aria-hidden="true"></span>
					<span>رایگان</span>
				</label>
			</div>
		`;
		card.querySelector(".item-remove").addEventListener("click", () => {
			card.remove();
			if (!els.itemsList.children.length) addItemRow();
			updateTotal();
		});
		card.querySelectorAll("input").forEach((input) => {
			input.addEventListener("input", updateTotal);
			input.addEventListener("change", updateTotal);
		});
		els.itemsList.appendChild(card);
		updateTotal();
	}

	function escapeAttr(value) {
		return String(value)
			.replaceAll("&", "&amp;")
			.replaceAll('"', "&quot;")
			.replaceAll("<", "&lt;")
			.replaceAll(">", "&gt;");
	}

	function renderFormats() {
		if (!els.formatGrid) return;
		els.formatGrid.innerHTML = "";
		const defs = state.metaFormats.length
			? state.metaFormats
			: [
					{ id: "web", label: "صفحه وب قابل چاپ", hint: "چاپ مستقیم به PDF از مرورگر" },
					{ id: "word", label: "قالب تصویری (Word)", hint: "خروجی DOCX با تصویر تم" },
				];
		defs.forEach((fmt) => {
			const btn = document.createElement("button");
			btn.type = "button";
			btn.className = "format-chip" + (state.formats.includes(fmt.id) ? " active" : "");
			btn.innerHTML = `<strong>${fmt.label}</strong><span>${fmt.hint || ""}</span>`;
			btn.addEventListener("click", () => {
				const idx = state.formats.indexOf(fmt.id);
				if (idx >= 0) {
					if (state.formats.length === 1) {
						toast("حداقل یک نوع خروجی باید انتخاب شود");
						return;
					}
					state.formats.splice(idx, 1);
				} else {
					state.formats.push(fmt.id);
				}
				renderFormats();
			});
			els.formatGrid.appendChild(btn);
		});
		const sub = $("themeSubhead");
		if (sub) {
			const hasWord = state.formats.includes("word");
			const hasWeb = state.formats.includes("web");
			if (hasWord && hasWeb) sub.textContent = "رنگ قالب تصویری / رنگ صفحه وب";
			else if (hasWeb) sub.textContent = "رنگ صفحه وب";
			else sub.textContent = "رنگ قالب تصویری";
		}
	}

	function renderThemes() {
		els.themeGrid.innerHTML = "";
		state.themes = normalizeThemes(state.themes);
		state.metaThemes.forEach((theme) => {
			const btn = document.createElement("button");
			btn.type = "button";
			const active = state.themes.includes(theme.id);
			btn.className = "theme-chip" + (active ? " active" : "");
			btn.setAttribute("aria-pressed", active ? "true" : "false");
			btn.innerHTML = `<img src="${theme.image}" alt="${theme.label}" loading="lazy" /><span>${theme.label}</span>`;
			btn.addEventListener("click", () => {
				const idx = state.themes.indexOf(theme.id);
				if (idx >= 0) {
					if (state.themes.length === 1) {
						toast("حداقل یک رنگ باید انتخاب شود");
						return;
					}
					state.themes.splice(idx, 1);
				} else {
					state.themes.push(theme.id);
				}
				renderThemes();
			});
			els.themeGrid.appendChild(btn);
		});
	}

	async function updateTotal() {
		try {
			const data = await api("/api/total", {
				method: "POST",
				body: JSON.stringify({
					items: collectItems(),
					off: Number(els.off.value || 0),
					tax: Number(els.tax.value || 0),
				}),
			});
			els.totalLine.textContent = `جمع کل: ${data.formatted} ریال`;
		} catch {
			els.totalLine.textContent = "جمع کل: ۰ ریال";
		}
	}

	function closeProfileMenu() {
		els.profileSelectWrap.classList.remove("open");
		els.profileSelectBtn.setAttribute("aria-expanded", "false");
		els.profileSelectMenu.hidden = true;
	}

	function fillProfiles(profiles, selectedId) {
		state.profiles = profiles || [];
		els.profileSelectMenu.innerHTML = "";
		let selected = selectedId || els.profileSelect.value || state.profiles[0]?.id || "";
		if (selected && !state.profiles.some((p) => p.id === selected)) {
			selected = state.profiles[0]?.id || "";
		}
		state.profiles.forEach((p) => {
			const li = document.createElement("li");
			li.role = "option";
			li.dataset.value = p.id;
			li.textContent = p.name + (p.is_default ? " ★" : "");
			if (p.id === selected) li.classList.add("active");
			li.addEventListener("click", async () => {
				els.profileSelect.value = p.id;
				els.profileSelectLabel.textContent = li.textContent;
				closeProfileMenu();
				[...els.profileSelectMenu.children].forEach((c) => c.classList.toggle("active", c === li));
				const profile = await api(`/api/profiles/${p.id}/activate`, { method: "POST" });
				applyProfile(profile);
			});
			els.profileSelectMenu.appendChild(li);
		});
		els.profileSelect.value = selected;
		const current = state.profiles.find((p) => p.id === selected);
		els.profileSelectLabel.textContent = current
			? current.name + (current.is_default ? " ★" : "")
			: "انتخاب پروفایل";
	}

	async function newFactor() {
		const boot = await api("/api/bootstrap");
		state.editingFactorId = null;
		state.formats = ["web"];
		els.filesPanel.hidden = true;
		els.filesList.innerHTML = "";
		fillProfiles(boot.profiles, boot.profile?.id);
		els.factorId.value = boot.factor_id;
		els.factorDate.value = boot.factor_date;
		els.title.value = "";
		els.addons.value = "";
		els.itemsList.innerHTML = "";
		addItemRow();
		applyProfile(boot.profile);
		renderFormats();
		renderThemes();
	}

	async function saveFactor(forceOverwrite = false) {
		const payload = collectPayload(forceOverwrite || state.editingFactorId === Number(els.factorId.value));
		if (!Number.isFinite(payload.factor_id)) {
			toast("آیدی باید عدد باشد");
			return;
		}
		try {
			const saved = await api("/api/factors", {
				method: "POST",
				body: JSON.stringify(payload),
			});
			state.editingFactorId = saved.factor_id;
			toast(`فاکتور با آیدی ${saved.factor_id} ذخیره شد`);
		} catch (err) {
			if (err.status === 409) {
				const suggested = err.data?.detail?.suggested_id;
				const ok = confirm(
					`${err.data?.detail?.message || "آیدی اشغال است"}\nبازنویسی شود؟` +
						(suggested ? `\nپیشنهاد: ${suggested}` : "")
				);
				if (ok) await saveFactor(true);
				return;
			}
			toast(err.message);
		}
	}

	async function openList() {
		const data = await api("/api/factors");
		if (!data.factors?.length) {
			toast("هنوز فاکتوری ذخیره نشده است");
			return;
		}
		const box = $("factorList");
		box.innerHTML = "";
		data.factors.forEach((f) => {
			const row = document.createElement("div");
			row.className = "factor-row";
			row.innerHTML = `
				<div>
					<strong>#${f.factor_id} — ${f.title || "بدون عنوان"}</strong>
					<small>${f.factor_date || ""} · ${f.status || ""}</small>
				</div>
				<div class="row-actions">
					<button type="button" class="btn tiny load">باز کردن</button>
					<button type="button" class="btn tiny danger del">حذف</button>
				</div>
			`;
			row.querySelector(".load").addEventListener("click", async () => {
				const factor = await api(`/api/factors/${f.factor_id}`);
				applyFactor(factor);
				els.listDialog.close();
			});
			row.querySelector(".del").addEventListener("click", async () => {
				if (!confirm(`فاکتور ${f.factor_id} حذف شود؟`)) return;
				await api(`/api/factors/${f.factor_id}`, { method: "DELETE" });
				if (state.editingFactorId === f.factor_id) await newFactor();
				row.remove();
			});
			box.appendChild(row);
		});
		els.listDialog.showModal();
	}

	function openProfileDialog(isNew) {
		state.profileEditingId = isNew ? null : els.profileSelect.value;
		$("profileDialogTitle").textContent = isNew ? "پروفایل جدید" : "ویرایش پروفایل";
		if (isNew) {
			$("pName").value = "";
			$("pSby").value = els.sby.value;
			$("pPhone").value = els.phone.value;
			$("pLocation").value = els.location.value;
			$("pOff").value = els.off.value;
			$("pTax").value = els.tax.value;
			$("pDefault").checked = false;
		} else {
			const current = state.profiles.find((p) => p.id === els.profileSelect.value);
			$("pName").value = current?.name || "";
			$("pSby").value = els.sby.value;
			$("pPhone").value = els.phone.value;
			$("pLocation").value = els.location.value;
			$("pOff").value = els.off.value;
			$("pTax").value = els.tax.value;
			$("pDefault").checked = !!current?.is_default;
		}
		els.profileDialog.showModal();
	}

	async function saveProfileFromDialog() {
		const payload = {
			id: state.profileEditingId,
			name: $("pName").value.trim(),
			sby: $("pSby").value.trim(),
			phone: $("pPhone").value.trim(),
			location: $("pLocation").value.trim(),
			off: Number($("pOff").value || 0),
			tax: Number($("pTax").value || 0),
			desc1: els.desc1.value.trim(),
			desc2: els.desc2.value.trim(),
			preferred_themes: normalizeThemes(state.themes),
			is_default: $("pDefault").checked,
		};
		if (!payload.name) {
			toast("نام پروفایل لازم است");
			return;
		}
		const saved = await api("/api/profiles", {
			method: "POST",
			body: JSON.stringify(payload),
		});
		const list = await api("/api/profiles");
		fillProfiles(list.profiles, saved.id);
		applyProfile(saved);
		els.profileDialog.close();
	}

	function initThemeMode() {
		const saved = localStorage.getItem("fm-theme");
		const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
		const mode = saved || (prefersDark ? "dark" : "light");
		document.documentElement.setAttribute("data-theme", mode);
		els.themeToggle.addEventListener("click", () => {
			const next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
			document.documentElement.setAttribute("data-theme", next);
			localStorage.setItem("fm-theme", next);
		});
	}

	async function init() {
		initThemeMode();
		setBusy(false);
		const meta = await api("/api/meta");
		state.metaThemes = meta.themes || [];
		state.metaFormats = meta.formats || [];
		renderFormats();
		renderThemes();

		els.profileSelectBtn.addEventListener("click", () => {
			const open = !els.profileSelectWrap.classList.contains("open");
			closeProfileMenu();
			if (open) {
				els.profileSelectWrap.classList.add("open");
				els.profileSelectBtn.setAttribute("aria-expanded", "true");
				els.profileSelectMenu.hidden = false;
			}
		});
		document.addEventListener("click", (ev) => {
			if (!els.profileSelectWrap.contains(ev.target)) closeProfileMenu();
		});

		$("btnAddItem").addEventListener("click", () => addItemRow());
		$("btnNew").addEventListener("click", () => newFactor().catch((e) => toast(e.message)));
		$("btnSave").addEventListener("click", () => saveFactor().catch((e) => toast(e.message)));
		$("btnOpen").addEventListener("click", () => openList().catch((e) => toast(e.message)));
		$("btnGenerate").addEventListener("click", () => generate());
		$("listCloseBtn").addEventListener("click", () => els.listDialog.close());
		$("toastOkBtn").addEventListener("click", () => els.toastDialog.close());
		$("busyCancel").addEventListener("click", () => {
			if (state.generateAbort) state.generateAbort.abort();
			setBusy(false);
		});
		$("btnProfileNew").addEventListener("click", () => openProfileDialog(true));
		$("btnProfileEdit").addEventListener("click", () => openProfileDialog(false));
		$("btnProfileDelete").addEventListener("click", async () => {
			const id = els.profileSelect.value;
			if (!id) return;
			if (!confirm("پروفایل حذف شود؟")) return;
			try {
				const data = await api(`/api/profiles/${id}`, { method: "DELETE" });
				const nextId = data.profiles[0]?.id;
				fillProfiles(data.profiles, nextId);
				if (nextId) {
					const profile = await api(`/api/profiles/${nextId}/activate`, { method: "POST" });
					applyProfile(profile);
				}
			} catch (e) {
				toast(e.message);
			}
		});
		$("profileCancelBtn").addEventListener("click", () => els.profileDialog.close());
		$("profileSaveBtn").addEventListener("click", () => {
			saveProfileFromDialog().catch((e) => toast(e.message));
		});

		["off", "tax"].forEach((id) => $(id).addEventListener("input", updateTotal));

		await newFactor();
	}

	init().catch((err) => {
		console.error(err);
		toast(err.message || "خطا در بارگذاری");
	});
})();

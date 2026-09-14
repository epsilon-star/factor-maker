from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox
import sys
import os
import jdatetime

from factor_maker_ui import (
	FactorMakerUI,
	ProfileDialog,
	load_app_stylesheet,
	apply_app_font,
	ProgressDialog,
	GenerateWorker,
)
from engine import (
	PersianFactor,
	APP_ROOT,
	image_map,
	DEFAULT_THEMES,
	THEME_LABELS,
)
import profiles as profile_store

TEMPLATE_PATH = os.path.join(APP_ROOT, "templates", "template.docx")
OUTPUT_DIR = os.path.join(APP_ROOT, "output")
IMAGES_DIR = os.path.join(APP_ROOT, "images")


def required_asset_errors(themes: list[str] | None = None):
	"""Return missing required asset paths for the selected themes."""
	missing = []
	if not os.path.isfile(TEMPLATE_PATH):
		missing.append(TEMPLATE_PATH)

	selected = themes if themes is not None else list(DEFAULT_THEMES)
	for theme_name in selected:
		image_name = image_map.get(theme_name)
		if not image_name:
			continue
		image_path = os.path.join(IMAGES_DIR, f"{image_name}.jpg")
		if not os.path.isfile(image_path):
			label = THEME_LABELS.get(theme_name, theme_name)
			missing.append(f"{image_path} ({label})")
	return missing


def is_system_dark() -> bool:
	scheme = QGuiApplication.styleHints().colorScheme()
	return scheme == Qt.ColorScheme.Dark


def apply_system_stylesheet(app: QApplication) -> None:
	stylesheet = load_app_stylesheet(APP_ROOT, dark=is_system_dark())
	if stylesheet:
		app.setStyleSheet(stylesheet)


class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("فاکتور ساز متانویا")
		self.setMinimumSize(1000, 720)
		self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

		os.makedirs(OUTPUT_DIR, exist_ok=True)
		os.makedirs(os.path.join(APP_ROOT, "templates"), exist_ok=True)
		os.makedirs(IMAGES_DIR, exist_ok=True)
		os.makedirs(os.path.join(APP_ROOT, "profiles"), exist_ok=True)

		self.ui = FactorMakerUI()
		self.setCentralWidget(self.ui)
		self._engine = PersianFactor(self.showMessage)
		self._worker = None
		self._progress_dialog = None

		self.ui.generate_btn.clicked.connect(self.makeFactor)
		self.ui.profile_changed.connect(self.on_profile_selected)
		self.ui.profile_new_requested.connect(self.on_profile_new)
		self.ui.profile_edit_requested.connect(self.on_profile_edit)
		self.ui.profile_delete_requested.connect(self.on_profile_delete)
		self.ui.total_inputs_changed.connect(self.update_total_preview)

		tt = jdatetime.datetime.now()
		self.ui.id_input.setText(str(tt.year)[1:] + "00" + "000")
		self.ui.date_input.setText(tt.strftime("%Y/%m/%d"))

		self.reload_profiles()
		self.update_total_preview()

		# Soft warning for default themes only (cyan/red optional)
		missing = required_asset_errors(list(DEFAULT_THEMES))
		if missing:
			listed = "\n".join(missing)
			self.showMessage(
				"فایل‌های مورد نیاز پیدا نشدند:\n"
				f"{listed}\n\n"
				"لطفا قالب و تصاویر را در پوشه برنامه قرار دهید."
			)

	def showMessage(self, message: str):
		msg_box = QMessageBox(self)
		msg_box.setWindowTitle("پیام")
		msg_box.setText(message)
		msg_box.exec()

	def reload_profiles(self, prefer_id: str | None = None):
		store = profile_store.load_store()
		profiles = store["profiles"]
		selected = prefer_id or store.get("last_profile_id")
		self.ui.set_profiles(profiles, selected)
		active = profile_store.get_profile(selected) if selected else None
		if not active:
			active = profile_store.get_active_profile()
		self.ui.apply_profile(active)
		profile_store.set_last_profile(active["id"])

	def on_profile_selected(self, profile_id: str):
		profile = profile_store.get_profile(profile_id)
		if not profile:
			return
		self.ui.apply_profile(profile)
		profile_store.set_last_profile(profile_id)
		self.update_total_preview()

	def on_profile_new(self):
		dialog = ProfileDialog(self, is_new=True)
		if dialog.exec() != ProfileDialog.DialogCode.Accepted:
			return
		data = dialog.result_profile()
		saved = profile_store.upsert_profile(data, make_default=data.get("is_default", False))
		self.reload_profiles(prefer_id=saved["id"])

	def on_profile_edit(self):
		pid = self.ui.current_profile_id()
		if not pid:
			return
		profile = profile_store.get_profile(pid)
		if not profile:
			return
		dialog = ProfileDialog(self, profile=profile, is_new=False)
		if dialog.exec() != ProfileDialog.DialogCode.Accepted:
			return
		data = dialog.result_profile()
		saved = profile_store.upsert_profile(data, make_default=data.get("is_default", False))
		self.reload_profiles(prefer_id=saved["id"])

	def on_profile_delete(self):
		pid = self.ui.current_profile_id()
		if not pid:
			return
		profiles = profile_store.list_profiles()
		if len(profiles) <= 1:
			self.showMessage("حداقل یک پروفایل باید باقی بماند")
			return
		profile = profile_store.get_profile(pid)
		name = (profile or {}).get("name", "")
		confirm = QMessageBox.question(
			self,
			"حذف پروفایل",
			f"پروفایل «{name}» حذف شود؟",
			QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
		)
		if confirm != QMessageBox.StandardButton.Yes:
			return
		if not profile_store.delete_profile(pid):
			self.showMessage("حذف پروفایل ممکن نشد")
			return
		self.reload_profiles()

	def update_total_preview(self):
		items, error = self.ui.get_table_items()
		if error or not items:
			self.ui.set_total_text("جمع کل: ۰ ریال")
			return
		total = self._engine.calculateTotal(
			list(items.values()),
			float(self.ui.tax_input.value()),
			float(self.ui.off_input.value()),
		)
		formatted = self._engine.formatNumber(str(total))
		self.ui.set_total_text(f"جمع کل: {formatted} ریال")

	def makeFactor(self):
		themes = self.ui.selected_themes()
		if not themes:
			self.showMessage("حداقل یک رنگ فاکتور را انتخاب کنید")
			return

		missing = required_asset_errors(themes)
		if missing:
			listed = "\n".join(missing)
			self.showMessage(f"فایل‌های مورد نیاز پیدا نشدند:\n{listed}")
			return

		if not len(self.ui.id_input.text()):
			self.showMessage("لطفا آیدی را مشخص کنید")
			return
		if not len(self.ui.date_input.text()):
			self.showMessage("لطفا تاریخ را مشخص کنید")
			return
		if not len(self.ui.title_input.text()):
			self.showMessage("لطفا عنوان را مشخص کنید")
			return
		if not len(self.ui.from_input.text()):
			self.showMessage("لطفا از را مشخص کنید")
			return
		if not len(self.ui.phone_input.text()):
			self.showMessage("لطفا تلفن را مشخص کنید")
			return
		if not len(self.ui.location_input.toPlainText()):
			self.showMessage("لطفا آدرس را مشخص کنید")
			return

		items, items_error = self.ui.get_table_items()
		if items_error:
			self.showMessage(items_error)
			return

		try:
			factor_id = int(self.ui.id_input.text().strip())
		except ValueError:
			self.showMessage("آیدی باید عدد باشد")
			return

		d = self._engine
		datas = {
			"factor_id": d.persianNumber(self.ui.id_input.text()),
			"factor_date": d.persianNumber(self.ui.date_input.text()),
			"factor_addons": d.persianNumber(self.ui.addons_input.text()),
			"title": self.ui.title_input.text(),
			"sby": self.ui.from_input.text(),
			"phone": self.ui.phone_input.text(),
			"location": self.ui.location_input.toPlainText(),
			"items": items,
			"off": float(self.ui.off_input.value()),
			"tax": float(self.ui.tax_input.value()),
			"factor_disc_1": d.persianNumber(self.ui.desc1_input.toPlainText()),
			"factor_disc_2": d.persianNumber(self.ui.desc2_input.toPlainText()),
		}

		# Silent engine messages while progress dialog is open; show result after
		self._pending_themes = themes
		self._engine.msgShow = None
		self.ui.generate_btn.setEnabled(False)

		self._progress_dialog = ProgressDialog(self)
		self._worker = GenerateWorker(
			self._engine,
			factor_id,
			TEMPLATE_PATH,
			datas,
			themes,
			parent=self,
		)
		self._worker.progress.connect(self._progress_dialog.update_progress)
		self._worker.finished_ok.connect(self._on_generate_finished)
		self._progress_dialog.show()
		self._worker.start()

	def _on_generate_finished(self, ok: bool):
		themes = getattr(self, "_pending_themes", [])
		if self._progress_dialog is not None:
			self._progress_dialog.accept()
			self._progress_dialog = None
		self.ui.generate_btn.setEnabled(True)
		self._engine.msgShow = self.showMessage

		if self._worker is not None:
			self._worker.wait(100)
			self._worker.deleteLater()
			self._worker = None

		if ok:
			theme_names = "، ".join(THEME_LABELS.get(t, t) for t in themes)
			self.showMessage(f"فاکتور با موفقیت ساخته شد\nرنگ‌ها: {theme_names}")
			pid = self.ui.current_profile_id()
			if pid:
				current = profile_store.get_profile(pid)
				if current:
					current["preferred_themes"] = themes
					profile_store.upsert_profile(current)
		else:
			self.showMessage("خطا در تولید فاکتور. لطفا دوباره تلاش کنید.")


def main():
	app = QApplication(sys.argv)
	app.setStyle("Fusion")
	app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
	apply_app_font(app)
	apply_system_stylesheet(app)

	# Follow OS light/dark changes live
	QGuiApplication.styleHints().colorSchemeChanged.connect(
		lambda *_: apply_system_stylesheet(app)
	)

	window = MainWindow()
	window.show()
	sys.exit(app.exec())


if __name__ == "__main__":
	main()

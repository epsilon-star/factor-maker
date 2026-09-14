from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
	QWidget, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout,
	QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QDoubleSpinBox,
	QScrollArea, QSizePolicy, QMessageBox, QFrame, QComboBox, QDialog,
	QDialogButtonBox, QFormLayout, QCheckBox, QAbstractItemView, QGridLayout,
)
import os

from engine import DEFAULT_THEMES, THEME_LABELS, image_map, FONT_PATH
import engine as engine_mod

MAX_ITEMS = 13

THEME_OBJECT_NAMES = {
	"blue": "ThemeBlue",
	"cyan": "ThemeCyan",
	"gray": "ThemeGray",
	"red": "ThemeRed",
}


def load_app_stylesheet(app_root: str, dark: bool = False) -> str:
	name = "styles_dark.qss" if dark else "styles.qss"
	path = os.path.join(app_root, name)
	try:
		with open(path, "r", encoding="utf-8") as fh:
			return fh.read()
	except OSError:
		# Fall back to light theme if dark file is missing
		if dark:
			return load_app_stylesheet(app_root, dark=False)
		return ""


def apply_app_font(app) -> str | None:
	"""Load fonts/IRANSansXV.ttf for the UI (and invoice engine). Returns family name."""
	family = None
	if os.path.isfile(FONT_PATH):
		font_id = QFontDatabase.addApplicationFont(FONT_PATH)
		if font_id != -1:
			families = QFontDatabase.applicationFontFamilies(font_id)
			if families:
				family = families[0]

	if not family:
		installed = set(QFontDatabase.families())
		for candidate in ("IRANSansX", "IRANSansXV", "IRANSans", "Vazirmatn", "Vazir", "Tahoma"):
			if candidate in installed:
				family = candidate
				break

	if family:
		app.setFont(QFont(family, 10))
		engine_mod.INVOICE_FONT_NAME = family
		return family

	app.setFont(QFont("Tahoma", 10))
	return None


def try_set_vazir_font(app) -> None:
	"""Backward-compatible alias."""
	apply_app_font(app)


class ProfileDialog(QDialog):
	def __init__(self, parent=None, profile: dict | None = None, is_new: bool = True):
		super().__init__(parent)
		self.setWindowTitle("پروفایل جدید" if is_new else "ویرایش پروفایل")
		self.setMinimumWidth(440)
		self._profile_id = (profile or {}).get("id") if profile and not is_new else None

		layout = QVBoxLayout(self)
		layout.setContentsMargins(16, 16, 16, 16)
		layout.setSpacing(12)
		form = QFormLayout()
		form.setSpacing(10)
		form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

		self.name_input = QLineEdit((profile or {}).get("name", ""))
		self.sby_input = QLineEdit((profile or {}).get("sby", ""))
		self.phone_input = QLineEdit((profile or {}).get("phone", ""))
		self.location_input = QTextEdit()
		self.location_input.setFixedHeight(70)
		self.location_input.setPlainText((profile or {}).get("location", ""))
		self.off_input = QDoubleSpinBox()
		self.tax_input = QDoubleSpinBox()
		self.off_input.setRange(0, 100)
		self.tax_input.setRange(0, 100)
		self.off_input.setSuffix(" %")
		self.tax_input.setSuffix(" %")
		self.off_input.setValue(float((profile or {}).get("off", 0) or 0))
		self.tax_input.setValue(float((profile or {}).get("tax", 0) or 0))
		self.desc1_input = QTextEdit()
		self.desc2_input = QTextEdit()
		self.desc1_input.setFixedHeight(50)
		self.desc2_input.setFixedHeight(50)
		self.desc1_input.setPlainText((profile or {}).get("desc1", ""))
		self.desc2_input.setPlainText((profile or {}).get("desc2", ""))
		self.default_check = QCheckBox("پروفایل پیش‌فرض")
		self.default_check.setChecked(bool((profile or {}).get("is_default", False)))

		preferred = list((profile or {}).get("preferred_themes") or DEFAULT_THEMES)
		self.theme_checks: dict[str, QCheckBox] = {}
		theme_row = QWidget()
		theme_layout = QHBoxLayout(theme_row)
		theme_layout.setContentsMargins(0, 0, 0, 0)
		theme_layout.setSpacing(10)
		for key in image_map:
			cb = QCheckBox(THEME_LABELS.get(key, key))
			cb.setChecked(key in preferred)
			self.theme_checks[key] = cb
			theme_layout.addWidget(cb)
		theme_layout.addStretch()

		form.addRow("نام پروفایل", self.name_input)
		form.addRow("از", self.sby_input)
		form.addRow("تلفن", self.phone_input)
		form.addRow("آدرس", self.location_input)
		form.addRow("تخفیف", self.off_input)
		form.addRow("مالیات", self.tax_input)
		form.addRow("توضیحات 1", self.desc1_input)
		form.addRow("توضیحات 2", self.desc2_input)
		form.addRow("رنگ‌های پیش‌فرض", theme_row)
		form.addRow("", self.default_check)

		layout.addLayout(form)
		buttons = QDialogButtonBox(
			QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
		)
		buttons.button(QDialogButtonBox.StandardButton.Save).setText("ذخیره")
		buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("انصراف")
		buttons.accepted.connect(self._on_accept)
		buttons.rejected.connect(self.reject)
		layout.addWidget(buttons)

	def _on_accept(self):
		if not self.name_input.text().strip():
			QMessageBox.warning(self, "پیام", "نام پروفایل را وارد کنید")
			return
		themes = [k for k, cb in self.theme_checks.items() if cb.isChecked()]
		if not themes:
			QMessageBox.warning(self, "پیام", "حداقل یک رنگ را انتخاب کنید")
			return
		self.accept()

	def result_profile(self) -> dict:
		return {
			"id": self._profile_id or "",
			"name": self.name_input.text().strip(),
			"sby": self.sby_input.text().strip(),
			"phone": self.phone_input.text().strip(),
			"location": self.location_input.toPlainText().strip(),
			"off": float(self.off_input.value()),
			"tax": float(self.tax_input.value()),
			"desc1": self.desc1_input.toPlainText(),
			"desc2": self.desc2_input.toPlainText(),
			"preferred_themes": [k for k, cb in self.theme_checks.items() if cb.isChecked()],
			"is_default": self.default_check.isChecked(),
		}


class FactorMakerUI(QWidget):
	profile_changed = Signal(str)
	profile_new_requested = Signal()
	profile_edit_requested = Signal()
	profile_delete_requested = Signal()
	total_inputs_changed = Signal()

	def __init__(self):
		super().__init__()
		self.setObjectName("FactorRoot")
		self.setWindowTitle("FactorMaker")
		self.setMinimumSize(1080, 720)
		self.theme_buttons: dict[str, QPushButton] = {}
		self._loading_profile = False
		self.init_ui()

	def init_ui(self):
		root = QVBoxLayout(self)
		root.setContentsMargins(16, 16, 16, 16)
		root.setSpacing(12)

		root.addWidget(self._build_header())

		body_scroll = QScrollArea()
		body_scroll.setWidgetResizable(True)
		body_scroll.setFrameShape(QFrame.Shape.NoFrame)
		body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
		body_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

		card = QFrame()
		card.setObjectName("ContentCard")
		card_layout = QHBoxLayout(card)
		card_layout.setContentsMargins(16, 16, 16, 16)
		card_layout.setSpacing(0)

		# In RTL, first added widget sits on the right.
		card_layout.addWidget(self._build_meta_column(), stretch=5)
		divider = QFrame()
		divider.setObjectName("SoftDivider")
		divider.setFrameShape(QFrame.Shape.NoFrame)
		card_layout.addSpacing(14)
		card_layout.addWidget(divider)
		card_layout.addSpacing(14)
		card_layout.addWidget(self._build_items_column(), stretch=7)

		body_scroll.setWidget(card)
		root.addWidget(body_scroll, stretch=1)
		root.addWidget(self._build_footer())
		self._wire_total_signals()

	def _build_header(self) -> QFrame:
		header = QFrame()
		header.setObjectName("HeaderBar")
		layout = QVBoxLayout(header)
		layout.setContentsMargins(16, 14, 16, 14)
		layout.setSpacing(12)

		top = QHBoxLayout()
		top.setSpacing(10)

		self.brand_title = QLabel("فاکتور ساز متانویا")
		self.brand_title.setObjectName("BrandTitle")
		top.addWidget(self.brand_title, alignment=Qt.AlignmentFlag.AlignVCenter)
		top.addStretch(1)

		profile_wrap = QFrame()
		profile_wrap.setObjectName("SectionBox")
		profile_row = QHBoxLayout(profile_wrap)
		profile_row.setContentsMargins(0, 0, 0, 0)
		profile_row.setSpacing(8)

		profile_label = QLabel("پروفایل")
		profile_label.setObjectName("FieldLabel")
		self.profile_combo = QComboBox()
		self.profile_combo.setMinimumWidth(170)
		self.profile_combo.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
		self.profile_combo.currentIndexChanged.connect(self._on_profile_combo_changed)

		self.profile_new_btn = QPushButton("جدید")
		self.profile_edit_btn = QPushButton("ویرایش")
		self.profile_delete_btn = QPushButton("حذف")
		self.profile_new_btn.setObjectName("SecondaryButton")
		self.profile_edit_btn.setObjectName("SecondaryButton")
		self.profile_delete_btn.setObjectName("DangerButton")
		self.profile_new_btn.clicked.connect(self.profile_new_requested.emit)
		self.profile_edit_btn.clicked.connect(self.profile_edit_requested.emit)
		self.profile_delete_btn.clicked.connect(self.profile_delete_requested.emit)

		profile_row.addWidget(profile_label)
		profile_row.addWidget(self.profile_combo)
		profile_row.addWidget(self.profile_new_btn)
		profile_row.addWidget(self.profile_edit_btn)
		profile_row.addWidget(self.profile_delete_btn)
		top.addWidget(profile_wrap)
		layout.addLayout(top)

		theme_row = QHBoxLayout()
		theme_row.setSpacing(8)
		theme_label = QLabel("رنگ فاکتور")
		theme_label.setObjectName("FieldLabel")
		theme_row.addWidget(theme_label)
		for key in image_map:
			btn = QPushButton(THEME_LABELS.get(key, key))
			btn.setCheckable(True)
			btn.setCursor(Qt.CursorShape.PointingHandCursor)
			btn.setChecked(key in DEFAULT_THEMES)
			btn.toggled.connect(lambda checked, k=key, b=btn: self._on_theme_button_toggled(b, k, checked))
			self.theme_buttons[key] = btn
			self._refresh_theme_chip_style(btn, key, btn.isChecked())
			theme_row.addWidget(btn)
		theme_row.addStretch(1)
		hint = QLabel("حداقل یک رنگ را انتخاب کنید")
		hint.setObjectName("HintLabel")
		theme_row.addWidget(hint)
		layout.addLayout(theme_row)
		return header

	def _refresh_theme_chip_style(self, btn: QPushButton, key: str, checked: bool):
		btn.setObjectName(THEME_OBJECT_NAMES.get(key, "ThemeChip") if checked else "ThemeChip")
		style = btn.style()
		if style is not None:
			style.unpolish(btn)
			style.polish(btn)
		btn.update()

	def _build_meta_column(self) -> QWidget:
		col = QWidget()
		layout = QVBoxLayout(col)
		layout.setContentsMargins(0, 0, 0, 0)
		layout.setSpacing(14)

		invoice_box = QFrame()
		invoice_box.setObjectName("SectionBox")
		invoice_layout = QVBoxLayout(invoice_box)
		invoice_layout.setContentsMargins(0, 0, 0, 0)
		invoice_layout.setSpacing(8)
		invoice_title = QLabel("اطلاعات فاکتور")
		invoice_title.setObjectName("SectionTitle")
		invoice_layout.addWidget(invoice_title)

		self.id_input = QLineEdit()
		self.date_input = QLineEdit()
		self.addons_input = QLineEdit()
		self.title_input = QLineEdit()
		invoice_layout.addLayout(self._field_row(self.id_input, "آیدی"))
		invoice_layout.addLayout(self._field_row(self.date_input, "تاریخ"))
		invoice_layout.addLayout(self._field_row(self.title_input, "عنوان"))
		invoice_layout.addLayout(self._field_row(self.addons_input, "افزونه‌ها"))
		layout.addWidget(invoice_box)

		seller_box = QFrame()
		seller_box.setObjectName("SectionBox")
		seller_layout = QVBoxLayout(seller_box)
		seller_layout.setContentsMargins(0, 0, 0, 0)
		seller_layout.setSpacing(8)
		seller_title = QLabel("اطلاعات فروشنده")
		seller_title.setObjectName("SectionTitle")
		seller_layout.addWidget(seller_title)

		self.from_input = QLineEdit()
		self.phone_input = QLineEdit()
		self.location_input = QTextEdit()
		self.location_input.setMinimumHeight(100)
		self.location_input.setMaximumHeight(140)
		seller_layout.addLayout(self._field_row(self.from_input, "از"))
		seller_layout.addLayout(self._field_row(self.phone_input, "تلفن"))
		seller_layout.addLayout(self._field_row(self.location_input, "آدرس"))
		layout.addWidget(seller_box)
		layout.addStretch(1)
		return col

	def _build_items_column(self) -> QWidget:
		col = QWidget()
		layout = QVBoxLayout(col)
		layout.setContentsMargins(0, 0, 0, 0)
		layout.setSpacing(10)

		items_title = QLabel("آیتم‌ها")
		items_title.setObjectName("SectionTitle")
		layout.addWidget(items_title)

		self.items_table = QTableWidget(0, 4)
		self.items_table.setHorizontalHeaderLabels(["عملیات", "قیمت (ریال)", "تعداد", "نام"])
		self.items_table.setAlternatingRowColors(True)
		self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # type: ignore
		self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # type: ignore
		self.items_table.verticalHeader().setVisible(False)
		self.items_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
		self.items_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
		self.items_table.setShowGrid(True)
		self.items_table.setMinimumHeight(280)
		self.items_table.itemChanged.connect(self._on_table_item_changed)
		layout.addWidget(self.items_table, stretch=1)

		self.add_item_btn = QPushButton("＋  افزودن آیتم")
		self.add_item_btn.setObjectName("AddItemButton")
		self.add_item_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.add_item_btn.clicked.connect(self.add_item_row)
		layout.addWidget(self.add_item_btn)

		rates = QGridLayout()
		rates.setHorizontalSpacing(12)
		rates.setVerticalSpacing(6)
		self.off_input = QDoubleSpinBox()
		self.tax_input = QDoubleSpinBox()
		self.off_input.setRange(0, 100)
		self.tax_input.setRange(0, 100)
		self.off_input.setSuffix(" %")
		self.tax_input.setSuffix(" %")
		rates.addLayout(self._field_row(self.off_input, "تخفیف"), 0, 0)
		rates.addLayout(self._field_row(self.tax_input, "مالیات"), 0, 1)
		layout.addLayout(rates)

		notes_title = QLabel("توضیحات")
		notes_title.setObjectName("SectionTitle")
		layout.addWidget(notes_title)

		self.desc1_input = QTextEdit()
		self.desc2_input = QTextEdit()
		self.desc1_input.setFixedHeight(58)
		self.desc2_input.setFixedHeight(58)
		layout.addLayout(self._field_row(self.desc1_input, "توضیحات ۱"))
		layout.addLayout(self._field_row(self.desc2_input, "توضیحات ۲"))
		return col

	def _build_footer(self) -> QFrame:
		footer = QFrame()
		footer.setObjectName("FooterBar")
		layout = QHBoxLayout(footer)
		layout.setContentsMargins(16, 12, 16, 12)
		layout.setSpacing(12)

		self.total_label = QLabel("جمع کل: ۰ ریال")
		self.total_label.setObjectName("TotalLabel")
		layout.addWidget(self.total_label)
		layout.addStretch(1)

		self.generate_btn = QPushButton("تولید فاکتور")
		self.generate_btn.setObjectName("PrimaryButton")
		self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.generate_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
		layout.addWidget(self.generate_btn)
		return footer

	def _field_row(self, widget: QWidget, label_text: str) -> QHBoxLayout:
		row = QHBoxLayout()
		row.setContentsMargins(0, 0, 0, 0)
		row.setSpacing(10)
		label = QLabel(label_text)
		label.setObjectName("FieldLabel")
		label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
		# In RTL: widget first (expands leftward), label on the right edge.
		row.addWidget(widget, stretch=1)
		row.addWidget(label, stretch=0)
		return row

	def _wire_total_signals(self):
		self.off_input.valueChanged.connect(lambda *_: self.total_inputs_changed.emit())
		self.tax_input.valueChanged.connect(lambda *_: self.total_inputs_changed.emit())

	def _on_table_item_changed(self, _item):
		if not self._loading_profile:
			self.total_inputs_changed.emit()

	def _on_theme_button_toggled(self, btn: QPushButton, key: str, checked: bool):
		if not checked and not self.selected_themes():
			btn.blockSignals(True)
			btn.setChecked(True)
			btn.blockSignals(False)
		self._refresh_theme_chip_style(btn, key, btn.isChecked())

	def _on_profile_combo_changed(self, _index: int):
		if self._loading_profile:
			return
		pid = self.profile_combo.currentData()
		if pid:
			self.profile_changed.emit(str(pid))

	def set_profiles(self, profiles: list[dict], selected_id: str | None = None):
		self._loading_profile = True
		self.profile_combo.blockSignals(True)
		self.profile_combo.clear()
		selected_index = 0
		for i, profile in enumerate(profiles):
			label = profile.get("name", "بدون نام")
			if profile.get("is_default"):
				label = f"{label} ★"
			self.profile_combo.addItem(label, profile["id"])
			if selected_id and profile["id"] == selected_id:
				selected_index = i
		if profiles:
			self.profile_combo.setCurrentIndex(selected_index)
		self.profile_combo.blockSignals(False)
		self._loading_profile = False

	def current_profile_id(self) -> str | None:
		data = self.profile_combo.currentData()
		return str(data) if data else None

	def apply_profile(self, profile: dict):
		"""Fill seller / rates / notes / themes without touching invoice items."""
		self._loading_profile = True
		self.from_input.setText(profile.get("sby", ""))
		self.phone_input.setText(profile.get("phone", ""))
		self.location_input.setPlainText(profile.get("location", ""))
		self.off_input.setValue(float(profile.get("off", 0) or 0))
		self.tax_input.setValue(float(profile.get("tax", 0) or 0))
		self.desc1_input.setPlainText(profile.get("desc1", ""))
		self.desc2_input.setPlainText(profile.get("desc2", ""))
		self.set_selected_themes(profile.get("preferred_themes") or DEFAULT_THEMES)
		self._loading_profile = False
		self.total_inputs_changed.emit()

	def selected_themes(self) -> list[str]:
		return [key for key, btn in self.theme_buttons.items() if btn.isChecked()]

	def set_selected_themes(self, themes: list[str]):
		wanted = set(themes) if themes else set(DEFAULT_THEMES)
		for key, btn in self.theme_buttons.items():
			btn.blockSignals(True)
			btn.setChecked(key in wanted)
			self._refresh_theme_chip_style(btn, key, btn.isChecked())
			btn.blockSignals(False)
		if not self.selected_themes():
			for key in DEFAULT_THEMES:
				if key in self.theme_buttons:
					btn = self.theme_buttons[key]
					btn.blockSignals(True)
					btn.setChecked(True)
					self._refresh_theme_chip_style(btn, key, True)
					btn.blockSignals(False)

	def set_total_text(self, text: str):
		self.total_label.setText(text)

	def get_table_items(self):
		"""Parse table rows into {name: [count, price]}.

		Returns (items, error_message). On success error_message is None.
		"""
		items = {}
		row_count = self.items_table.rowCount()
		if row_count == 0:
			return None, "لطفا حداقل یک آیتم اضافه کنید"

		for row in range(row_count):
			price_item = self.items_table.item(row, 1)
			count_item = self.items_table.item(row, 2)
			name_item = self.items_table.item(row, 3)
			price_text = (price_item.text() if price_item else "").strip()
			count_text = (count_item.text() if count_item else "").strip()
			name = (name_item.text() if name_item else "").strip()

			if not name:
				return None, f"نام آیتم در ردیف {row + 1} خالی است"

			if name in items:
				return None, f"نام آیتم تکراری است: {name}"

			try:
				count = int(count_text)
				price = int(price_text)
			except ValueError:
				return None, f"تعداد/قیمت در ردیف {row + 1} عدد معتبر نیست"

			if count <= 0:
				return None, f"تعداد در ردیف {row + 1} باید بیشتر از صفر باشد"
			if price < 0:
				return None, f"قیمت در ردیف {row + 1} نمی‌تواند منفی باشد"

			items[name] = [count, price]

		return items, None

	def add_item_row(self):
		if self.items_table.rowCount() >= MAX_ITEMS:
			QMessageBox.information(self, "پیام", f"حداکثر {MAX_ITEMS} آیتم مجاز است")
			return
		row = self.items_table.rowCount()
		self.items_table.blockSignals(True)
		self.items_table.insertRow(row)
		self.items_table.setItem(row, 1, QTableWidgetItem("0"))
		self.items_table.setItem(row, 2, QTableWidgetItem("1"))
		self.items_table.setItem(row, 3, QTableWidgetItem(""))
		self.items_table.blockSignals(False)
		self._set_delete_button(row)
		self.total_inputs_changed.emit()

	def _set_delete_button(self, row: int):
		btn = QPushButton("حذف")
		btn.setObjectName("RowDeleteButton")
		btn.setCursor(Qt.CursorShape.PointingHandCursor)
		btn.clicked.connect(lambda *_: self._delete_row_by_button(btn))
		self.items_table.setCellWidget(row, 0, btn)

	def _delete_row_by_button(self, button: QPushButton):
		for row in range(self.items_table.rowCount()):
			if self.items_table.cellWidget(row, 0) is button:
				self.items_table.removeRow(row)
				self._rebind_delete_buttons()
				self.total_inputs_changed.emit()
				return

	def _rebind_delete_buttons(self):
		for row in range(self.items_table.rowCount()):
			widget = self.items_table.cellWidget(row, 0)
			if isinstance(widget, QPushButton):
				try:
					widget.clicked.disconnect()
				except RuntimeError:
					pass
				widget.clicked.connect(lambda *_ , b=widget: self._delete_row_by_button(b))
			else:
				self._set_delete_button(row)

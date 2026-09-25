from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
	QWidget, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout,
	QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QDoubleSpinBox,
	QScrollArea, QSizePolicy, QMessageBox, QFrame, QComboBox, QDialog,
	QDialogButtonBox, QFormLayout, QCheckBox, QAbstractItemView, QGridLayout,
	QAbstractSpinBox, QProgressBar,
)
import os

from engine import DEFAULT_THEMES, THEME_LABELS, image_map, FONT_PATH
import engine as engine_mod

MAX_ITEMS = 13

# Physical right — ignores RTL mirroring of AlignLeft/AlignRight
ALIGN_RIGHT = Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
ALIGN_RIGHT_TOP = Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop

THEME_OBJECT_NAMES = {
	"blue": "ThemeBlue",
	"cyan": "ThemeCyan",
	"red": "ThemeRed",
	"yellow": "ThemeYellow",
	"green": "ThemeGreen",
	"gray": "ThemeGray",
}


def load_app_stylesheet(app_root: str, dark: bool = False) -> str:
	name = "styles_dark.qss" if dark else "styles.qss"
	path = os.path.join(app_root, name)
	try:
		with open(path, "r", encoding="utf-8") as fh:
			css = fh.read()
	except OSError:
		if dark:
			return load_app_stylesheet(app_root, dark=False)
		return ""

	# Resolve icon paths so QSS works regardless of CWD
	icons_dir = os.path.join(app_root, "icons").replace("\\", "/")
	css = css.replace("icons/", f"{icons_dir}/")
	return css


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


def style_combo(combo: QComboBox) -> QComboBox:
	combo.setEditable(False)
	combo.setMaxVisibleItems(12)
	combo.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
	combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
	return combo


def style_spin(spin: QDoubleSpinBox) -> QDoubleSpinBox:
	spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.UpDownArrows)
	spin.setAlignment(ALIGN_RIGHT)
	spin.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
	spin.setKeyboardTracking(False)
	spin.setDecimals(2)
	spin.setSingleStep(1.0)
	return spin


def style_text_input(widget: QWidget) -> QWidget:
	"""Force text to the physical right inside inputs (Persian RTL)."""
	widget.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
	if isinstance(widget, QLineEdit):
		widget.setAlignment(ALIGN_RIGHT)
	elif isinstance(widget, QTextEdit):
		align = Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignRight

		def _force_right_align(edit: QTextEdit = widget, flag: Qt.AlignmentFlag = align):
			edit.blockSignals(True)
			try:
				edit.setAlignment(flag)
				doc = edit.document()
				opt = doc.defaultTextOption()
				opt.setAlignment(flag)
				opt.setTextDirection(Qt.LayoutDirection.RightToLeft)
				doc.setDefaultTextOption(opt)
			finally:
				edit.blockSignals(False)

		_force_right_align()
		widget.textChanged.connect(_force_right_align)
	return widget


def style_label(label: QLabel, *, title: bool = False) -> QLabel:
	label.setAlignment(ALIGN_RIGHT)
	label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
	if title:
		label.setObjectName("SectionTitle")
	return label


def force_layout_right(layout) -> None:
	"""Pin a layout's child alignment to the physical right."""
	layout.setAlignment(ALIGN_RIGHT)


class FactorListDialog(QDialog):
	"""Pick a saved factor to load or delete."""

	def __init__(self, parent=None, factors: list[dict] | None = None):
		super().__init__(parent)
		self.setWindowTitle("فاکتورهای ذخیره‌شده")
		self.setMinimumSize(560, 360)
		self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		self._selected_id: int | None = None
		self._factors = factors or []

		layout = QVBoxLayout(self)
		layout.setContentsMargins(16, 16, 16, 16)
		layout.setSpacing(12)

		hint = style_label(QLabel("یک فاکتور را انتخاب کنید"))
		hint.setObjectName("HintLabel")
		layout.addWidget(hint, alignment=ALIGN_RIGHT)

		self.table = QTableWidget(0, 4)
		self.table.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
		self.table.setHorizontalHeaderLabels(["آیدی", "تاریخ", "عنوان", "وضعیت"])
		self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # type: ignore
		self.table.horizontalHeader().setDefaultAlignment(ALIGN_RIGHT)
		self.table.verticalHeader().setVisible(False)
		self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
		self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
		self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
		self.table.doubleClicked.connect(self._accept_load)
		layout.addWidget(self.table, stretch=1)

		status_map = {"draft": "پیش‌نویس", "generated": "تولیدشده"}
		for factor in self._factors:
			row = self.table.rowCount()
			self.table.insertRow(row)
			values = [
				str(factor.get("factor_id", "")),
				str(factor.get("factor_date", "")),
				str(factor.get("title", "")),
				status_map.get(str(factor.get("status", "")), str(factor.get("status", ""))),
			]
			for col, text in enumerate(values):
				item = QTableWidgetItem(text)
				item.setTextAlignment(int(ALIGN_RIGHT))
				self.table.setItem(row, col, item)

		buttons = QHBoxLayout()
		self.load_btn = QPushButton("بارگذاری")
		self.load_btn.setObjectName("PrimaryButton")
		self.delete_btn = QPushButton("حذف")
		self.delete_btn.setObjectName("DangerButton")
		self.cancel_btn = QPushButton("انصراف")
		self.cancel_btn.setObjectName("SecondaryButton")
		self.load_btn.clicked.connect(self._accept_load)
		self.delete_btn.clicked.connect(self._accept_delete)
		self.cancel_btn.clicked.connect(self.reject)
		buttons.addWidget(self.cancel_btn)
		buttons.addStretch(1)
		buttons.addWidget(self.delete_btn)
		buttons.addWidget(self.load_btn)
		layout.addLayout(buttons)

		self._action = None  # "load" | "delete"

	def _current_factor_id(self) -> int | None:
		rows = self.table.selectionModel().selectedRows()
		if not rows:
			return None
		item = self.table.item(rows[0].row(), 0)
		if not item:
			return None
		try:
			return int(item.text())
		except ValueError:
			return None

	def _accept_load(self):
		fid = self._current_factor_id()
		if fid is None:
			QMessageBox.information(self, "پیام", "لطفا یک فاکتور را انتخاب کنید")
			return
		self._selected_id = fid
		self._action = "load"
		self.accept()

	def _accept_delete(self):
		fid = self._current_factor_id()
		if fid is None:
			QMessageBox.information(self, "پیام", "لطفا یک فاکتور را انتخاب کنید")
			return
		confirm = QMessageBox.question(
			self,
			"حذف فاکتور",
			f"فاکتور با آیدی {fid} حذف شود؟",
			QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
		)
		if confirm != QMessageBox.StandardButton.Yes:
			return
		self._selected_id = fid
		self._action = "delete"
		self.accept()

	def result_action(self) -> tuple[str | None, int | None]:
		return self._action, self._selected_id


class ProgressDialog(QDialog):
	"""Modal progress window shown while generating invoices."""

	def __init__(self, parent=None):
		super().__init__(parent)
		self.setWindowTitle("در حال تولید فاکتور")
		self.setModal(True)
		self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		self.setMinimumWidth(420)
		self.setWindowFlags(
			Qt.WindowType.Dialog
			| Qt.WindowType.CustomizeWindowHint
			| Qt.WindowType.WindowTitleHint
		)

		layout = QVBoxLayout(self)
		layout.setContentsMargins(20, 20, 20, 20)
		layout.setSpacing(14)

		title = QLabel("لطفا صبر کنید…")
		title.setObjectName("SectionTitle")
		title.setAlignment(Qt.AlignmentFlag.AlignCenter)
		layout.addWidget(title)

		self.status_label = QLabel("شروع تولید فاکتور")
		self.status_label.setObjectName("HintLabel")
		self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		self.status_label.setWordWrap(True)
		layout.addWidget(self.status_label)

		self.bar = QProgressBar()
		self.bar.setRange(0, 100)
		self.bar.setValue(0)
		self.bar.setTextVisible(True)
		self.bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
		self.bar.setMinimumHeight(22)
		layout.addWidget(self.bar)

		self.detail_label = QLabel("")
		self.detail_label.setObjectName("HintLabel")
		self.detail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
		layout.addWidget(self.detail_label)

	def update_progress(self, current: int, total: int, message: str):
		total = max(total, 1)
		pct = int(round(100 * current / total))
		self.bar.setValue(min(max(pct, 0), 100))
		self.status_label.setText(message)
		self.detail_label.setText(f"{current} از {total}")


class GenerateWorker(QThread):
	progress = Signal(int, int, str)
	finished_ok = Signal(bool)

	def __init__(self, engine, factor_id, template_path, datas, themes, parent=None):
		super().__init__(parent)
		self._engine = engine
		self._factor_id = factor_id
		self._template_path = template_path
		self._datas = datas
		self._themes = themes

	def run(self):
		ok = self._engine.makeFactor(
			self._factor_id,
			self._template_path,
			self._datas,
			themes=self._themes,
			progress=lambda c, t, m: self.progress.emit(c, t, m),
		)
		self.finished_ok.emit(bool(ok))


class ProfileDialog(QDialog):
	def __init__(self, parent=None, profile: dict | None = None, is_new: bool = True):
		super().__init__(parent)
		self.setWindowTitle("پروفایل جدید" if is_new else "ویرایش پروفایل")
		self.setMinimumWidth(440)
		self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		self._profile_id = (profile or {}).get("id") if profile and not is_new else None

		layout = QVBoxLayout(self)
		layout.setContentsMargins(16, 16, 16, 16)
		layout.setSpacing(12)
		form = QFormLayout()
		form.setSpacing(10)
		form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
		form.setFormAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)

		self.name_input = style_text_input(QLineEdit((profile or {}).get("name", "")))
		self.sby_input = style_text_input(QLineEdit((profile or {}).get("sby", "")))
		self.phone_input = style_text_input(QLineEdit((profile or {}).get("phone", "")))
		self.location_input = style_text_input(QTextEdit())
		self.location_input.setFixedHeight(70)
		self.location_input.setPlainText((profile or {}).get("location", ""))
		self.off_input = style_spin(QDoubleSpinBox())
		self.tax_input = style_spin(QDoubleSpinBox())
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
	save_requested = Signal()
	open_list_requested = Signal()
	new_factor_requested = Signal()

	def __init__(self):
		super().__init__()
		self.setObjectName("FactorRoot")
		self.setWindowTitle("FactorMaker")
		self.setMinimumSize(1080, 720)
		self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
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

		profile_label = style_label(QLabel("پروفایل"))
		profile_label.setObjectName("FieldLabel")
		self.profile_combo = style_combo(QComboBox())
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

		profile_row.addWidget(profile_label, alignment=ALIGN_RIGHT)
		profile_row.addWidget(self.profile_combo)
		profile_row.addWidget(self.profile_new_btn)
		profile_row.addWidget(self.profile_edit_btn)
		profile_row.addWidget(self.profile_delete_btn)
		top.addWidget(profile_wrap)
		layout.addLayout(top)

		theme_row = QHBoxLayout()
		theme_row.setSpacing(8)
		theme_label = style_label(QLabel("رنگ فاکتور"))
		theme_label.setObjectName("FieldLabel")
		theme_row.addWidget(theme_label, alignment=ALIGN_RIGHT)
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
		hint = style_label(QLabel("حداقل یک رنگ را انتخاب کنید"))
		hint.setObjectName("HintLabel")
		theme_row.addWidget(hint, alignment=ALIGN_RIGHT)
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
		col.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		layout = QVBoxLayout(col)
		layout.setContentsMargins(0, 0, 0, 0)
		layout.setSpacing(14)
		force_layout_right(layout)

		invoice_box = QFrame()
		invoice_box.setObjectName("SectionBox")
		invoice_box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		invoice_layout = QVBoxLayout(invoice_box)
		invoice_layout.setContentsMargins(0, 0, 0, 0)
		invoice_layout.setSpacing(8)
		force_layout_right(invoice_layout)
		invoice_title = style_label(QLabel("اطلاعات فاکتور"), title=True)
		invoice_layout.addWidget(invoice_title, alignment=ALIGN_RIGHT)

		self.id_input = style_text_input(QLineEdit())
		self.date_input = style_text_input(QLineEdit())
		self.addons_input = style_text_input(QLineEdit())
		self.title_input = style_text_input(QLineEdit())
		invoice_layout.addWidget(self._field_row(self.id_input, "آیدی"))
		invoice_layout.addWidget(self._field_row(self.date_input, "تاریخ"))
		invoice_layout.addWidget(self._field_row(self.title_input, "عنوان"))
		invoice_layout.addWidget(self._field_row(self.addons_input, "افزونه‌ها"))
		layout.addWidget(invoice_box)

		seller_box = QFrame()
		seller_box.setObjectName("SectionBox")
		seller_box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		seller_layout = QVBoxLayout(seller_box)
		seller_layout.setContentsMargins(0, 0, 0, 0)
		seller_layout.setSpacing(8)
		force_layout_right(seller_layout)
		seller_title = style_label(QLabel("اطلاعات فروشنده"), title=True)
		seller_layout.addWidget(seller_title, alignment=ALIGN_RIGHT)

		self.from_input = style_text_input(QLineEdit())
		self.phone_input = style_text_input(QLineEdit())
		self.location_input = style_text_input(QTextEdit())
		self.location_input.setMinimumHeight(100)
		self.location_input.setMaximumHeight(140)
		seller_layout.addWidget(self._field_row(self.from_input, "از"))
		seller_layout.addWidget(self._field_row(self.phone_input, "تلفن"))
		seller_layout.addWidget(self._field_row(self.location_input, "آدرس"))
		layout.addWidget(seller_box)
		layout.addStretch(1)
		return col

	def _build_items_column(self) -> QWidget:
		col = QWidget()
		col.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		layout = QVBoxLayout(col)
		layout.setContentsMargins(0, 0, 0, 0)
		layout.setSpacing(10)
		force_layout_right(layout)

		items_title = style_label(QLabel("آیتم‌ها"), title=True)
		layout.addWidget(items_title, alignment=ALIGN_RIGHT)

		# LTR columns left→right: عملیات | رایگان | قیمت | تعداد | نام
		self.items_table = QTableWidget(0, 5)
		self.items_table.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
		self.items_table.setHorizontalHeaderLabels(
			["عملیات", "رایگان", "قیمت (ریال)", "تعداد", "نام"]
		)
		self.items_table.setAlternatingRowColors(True)
		header = self.items_table.horizontalHeader()
		header.setSectionResizeMode(QHeaderView.Stretch)  # type: ignore
		header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # type: ignore
		header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # type: ignore
		header.setDefaultAlignment(ALIGN_RIGHT)
		header.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
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
		self.add_item_btn.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		self.add_item_btn.clicked.connect(self.add_item_row)
		layout.addWidget(self.add_item_btn)

		rates = QHBoxLayout()
		rates.setSpacing(12)
		force_layout_right(rates)
		self.off_input = style_spin(QDoubleSpinBox())
		self.tax_input = style_spin(QDoubleSpinBox())
		self.off_input.setRange(0, 100)
		self.tax_input.setRange(0, 100)
		self.off_input.setSuffix(" %")
		self.tax_input.setSuffix(" %")
		rates.addWidget(self._field_row(self.off_input, "تخفیف"), stretch=1)
		rates.addWidget(self._field_row(self.tax_input, "مالیات"), stretch=1)
		layout.addLayout(rates)

		notes_title = style_label(QLabel("توضیحات"), title=True)
		layout.addWidget(notes_title, alignment=ALIGN_RIGHT)

		self.desc1_input = style_text_input(QTextEdit())
		self.desc2_input = style_text_input(QTextEdit())
		self.desc1_input.setFixedHeight(58)
		self.desc2_input.setFixedHeight(58)
		layout.addWidget(self._field_row(self.desc1_input, "توضیحات ۱"))
		layout.addWidget(self._field_row(self.desc2_input, "توضیحات ۲"))
		return col

	def _build_footer(self) -> QFrame:
		footer = QFrame()
		footer.setObjectName("FooterBar")
		footer.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		layout = QHBoxLayout(footer)
		layout.setContentsMargins(16, 12, 16, 12)
		layout.setSpacing(12)

		self.total_label = style_label(QLabel("جمع کل: ۰ ریال"))
		self.total_label.setObjectName("TotalLabel")
		layout.addWidget(self.total_label, alignment=ALIGN_RIGHT)
		layout.addStretch(1)

		self.new_factor_btn = QPushButton("فاکتور جدید")
		self.new_factor_btn.setObjectName("SecondaryButton")
		self.new_factor_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.new_factor_btn.clicked.connect(self.new_factor_requested.emit)

		self.open_list_btn = QPushButton("باز کردن")
		self.open_list_btn.setObjectName("SecondaryButton")
		self.open_list_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.open_list_btn.clicked.connect(self.open_list_requested.emit)

		self.save_btn = QPushButton("ذخیره")
		self.save_btn.setObjectName("SecondaryButton")
		self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.save_btn.clicked.connect(self.save_requested.emit)

		self.generate_btn = QPushButton("تولید فاکتور")
		self.generate_btn.setObjectName("PrimaryButton")
		self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
		self.generate_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

		layout.addWidget(self.new_factor_btn)
		layout.addWidget(self.open_list_btn)
		layout.addWidget(self.save_btn)
		layout.addWidget(self.generate_btn)
		return footer

	def _field_row(self, widget: QWidget, label_text: str) -> QWidget:
		"""Label on the right, control fills the rest — text inside control is right-aligned."""
		wrap = QWidget()
		wrap.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
		row = QHBoxLayout(wrap)
		row.setContentsMargins(0, 0, 0, 0)
		row.setSpacing(10)
		force_layout_right(row)

		label = style_label(QLabel(label_text))
		label.setObjectName("FieldLabel")
		label.setMinimumWidth(78)

		row.addWidget(label, stretch=0, alignment=ALIGN_RIGHT)
		row.addWidget(widget, stretch=1)
		return wrap

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
		"""Parse table rows into {name: [count, price, is_free]}.

		Columns: 0=عملیات, 1=رایگان, 2=قیمت, 3=تعداد, 4=نام
		Returns (items, error_message). On success error_message is None.
		"""
		items = {}
		row_count = self.items_table.rowCount()
		if row_count == 0:
			return None, "لطفا حداقل یک آیتم اضافه کنید"

		for row in range(row_count):
			price_item = self.items_table.item(row, 2)
			count_item = self.items_table.item(row, 3)
			name_item = self.items_table.item(row, 4)
			price_text = (price_item.text() if price_item else "").strip()
			count_text = (count_item.text() if count_item else "").strip()
			name = (name_item.text() if name_item else "").strip()
			is_free = self._row_is_free(row)

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
			if is_free and price <= 0:
				return None, f"برای ردیف رایگان {row + 1} قیمت واقعی را وارد کنید"

			items[name] = [count, price, is_free]

		return items, None

	def clear_items(self):
		self.items_table.setRowCount(0)
		self.total_inputs_changed.emit()

	def set_table_items(self, items: dict | list):
		"""Load engine dict or stored list into the items table."""
		self.items_table.blockSignals(True)
		self.items_table.setRowCount(0)
		rows = []
		if isinstance(items, dict):
			for name, entry in items.items():
				rows.append(
					{
						"name": name,
						"count": int(entry[0]),
						"price": int(entry[1]),
						"free": bool(entry[2]) if len(entry) > 2 else False,
					}
				)
		else:
			rows = list(items or [])

		for row_data in rows:
			if self.items_table.rowCount() >= MAX_ITEMS:
				break
			row = self.items_table.rowCount()
			self.items_table.insertRow(row)
			price_item = QTableWidgetItem(str(int(row_data.get("price") or 0)))
			count_item = QTableWidgetItem(str(int(row_data.get("count") or 1)))
			name_item = QTableWidgetItem(str(row_data.get("name") or ""))
			for item in (price_item, count_item, name_item):
				item.setTextAlignment(int(ALIGN_RIGHT))
			self.items_table.setItem(row, 2, price_item)
			self.items_table.setItem(row, 3, count_item)
			self.items_table.setItem(row, 4, name_item)
			self._set_row_controls(row, free=bool(row_data.get("free")))
		self.items_table.blockSignals(False)
		self.total_inputs_changed.emit()

	def collect_form_snapshot(self) -> dict:
		"""Current form values for save/load (items may be empty/invalid)."""
		items, _error = self.get_table_items()
		item_rows = []
		if items:
			for name, entry in items.items():
				item_rows.append(
					{
						"name": name,
						"count": int(entry[0]),
						"price": int(entry[1]),
						"free": bool(entry[2]) if len(entry) > 2 else False,
					}
				)
		else:
			# Soft collect even incomplete rows for drafts
			for row in range(self.items_table.rowCount()):
				name = (self.items_table.item(row, 4).text() if self.items_table.item(row, 4) else "").strip()
				if not name:
					continue
				try:
					price = int((self.items_table.item(row, 2).text() if self.items_table.item(row, 2) else "0").strip() or "0")
					count = int((self.items_table.item(row, 3).text() if self.items_table.item(row, 3) else "1").strip() or "1")
				except ValueError:
					continue
				item_rows.append(
					{
						"name": name,
						"count": count,
						"price": price,
						"free": self._row_is_free(row),
					}
				)

		fid_text = self.id_input.text().strip()
		try:
			factor_id = int(fid_text) if fid_text else None
		except ValueError:
			factor_id = None

		return {
			"factor_id": factor_id,
			"factor_date": self.date_input.text().strip(),
			"title": self.title_input.text().strip(),
			"addons": self.addons_input.text().strip(),
			"sby": self.from_input.text().strip(),
			"phone": self.phone_input.text().strip(),
			"location": self.location_input.toPlainText().strip(),
			"off": float(self.off_input.value()),
			"tax": float(self.tax_input.value()),
			"desc1": self.desc1_input.toPlainText(),
			"desc2": self.desc2_input.toPlainText(),
			"items": item_rows,
			"themes": self.selected_themes(),
		}

	def apply_factor_snapshot(self, data: dict):
		"""Fill the whole form from a saved factor."""
		self._loading_profile = True
		if data.get("factor_id") is not None:
			self.id_input.setText(str(data["factor_id"]))
		self.date_input.setText(str(data.get("factor_date") or ""))
		self.title_input.setText(str(data.get("title") or ""))
		self.addons_input.setText(str(data.get("addons") or ""))
		self.from_input.setText(str(data.get("sby") or ""))
		self.phone_input.setText(str(data.get("phone") or ""))
		self.location_input.setPlainText(str(data.get("location") or ""))
		self.off_input.setValue(float(data.get("off") or 0))
		self.tax_input.setValue(float(data.get("tax") or 0))
		self.desc1_input.setPlainText(str(data.get("desc1") or ""))
		self.desc2_input.setPlainText(str(data.get("desc2") or ""))
		self.set_selected_themes(data.get("themes") or DEFAULT_THEMES)
		self._loading_profile = False
		self.set_table_items(data.get("items") or [])

	def add_item_row(self, *, free: bool = False):
		if self.items_table.rowCount() >= MAX_ITEMS:
			QMessageBox.information(self, "پیام", f"حداکثر {MAX_ITEMS} آیتم مجاز است")
			return
		row = self.items_table.rowCount()
		self.items_table.blockSignals(True)
		self.items_table.insertRow(row)
		price_item = QTableWidgetItem("0")
		count_item = QTableWidgetItem("1")
		name_item = QTableWidgetItem("")
		for item in (price_item, count_item, name_item):
			item.setTextAlignment(int(ALIGN_RIGHT))
		self.items_table.setItem(row, 2, price_item)
		self.items_table.setItem(row, 3, count_item)
		self.items_table.setItem(row, 4, name_item)
		self.items_table.blockSignals(False)
		self._set_row_controls(row, free=free)
		self.total_inputs_changed.emit()

	def _row_is_free(self, row: int) -> bool:
		widget = self.items_table.cellWidget(row, 1)
		if isinstance(widget, QWidget):
			cb = widget.findChild(QCheckBox)
			if cb is not None:
				return cb.isChecked()
		return False

	def _set_row_controls(self, row: int, *, free: bool = False):
		# Delete button
		btn = QPushButton("حذف")
		btn.setObjectName("RowDeleteButton")
		btn.setCursor(Qt.CursorShape.PointingHandCursor)
		btn.clicked.connect(lambda *_: self._delete_row_by_button(btn))
		self.items_table.setCellWidget(row, 0, btn)

		# Free checkbox centered in cell
		wrap = QWidget()
		wrap_layout = QHBoxLayout(wrap)
		wrap_layout.setContentsMargins(0, 0, 0, 0)
		wrap_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
		cb = QCheckBox()
		cb.setChecked(free)
		cb.setToolTip("این ردیف رایگان است — قیمت نمایش داده می‌شود ولی در جمع حساب نمی‌شود")
		cb.toggled.connect(lambda *_: self.total_inputs_changed.emit())
		wrap_layout.addWidget(cb)
		self.items_table.setCellWidget(row, 1, wrap)

	def _set_delete_button(self, row: int):
		self._set_row_controls(row, free=self._row_is_free(row))

	def _delete_row_by_button(self, button: QPushButton):
		for row in range(self.items_table.rowCount()):
			widget = self.items_table.cellWidget(row, 0)
			if widget is button:
				self.items_table.removeRow(row)
				self._rebind_row_controls()
				self.total_inputs_changed.emit()
				return

	def _rebind_row_controls(self):
		for row in range(self.items_table.rowCount()):
			free = self._row_is_free(row)
			self._set_row_controls(row, free=free)

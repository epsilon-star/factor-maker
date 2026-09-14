from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
	QWidget, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout,
	QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QDoubleSpinBox,
	QScrollArea, QSizePolicy, QMessageBox, QFrame, QComboBox, QDialog,
	QDialogButtonBox, QFormLayout, QCheckBox, QAbstractItemView,
)

from engine import DEFAULT_THEMES, THEME_LABELS, image_map

MAX_ITEMS = 13


def load_app_stylesheet(app_root: str) -> str:
	import os
	path = os.path.join(app_root, "styles.qss")
	try:
		with open(path, "r", encoding="utf-8") as fh:
			return fh.read()
	except OSError:
		return ""


def try_set_vazir_font(app) -> None:
	"""Use Vazir when installed; otherwise leave the Qt default."""
	families = set(QFontDatabase.families())
	for family in ("Vazirmatn", "Vazir", "Tahoma"):
		if family in families or family == "Tahoma":
			app.setFont(QFont(family))
			return


class ProfileDialog(QDialog):
	def __init__(self, parent=None, profile: dict | None = None, is_new: bool = True):
		super().__init__(parent)
		self.setWindowTitle("پروفایل جدید" if is_new else "ویرایش پروفایل")
		self.setMinimumWidth(420)
		self._profile_id = (profile or {}).get("id") if profile and not is_new else None

		layout = QVBoxLayout(self)
		form = QFormLayout()

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
		self.setMinimumSize(980, 680)
		self.theme_buttons: dict[str, QPushButton] = {}
		self._loading_profile = False
		self.init_ui()

	def init_ui(self):
		root = QVBoxLayout(self)
		root.setContentsMargins(14, 14, 14, 14)
		root.setSpacing(12)

		root.addWidget(self._build_header())

		body_scroll = QScrollArea()
		body_scroll.setWidgetResizable(True)
		body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
		body = QWidget()
		body_layout = QHBoxLayout(body)
		body_layout.setContentsMargins(0, 0, 0, 0)
		body_layout.setSpacing(12)
		body_layout.addWidget(self._build_left_column(), stretch=2)
		body_layout.addWidget(self._build_right_column(), stretch=3)
		body_scroll.setWidget(body)
		root.addWidget(body_scroll, stretch=1)

		root.addWidget(self._build_footer())

		self._wire_total_signals()

	def _build_header(self) -> QFrame:
		header = QFrame()
		header.setObjectName("HeaderBar")
		layout = QVBoxLayout(header)
		layout.setSpacing(10)

		top = QHBoxLayout()
		self.brand_title = QLabel("فاکتور ساز متانویا")
		self.brand_title.setObjectName("BrandTitle")
		top.addWidget(self.brand_title)
		top.addStretch()

		profile_label = QLabel("پروفایل")
		profile_label.setObjectName("FieldLabel")
		self.profile_combo = QComboBox()
		self.profile_combo.setMinimumWidth(180)
		self.profile_combo.currentIndexChanged.connect(self._on_profile_combo_changed)

		self.profile_new_btn = QPushButton("جدید")
		self.profile_edit_btn = QPushButton("ویرایش")
		self.profile_delete_btn = QPushButton("حذف")
		self.profile_delete_btn.setObjectName("DangerButton")
		self.profile_new_btn.clicked.connect(self.profile_new_requested.emit)
		self.profile_edit_btn.clicked.connect(self.profile_edit_requested.emit)
		self.profile_delete_btn.clicked.connect(self.profile_delete_requested.emit)

		top.addWidget(profile_label)
		top.addWidget(self.profile_combo)
		top.addWidget(self.profile_new_btn)
		top.addWidget(self.profile_edit_btn)
		top.addWidget(self.profile_delete_btn)
		layout.addLayout(top)

		theme_row = QHBoxLayout()
		theme_label = QLabel("رنگ فاکتور")
		theme_label.setObjectName("FieldLabel")
		theme_row.addWidget(theme_label)
		for key in image_map:
			btn = QPushButton(THEME_LABELS.get(key, key))
			btn.setObjectName("ThemeChip")
			btn.setCheckable(True)
			btn.setProperty("theme", key)
			btn.setChecked(key in DEFAULT_THEMES)
			btn.toggled.connect(self._on_theme_toggled)
			self.theme_buttons[key] = btn
			theme_row.addWidget(btn)
		theme_row.addStretch()
		layout.addLayout(theme_row)
		return header

	def _build_left_column(self) -> QFrame:
		panel = QFrame()
		panel.setObjectName("Panel")
		layout = QVBoxLayout(panel)
		layout.setSpacing(10)

		invoice_title = QLabel("اطلاعات فاکتور")
		invoice_title.setObjectName("SectionTitle")
		layout.addWidget(invoice_title)

		self.id_input = QLineEdit()
		self.date_input = QLineEdit()
		self.addons_input = QLineEdit()
		self.title_input = QLineEdit()
		layout.addLayout(self._labeled_row(self.id_input, "آیدی"))
		layout.addLayout(self._labeled_row(self.date_input, "تاریخ"))
		layout.addLayout(self._labeled_row(self.title_input, "عنوان"))
		layout.addLayout(self._labeled_row(self.addons_input, "افزونه‌ها"))

		seller_title = QLabel("اطلاعات فروشنده")
		seller_title.setObjectName("SectionTitle")
		layout.addWidget(seller_title)

		self.from_input = QLineEdit()
		self.phone_input = QLineEdit()
		self.location_input = QTextEdit()
		self.location_input.setFixedHeight(90)
		layout.addLayout(self._labeled_row(self.from_input, "از"))
		layout.addLayout(self._labeled_row(self.phone_input, "تلفن"))
		layout.addLayout(self._labeled_row(self.location_input, "آدرس"))

		layout.addStretch()
		return panel

	def _build_right_column(self) -> QFrame:
		panel = QFrame()
		panel.setObjectName("Panel")
		layout = QVBoxLayout(panel)
		layout.setSpacing(10)

		items_title = QLabel("آیتم‌ها")
		items_title.setObjectName("SectionTitle")
		layout.addWidget(items_title)

		self.items_table = QTableWidget(0, 4)
		self.items_table.setHorizontalHeaderLabels(["عملیات", "قیمت (ریال)", "تعداد", "نام"])
		self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # type: ignore
		self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # type: ignore
		self.items_table.verticalHeader().setVisible(False)
		self.items_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
		self.items_table.setMinimumHeight(260)
		self.items_table.itemChanged.connect(self._on_table_item_changed)
		layout.addWidget(self.items_table)

		self.add_item_btn = QPushButton("افزودن آیتم")
		self.add_item_btn.clicked.connect(self.add_item_row)
		layout.addWidget(self.add_item_btn)

		rates = QHBoxLayout()
		self.off_input = QDoubleSpinBox()
		self.tax_input = QDoubleSpinBox()
		self.off_input.setRange(0, 100)
		self.tax_input.setRange(0, 100)
		self.off_input.setSuffix(" %")
		self.tax_input.setSuffix(" %")
		rates.addLayout(self._labeled_row(self.off_input, "تخفیف"))
		rates.addLayout(self._labeled_row(self.tax_input, "مالیات"))
		layout.addLayout(rates)

		notes_title = QLabel("توضیحات")
		notes_title.setObjectName("SectionTitle")
		layout.addWidget(notes_title)

		self.desc1_input = QTextEdit()
		self.desc2_input = QTextEdit()
		self.desc1_input.setFixedHeight(56)
		self.desc2_input.setFixedHeight(56)
		layout.addLayout(self._labeled_row(self.desc1_input, "توضیحات ۱"))
		layout.addLayout(self._labeled_row(self.desc2_input, "توضیحات ۲"))
		return panel

	def _build_footer(self) -> QFrame:
		footer = QFrame()
		footer.setObjectName("FooterBar")
		layout = QHBoxLayout(footer)

		self.total_label = QLabel("جمع کل: ۰ ریال")
		self.total_label.setObjectName("TotalLabel")
		layout.addWidget(self.total_label)
		layout.addStretch()

		self.generate_btn = QPushButton("تولید فاکتور")
		self.generate_btn.setObjectName("PrimaryButton")
		self.generate_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
		layout.addWidget(self.generate_btn)
		return footer

	def _labeled_row(self, widget: QWidget, label_text: str) -> QHBoxLayout:
		row = QHBoxLayout()
		row.setSpacing(8)
		label = QLabel(label_text)
		label.setObjectName("FieldLabel")
		label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
		row.addWidget(widget, stretch=1)
		row.addWidget(label)
		return row

	def _wire_total_signals(self):
		self.off_input.valueChanged.connect(lambda *_: self.total_inputs_changed.emit())
		self.tax_input.valueChanged.connect(lambda *_: self.total_inputs_changed.emit())

	def _on_table_item_changed(self, _item):
		if not self._loading_profile:
			self.total_inputs_changed.emit()

	def _on_theme_toggled(self, checked: bool):
		sender = self.sender()
		if checked or self.selected_themes():
			return
		# Keep at least one theme selected
		if isinstance(sender, QPushButton):
			sender.blockSignals(True)
			sender.setChecked(True)
			sender.blockSignals(False)

	def _on_profile_combo_changed(self, _index: int):
		if self._loading_profile:
			return
		pid = self.profile_combo.currentData()
		if pid:
			self.profile_changed.emit(str(pid))

	# --- Public helpers used by MainWindow ---

	def set_profiles(self, profiles: list[dict], selected_id: str | None = None):
		self._loading_profile = True
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
			btn.blockSignals(False)
		if not self.selected_themes():
			for key in DEFAULT_THEMES:
				if key in self.theme_buttons:
					self.theme_buttons[key].blockSignals(True)
					self.theme_buttons[key].setChecked(True)
					self.theme_buttons[key].blockSignals(False)

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
		btn.clicked.connect(lambda *_ , r=row: self._delete_row_by_button(btn))
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

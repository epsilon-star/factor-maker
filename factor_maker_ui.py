from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
	QWidget, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout,
	QFormLayout, QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QDoubleSpinBox, QScrollArea,
	QSizePolicy, QMessageBox
)

MAX_ITEMS = 13


class FactorMakerUI(QWidget):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("FactorMaker")
		self.setMinimumSize(800, 600)
		self.init_ui()

	def init_ui(self):
		layout = QVBoxLayout()
		scroll = QScrollArea()
		container = QWidget()
		form_layout = QFormLayout()

		def add_row_with_label_after(widget, label_text):
			row_widget = QWidget()
			row_layout = QHBoxLayout()
			row_layout.setContentsMargins(0, 0, 0, 0)
			row_layout.addWidget(widget)
			label = QLabel(label_text)
			label.setAlignment(Qt.AlignmentFlag.AlignCenter)
			row_layout.addWidget(label)
			row_layout.setAlignment(label, Qt.AlignmentFlag.AlignRight)
			row_widget.setLayout(row_layout)
			form_layout.addRow(row_widget)

		self.app_title = QLabel("فاکتور ساز متانویا")

		self.id_input = QLineEdit()
		self.date_input = QLineEdit()
		self.addons_input = QLineEdit()
		self.title_input = QLineEdit()
		self.from_input = QLineEdit()
		self.phone_input = QLineEdit()

		add_row_with_label_after(self.app_title, "")
		add_row_with_label_after(self.id_input, "ایدی")
		add_row_with_label_after(self.date_input, "تاریخ")
		add_row_with_label_after(self.addons_input, "افزونه ها")
		add_row_with_label_after(self.title_input, "عنوان")
		add_row_with_label_after(self.from_input, "از")
		add_row_with_label_after(self.phone_input, "تلفن")

		self.location_input = QTextEdit()
		self.location_input.setFixedHeight(60)
		add_row_with_label_after(self.location_input, "آدرس")

		self.items_table = QTableWidget(0, 4)
		self.items_table.setHorizontalHeaderLabels(["عملیات", "قیمت (ریال)", "تعداد", "نام"])
		self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # type: ignore
		self.items_table.verticalHeader().setVisible(False)
		self.items_table.doubleClicked.connect(self.tableDConnect)
		add_row_with_label_after(self.items_table, "آیتم ها")

		self.add_item_btn = QPushButton("افزودن آیتم")
		self.add_item_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		add_item_btn_widget = QWidget()
		add_item_btn_layout = QHBoxLayout()
		add_item_btn_layout.setContentsMargins(0, 0, 0, 0)
		add_item_btn_layout.addWidget(self.add_item_btn)
		add_item_btn_widget.setLayout(add_item_btn_layout)
		form_layout.addRow(add_item_btn_widget)
		self.add_item_btn.clicked.connect(self.add_item_row)

		self.off_input = QDoubleSpinBox()
		self.tax_input = QDoubleSpinBox()
		self.off_input.setRange(0, 100)
		self.tax_input.setRange(0, 100)
		self.off_input.setSuffix(" %")
		self.tax_input.setSuffix(" %")
		add_row_with_label_after(self.off_input, "تخفیف")
		add_row_with_label_after(self.tax_input, "مالیات")

		self.desc1_input = QTextEdit()
		self.desc2_input = QTextEdit()
		self.desc1_input.setFixedHeight(50)
		self.desc2_input.setFixedHeight(50)
		add_row_with_label_after(self.desc1_input, "توضیحات 1")
		add_row_with_label_after(self.desc2_input, "توضیحات 2")

		self.generate_btn = QPushButton("تولید فاکتور")
		self.generate_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		generate_btn_widget = QWidget()
		generate_btn_layout = QHBoxLayout()
		generate_btn_layout.setContentsMargins(0, 0, 0, 0)
		generate_btn_layout.addWidget(self.generate_btn)
		generate_btn_widget.setLayout(generate_btn_layout)
		form_layout.addRow(generate_btn_widget)

		container.setLayout(form_layout)
		scroll.setWidget(container)
		scroll.setWidgetResizable(True)
		layout.addWidget(scroll)
		self.setLayout(layout)

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

	def tableDConnect(self, event):
		if event.column() == 0:
			self.items_table.removeRow(event.row())

	def add_item_row(self):
		if self.items_table.rowCount() >= MAX_ITEMS:
			QMessageBox.information(self, "پیام", f"حداکثر {MAX_ITEMS} آیتم مجاز است")
			return
		row = self.items_table.rowCount()
		self.items_table.insertRow(row)
		self.items_table.setItem(row, 0, QTableWidgetItem("حذف"))
		self.items_table.setItem(row, 1, QTableWidgetItem("0"))
		self.items_table.setItem(row, 2, QTableWidgetItem("1"))
		self.items_table.setItem(row, 3, QTableWidgetItem(""))

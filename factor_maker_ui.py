import PySide6
from PySide6.QtCore import * #type: ignore
from PySide6.QtWidgets import (
	QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout,
	QFormLayout, QTableWidget, QTableWidgetItem, QPushButton, QHeaderView, QDoubleSpinBox, QScrollArea,
	QSizePolicy
)
import sys
	
class FactorMakerUI(QWidget):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("FactorMaker")
		self.setMinimumSize(800, 600)
		self.init_ui()

	def init_ui(self):
		# layout = QVBoxLayout()

		# # Scroll Area in case content grows
		# scroll = QScrollArea()
		# container = QWidget()
		# form_layout = QFormLayout()

		# # Row 1 - Basic Text Inputs
		# self.id_input = QLineEdit()
		# self.id_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.date_input = QLineEdit()
		# self.date_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.addons_input = QLineEdit()
		# self.addons_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.title_input = QLineEdit()
		# self.title_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.from_input = QLineEdit()
		# self.from_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.phone_input = QLineEdit()
		# self.phone_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# # Helper to add row with label after widget and right-aligned label
		# def add_row_with_label_after(widget, label_text):
		#     row_widget = QWidget()
		#     row_layout = QHBoxLayout()
		#     row_layout.setContentsMargins(0, 0, 0, 0)
		#     row_layout.addWidget(widget)
		#     label = QLabel(label_text)
		#     label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
		#     row_layout.addWidget(label)
		#     row_layout.setAlignment(label, Qt.AlignmentFlag.AlignRight)
		#     row_widget.setLayout(row_layout)
		#     form_layout.addRow(row_widget)

		# add_row_with_label_after(self.id_input, "ایدی")
		# add_row_with_label_after(self.date_input, "تاریخ")
		# add_row_with_label_after(self.addons_input, "افزونه ها")
		# add_row_with_label_after(self.title_input, "عنوان")
		# add_row_with_label_after(self.from_input, "از")
		# add_row_with_label_after(self.phone_input, "تلفن")

		# # Row 2 - Full-width Location
		# self.location_input = QTextEdit()
		# self.location_input.setFixedHeight(60)
		# form_layout.addRow("آدرس", self.location_input)

		# # Row 3 - Table of Items
		# self.items_table = QTableWidget(0, 3)
		# self.items_table.setHorizontalHeaderLabels(["قیمت", "تعداد", "نام"])
		# self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # type: ignore
		# form_layout.addRow(QLabel("آیتم ها"), self.items_table)

		# add_item_btn = QPushButton("افزودن آیتم")
		# add_item_btn.clicked.connect(self.add_item_row)
		# form_layout.addRow("", add_item_btn)

		# # Row 4 - Off & Tax
		# self.off_input = QDoubleSpinBox()
		# self.tax_input = QDoubleSpinBox()
		# self.off_input.setMaximum(1000000)
		# self.tax_input.setMaximum(1000000)
		# form_layout.addRow("تخفیف", self.off_input)
		# form_layout.addRow("مالیات", self.tax_input)

		# # Row 5 - Descriptions
		# self.desc1_input = QLineEdit()
		# self.desc2_input = QTextEdit()
		# self.desc2_input.setFixedHeight(40)
		# form_layout.addRow("توضیحات 1", self.desc1_input)
		# form_layout.addRow("توضیحات 2", self.desc2_input)

		# # Row 6 - Generate Button
		# self.generate_btn = QPushButton("تولید فاکتور")
		# form_layout.addRow("", self.generate_btn)

		# # Final assembly
		# container.setLayout(form_layout)
		# scroll.setWidget(container)
		# scroll.setWidgetResizable(True)
		# layout.addWidget(scroll)
		# self.setLayout(layout)


		layout = QVBoxLayout()
		# Scroll Area in case content grows
		scroll = QScrollArea()
		container = QWidget()
		form_layout = QFormLayout()

		# Helper to add row with label after widget and right-aligned label
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

		# Row 0 - Title Of app
		self.app_title = QLabel("فاکتور ساز متانویا")

		# Row 1 - Basic Text Inputs
		self.id_input = QLineEdit()
		self.date_input = QLineEdit()
		self.addons_input = QLineEdit()
		self.title_input = QLineEdit()
		self.from_input = QLineEdit()
		self.phone_input = QLineEdit()
		# self.id_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.date_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.addons_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.title_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.from_input.setAlignment(Qt.AlignmentFlag.AlignRight)
		# self.phone_input.setAlignment(Qt.AlignmentFlag.AlignRight)

		add_row_with_label_after(self.app_title, "")
		add_row_with_label_after(self.id_input, "ایدی")
		add_row_with_label_after(self.date_input, "تاریخ")
		add_row_with_label_after(self.addons_input, "افزونه ها")
		add_row_with_label_after(self.title_input, "عنوان")
		add_row_with_label_after(self.from_input, "از")
		add_row_with_label_after(self.phone_input, "تلفن")

		# Row 2 - Full-width Location
		self.location_input = QTextEdit()
		self.location_input.setFixedHeight(60)
		add_row_with_label_after(self.location_input, "آدرس")

		# Row 3 - Table of Items
		self.items_table = QTableWidget(0, 4)
		self.items_table.setHorizontalHeaderLabels(["عملیات","قیمت (ریال)", "تعداد", "نام"])
		self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # type: ignore

		# Remove the vertical header (row numbers) since moving it to the right is not supported
		self.items_table.verticalHeader().setVisible(False)
		self.items_table.doubleClicked.connect(self.tableDConnect)

		add_row_with_label_after(self.items_table, "آیتم ها")

		add_item_btn = QPushButton("افزودن آیتم")
		rem_item_btn = QPushButton("حذف آیتم")
		# add_row_with_label_after(add_item_btn, "")

		self.add_item_btn = QPushButton("افزودن آیتم")
		# self.rem_item_btn = QPushButton("حذف آیتم")
		self.add_item_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		# self.rem_item_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		add_item_btn_widget = QWidget()
		add_item_btn_layout = QHBoxLayout()
		add_item_btn_layout.setContentsMargins(0, 0, 0, 0)
		add_item_btn_layout.addWidget(self.add_item_btn)
		# add_item_btn_layout.addWidget(self.rem_item_btn)
		add_item_btn_widget.setLayout(add_item_btn_layout)
		form_layout.addRow(add_item_btn_widget)
		self.add_item_btn.clicked.connect(self.add_item_row)
		# self.rem_item_btn.clicked.connect(self.rem_item_row)

		# Row 4 - Off & Tax
		self.off_input = QDoubleSpinBox()
		self.tax_input = QDoubleSpinBox()
		self.off_input.setMaximum(1000000)
		self.tax_input.setMaximum(1000000)
		add_row_with_label_after(self.off_input, "تخفیف")
		add_row_with_label_after(self.tax_input, "مالیات")

		# Row 5 - Descriptions
		self.desc1_input = QTextEdit()
		self.desc2_input = QTextEdit()
		self.desc1_input.setFixedHeight(50)
		self.desc2_input.setFixedHeight(50)
		add_row_with_label_after(self.desc1_input, "توضیحات 1")
		add_row_with_label_after(self.desc2_input, "توضیحات 2")

		# Row 6 - Generate Button
		self.generate_btn = QPushButton("تولید فاکتور")
		self.generate_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		generate_btn_widget = QWidget()
		generate_btn_layout = QHBoxLayout()
		generate_btn_layout.setContentsMargins(0, 0, 0, 0)
		generate_btn_layout.addWidget(self.generate_btn)
		generate_btn_widget.setLayout(generate_btn_layout)
		form_layout.addRow(generate_btn_widget)

		# Final assembly
		container.setLayout(form_layout)
		scroll.setWidget(container)
		scroll.setWidgetResizable(True)
		layout.addWidget(scroll)
		self.setLayout(layout)

	def get_table_items(self):
		items = {}
		for row in range(self.items_table.rowCount()):
			price_item = self.items_table.item(row, 1)
			count_item = self.items_table.item(row, 2)
			name_item = self.items_table.item(row, 3)
			price = price_item.text() if price_item else ""
			count = count_item.text() if count_item else ""
			name = name_item.text() if name_item else ""
			items[str(name)] = [int(count), int(price)]
		return items

	def tableDConnect(self,event):
		if event.column() == 0:
			self.items_table.removeRow(event.row())
			# print("Remove")

	def add_item_row(self):
		if self.items_table.rowCount() < 13:
			row = self.items_table.rowCount()
			self.items_table.insertRow(row)
			self.items_table.setItem(row, 0, QTableWidgetItem("حذف"))  # Remove
			self.items_table.setItem(row, 1, QTableWidgetItem("0"))  # Price
			self.items_table.setItem(row, 2, QTableWidgetItem("1"))  # Quantity
			self.items_table.setItem(row, 3, QTableWidgetItem(""))  # Name

	def rem_item_row(self):
		if self.items_table.rowCount() > 0: self.items_table.removeRow(self.items_table.rowCount()-1)

# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = FactorMakerUI()
#     window.show()
#     sys.exit(app.exec())
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox
import sys
from factor_maker_ui import FactorMakerUI
from engine import PersianFactor, APP_ROOT, image_map
import jdatetime
import os

TEMPLATE_PATH = os.path.join(APP_ROOT, "templates", "template.docx")
OUTPUT_DIR = os.path.join(APP_ROOT, "output")
IMAGES_DIR = os.path.join(APP_ROOT, "images")


def required_asset_errors():
    """Return a list of missing required asset paths."""
    missing = []
    if not os.path.isfile(TEMPLATE_PATH):
        missing.append(TEMPLATE_PATH)
    for theme_file in image_map.values():
        image_path = os.path.join(IMAGES_DIR, f"{theme_file}.jpg")
        if not os.path.isfile(image_path):
            missing.append(image_path)
    return missing


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(os.path.join(APP_ROOT, "templates"), exist_ok=True)
        os.makedirs(IMAGES_DIR, exist_ok=True)

        self.ui = FactorMakerUI()
        self.setCentralWidget(self.ui)

        self.ui.generate_btn.clicked.connect(self.makeFactor)

        tt = jdatetime.datetime.now()

        # first inputs
        self.ui.id_input.setText(str(tt.year)[1:] + "00" +  "000")
        self.ui.date_input.setText(tt.strftime("%Y/%m/%d"))
        self.ui.from_input.setText("تیم توسعه دهنده متانویا")
        self.ui.phone_input.setText("09000000000")

        missing = required_asset_errors()
        if missing:
            listed = "\n".join(missing)
            self.showMessage(
                "فایل‌های مورد نیاز پیدا نشدند:\n"
                f"{listed}\n\n"
                "لطفا قالب و تصاویر را در پوشه برنامه قرار دهید."
            )

    def showMessage(self, message:str):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("پیام")
        msg_box.setText(message)
        msg_box.exec()

    def makeFactor(self):
        missing = required_asset_errors()
        if missing:
            listed = "\n".join(missing)
            self.showMessage(f"فایل‌های مورد نیاز پیدا نشدند:\n{listed}")
            return

        d = PersianFactor(self.showMessage)

        if not len(self.ui.id_input.text()):
            self.showMessage("لطفا ایدی را مشخص کنید")
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
            self.showMessage("ایدی باید عدد باشد")
            return

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
            "factor_disc_2": d.persianNumber(self.ui.desc2_input.toPlainText())
        }
        try:
            d.makeFactor(factor_id, TEMPLATE_PATH, datas)
        except Exception as exc:
            self.showMessage(f"خطا در تولید فاکتور\n{exc}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

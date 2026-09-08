from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox
import sys
from factor_maker_ui import FactorMakerUI  # Make sure the class is in factor_maker_ui.py
from engine import PersianFactor
import time
import jdatetime
import os
import sys
import traceback

NDEBUG = (int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[1].lower() == '--ndebug' else 0) or 0

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # directory check
        dir_output = os.getcwd().replace('\\','/') + "/output"
        dir_templates = os.getcwd().replace('\\','/') + "/templates"
        dir_images = os.getcwd().replace('\\','/') + "/images"
        
        if not os.path.exists(dir_output): os.mkdir(dir_output)
        if not os.path.exists(dir_templates): os.mkdir(dir_templates)
        if not os.path.exists(dir_images): os.mkdir(dir_images)

        self.ui = FactorMakerUI()
        self.setCentralWidget(self.ui)
        # self.setFixedSize(self.ui.size())

        self.ui.generate_btn.clicked.connect(self.makeFactor)

        tt = jdatetime.datetime.now()

        # first inputs
        self.ui.id_input.setText(str(tt.year)[1:] + "00" +  "000")
        self.ui.date_input.setText(tt.strftime("%Y/%m/%d"))
        self.ui.from_input.setText("تیم توسعه دهنده متانویا")
        self.ui.phone_input.setText("09000000000")

    def showMessage(self, message:str):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("پیام")
        msg_box.setText(message)
        msg_box.exec()

    def makeFactor(self):
        # try:
            d = PersianFactor(self.showMessage)

            if not len(self.ui.id_input.text()): self.showMessage("لطفا ایدی را مشخص کنید")
            elif not len(self.ui.date_input.text()): self.showMessage("لطفا تاریخ را مشخص کنید")
            # elif not len(self.ui.addons_input.text()): self.showMessage("لطفا افزونه ها را مشخص کنید")
            elif not len(self.ui.title_input.text()): self.showMessage("لطفا عنوان را مشخص کنید")
            elif not len(self.ui.from_input.text()): self.showMessage("لطفا از را مشخص کنید")
            elif not len(self.ui.phone_input.text()): self.showMessage("لطفا تلفن را مشخص کنید")
            elif not len(self.ui.location_input.toPlainText()): self.showMessage("لطفا آدرس را مشخص کنید")
            else:
                datas = {
                    "factor_id":d.persianNumber(self.ui.id_input.text()),
                    "factor_date":d.persianNumber(self.ui.date_input.text()),
                    "factor_addons":d.persianNumber(self.ui.addons_input.text()),

                    "title":self.ui.title_input.text(),
                    "sby":self.ui.from_input.text(),
                    "phone":self.ui.phone_input.text(),
                    "location":self.ui.location_input.toPlainText(),

                    "items":self.ui.get_table_items(),
                    "off":float(d.persianNumber(self.ui.off_input.text())),
                    "tax":float(d.persianNumber(self.ui.tax_input.text())),

                    "factor_disc_1":d.persianNumber(self.ui.desc1_input.toPlainText()),
                    "factor_disc_2":d.persianNumber(self.ui.desc2_input.toPlainText())
                }
                d.makeFactor(int(self.ui.id_input.text()),"templates/template.docx",datas)
        # except Exception as e:
            # if not NDEBUG:
                # self.showMessage("خطا در تولید فاکتور\nلطفا از تگ NDEBUG استفاده کنید")
            # else: 
                # self.showMessage(f"خطا در تولید فاکتور\n\n{sys.exc_info()}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

try:
    from fpdf import FPDF
    from PIL import Image
    from barcode import EAN13, EAN8, Code39
    from barcode.writer import ImageWriter
    from PyQt5 import uic, QtWidgets, QtCore, QtGui
    from PyQt5.QtGui import QIcon
    from os.path import expanduser
    import sys
    import traceback
    import logging

except Exception as ex:
    from datetime import datetime
    f = open("traceback.txt", "a")
    f.write(f"[{datetime.now()}] --- {ex}\n")
    f.close()

WIDTH = 105 #mm
HEIGHT = 148

PATH = {"LOGO": "components/logo.png",
        "PN": "components/pn.png",
        "BAR": "components/bar.png",
        "QTY": "components/qty.png",
        "DATE": "components/date.png",
        "RIF": "components/rif.png",
        "DEST": "components/dest.png",
        "OPER": "components/oper.png",
        "AEE": "components/aee.png",
        "ICON": "components/icon.png"}

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi("interface.ui", self)
        self.setWindowIcon(QIcon(PATH["ICON"]))
        # Cartellino
        self.pnInput = self.findChild(QtWidgets.QLineEdit, "pnInput")
        self.revInput = self.findChild(QtWidgets.QLineEdit, "revInput")
        self.qtyInput = self.findChild(QtWidgets.QLineEdit, "qtyInput")
        self.dateEdit = self.findChild(QtWidgets.QDateEdit, "dateEdit")
        self.refInput = self.findChild(QtWidgets.QLineEdit, "refInput")
        self.destInput = self.findChild(QtWidgets.QLineEdit, "destInput")
        self.operatorInput = self.findChild(QtWidgets.QLineEdit, "operatorInput")
        self.imageLabel = self.findChild(QtWidgets.QLabel, "imageLabel")
        self.resetBtn = self.findChild(QtWidgets.QPushButton, "resetBtn")
        self.resetBtn.clicked.connect(self.clear)
        self.genBtn = self.findChild(QtWidgets.QPushButton, "genBtn")
        self.genBtn.clicked.connect(self.gen_pdf)
        self.setFixedSize(self.width(), self.height())
        # Shipping list
        self.listInput = self.findChild(QtWidgets.QLineEdit, "listInput")
        self.clearTableBtn = self.findChild(QtWidgets.QPushButton, "clearTableBtn")
        self.undoBtn = self.findChild(QtWidgets.QPushButton, "undoBtn")
        self.exportBtn = self.findChild(QtWidgets.QPushButton, "exportBtn")
        self.shippingTable = self.findChild(QtWidgets.QTableWidget, "shippingTable")

        pixmap = QtGui.QPixmap(PATH["AEE"]).scaledToWidth(self.imageLabel.width())
        self.imageLabel.setPixmap(pixmap)
        self.show()
        self.dateEdit.setDateTime(QtCore.QDateTime.currentDateTime())

    def clear(self):
        self.pnInput.clear()
        self.revInput.clear()
        self.qtyInput.clear()
        self.dateEdit.clear()
        self.refInput.clear()
        self.destInput.clear()
        self.operatorInput.clear()
        self.dateEdit.setDateTime(QtCore.QDateTime.currentDateTime())

    def gen_pdf(self):
        pn = self.pnInput.text()
        rev = self.revInput.text()
        qty = self.qtyInput.text()
        date = self.dateEdit.text()
        ref = self.refInput.text()
        dest = self.destInput.text()
        operator = self.operatorInput.text()

        d = {"Product number": pn, "Revisione": rev, "Quantità": qty}
        empty = ""
        for entry in d:
            if d[entry] == "":
                empty += f" {entry}\n"
        if empty != "":
            message = f"Non hai inserito dati nei campi\n\n{empty}\n vuoi procedere comunque?"
            if warning_dialog(message, dialog_type="yes_no") == QtWidgets.QMessageBox.No:
                return
        dir = str(QtWidgets.QFileDialog.getSaveFileName(self, "Salva file con nome",
                                                        f"{expanduser('~')}\\Desktop",
                                                        "PDF (*.pdf);;All Files (*)",
                                                        ))
        if dir == "(\'\', \'\')":
            return
        dir = dir.split(",")
        dir = dir[0]
        dir = dir.lstrip("(\'")
        dir = dir.rstrip("'")
        gen_pdf(pn, rev, qty, date, ref, dest, operator, dir)
        message = f"Cartellino generato in {dir}"
        QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "CIP-Gen", message).exec_()


def gen_barcode(pdf: FPDF, pn: str, rev: str, qty: str):
    PN = f"{pn}-{rev}"
    PN.upper()
    with open("tmp.png", "wb") as tmp_file:
        Code39(PN, writer=ImageWriter()).write(tmp_file)
    dim = scale_image_dim("tmp.png", 15)
    pdf.image("tmp.png", x=HEIGHT-50, y=63, w=dim["width"], h=dim["height"])

    while len(qty) != 7:
        qty = f"0{qty}"
    with open("tmp.png", "wb") as tmp_file:
        Code39(qty, writer=ImageWriter()).write(tmp_file)
    dim = scale_image_dim("tmp.png", 15)
    pdf.image("tmp.png", x=HEIGHT-36.3, y=83, w=dim["width"], h=dim["height"])


def add_image(pdf: FPDF, img_path: str, x, y, w_ext=0, h_ext=0, scale=8):
    dim = scale_image_dim(img_path, scale)
    pdf.image(img_path, x=x, y=y, w=dim['width']+w_ext, h=dim['height']+h_ext)


def scale_image_dim(img_path: str, scale_factor: int):
    image = Image.open(img_path)
    size = image.size
    dim = {"width": size[0] / scale_factor,
           "height": size[1] / scale_factor}
    return dim


def gen_pdf(pn: str, rev: str, qty: str, date: str, ref: str, dest: str, oper: str, dir: str):
    pdf = FPDF('P', 'mm', [WIDTH, HEIGHT])
    pdf.add_page('horizontal')
    pdf.set_font('Arial', 'B', 16)

    # LOGO
    dim = scale_image_dim(PATH["LOGO"], 8)
    add_image(pdf, PATH["LOGO"], x=(HEIGHT - dim['width']) / 2, y=2)

    # PN
    add_image(pdf, PATH["PN"], x=2, y=40)
    add_image(pdf, PATH["BAR"], x=15, y=44, w_ext=85)
    pdf.text(x=18, y=43, txt=f"{pn.upper()}-{rev.upper()}")

    # QTY
    add_image(pdf, PATH["QTY"], x=1.5, y=52)
    add_image(pdf, PATH["BAR"], x=15, y=56, w_ext=5)
    pdf.text(x=18, y=55, txt=qty)

    # DATE
    add_image(pdf, PATH["DATE"], x=70, y=52.2)
    add_image(pdf, PATH["BAR"], x=87, y=56, w_ext=11)
    pdf.text(x=90, y=55, txt=date)

    # RIF
    add_image(pdf, PATH["RIF"], x=1.5, y=64, scale=9)
    add_image(pdf, PATH["BAR"], x=27, y=68.5, w_ext=30)
    pdf.text(x=30, y=67.5, txt=ref)

    # DEST
    add_image(pdf, PATH["DEST"], x=1, y=76, scale=7)
    add_image(pdf, PATH["BAR"], x=15, y=80, w_ext=42)
    pdf.text(x=18, y=79, txt=dest)

    # OPER
    add_image(pdf, PATH["OPER"], x=1, y=88, scale=7)
    add_image(pdf, PATH["BAR"], x=15, y=92, w_ext=36)
    pdf.text(x=18, y=91, txt=oper)

    # barcodes
    gen_barcode(pdf, pn=pn, rev=rev, qty=qty)

    # print pdf
    pdf.output(dir, 'F')


def warning_dialog(message: str, dialog_type: str):
    dialog = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Warning", message)
    if dialog_type == "yes_no":
        dialog.setStandardButtons(QtWidgets.QMessageBox.Yes)
        dialog.addButton(QtWidgets.QMessageBox.No)
        dialog.setDefaultButton(QtWidgets.QMessageBox.No)
        return dialog.exec_()
    elif dialog_type == "ok":
        dialog.setStandardButtons(QtWidgets.QMessageBox.Ok)
        return dialog.exec_()
    elif dialog_type == "yes_no_chk":
        dialog.setStandardButtons(QtWidgets.QMessageBox.Yes)
        dialog.addButton(QtWidgets.QMessageBox.No)
        dialog.setDefaultButton(QtWidgets.QMessageBox.No)
        cb = QtWidgets.QCheckBox("Click here to load cannolo image.")
        dialog.setCheckBox(cb)
        result = [dialog.exec_(), True if cb.isChecked() else False]
        return result
    elif dialog_type == "yes_no_cancel":
        dialog.addButton(QtWidgets.QMessageBox.Yes)
        dialog.addButton(QtWidgets.QMessageBox.No)
        dialog.addButton(QtWidgets.QMessageBox.Cancel)
        dialog.setDefaultButton(QtWidgets.QMessageBox.Cancel)
        return dialog.exec_()


def main():
    # noinspection PyBroadException
    try:
        app = QtWidgets.QApplication(sys.argv)
        window = Ui()
        app.exec_()

    except KeyboardInterrupt:
        print("KeyboardInterrupt")

    except BaseException:
        print(traceback.print_exc(file=sys.stdout))

def save_traceback():
    from datetime import datetime
    f = open("traceback.txt", "a")
    f.write(f"[{datetime.now()}] --- {ex}\n")
    f.close()

if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        save_traceback()
    except OSError as ex:
        save_traceback()

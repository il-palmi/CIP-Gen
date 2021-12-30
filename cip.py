try:
    from fpdf import FPDF
    from PIL import Image
    from barcode import Code128
    from barcode.writer import ImageWriter
    from PyQt5 import uic, QtWidgets, QtCore, QtGui
    from PyQt5.QtGui import QIcon
    from os.path import expanduser
    import sys
    import traceback
    import logging
    import xlsxwriter
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
        self.lastInput = None

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
        self.listInput.returnPressed.connect(self.add_to_list)
        self.clearTableBtn = self.findChild(QtWidgets.QPushButton, "clearTableBtn")
        self.clearTableBtn.clicked.connect(self.clear_table)
        self.undoBtn = self.findChild(QtWidgets.QPushButton, "undoBtn")
        self.undoBtn.clicked.connect(self.undo)
        self.exportBtn = self.findChild(QtWidgets.QPushButton, "exportBtn")
        self.exportBtn.clicked.connect(self.export_table)
        self.shippingTable = self.findChild(QtWidgets.QTableWidget, "shippingTable")
        headerLenght = self.shippingTable.horizontalHeader().width()
        self.shippingTable.setColumnWidth(0, headerLenght-100-100)
        self.shippingTable.setColumnWidth(1, 100)
        self.shippingTable.setColumnWidth(2, 100)
        self.shippingTable.horizontalHeader().setStretchLastSection(True)

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
        dir = save_file_dialog(self, title="Salva cartellino", format="pdf")
        if dir is None:
            return
        gen_pdf(pn, rev, qty, date, ref, dest, operator, dir)
        message = f"Cartellino generato in {dir}"
        QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "CIP-Gen", message).exec_()

    def add_to_list(self):
        input = self.listInput.text()
        if not '//' in input:
            message = "Codice non valido."
            warning_dialog(message, dialog_type="ok")
            return
        input = input.split("//")
        try:
            int(input[1])
        except:
            message = "Codice non valido."
            warning_dialog(message, dialog_type="ok")
            return
        pn = input[0]
        qty = input[1]
        self.lastInput = f"{pn}//{qty}"
        for row in range(self.shippingTable.rowCount()):
            if self.shippingTable.item(row, 0).text() == pn:
                self.shippingTable.item(row, 0).setText(pn)
                tot = int(self.shippingTable.item(row, 1).text()) + int(qty)
                self.shippingTable.item(row, 1).setText(f"{tot}")
                envelopes = int(self.shippingTable.item(row, 2).text()) + 1
                self.shippingTable.item(row, 2).setText(f"{envelopes}")
                return
        self.shippingTable.setRowCount(self.shippingTable.rowCount() + 1)
        row = self.shippingTable.rowCount()
        self.shippingTable.setItem(row-1, 0, QtWidgets.QTableWidgetItem(pn))
        self.shippingTable.setItem(row-1, 1, QtWidgets.QTableWidgetItem(qty))
        self.shippingTable.setItem(row-1, 2, QtWidgets.QTableWidgetItem("1"))

    def clear_table(self):
        self.shippingTable.setRowCount(0)

    def undo(self):
        if self.lastInput is None:
            return
        for row in range(self.shippingTable.rowCount()):
            pn = self.lastInput.split("//")[0]
            qty = self.lastInput.split("//")[1]
            if self.shippingTable.item(row, 0).text() == pn:
                tot = int(self.shippingTable.item(row, 1).text()) - int(qty)
                self.shippingTable.item(row, 1).setText(f"{tot}")
                envelopes = int(self.shippingTable.item(row, 2).text()) - 1
                self.shippingTable.item(row, 2).setText(f"{envelopes}")
                self.lastInput = None
                return

    def export_table(self):
        d = save_file_dialog(self, title="Salva shipping list", format="excel")
        if d is None:
            return
        try:
            workbook = xlsxwriter.Workbook(d)
        except Exception:
            print("asd")
        worksheet = workbook.add_worksheet()
        excel_rows = []


        for row in range(self.shippingTable.rowCount()):
            excel_rows.append([self.shippingTable.item(row, 0).text(),
                               self.shippingTable.item(row, 1).text(),
                               self.shippingTable.item(row, 2).text()])

        bold_format = workbook.add_format({'bold': True, 'italic': False})
        worksheet.write_row(0, 0, ["CODICE", "QUANTITA'", "N° BUSTE"], bold_format)
        ps_max_width = len("CODICE")
        qty_max_width = len("QUANTITA'")
        num_max_width = len("N° BUSTE")
        for idx, row in enumerate(excel_rows):
            worksheet.write_row(idx+1, 0, row)
            if len(row[0]) > ps_max_width:
                ps_max_width = len(row[0])
            if len(row[1]) > qty_max_width:
                ps_max_width = len(row[1])
            if len(row[2]) > num_max_width:
                ps_max_width = len(row[2])
        worksheet.set_column(0, 0, ps_max_width+2)
        worksheet.set_column(1, 1, qty_max_width+2)
        worksheet.set_column(2, 2, num_max_width+2)
        try:
            workbook.close()
        except BaseException:
            message = "Impossibile salvare il file. Controllare che non sia in uso su altri programmi."
            warning_dialog(message, dialog_type="ok")


def gen_barcode(pdf: FPDF, pn: str, rev: str, qty: str):
    PN = f"{pn}-{rev}"
    text = f"{PN}//{qty}"
    with open("tmp.png", "wb") as tmp_file:
        Code128(text, writer=ImageWriter()).write(tmp_file)
    dim = scale_image_dim("tmp.png", 15)
    pdf.image("tmp.png", x=HEIGHT-(dim["width"])/2-33, y=73, w=dim["width"], h=dim["height"])


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
    add_image(pdf, PATH["BAR"], x=27, y=68.5, w_ext=8)
    pdf.text(x=30, y=67.5, txt=ref)

    # DEST
    add_image(pdf, PATH["DEST"], x=1, y=76, scale=7)
    add_image(pdf, PATH["BAR"], x=15, y=80, w_ext=20)
    pdf.text(x=18, y=79, txt=dest)

    # OPER
    add_image(pdf, PATH["OPER"], x=1, y=88, scale=7)
    add_image(pdf, PATH["BAR"], x=15, y=92, w_ext=20)
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


def save_file_dialog(window: Ui, title: str, format: str):
    if format == "pdf":
        format = "PDF (*.pdf)"
    elif format == "excel":
        format = "Excel File (*.xlsx)"
    else:
        print(f"Error: {format} type is not supported.")
        return
    dir = QtWidgets.QFileDialog.getSaveFileName(window, title,
                                                    f"{expanduser('~')}\\Desktop",
                                                    f"{format};;All Files (*)",
                                                    )
    if dir[0] == '':
        return None
    return dir[0]

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

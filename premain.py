from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from main_window import Ui as mainWindow

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        print("Program running...")
        uic.loadUi('ui/premain_1.ui', self)
        self.setWindowTitle("Masukkan identitas")
        self.cudate = QtCore.QDate.currentDate()
        self.identity = []
        self.dateEdit.setDate(self.cudate)

        #self.nameID.setText("\xBC")
        #self.ageID.setText("\xBD")
        #self.eduID.setText("\xBE")

        self.setFixedSize(412, 380)
        self.verticalLayout.setContentsMargins(20,0,20,0)
        self.show()
        self.pushButton.clicked.connect(self.next)

    def next(self):
        if (self.nameID.text() is not '' and self.ageID.text() is not '' and self.eduID.text() is not '' and self.eduID_2.text() is not '' and self.eduID_3.text() is not '' and self.posID.text() is not '' and self.phoneID.text() is not ''):
            self.identity.append(self.nameID.text())
            self.identity.append(self.ageID.text())
            self.identity.append(self.eduID.text())
            self.identity.append(self.eduID_2.text())
            self.identity.append(self.eduID_3.text())
            self.identity.append(self.posID.text())
            self.identity.append("\'"+self.phoneID.text())
            #print("pass")
            self.identity.append(self.dateEdit.date().toPyDate())

            self.MainWindow = QtWidgets.QMainWindow()
            self.MainWindow.ui = mainWindow(self.identity)
            self.close()
        else:
            self.buttonReply = QtWidgets.QMessageBox
            self.warning = self.buttonReply.question(self, 'PERINGATAN', "Mohon isi semua kolom", QtWidgets.QMessageBox.Ok)
            #if self.warning == QtWidgets.QMessageBox.Ok:
                #print('Yes clicked.')
                #self.buttonReply.close()
                #print("Modar")
            #print("empty val")
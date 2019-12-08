# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'subwindow.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from ist3 import MyApp as IST3
from tiu import MyApp as TIU
from adkudag import MyApp as Adkudag
from disc import MyApp as DISC
from ist5 import MyApp as IST5
from AA1 import Ui as ArmyAlpha1
from PyQt5 import QtMultimedia
import xlsxwriter

class Ui_Dialog(QtWidgets.QDialog):
    def __init__(self, listdat, workbook, parentWin = None):
        self.listdata = listdat
        self.parentWin = parentWin

        #print(self.listdata)
        self.workbook = workbook
        if self.listdata[0] == "Tes TIU":
            self.filename = "data/instruction/TIU.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)
        elif self.listdata[0] == "Tes IST3":
            self.filename = "data/instruction/IST3.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)
        elif self.listdata[0] == "Tes IST5":
            self.filename = "data/instruction/IST5.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)
        elif self.listdata[0] == "Tes B-S":
            self.filename = "data/instruction/ADKUDAG.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)
        elif self.listdata[0] == "Tes Army Alpha":
            self.filename = "data/instruction/AA_1.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)
        elif self.listdata[0] == "Tes DISC":
            self.filename = "data/instruction/DISC.wav"
            self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
            self.player = QtMultimedia.QSound(self.url)

        self.namedate = [listdat[3], listdat[4]]
        self.mainWin = listdat[5]
        #print(str(self.namedate))

    def setupUi(self, Dialog):
        #Dialog.setWindowTitle(self.listdata[0])
        #self.setWindowTitle(self.listdata[0])
        self.labelfont = QtGui.QFont()
        self.labelfont.setBold(True)

        self.dial = Dialog
        self.dial.setWindowFlag(QtCore.Qt.WindowCloseButtonHint,False)
        self.dial.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
        self.dial.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)

        Dialog.setObjectName("Dialog")
        Dialog.setFixedSize(800, 500)
        #Dialog.setWindowFlag(self.windowFlags() & ~QtCore.Qt.WindowCloseButtonHint)
        self.verticalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 520, 351))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label.setObjectName("label")
        #self.label.setFont(QFont("Serif", 12))
        self.label.setStyleSheet("font: 12pt Serif; font-weight: bold")
        self.horizontalLayout.addWidget(self.label)
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_2.setAlignment(QtCore.Qt.AlignRight)
        self.label_2.setObjectName("label_2")
        self.label_2.setStyleSheet("font: 12pt Serif; font-weight: bold")
        #self.label_2.setVisible(False)
        #self.label_2.setFont(QFont("Serif", 12))
        self.horizontalLayout.addWidget(self.label_2)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.label_3 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.label_3.setStyleSheet("font: 10pt Serif")#; font-weight: bold")
        #self.label_3.setFont(QFont("Serif", 12))
        self.label_3.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addWidget(self.label_3)

        if self.listdata[0]=='Tes TIU':
            self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap = QtGui.QPixmap('data/instruction/TIU1.png').scaledToHeight(75)
            self.label_4.setPixmap(self.pixmap)
            self.verticalLayout.addWidget(self.label_4)

            self.label_5 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.label_5.setText("Bila A diputar diperoleh B. Bila C diputar diperoleh gambar .................. \n"+
                                 "Carilah gambar tersebut dan berilah coretan di bawahnya.")
            self.label_5.setObjectName("label_5")
            self.label_5.setStyleSheet("font: 10pt Serif")#; font-weight: bold")
            self.label_5.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
            self.verticalLayout.addWidget(self.label_5)

            self.label_6 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap_2 = QtGui.QPixmap('data/instruction/TIU2.png').scaledToHeight(75)
            self.label_6.setPixmap(self.pixmap_2)
            self.verticalLayout.addWidget(self.label_6)
        elif self.listdata[0]=='Tes IST5':
            self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap = QtGui.QPixmap('data/instruction/IST5.png').scaledToHeight(60)
            self.label_4.setPixmap(self.pixmap)
            self.verticalLayout.addWidget(self.label_4)

            self.label_5 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.label_5.setText("Dengan sepeda Husin dapat mencapai 15km dalam waktu 1 jam. Berapakah yang dapat\n"+
                                "ia capai dalam waktu 4 jam?\n"+
                                "Jawabannya ialah: 60\n"+
                                "Maka untuk menunjukkan jawaban itu angka 6 dan 0 seharusnya yang dicoret.")
            self.label_5.setObjectName("label_5")
            self.label_5.setStyleSheet("font: 10pt Serif")#; font-weight: bold")
            self.label_5.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
            self.verticalLayout.addWidget(self.label_5)

        elif self.listdata[0]=='Tes IST3':
            self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap = QtGui.QPixmap('data/instruction/IST3.png').scaledToHeight(30)
            self.label_4.setPixmap(self.pixmap)
            self.verticalLayout.addWidget(self.label_4)

            self.label_5 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.label_5.setText("Gelap ialah lawannya dari terang, maka untuk basah lawannya ialah kering\n" +
                                 "Maka jawabannya ialah: e) kering.\n" +
                                 "Oleh karena itu huruf e seharusnya dicoret.")
            self.label_5.setObjectName("label_5")
            self.label_5.setStyleSheet("font: 10pt Serif")#; font-weight: bold")
            self.label_5.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
            self.verticalLayout.addWidget(self.label_5)

            self.label_6 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap_2 = QtGui.QPixmap('data/instruction/IST3_2.png').scaledToHeight(35)
            self.label_6.setPixmap(self.pixmap_2)
            self.verticalLayout.addWidget(self.label_6)
        elif self.listdata[0] == 'Tes B-S':
            self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget)
            self.pixmap = QtGui.QPixmap('data/instruction/adkudag.png').scaledToHeight(130)
            self.label_4.setPixmap(self.pixmap)
            self.verticalLayout.addWidget(self.label_4)

        self.instr = QtWidgets.QPushButton(self.verticalLayoutWidget)#play instruction sound
        self.instr.setObjectName("instr")
        self.verticalLayout.addWidget(self.instr)

        self.pButton = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pButton.setObjectName("pButton")
        self.verticalLayout.addWidget(self.pButton)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Instruksi "+self.listdata[0]))
        self.label.setText(_translate("Dialog", self.listdata[0]))
        if self.listdata[1]!= '':
            self.label_2.setText(_translate("Dialog", "Waktu: "+self.listdata[1]))
        else:
            self.label_2.setText(_translate("Dialog", self.listdata[1]))
        self.label_2.setVisible(False)
        self.label_3.setText(_translate("Dialog", self.listdata[2]))
        self.instr.setText(_translate("Dialog", "Instruksi"))
        self.instr.clicked.connect(lambda: self.play(strin="gas"))
        self.pButton.setText(_translate("Dialog", "Mulai Tes"))
        self.pButton.clicked.connect(lambda: self.onclick(strin="gas"))
        #self.exButton.setText(_translate("Dialog", "Tutup"))

    def play(self, strin):
        self.instr.setEnabled(False)
        self.player.play()

    def onclick(self, strin):
        if not self.player.isFinished():
            self.player.stop()

        self.workbook = [self.workbook, self.namedate, self.mainWin]
        self.pButton.setEnabled(False)
        if self.listdata[0]=="Tes TIU":
            self.window = QtWidgets.QWidget()
            self.window.ui = TIU(self.workbook, self.parentWin)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''
            #self.window.showMaximized()
        elif self.listdata[0]=="Tes IST3":
            self.window = QtWidgets.QWidget()
            self.window.ui = IST3(self.workbook, self.parentWin)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''
        elif self.listdata[0]=="Tes IST5":
            self.window = QtWidgets.QWidget()
            self.window.ui = IST5(self.workbook, self.parentWin)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''
        elif self.listdata[0]=="Tes B-S":
            self.window = QtWidgets.QWidget()
            self.window.ui = Adkudag(self.workbook, self.parentWin)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''
        elif self.listdata[0]=="Tes Army Alpha":
            self.window = QtWidgets.QWidget()
            #print(self.listdata[0])
            self.window.ui = ArmyAlpha1(self.workbook)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''
        elif self.listdata[0]=="Tes DISC":
            self.window = QtWidgets.QWidget()
            self.window.ui = DISC(self.workbook, self.parentWin)
            '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.window.showMaximized()'''

        self.dial.close()
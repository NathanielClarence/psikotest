from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
from PyQt5 import QtCore
import sys
import csv
import collections
import xlsxwriter
import datetime

quest = []
discCat = [
    ['s', 'i','c','d'],
    ['i', 'c','d','s'],
    ['c', 'd','s','i'],
    ['c', 's','d','i'],#
    ['i', 'c','d','s'],
    ['d', 's','i','c'],
    ['c', 's','d','i'],
    ['d', 'i','s','c'],#
    ['i', 's','d','c'],
    ['d', 'c','i','s'],
    ['i', 's','c','d'],
    ['i', 'd','c','s'],#
    ['d', 'i','s','c'],
    ['c', 'd','i','s'],
    ['s', 'i','c','d'],
    ['i', 's','c','d'],#ungu, merah, biru, hijau
    ['c', 's','i','d'],
    ['i', 's','c','d'],
    ['c', 'd','i','s'],
    ['d', 'c','s','i'],#
    ['i', 's','d','c'],
    ['i', 'c','d','s'],
    ['i', 'c','d','s'],
    ['d', 's','i','c']#
]

with open("data/question/disc.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        opts = []
        for col in row:
            #print(row[col])
            opts.append(col)
        quest.append(opts)
        #quest.append(opts)

#print(quest)

class MyApp(QWidget):
    def __init__(self, workbook, parentWin):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes DISC")
        self.q_answer_true = []
        self.q_answer_false = []

        self.parentWin = parentWin

        self.workbook = workbook[0]
        self.namedate = workbook[1]

        self.choices = []
        self.initUI()

    def createLayout_group(self, num):#question, num):
        sgroupbox = QGroupBox(str(num+1), self)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)
        layout_groupbox = QHBoxLayout(sgroupbox)
        self.q_choice = []
        q_ans = quest[num]
        for i in range(len(q_ans)):
            ans = QLineEdit()
            item = QLabel(q_ans[i])

            #if i ==0
            #item = QRadioButton(q_ans[i])
            #self.q_choice.append(item)
            ans.setMaxLength(1)
            ans.setMaximumWidth(20)
            self.q_choice.append(ans)
            layout_groupbox.addWidget(ans)#item change to layoutchoice
            layout_groupbox.addWidget(item)
        layout_groupbox.addStretch(1)
        self.choices.append(self.q_choice)

        return sgroupbox

    def createLayout_Container(self):
        self.scrollarea = QScrollArea(self)
        #self.scrollarea.setFixedWidth(780)
        self.scrollarea.setWidgetResizable(True)

        widget = QWidget()
        self.scrollarea.setWidget(widget)
        self.layout_SArea = QVBoxLayout(widget)

        #groupbox (jumlah soal)
        #num = 0
        for i in range(len(quest)):
            self.layout_SArea.addWidget((self.createLayout_group(i)))
        self.layout_SArea.addStretch(1)

    def initUI(self):
        self.createLayout_Container()
        self.layout_All = QVBoxLayout(self)
        self.layout_All.addWidget(self.scrollarea)
        self.pushButton = QPushButton()
        self.pushButton.setObjectName("pushButton")
        self.layout_All.addWidget(self.pushButton)
        self.pushButton.setText("Selesai")
        self.pushButton.clicked.connect(self.on_click)
        try:
            self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
        except Exception as e:
            print(e)
        #self.showMaximized()
        self.showFullScreen()

    def on_click(self):
        self.enableClear=False

        for x in range(len(self.choices)):
            self.check = []
            self.l_num =0
            self.m_num =0
            self.empt = 0
            for y in self.choices[x]:
                self.check.append(y.text().lower())
            for z in self.check:
                if z =='l':
                    self.l_num+=1
                elif z == 'm':
                    self.m_num+=1
                else:
                    self.empt +=1
            if self.l_num==1 and self.m_num ==1 and self.empt ==2:
                self.enableClear=True
            else:
                self.enableClear=False
                break


        #print("saved")

        if self.enableClear:
            self.parentWin.autosave(ans_disc=self.choices)
            self.close()
        else:
            self.buttonReply = QtWidgets.QMessageBox
            self.warning = self.buttonReply.question(self, 'PERINGATAN', "Tolong cek kembali kolom:\n"+
                                                     "1. Cek kembali kolom yang berisi duplikat\n"+
                                                     "2. Cek kembali apabila masih ada kolom yang belum terisi",
                                                     QtWidgets.QMessageBox.Ok)

'''if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    sys.exit(app.exec_())'''
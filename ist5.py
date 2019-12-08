from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import csv
from PyQt5 import QtCore
import xlsxwriter
import datetime

opts=['0','1','2','3','4',
      '5','6','7','8','9']

quest = []

DURATION_INT = 600

with open("data/question/ist5.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        quest.append(row[0])

class MyApp(QWidget):
    def __init__(self, workbook, parentWin):
        super(MyApp, self).__init__()

        self.setWindowTitle("Tes IST5")

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        self.parentWin = parentWin
        #self.worksheet = self.workbook.add_worksheet("IST5")

        self.choices = []
        self.initUI()
        self.ttimer.setVisible(False)
        self.timer_start()
        self.update_gui()

    def timer_start(self):
        self.time_left_int = DURATION_INT

        self.my_qtimer = QtCore.QTimer(self)
        self.my_qtimer.timeout.connect(self.timer_timeout)
        self.my_qtimer.start(1000)

        self.update_gui()

    def update_gui(self):
        self.ttimer.setText("Waktu tersisa: " + str(int(self.time_left_int / 60)) + ":" + "{0:0=2d}".format(
            int(self.time_left_int % 60)))

    def timer_timeout(self):
        self.time_left_int -= 1

        if self.time_left_int == 0:
            print("timeout")
            self.on_click()

        self.update_gui()

    def createLayout_group(self, question, num):
        sgroupbox = QGroupBox(str(num+1), self)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)

        question = '-\n'.join(question[i:i+80] for i in range(0, len(question), 80))

        layout_vgroupbox = QVBoxLayout(sgroupbox)
        layout_vgroupbox.addWidget(QLabel(question))

        layout_groupbox = QHBoxLayout()
        layout_vgroupbox.addLayout(layout_groupbox)
        self.q_choice = []
        for i in range(len(opts)):
            item = QCheckBox(opts[i])
            self.q_choice.append(item)
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
        num = 0
        for q in quest:
            self.layout_SArea.addWidget(self.createLayout_group(q, num))
            num +=1
        self.layout_SArea.addStretch(1)

    def initUI(self):
        self.createLayout_Container()
        self.layout_All = QVBoxLayout(self)
        #
        self.ttimer = QLabel(self)
        self.ttimer.setText("Hello")
        self.ttimer.setFont(QFont("Serif", 20))
        self.layout_All.addWidget(self.ttimer)
        #
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
        #print("saved")

        self.q_answer = []
        for x in range(len(quest)):
            chosen = []
            for u in self.choices[x]:
                if u.isChecked():
                    chosen.append(u.text())
            self.q_answer.append(set(chosen))

        self.parentWin.autosave(ans_ist5=self.q_answer)
        self.close()

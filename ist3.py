from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
import sys
import csv
import xlsxwriter
import datetime

quest = {}

with open("data/question/ist3.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        opts = []
        for col in range(1,len(row)):
            #print(row[col])
            opts.append(row[col])
        quest[row[0]] = opts
    #print(str(quest))
    #print(str(opts))

DURATION_INT= 300

class MyApp(QWidget):
    def __init__(self, workbook, parentWin):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes IST3")

        self.q_answer = []
        self.choices = []
        self.parentWin = parentWin

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        #self.worksheet = self.workbook.add_worksheet("IST3")

        '''self.answer_k = ['C','E','D','D','D',
                         'B','D','B','E','D',
                         'C','C','C','C','D',
                         'C','C','E','E','E']'''
        '''window_width = 800
        window_height = 600
        self.setFixedSize(window_width, window_height)'''
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
        self.ttimer.setText("Waktu tersisa: "+str(int(self.time_left_int/60))+":"+"{0:0=2d}".format(int(self.time_left_int%60)))

    def timer_timeout(self):
        self.time_left_int -= 1

        if self.time_left_int == 0:
            #print("timeout")
            self.on_click()

        self.update_gui()

    def createLayout_group(self, question, num):
        sgroupbox = QGroupBox(question, self)
        #font test later
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)

        layout_groupbox = QHBoxLayout(sgroupbox)
        self.q_choice = []
        q_ans = quest.get(question)
        for i in range(len(q_ans)):
            item = QRadioButton(q_ans[i])
            self.q_choice.append(item)
            if i == 0:
                item.toggled.connect(lambda: self.chosen(number=num, ans='A'))
            elif i == 1:
                item.toggled.connect(lambda: self.chosen(number=num, ans='B'))
            elif i == 2:
                item.toggled.connect(lambda: self.chosen(number=num, ans='C'))
            elif i == 3:
                item.toggled.connect(lambda: self.chosen(number=num, ans='D'))
            elif i == 4:
                item.toggled.connect(lambda: self.chosen(number=num, ans='E'))
            layout_groupbox.addWidget(item)
        layout_groupbox.addStretch(1)
        self.choices.append(self.q_choice)
        self.q_answer.append("M")

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
        #self.layout_All.setSpacing(0)
        try:
            self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
        except Exception as e:
            print(e)
        #self.showMaximized()
        self.showFullScreen()

    def chosen(self, number, ans):
        self.q_answer[number]=ans
        #print(str(self.q_answer))

    def on_click(self):
        #print("saved")
        self.parentWin.autosave(ans_ist3= self.q_answer)
        self.close()
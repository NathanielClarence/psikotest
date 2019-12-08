from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5 import QtCore
import csv
import xlsxwriter
import datetime

q1 = []
q2 =  []
with open("data/question/adkudag.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        q1.append(row[0])
        q2.append(row[1])
    #print(q1)
    #print(q2)

DURATION_INT = 300

class MyApp(QWidget):
    def __init__(self, workbook, parentWin):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes ADKUDAG")
        self.q_answer = []
        self.choices = []
        self.special1 = ['670\xBE', '43\xBC', '2234\xBC', '8320\xBC', '21.24\xBD', '8934\xBC', '845\xBD']
        self.special2 = ['670\xBE', '43\xBD', '2234\xBC', '8320\xBC', '21.24\xBD', '8934\xBC', '845\xBD']

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        self.parentWin = parentWin
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
        self.ttimer.setText("Waktu tersisa: " + str(int(self.time_left_int / 60)) + ":" + "{0:0=2d}".format(
            int(self.time_left_int % 60)))

    def timer_timeout(self):
        self.time_left_int -= 1

        if self.time_left_int == 0:
            #print("timeout")
            self.on_click()

        self.update_gui()

    def createLayout_group(self, question, usethis = None, sp = False):
        if sp:
            sgroupbox = QGroupBox(self.special1[question] + "   " + self.special2[question], self)
        else:
            sgroupbox = QGroupBox(str(q1[question])+"   "+str(q2[question]), self)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)

        self.q_choice = []
        layout_groupbox = QHBoxLayout(sgroupbox)
        q_ans = ['B', 'S']
        for i in range(len(q_ans)):
            item = QRadioButton(q_ans[i])
            self.q_choice.append(item)
            if i == 0:
                item.toggled.connect(lambda: self.chosen(number=usethis, ans=0))
            elif i == 1:
                item.toggled.connect(lambda: self.chosen(number=usethis, ans=1))
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
        try:
            for q in range(len(q1)):
                if q1[q] == "-":
                    '''label = QLabel(self)
                    pixmap = QPixmap('data/question/adkudag_support/'+str(num)+'.png').scaledToHeight(20)
                    #pixmap = pixmap.scaledToHeight(15)
                    label.setPixmap(pixmap)
                    self.layout_SArea.addWidget(label)'''
                    self.layout_SArea.addWidget(self.createLayout_group(num, q, sp = True))
                    num += 1
                else:
                    self.layout_SArea.addWidget(self.createLayout_group(q, q))
        except Exception as e:
            print(e)
            #num +=1
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

    def chosen(self, number, ans):
        self.q_answer[number]=self.choices[number][ans].text()

    def on_click(self):
        #print("saved")
        self.parentWin.autosave(ans_adkudag= self.q_answer)
        self.close()
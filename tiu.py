#from PyQt5 import QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5 import QtCore
import xlsxwriter
import datetime

DURATION_INT= 480

class MyApp(QWidget):
    def __init__(self, workbook, parentWin):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes TIU")
        self.q_answer = []
        self.choices = []

        self.parentWin = parentWin
        self.workbook = workbook[0]
        self.namedate = workbook[1]
        #self.worksheet = self.workbook.add_worksheet("TIU")

        self.answer_k = ['B','E','D','D','E',
                         'D','E','A','E','C',
                         'E','D','E','C','B',
                         'C','B','C','D','D',
                         'C','C','D','D','E',
                         'A','B','E','C','E']

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

    def createLayout_group(self, num):
        sgroupbox = QGroupBox(str(num+1), self)
        #gb = QVBoxLayout(sgroupbox)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)
        gb = QVBoxLayout(sgroupbox)
        gb.addStretch(1)

        label = QLabel()
        #print(str(num+1)+'_'+ str(i+1)+ '.png')
        pixmap = QPixmap('data/question/tiu/' + str(num + 1) + '.png').scaledToHeight(50)
        label.setPixmap(pixmap)
        gb.addWidget(label)
        #gb.addStretch(1)
        #gb.addWidget()

        layout_groupbox = QHBoxLayout()
        self.q_choice = []
        q_ans = ['a','b','c','d','e']
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
            label = QLabel()
            #print(str(num+1)+'_'+ str(i+1)+ '.png')
            pixmap = QPixmap('data/question/tiu/tiu_opts/' + str(num+1)+'_'+ str(i+1)+ '.png').scaledToHeight(50)
            label.setPixmap(pixmap)
            layout_groupbox.addWidget(label)
        gb.addLayout(layout_groupbox)
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
        for q in range(30):
            self.layout_SArea.addWidget(self.createLayout_group(num))
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

    def chosen(self, number, ans):
        self.q_answer[number]=ans

    def on_click(self):
        #print("saved")
        try:
            self.parentWin.autosave(ans_tiu= self.q_answer)
        except Exception as e:
            print(e)

        self.close()
        #print(str(self.choices))
        #go save here
        #for x in self.choices:
        #    for y in x:
        #        print(y.text())

#if __name__ == "__main__":
#    app = QApplication(sys.argv)
#    window = MyApp()
#    sys.exit(app.exec_())
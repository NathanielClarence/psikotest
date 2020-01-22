from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
from PyQt5 import QtGui
import sys
import xlsxwriter
import datetime

DURATION_INT = 28
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA12.ui', self)

        self.pixmap = QtGui.QPixmap(
            QtCore.QDir.current().absoluteFilePath('data/question/AA/12.png'))  # .scaledToHeight(70)
        self.label.setPixmap(self.pixmap)
        self.label.setGeometry(QtCore.QRect(1, 1, 725, 83))

        self.res = res

        self.workbook = workbook[0]

        self.namedate = workbook[1]
        self.mainWin = workbook[2]

        try:
            self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
            self.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
            self.setWindowFlag(QtCore.Qt.Window)
        except Exception as e:
            print(e)
        # self.showMaximized()
        self.showFullScreen()
        self.startEx.clicked.connect(self.start)
        self.nextEx.clicked.connect(self.done)
        self.ttimer.setText("0:25")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/12.wav"
        self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
        self.player = QtMultimedia.QSound(self.url)

        self.point =0

        self.pushButton.clicked.connect(lambda:self.add(self.pushButton))
        self.pushButton_2.clicked.connect(lambda: self.min(self.pushButton_2))
        self.pushButton_3.clicked.connect(lambda: self.min(self.pushButton_3))
        self.pushButton_4.clicked.connect(lambda: self.min(self.pushButton_4))
        self.pushButton_5.clicked.connect(lambda: self.add(self.pushButton_5))
        self.pushButton_6.clicked.connect(lambda: self.min(self.pushButton_6))
        self.pushButton_7.clicked.connect(lambda: self.min(self.pushButton_7))
        self.pushButton_8.clicked.connect(lambda: self.min(self.pushButton_8))
        self.pushButton_9.clicked.connect(lambda: self.add(self.pushButton_9))

    def add(self, btn):
        self.point+=1
        btn.setEnabled(False)
        #print(self.point)

    def min(self, btn):
        self.point -=1
        btn.setEnabled(False)
       # print(self.point)

    def timer_start(self):
        self.time_left_int = DURATION_INT

        self.my_qtimer = QtCore.QTimer(self)
        self.my_qtimer.timeout.connect(self.timer_timeout)
        self.my_qtimer.start(1000)

        self.update_gui(self.run)

    def update_gui(self, run):
        if run:
            self.ttimer.setText(str(int(self.time_left_int / 60)) + ":" + "{0:0=2d}".format(int(self.time_left_int % 60)))
        else:
            self.ttimer.setText("0:00")

    def timer_timeout(self):
        self.time_left_int -= 1

        if self.time_left_int == 0:
           # print("timeout")
            self.done()

        self.update_gui(self.run)

    def start(self):
        self.audioTimer()
        self.startEx.setEnabled(False)
        self.player.play()
        #print(self.player.fileName())

        #if self.player.isFinished():

    def startTest(self):
        if self.player.isFinished():
            self.a_timer.stop()
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)
            self.pushButton_7.setEnabled(True)
            self.pushButton_8.setEnabled(True)
            self.pushButton_9.setEnabled(True)
            self.nextEx.setEnabled(True)
            self.timer_start()
            self.update_gui(self.run)

    def audioTimer(self):
        #self.a_timer_cnt = 0
        self.a_timer = QtCore.QTimer(self)
        self.a_timer.timeout.connect(self.startTest)
        self.a_timer.start(1000)

    def done(self):
        self.run = False
        self.my_qtimer.stop()
        if (self.point==3):
            #print("1point")
            #add point >
            self.res[11]=1
        #else:
            #print("false input")

        self.pushButton.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.pushButton_4.setEnabled(False)
        self.pushButton_5.setEnabled(False)
        self.pushButton_6.setEnabled(False)
        self.pushButton_7.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_9.setEnabled(False)

        #print(sum(self.res))
        self.mainWin.setAAResult(_num = sum(self.res))
        self.mainWin.autosave(ans_aa= self.res)

        self.close()
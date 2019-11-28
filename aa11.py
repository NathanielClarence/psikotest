from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
import sys
#from math import sqrt, pow
import math
from aa12 import Ui as ArmyAlpha12
import xlsxwriter

DURATION_INT = 20
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA11.ui', self)
        self.res = res
        self.workbook = workbook
        self.show()
        self.startEx.clicked.connect(self.start)
        self.nextEx.clicked.connect(self.done)
        self.ttimer.setText("0:20")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/11.wav"
        self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
        self.player = QtMultimedia.QSound(self.url)

        self.cb_list = []
        self.cb_list.append(self.checkBox)
        self.cb_list.append(self.checkBox_2)
        self.cb_list.append(self.checkBox_3)
        self.cb_list.append(self.checkBox_4)
        self.cb_list.append(self.checkBox_5)
        self.cb_list.append(self.checkBox_6)
        self.cb_list.append(self.checkBox_7)
        self.cb_list.append(self.checkBox_8)
        self.cb_list.append(self.checkBox_9)
        self.cb_list.append(self.checkBox_10)
        self.cb_list.append(self.checkBox_11)
        self.cb_list.append(self.checkBox_12)
        self.cb_list.append(self.checkBox_13)
        self.cb_list.append(self.checkBox_14)
        self.cb_list.append(self.checkBox_15)
        self.cb_list.append(self.checkBox_16)

        self.cb_list2 = {}
        for x in self.cb_list:
            self.cb_list2[x] = False

        self.checkBox.stateChanged.connect(lambda: self.select(self.cb_list[0]))
        self.checkBox_2.stateChanged.connect(lambda: self.select(self.cb_list[1]))
        self.checkBox_3.stateChanged.connect(lambda: self.select(self.cb_list[2]))
        self.checkBox_4.stateChanged.connect(lambda: self.select(self.cb_list[3]))
        self.checkBox_5.stateChanged.connect(lambda: self.select(self.cb_list[4]))
        self.checkBox_6.stateChanged.connect(lambda: self.select(self.cb_list[5]))
        self.checkBox_7.stateChanged.connect(lambda: self.select(self.cb_list[6]))
        self.checkBox_8.stateChanged.connect(lambda: self.select(self.cb_list[7]))
        self.checkBox_9.stateChanged.connect(lambda: self.select(self.cb_list[8]))
        self.checkBox_10.stateChanged.connect(lambda: self.select(self.cb_list[9]))
        self.checkBox_11.stateChanged.connect(lambda: self.select(self.cb_list[10]))
        self.checkBox_12.stateChanged.connect(lambda: self.select(self.cb_list[11]))
        self.checkBox_13.stateChanged.connect(lambda:self.select(self.cb_list[12]))
        self.checkBox_14.stateChanged.connect(lambda: self.select(self.cb_list[13]))
        self.checkBox_15.stateChanged.connect(lambda: self.select(self.cb_list[14]))
        self.checkBox_16.stateChanged.connect(lambda: self.select(self.cb_list[15]))

        self.selected = 0

    def select(self, btn):
        #print(btn.text())
        if btn.isChecked():
            self.selected+=1
        else:
            self.selected-=1
        #print(self.selected)

        if self.selected >= 2:
            for x in self.cb_list:
                if not x.isChecked():
                    x.setEnabled(False)
        else:
            for x in self.cb_list:
                x.setEnabled(True)

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
        self.startTest()
        #if self.player.isFinished():

    def startTest(self):
        if self.player.isFinished():
            self.a_timer.stop()
            for x in self.cb_list:
                x.setEnabled(True)
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
        self.ans1 = 0
        self.ans2 = 0
        for i in self.cb_list:
            if i.isChecked() and not self.cb_list2[i]:
                #print("dds")
                self.cb_list2[i]=True
                self.ans1 = int(i.text())
                #print("ssss")
                #
                break
        #self.cb_list.remove(i)
        #print(self.ans1)
        for j in self.cb_list:
            if j.isChecked() and not self.cb_list2[j]:
                self.ans2 = int(j.text())
                #self.cb_list.remove(i)
                break
        #print(self.cb_list2[i])
        self.cb_list2[i]=False
        #print(self.cb_list2[i])
        #print(self.ans2)
        #self.cb_list.remove(i)
        #print(self.ans2)
        #self.ans = self.ans1-self.ans2
        #print(self.ans)
        #print(int(math.sqrt(math.pow(self.ans, 2))))
        if int(math.sqrt(math.pow(self.ans1-self.ans2, 2)))==3:
            #print("1point")
            self.res[10]=1
        #else:
            #print("false input")
        for x in self.cb_list:
            x.setEnabled(False)
        #self.nextEx.setEnabled(False)

        #go to next problem AA2
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha12(self.res, self.workbook)
        self.close()
        #close this win if possible

#if __name__ == "__main__":
#    app = QtWidgets.QApplication(sys.argv)
#    window = Ui()
#    sys.exit(app.exec_())
from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
import sys
from aa10 import Ui as ArmyAlpha10
import xlsxwriter

DURATION_INT = 18
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA9.ui', self)
        self.res = res
        self.workbook = workbook
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
        self.ttimer.setText("0:15")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/09.wav"
        self.url = QtCore.QDir.current().absoluteFilePath(self.filename)
        self.player = QtMultimedia.QSound(self.url)

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
            #print("timeout")
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
        self.checker = True
        for i in range(len(self.cb_list)):
            if ((i == 4 or i==10) and not self.cb_list[i].isChecked()):
                self.checker = False
                break
            elif (self.cb_list[i].isChecked() and (i!=4 and i!=10)):
                self.checker = False
                break
        if (self.checker):
            #print("1point")
            self.res[8]=1
            #add point >
        #else:
            #print("false input")

        for x in self.cb_list:
            x.setEnabled(False)
        #self.nextEx.setEnabled(False)

        #go to next problem AA2
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha10(self.res, self.workbook)
        self.close()
        #close this win if possible

#if __name__ == "__main__":
#    app = QtWidgets.QApplication(sys.argv)
#    window = Ui()
#    sys.exit(app.exec_())
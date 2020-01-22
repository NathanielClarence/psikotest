from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
import sys
from aa4 import Ui as ArmyAlpha4
import xlsxwriter

DURATION_INT = 8
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA3.ui', self)
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
        self.ttimer.setText("0:05")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/03.wav"
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
            self.radioButton.setEnabled(True)
            self.radioButton_2.setEnabled(True)
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
        if (self.radioButton.isChecked()):
            #print("1point")
            self.res[2] = 1
            #add point >
        #else:
           # print("false input")

        self.radioButton.setEnabled(False)
        self.radioButton_2.setEnabled(False)
        #self.nextEx.setEnabled(False)
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha4(self.res, self.workbook)
        self.close()
        #go to next problem AA2
        #close this win if possible

'''if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    sys.exit(app.exec_())'''
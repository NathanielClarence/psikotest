from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
from PyQt5 import QtGui
import sys
from aa3 import Ui as ArmyAlpha3
import xlsxwriter

DURATION_INT = 5
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA2.ui', self)
        #self.label_2
        #print("ss")

        #print("ss")
        #self.label_2.setGeometry(QtCore.QRect(1,1,607,73))
        self.res = res
        self.workbook = workbook
        self.show()
        self.pixmap = QtGui.QPixmap(QtCore.QDir.current().absoluteFilePath('data/question/AA/2.png'))#.scaledToHeight(70)
        self.label_2.setPixmap(self.pixmap)
        self.label_2.setGeometry(QtCore.QRect(1, 1, 607, 73))
        self.startEx.clicked.connect(self.start)
        self.nextEx.clicked.connect(self.done)
        self.ttimer.setText("0:05")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/02.wav"
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
            self.lineEdit_2.setEnabled(True)
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
        no2 = self.lineEdit_2.text()
        if (no2 == '654'):
            #print("1point")
            self.res[1]=1
            #add point >
            #print(self.res)
        #else:
            #print("false input")

        self.lineEdit_2.setEnabled(False)
        #self.nextEx.setEnabled(False)
        #go to next problem AA2
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha3(self.res, self.workbook)
        self.close()
        #close this win if possible

'''if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    sys.exit(app.exec_())'''
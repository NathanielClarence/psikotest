from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
import sys
from PyQt5 import QtGui
from aa7 import Ui as ArmyAlpha7
import xlsxwriter

DURATION_INT = 10
class Ui(QtWidgets.QMainWindow):
    def __init__(self, res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA6.ui', self)
        self.pixmap = QtGui.QPixmap(
            QtCore.QDir.current().absoluteFilePath('data/question/AA/6.png'))  # .scaledToHeight(70)
        self.label.setPixmap(self.pixmap)
        self.label.setGeometry(QtCore.QRect(40, 20, 521, 201))
        self.res = res
        self.workbook = workbook
        self.show()
        self.startEx.clicked.connect(self.start)
        self.nextEx.clicked.connect(self.done)
        self.ttimer.setText("0:10")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/06.wav"
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
            self.lineEdit_5.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_4.setEnabled(True)
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
        no1 = self.lineEdit_5.text()
        no2 = self.lineEdit_2.text()
        no3 = self.lineEdit_3.text()
        no4 = self.lineEdit_4.text()
        if (no3.lower()=='x' and no4.lower()=='a' and no1 == '' and no2==''):
            #print("1point")
            self.res[5]=1
            #add point >
        #else:
            #print("false input")

        self.lineEdit_5.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_3.setEnabled(False)
        self.lineEdit_4.setEnabled(False)
        #self.nextEx.setEnabled(False)

        #go to next problem AA2
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha7(self.res, self.workbook)
        #self.nextwindow.ui = ArmyAlpha7([1,1,1,1,1,1,0,0,0,0,0,0], xlsxwriter.Workbook('data/result/test.xlsx'))
        self.close()
        #close this win if possible

'''if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    sys.exit(app.exec_())'''
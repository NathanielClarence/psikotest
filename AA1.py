from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
import sys
from aa2 import Ui as ArmyAlpha2
import xlsxwriter

DURATION_INT = 5
class Ui(QtWidgets.QMainWindow):
    def __init__(self, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA1.ui', self)
        self.res = [0,0,0,0,0,0,
                    0,0,0,0,0,0]
        #print('aa1')
        self.workbook = workbook
        #self.worksheet = self.workbook.add_worksheet("AA")
        #self.ttimer.setVisible(False)

        self.show()
        self.startEx.clicked.connect(self.start)
        self.nextEx.clicked.connect(self.done)
        self.ttimer.setText("0:05")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/01.wav"
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
            self.run = False
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
            self.lineEdit.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_4.setEnabled(True)
            self.lineEdit_5.setEnabled(True)
            self.nextEx.setEnabled(True)
            self.timer_start()
            self.update_gui(self.run)

    def audioTimer(self):
        #self.a_timer_cnt = 0
        self.a_timer = QtCore.QTimer(self)
        self.a_timer.timeout.connect(self.startTest)
        self.a_timer.start(1000)

    def done(self):
        #self.run = False
        #print("ada")
        self.my_qtimer.stop()
        #print("sesuatu")

        no1 = self.lineEdit.text()
        #print(no1+'aa')
        no2 = self.lineEdit_2.text()
        #print(no2+'aa')
        no3 = self.lineEdit_3.text()
        no4 = self.lineEdit_4.text()
        no5 = self.lineEdit_5.text()
        if (no3.lower()=='x' and no5.lower()=='a' and no1 == '' and no2=='' and no4==''):
            self.res[0]=1
            #print("1point")
            #add point >
            #self.res[0]=1
        #else:
            #pass
            #print("false input")

        self.lineEdit.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_3.setEnabled(False)
        self.lineEdit_4.setEnabled(False)
        self.lineEdit_5.setEnabled(False)
        #self.nextEx.setEnabled(False)

        #go to next problem AA2
        #print("go")
        self.nextwindow = QtWidgets.QWidget()
        #print("ready")
        self.nextwindow.ui = ArmyAlpha2(self.res, self.workbook)
        #print()
        self.close()
        #close this win if possible

'''if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    sys.exit(app.exec_())'''
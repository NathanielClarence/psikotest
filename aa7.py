from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore
from PyQt5 import QtMultimedia
from PyQt5 import QtGui
import sys
from aa8 import Ui as ArmyAlpha8
import xlsxwriter

DURATION_INT = 13
class Ui(QtWidgets.QMainWindow):
    def __init__(self,res, workbook):
        super(Ui, self).__init__()
        uic.loadUi('ui/AA7.ui', self)

        self.pixmap = QtGui.QPixmap(
            QtCore.QDir.current().absoluteFilePath('data/question/AA/7.png'))  # .scaledToHeight(70)
        self.label.setPixmap(self.pixmap)
        self.label.setGeometry(QtCore.QRect(40, 20, 611, 201))

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
        self.ttimer.setText("0:10")
        self.ttimer.setVisible(False)
        self.run = True
        self.filename = "data/instruction/AA/07.wav"
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
            self.lineEdit.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_4.setEnabled(True)
            self.lineEdit_5.setEnabled(True)
            self.lineEdit_6.setEnabled(True)
            self.lineEdit_7.setEnabled(True)
            self.lineEdit_8.setEnabled(True)
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
        no1 = self.lineEdit.text()
        no2 = self.lineEdit_2.text()
        no3 = self.lineEdit_3.text()
        no4 = self.lineEdit_4.text()
        no5 = self.lineEdit_5.text()
        no6 = self.lineEdit_6.text()
        no7 = self.lineEdit_7.text()
        no8 = self.lineEdit_8.text()
        if (no3.lower()=='a' and no7.lower()=='x' and no1 == '' and no2=='' and no4=='' and no5=='' and no6 =='' and no8==''):
           #print("1point")
            #add point >
            self.res[6]=1
        #else:
           # print("false input")

        self.lineEdit.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_3.setEnabled(False)
        self.lineEdit_4.setEnabled(False)
        self.lineEdit_5.setEnabled(False)
        self.lineEdit_6.setEnabled(False)
        self.lineEdit_7.setEnabled(False)
        self.lineEdit_8.setEnabled(False)
        #self.nextEx.setEnabled(False)

        #go to next problem AA2
        self.nextwindow = QtWidgets.QWidget()
        self.nextwindow.ui = ArmyAlpha8(self.res, self.workbook)
        self.close()
        #close this win if possible

'''if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    sys.exit(app.exec_())'''
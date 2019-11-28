#from main import Ui_MainWindow
from PyQt5 import QtWidgets
from premain import Ui

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)

    premain = QtWidgets.QWidget()
    premain.ui = Ui()
    #premain.show()

#    MainWindow = QtWidgets.QMainWindow()
#    ui = Ui_MainWindow()
#    ui.setupUi(MainWindow)
#    MainWindow.show()

    sys.exit(app.exec_())
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5 import QtCore
import csv
import xlsxwriter
import datetime

q1 = []
q2 =  []
with open("data/question/adkudag.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        q1.append(row[0])
        q2.append(row[1])
    #print(q1)
    #print(q2)

DURATION_INT = 300

class MyApp(QWidget):
    def __init__(self, workbook):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes ADKUDAG")
        self.q_answer = []
        self.choices = []
        self.special1 = ['670\xBE', '43\xBC', '2234\xBC', '8320\xBC', '21.24\xBD', '8934\xBC', '845\xBD']
        self.special2 = ['670\xBE', '43\xBD', '2234\xBC', '8320\xBC', '21.24\xBD', '8934\xBC', '845\xBD']

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        self.worksheet = self.workbook.add_worksheet("ADKUDAG")

        self.answer_k = ['s','s','b','b','s','s','b','s','b','b',
                         's','b','s','s','s','b','s','s','s','s',
                         'b','b','b','b','s','s','s','b','b','b',
                         'b','b','b','b','b','b','b','b','s','b',
                         's','b','s','b','b','b','b','s','s','b',
                         'b','b','b','b','b','s','s','b','b','s',
                         's','b','b','b','b','s','s','b','s','s',
                         's','b','b','s','b','s','b','s','s','s',
                         's','s','b','b','b','s','b','s','s','b',
                         's','s','s','s','b','s','s','s','s','s',
                         's','b','s','b','s','b','s','s','b','b',
                         'b','s','s','s','b','b','s','s','b','b',
                         's','b','s','b','b','b','s','s','s','b',
                         'b','b','s','b','s','b','s','b','s','s',
                         'b','b','b','s','s','b','b','b','b','b']

        #print(len(self.answer_k))

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

    def createLayout_group(self, question, usethis = None, sp = False):
        if sp:
            sgroupbox = QGroupBox(self.special1[question] + "   " + self.special2[question], self)
        else:
            sgroupbox = QGroupBox(str(q1[question])+"   "+str(q2[question]), self)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)

        self.q_choice = []
        layout_groupbox = QHBoxLayout(sgroupbox)
        q_ans = ['B', 'S']
        for i in range(len(q_ans)):
            item = QRadioButton(q_ans[i])
            self.q_choice.append(item)
            if i == 0:
                item.toggled.connect(lambda: self.chosen(number=usethis, ans=0))
            elif i == 1:
                item.toggled.connect(lambda: self.chosen(number=usethis, ans=1))
            layout_groupbox.addWidget(item)
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
        try:
            for q in range(len(q1)):
                if q1[q] == "-":
                    '''label = QLabel(self)
                    pixmap = QPixmap('data/question/adkudag_support/'+str(num)+'.png').scaledToHeight(20)
                    #pixmap = pixmap.scaledToHeight(15)
                    label.setPixmap(pixmap)
                    self.layout_SArea.addWidget(label)'''
                    self.layout_SArea.addWidget(self.createLayout_group(num, q, sp = True))
                    num += 1
                else:
                    self.layout_SArea.addWidget(self.createLayout_group(q, q))
        except Exception as e:
            print(e)
            #num +=1
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
        #print(number)
        #print("answer saved for "+str(number))
        #print(hex(id(ans)))
        #print(self.choices[number][ans].text())
        self.q_answer[number]=self.choices[number][ans].text()
        print(str(self.q_answer))
        #print(str(self.q_answer))
        #print(str(self.q_answer))

    def on_click(self):
        #print("saved")

        self.res = []
        for x in range(len(self.q_answer)):
            if self.q_answer[x].lower()==self.answer_k[x]:
                self.res.append(1)
            else:
                self.res.append(0)
        #print(self.res)

        row = 0
        col = 0
        self.header = self.workbook.add_format({'bold': True})
        self.fill = self.workbook.add_format({'bg_color': 'lime'})
        self.desired = self.workbook.add_format({'bg_color': 'cyan'})
        self.percentage = self.workbook.add_format({'num_format': '0.00%', 'bg_color':'lime'})
        self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
        self.worksheet.write(0, 1, 'Jawaban', self.header)
        self.worksheet.write(0, 2, 'Skor', self.header)
        self.worksheet.write(0, 3, 'Percentile', self.header)
        row += 1
        for x in range(len(self.answer_k)):
            self.worksheet.write(row, col, self.answer_k[x], self.desired)
            self.worksheet.write(row, col + 1, self.q_answer[x], self.fill)
            self.worksheet.write(row, col + 2, self.res[x], self.fill)
            row += 1

        self.worksheet.write(row, 1, "Total", self.header)
        self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)
        self.worksheet.write(1,3, "=(C"+str(row+1)+"/150)", self.percentage)
        self.worksheet.write(1, 5, 'Scale', self.header)
        self.worksheet.write_formula(1, 6,
                                     '=IF(D2<60%,"Kurang",IF(D2<81%,"Cukup",IF(D2<101%,"Baik")))',
                                     self.fill)

        '''self.worksheet.write(1, 8, "Nama", self.header)
        self.worksheet.write(1, 9, self.namedate[0], self.fill)
        self.worksheet.write(2, 8, "Usia", self.header)
        self.worksheet.write(2, 9, (datetime.date.today().year - self.namedate[1].year), self.fill)
        self.worksheet.write(2, 10, str(self.namedate[1]))
'''
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
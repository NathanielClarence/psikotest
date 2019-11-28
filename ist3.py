from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
import sys
import csv
import xlsxwriter
import datetime

quest = {}

with open("data/question/ist3.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        opts = []
        for col in range(1,len(row)):
            #print(row[col])
            opts.append(row[col])
        quest[row[0]] = opts
    #print(str(quest))
    #print(str(opts))

DURATION_INT= 300

class MyApp(QWidget):
    def __init__(self, workbook):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes IST3")

        self.q_answer = []
        self.choices = []

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        self.worksheet = self.workbook.add_worksheet("IST3")

        self.answer_k = ['C','E','D','D','D',
                         'B','D','B','E','D',
                         'C','C','C','C','D',
                         'C','C','E','E','E']
        window_width = 800
        window_height = 600
        self.setFixedSize(window_width, window_height)
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
        self.ttimer.setText("Waktu tersisa: "+str(int(self.time_left_int/60))+":"+"{0:0=2d}".format(int(self.time_left_int%60)))

    def timer_timeout(self):
        self.time_left_int -= 1

        if self.time_left_int == 0:
            #print("timeout")
            self.on_click()

        self.update_gui()

    def createLayout_group(self, question, num):
        sgroupbox = QGroupBox(question, self)
        #font test later
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)

        layout_groupbox = QHBoxLayout(sgroupbox)
        self.q_choice = []
        q_ans = quest.get(question)
        for i in range(len(q_ans)):
            item = QRadioButton(q_ans[i])
            self.q_choice.append(item)
            if i == 0:
                item.toggled.connect(lambda: self.chosen(number=num, ans='A'))
            elif i == 1:
                item.toggled.connect(lambda: self.chosen(number=num, ans='B'))
            elif i == 2:
                item.toggled.connect(lambda: self.chosen(number=num, ans='C'))
            elif i == 3:
                item.toggled.connect(lambda: self.chosen(number=num, ans='D'))
            elif i == 4:
                item.toggled.connect(lambda: self.chosen(number=num, ans='E'))
            layout_groupbox.addWidget(item)
        layout_groupbox.addStretch(1)
        self.choices.append(self.q_choice)
        self.q_answer.append("M")

        return sgroupbox

    def createLayout_Container(self):
        self.scrollarea = QScrollArea(self)
        self.scrollarea.setFixedWidth(780)
        self.scrollarea.setWidgetResizable(True)

        widget = QWidget()
        self.scrollarea.setWidget(widget)
        self.layout_SArea = QVBoxLayout(widget)

        #groupbox (jumlah soal)
        num = 0
        for q in quest:
            self.layout_SArea.addWidget(self.createLayout_group(q, num))
            num +=1
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
        #self.layout_All.setSpacing(0)
        self.show()

    def chosen(self, number, ans):
        #print(number)
        #print("answer saved for "+str(number))
        #print(hex(id(ans)))
        #print(self.choices[number][ans].text())
        #self.q_answer[number]=self.choices[number][ans].text()
        self.q_answer[number]=ans
        #print(str(self.q_answer))

    def on_click(self):
        #print("saved")
        self.res = []
        for i in range(len(self.q_answer)):
            if self.q_answer[i]==self.answer_k[i]:
                self.res.append(1)
            else:
                self.res.append(0)
        #print(self.res)
        #print(sum(self.res))

        row = 0
        col = 0
        self.header = self.workbook.add_format({'bold': True})
        self.fill = self.workbook.add_format({'bg_color': 'lime'})
        self.desired = self.workbook.add_format({'bg_color': 'cyan'})
        self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
        self.worksheet.write(0, 1, 'Jawaban', self.header)
        self.worksheet.write(0, 2, 'Skor', self.header)
        row += 1
        for x in range(len(self.answer_k)):
            self.worksheet.write(row, col, self.answer_k[x], self.desired)
            self.worksheet.write(row, col + 1, self.q_answer[x], self.fill)
            self.worksheet.write(row, col + 2, self.res[x], self.fill)
            row += 1

        self.worksheet.write(row, 1, "Total", self.header)
        self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)

        self.worksheet.write(1, 5, 'RS', self.header)
        self.worksheet.write(1, 6, '=C' + str(row + 1), self.fill)
        self.worksheet.write(2, 5, 'WS', self.header)
        self.worksheet.write(2, 6, '=(G2*8)+44', self.fill)
        self.worksheet.write(2, 7, "=IF(G3<=78,1,IF(G3<=107,2,IF(G3<=136,3,IF(G3<=165,4,5))))")
        self.worksheet.write(3, 5, 'Scale', self.header)
        self.worksheet.write_formula(3, 6,
                                     '=IF(G3<=78,"(1) Rendah",IF(G3<=107,"(2) Rata-rata Bawah",IF(G3<=136, "(3) Sedang", IF(G3<=165, "(4) Rata-rata atas", "(5) Tinggi"))))',
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
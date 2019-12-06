from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
from PyQt5 import QtCore
import sys
import csv
import collections
import xlsxwriter
import datetime

quest = []
discCat = [
    ['s', 'i','c','d'],
    ['i', 'c','d','s'],
    ['c', 'd','s','i'],
    ['c', 's','d','i'],#
    ['i', 'c','d','s'],
    ['d', 's','i','c'],
    ['c', 's','d','i'],
    ['d', 'i','s','c'],#
    ['i', 's','d','c'],
    ['d', 'c','i','s'],
    ['i', 's','c','d'],
    ['i', 'd','c','s'],#
    ['d', 'i','s','c'],
    ['c', 'd','i','s'],
    ['s', 'i','c','d'],
    ['i', 's','c','d'],#ungu, merah, biru, hijau
    ['c', 's','i','d'],
    ['i', 's','c','d'],
    ['c', 'd','i','s'],
    ['d', 'c','s','i'],#
    ['i', 's','d','c'],
    ['i', 'c','d','s'],
    ['i', 'c','d','s'],
    ['d', 's','i','c']#
]

with open("data/question/disc.csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter = ',')
    line = 0
    for row in csv_reader:
        opts = []
        for col in row:
            #print(row[col])
            opts.append(col)
        quest.append(opts)
        #quest.append(opts)

#print(quest)

class MyApp(QWidget):
    def __init__(self, workbook):
        super(MyApp, self).__init__()
        self.setWindowTitle("Tes DISC")
        self.q_answer_true = []
        self.q_answer_false = []

        self.workbook = workbook[0]
        self.namedate = workbook[1]
        self.worksheet = self.workbook.add_worksheet("DISC")
        self.res = {
            'D':0,
            'I':0,
            'S':0,
            'C':0
        }

        self.choices = []
        '''window_width = 800
        window_height = 600
        self.setFixedSize(window_width, window_height)'''
        self.initUI()

    def createLayout_group(self, num):#question, num):
        sgroupbox = QGroupBox(str(num+1), self)
        self.font = QFont("Serif", 12)
        self.font.setBold(True)
        sgroupbox.setFont(self.font)
        layout_groupbox = QHBoxLayout(sgroupbox)
        self.q_choice = []
        q_ans = quest[num]
        for i in range(len(q_ans)):
            ans = QLineEdit()
            item = QLabel(q_ans[i])

            #if i ==0
            #item = QRadioButton(q_ans[i])
            #self.q_choice.append(item)
            ans.setMaxLength(1)
            ans.setMaximumWidth(20)
            self.q_choice.append(ans)
            layout_groupbox.addWidget(ans)#item change to layoutchoice
            layout_groupbox.addWidget(item)
        layout_groupbox.addStretch(1)
        self.choices.append(self.q_choice)

        return sgroupbox

    def createLayout_Container(self):
        self.scrollarea = QScrollArea(self)
        #self.scrollarea.setFixedWidth(780)
        self.scrollarea.setWidgetResizable(True)

        widget = QWidget()
        self.scrollarea.setWidget(widget)
        self.layout_SArea = QVBoxLayout(widget)

        #groupbox (jumlah soal)
        #num = 0
        for i in range(len(quest)):
            self.layout_SArea.addWidget((self.createLayout_group(i)))
        self.layout_SArea.addStretch(1)

    def initUI(self):
        self.createLayout_Container()
        self.layout_All = QVBoxLayout(self)
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

    def calculate(self, num, row, ans):
        if discCat[num][row]=='d':
            if ans == 'm':
                if num != 0 and num != 3 and num != 4 and num != 14:
                    self.res['D']+=1
                #exc 1,4,5,15
            elif ans == 'l':
                if num != 7 and num != 12 and num !=13:
                    self.res['D']-=1
                #exc 8, 13,14
        elif discCat[num][row]=='i':
            if ans == 'm':
                if num != 4 and num !=5 and num !=9 and num != 11 and num !=14 and num !=16 and num != 20:
                    self.res['I']+=1
                #exc 5,6, 10,12,15,17,21
            elif ans == 'l':
                if num != 0 and num != 7 and num != 14 and num != 15 and num != 17:
                    self.res['I'] -= 1
                #exc 1, 8, 15, 16, 18
        elif discCat[num][row]=='s':
            if ans == 'm':
                if num != 1 and num !=7 and num != 15 and num!=19 and num !=22:
                    self.res['S']+=1
                #exc 2, 8, 16, 20, 23
            elif ans == 'l':
                if num != 2 and num != 6 and num != 13 and num != 14 and num !=17:
                    self.res['S']-=1
                #exc 3, 7, 14, 15, 18
        elif discCat[num][row]=='c':
            if ans == 'm':
                if num != 2 and num != 5 and num != 6 and num != 7 and num != 10 and num !=12 and num != 15 and num != 17 and num !=21:
                    self.res['C']+=1
                #exc 3, 6, 7, 8, 11, 13, 16, 18, 22
            elif ans == 'l':
                if num != 3 and num != 8 and num != 9 and num !=11 and num != 16 and num !=18 and num != 19 and num != 22:
                    self.res['C']-=1
                #exc 4, 9, 10, 12, 17, 19, 20, 23
        #print(discCat[num][row]+' '+self.res.get(discCat[num][row].upper()))

    def on_click(self):
        self.enableClear=False

        for x in range(len(self.choices)):
            self.check = []
            self.l_num =0
            self.m_num =0
            self.empt = 0
            for y in self.choices[x]:
                self.check.append(y.text().lower())
            for z in self.check:
                if z =='l':
                    self.l_num+=1
                elif z == 'm':
                    self.m_num+=1
                else:
                    self.empt +=1
            if self.l_num==1 and self.m_num ==1 and self.empt ==2:
                self.enableClear=True
            else:
                self.enableClear=False
                break


        #print("saved")

        if self.enableClear:
            row = 0

            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.suitable = self.workbook.add_format({'bg_color': 'lime'})
            self.unsuitable = self.workbook.add_format({'bg_color': 'orange'})
            self.purp = self.workbook.add_format({'bg_color':'red'})
            self.cyan = self.workbook.add_format({'bg_color':'cyan'})
            #self.worksheet.write(0, 0, 'Paling Sesuai',self.header)
            #self.worksheet.write(0, 1, 'Paling Tidak Sesuai', self.header)
            row += 1

            #print("masuk")

            for x in range(len(self.choices)):
                #print(self.choices[x])
                col = 1
                for y in range(len(self.choices[x])):
                    #print(self.choices[x][y].text())
                    if discCat[x][y]=='d':
                        self.bg = self.purp
                    elif discCat[x][y] =='i':
                        self.bg = self.unsuitable
                    elif discCat[x][y] =='s':
                        self.bg = self.cyan
                    elif discCat[x][y] =='c':
                        self.bg = self.suitable

                    if self.choices[x][y].text().lower() == 'l':
                        self.worksheet.write(row, col, self.choices[x][y].text().upper(), self.bg)
                        self.calculate(x, y, 'l')
                    elif self.choices[x][y].text().lower() =='m':
                        self.worksheet.write(row, col, self.choices[x][y].text().upper(), self.bg)
                        self.calculate(x, y, 'm')
                    else :
                        self.worksheet.write(row, col, self.choices[x][y].text().upper(), self.bg)
                    col+=1
                    #self.worksheet.write(row, col, quest[x][y])
                    #col+=1
                row+=2

            #print(self.res)
            self.sorted_res = sorted(self.res.items(), key=lambda kv: kv[1], reverse=True)
            #self.sortedRes = collections.OrderedDict(self.sorted_res)
            #print(self.sorted_res)
            #print(self.sortedRes)
            self.worksheet.write(11, 6, "Result", self.header)
            #col = 7
            self.strRes = []
            for x in self.sorted_res:
                self.strRes.append(x[0])
            self.strRes = "".join(self.strRes)
            self.worksheet.write(11, 7, self.strRes)

            self.worksheet.write(0, 7, 'M', self.header)
            self.worksheet.write(0, 8, 'L', self.header)
            self.worksheet.write(0, 9, 'M-L', self.header)

            self.worksheet.write(1, 6, 'D', self.header)
            self.worksheet.write(1, 7, '=IF(D4="M",1,0)+IF(C6="M",1,0)+IF(B12="M",1,0)+IF(D14="M",1,0)+IF(B16="M",1,0)+IF(D18="M",1,0)+IF(B20="M",1,0)+IF(E22="M",1,0)+IF(C24="M",1,0)+IF(B26="M",1,0)+IF(C28="M",1,0)+IF(E32="M",1,0)+IF(E34="M",1,0)+IF(E36="M",1,0)+IF(C38="M",1,0)+IF(B40="M",1,0)+IF(D42="M",1,0)+IF(D44="M",1,0)+IF(D46="M",1,0)+IF(B48="M",1,0)')
            self.worksheet.write(1, 8, '=IF(E2="L",1,0)+IF(D4="L",1,0)+IF(C6="L",1,0)+IF(D8="L",1,0)+IF(D10="L",1,0)+IF(B12="L",1,0)+IF(D14="L",1,0)+IF(D18="L",1,0)+IF(B20="L",1,0)+IF(E22="L",1,0)+IF(C24="L",1,0)+IF(E30="L",1,0)+IF(E32="L",1,0)+IF(E34="L",1,0)+IF(E36="L",1,0)+IF(C38="L",1,0)+IF(B40="L",1,0)+IF(D42="L",1,0)+IF(D44="L",1,0)+IF(D46="L",1,0)+IF(B48="L",1,0)')
            self.worksheet.write(1, 9, '=+H2-I2')

            self.worksheet.write(3, 6, 'I', self.header)
            self.worksheet.write(3, 7, '=IF(C2="M",1,0)+IF(B4="M",1,0)+IF(E6="M",1,0)+IF(E8="M",1,0)+IF(E14="M",1,0)+IF(C16="M",1,0)+IF(B18="M",1,0)+IF(B22="M",1,0)+IF(C26="M",1,0)+IF(D28="M",1,0)+IF(B32="M",1,0)+IF(B36="M",1,0)+IF(D38="M",1,0)+IF(E40="M",1,0)+IF(B44="M",1,0)+IF(B46="M",1,0)+IF(D48="M",1,0)')
            self.worksheet.write(3, 8, '=IF(B4="L",1,0)+IF(E6="L",1,0)+IF(E8="L",1,0)+IF(B10="L",1,0)+IF(D12="L",1,0)+IF(E14="L",1,0)+IF(B18="L",1,0)+IF(D20="L",1,0)+IF(B22="L",1,0)+IF(B24="L",1,0)+IF(C26="L",1,0)+IF(D28="L",1,0)+IF(D34="L",1,0)+IF(D38="L",1,0)+IF(E40="L",1,0)+IF(B42="L",1,0)+IF(B44="L",1,0)+IF(B46="L",1,0)+IF(D48="L",1,0)')
            self.worksheet.write(3, 9, '=+H4-I4')

            self.worksheet.write(5, 6, 'S', self.header)
            self.worksheet.write(5,7, '=IF(B2="M",1,0)+IF(D6="M",1,0)+IF(C8="M",1,0)+IF(E10="M",1,0)+IF(C12="M",1,0)+IF(C14="M",1,0)+IF(C18="M",1,0)+IF(E20="M",1,0)+IF(C22="M",1,0)+IF(E24="M",1,0)+IF(D26="M",1,0)+IF(E28="M",1,0)+IF(B30="M",1,0)+IF(C34="M",1,0)+IF(C36="M",1,0)+IF(E38="M",1,0)+IF(B42="M",1,0)+IF(E44="M",1,0)+IF(C48="M",1,0)')
            self.worksheet.write(5,8, '=IF(B2="L",1,0)+IF(E4="L",1,0)+IF(C8="L",1,0)+IF(E10="L",1,0)+IF(C12="L",1,0)+IF(D16="L",1,0)+IF(C18="L",1,0)+IF(E20="L",1,0)+IF(C22="L",1,0)+IF(E24="L",1,0)+IF(D26="L",1,0)+IF(C32="L",1,0)+IF(C34="L",1,0)+IF(E38="L",1,0)+IF(D40="L",1,0)+IF(C42="L",1,0)+IF(E44="L",1,0)+IF(E46="L",1,0)+IF(C48="L",1,0)')
            self.worksheet.write(5, 9, '=+H6-I6')

            self.worksheet.write(7, 6, 'C', self.header)
            self.worksheet.write(7, 7, '=IF(D2="M",1,0)+IF(C4="M",1,0)+IF(B8="M",1,0)+IF(C10="M",1,0)+IF(E18="M",1,0)+IF(C20="M",1,0)+IF(D24="M",1,0)+IF(B28="M",1,0)+IF(D30="M",1,0)+IF(B34="M",1,0)+IF(B38="M",1,0)+IF(C40="M",1,0)+IF(E42="M",1,0)+IF(C46="M",1,0)+IF(E48="M",1,0)')
            self.worksheet.write(7, 8, '=IF(D2="L",1,0)+IF(C4="L",1,0)+IF(B6="L",1,0)+IF(C10="L",1,0)+IF(E12="L",1,0)+IF(B14="L",1,0)+IF(E16="L",1,0)+IF(D22="L",1,0)+IF(E26="L",1,0)+IF(B28="L",1,0)+IF(D30="L",1,0)+IF(D32="L",1,0)+IF(D36="L",1,0)+IF(E42="L",1,0)+IF(C44="L",1,0)+IF(E48="L",1,0)')
            self.worksheet.write(7, 9, '=+H8-I8')

            self.worksheet.write(8, 6, 'subtotal', self.header)
            self.worksheet.write(8, 7, '=SUM(H2:H8)')
            self.worksheet.write(8, 8, '=SUM(I2:I8)')
            self.worksheet.write(8, 9, '=SUM(J2:J8)')

            self.worksheet.write(9, 6, 'circle', self.header)
            self.worksheet.write(9, 7, '=IF(E2="M",1,0)+IF(E4="M",1,0)+IF(B6="M",1,0)+IF(D8="M",1,0)+IF(B10="M",1,0)+IF(D10="M",1,0)+IF(D12="M",1,0)+IF(E12="M",1,0)+IF(B14="M",1,0)+IF(D16="M",1,0)+IF(E16="M",1,0)+IF(D20="M",1,0)+IF(D22="M",1,0)+IF(B24="M",1,0)+IF(E26="M",1,0)+IF(C30="M",1,0)+IF(E30="M",1,0)+IF(C32="M",1,0)+IF(D32="M",1,0)+IF(D34="M",1,0)+IF(D36="M",1,0)+IF(D40="M",1,0)+IF(C42="M",1,0)+IF(C44="M",1,0)+IF(E46="M",1,0)')
            self.worksheet.write(9, 8, '=IF(C2="L",1,0)+IF(D6="L",1,0)+IF(B8="L",1,0)+IF(C14="L",1,0)+IF(B16="L",1,0)+IF(C16="L",1,0)+IF(E18="L",1,0)+IF(C20="L",1,0)+IF(D24="L",1,0)+IF(B26="L",1,0)+IF(C28="L",1,0)+IF(E28="L",1,0)+IF(B30="L",1,0)+IF(C30="L",1,0)+IF(B32="L",1,0)+IF(B34="L",1,0)+IF(B36="L",1,0)+IF(C36="L",1,0)+IF(B38="L",1,0)+IF(C40="L",1,0)+IF(C46="L",1,0)')
            self.worksheet.write(9, 9, '=SUM(H10:I10)')

            self.worksheet.write(10, 6, 'total', self.header)
            self.worksheet.write(10, 7, '=SUM(H9:H10)')
            self.worksheet.write(10, 8, '=SUM(I9:I10)')
            self.worksheet.write(10, 9, '=SUM(H11:I11)')

            self.close()
        else:
            self.buttonReply = QtWidgets.QMessageBox
            self.warning = self.buttonReply.question(self, 'PERINGATAN', "Tolong cek kembali kolom:\n"+
                                                     "1. Cek kembali kolom yang berisi duplikat\n"+
                                                     "2. Cek kembali apabila masih ada kolom yang belum terisi",
                                                     QtWidgets.QMessageBox.Ok)

'''if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    sys.exit(app.exec_())'''
from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore, QtGui
from math import ceil, floor
from subwindow import Ui_Dialog
import xlsxwriter
import datetime
import XLSDataInputTest

class Ui(QtWidgets.QMainWindow):
    def __init__(self, identity):
        super(Ui, self).__init__()
        uic.loadUi('ui/main.ui', self)
        self.setWindowTitle("Psikotes")
        self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
        self.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
        self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        #self.showMaximized()
        self.showFullScreen()
        #print(name)

        self.name = identity[0]
        self.date = identity[7]
        self.ident = identity
        self.AAresult=0

        self.tiu_res=0
        self.ist3_res=0
        self.ist5_res=0
        self.adkudag_res =0
        self.strRes = "Blm"

        #print(datetime.date.today())
        self.workname = 'data/result/'+self.name+'_'+str(datetime.date.today())+'.xlsx'
        self.workbook = xlsxwriter.Workbook(self.workname)
        self.workbook.close()

        self.discCat = [
            ['s', 'i', 'c', 'd'],
            ['i', 'c', 'd', 's'],
            ['c', 'd', 's', 'i'],
            ['c', 's', 'd', 'i'],  #
            ['i', 'c', 'd', 's'],
            ['d', 's', 'i', 'c'],
            ['c', 's', 'd', 'i'],
            ['d', 'i', 's', 'c'],  #
            ['i', 's', 'd', 'c'],
            ['d', 'c', 'i', 's'],
            ['i', 's', 'c', 'd'],
            ['i', 'd', 'c', 's'],  #
            ['d', 'i', 's', 'c'],
            ['c', 'd', 'i', 's'],
            ['s', 'i', 'c', 'd'],
            ['i', 's', 'c', 'd'],  # ungu, merah, biru, hijau
            ['c', 's', 'i', 'd'],
            ['i', 's', 'c', 'd'],
            ['c', 'd', 'i', 's'],
            ['d', 'c', 's', 'i'],  #
            ['i', 's', 'd', 'c'],
            ['i', 'c', 'd', 's'],
            ['i', 'c', 'd', 's'],
            ['d', 's', 'i', 'c']  #
        ]

        #self.show()

        self.pushButton.setText("Tes 1")
        self.pushButton.clicked.connect(lambda: self.on_click(1))

        self.pushButton_2.setText( "Tes 2")
        self.pushButton_2.clicked.connect(lambda: self.on_click(2))

        self.pushButton_3.setText( "Tes 3")
        self.pushButton_3.clicked.connect(lambda: self.on_click(3))

        self.pushButton_4.setText( "Tes 4")
        self.pushButton_4.clicked.connect(lambda: self.on_click(4))

        self.pushButton_5.setText( "Tes 5")
        self.pushButton_5.clicked.connect(lambda: self.on_click(5))

        self.pushButton_6.setText("Tes 6")
        self.pushButton_6.clicked.connect(lambda: self.on_click(6))

        self.pushButton_7.setText( "Save and Exit")
        self.pushButton_7.clicked.connect(self.saveandquit)

    def autosave(self, ans_tiu = None, ans_ist3 = None, ans_ist5 = None, ans_adkudag = None, ans_aa = None, ans_disc = None):
        #CREATE WORKBOOK
        try:
            self.workbook = xlsxwriter.Workbook(self.workname)

            self.worksheet = self.workbook.add_worksheet("Identitas")
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.worksheet.write(0, 0, "Nama", self.header)
            self.worksheet.write(1, 0, "Usia", self.header)
            self.worksheet.write(2, 0, "Pendidikan Terakhir", self.header)
            self.worksheet.write(3, 0, "Prodi/Jurusan", self.header)
            self.worksheet.write(4, 0, "Universitas", self.header)
            self.worksheet.write(5, 0, "Posisi Dilamar", self.header)
            self.worksheet.write(6, 0, "No Telepon", self.header)
            self.worksheet.write(7, 0, "Tanggal Tes", self.header)

            self.row = 0

            for x in self.ident:
                if self.row == 6:
                    self.worksheet.write(self.row, 1, "'" + x, self.fill)
                elif self.row == 7:
                    self.worksheet.write(self.row, 1, str(x), self.fill)
                else:
                    self.worksheet.write(self.row, 1, x, self.fill)

                self.row += 1

        except Exception as e:
            print(e)

        #SAVE SHEET TIU
        try:
            self.worksheet = self.workbook.add_worksheet("TIU")
            self.key_tiu = ['B', 'E', 'D', 'D', 'E',
                            'D', 'E', 'A', 'E', 'C',
                            'E', 'D', 'E', 'C', 'B',
                            'C', 'B', 'C', 'D', 'D',
                            'C', 'C', 'D', 'D', 'E',
                            'A', 'B', 'E', 'C', 'E']
            if ans_tiu != None:
                self.ans_tiu = ans_tiu
            self.res = []

            for i in range(len(self.ans_tiu)):
                if self.ans_tiu[i] == self.key_tiu[i]:
                    self.res.append(1)
                else:
                    self.res.append(0)

            row = 0
            col = 0
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.desired = self.workbook.add_format({'bg_color': 'cyan'})
            self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
            self.worksheet.write(0, 1, 'Jawaban', self.header)
            self.worksheet.write(0, 2, 'Skor', self.header)
            row += 1
            for x in range(len(self.key_tiu)):
                self.worksheet.write(row, col, self.key_tiu[x], self.desired)
                self.worksheet.write(row, col + 1, self.ans_tiu[x], self.fill)
                self.worksheet.write(row, col + 2, self.res[x], self.fill)
                row += 1

            self.worksheet.write(row, 1, "Total", self.header)
            self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)
            # self.workbook.close()

            self.worksheet.write(1, 5, 'RS', self.header)
            self.worksheet.write(1, 6, '=C' + str(row + 1), self.fill)
            self.worksheet.write(2, 5, 'WS', self.header)
            self.worksheet.write(2, 6, '=IF(G2<=1,2,IF(G2<=5,4,IF(G2<=7,5,IF(G2<=9,6,IF(G2<=11,7,IF(G2<=13,8,IF(G2<=15,9,' +
                                 'IF(G2<=17,10,IF(G2<=19,11,IF(G2<=21,12,IF(G2<=23,13,IF(G2<=25,14,IF(G3<=27,15,IF' +
                                 '(G2<=29,16,17))))))))))))))', self.fill)
            self.worksheet.write(2, 7, '=IF(G3<=2,1, IF(G3<=6,2,IF(G3<=11,3,IF(G3<=15,4,5))))')
            self.worksheet.write(3, 5, 'Scale', self.header)
            self.worksheet.write_formula(3, 6,
                                         '=IF(G3<=6,"(1) Rendah",IF(G3<=12,"(2) Rata-rata Bawah",IF(G3<=18, "(3) Sedang", IF(G3<=24, "(4) Rata-rata atas", "(5) Tinggi"))))',
                                         self.fill)

            self.tiu_res = ceil(sum(self.res)+3/2)

        except Exception as e:
            print("Save TIU gagal")
            print(e)

        #SAVE SHEET IST3
        try:
            self.worksheet = self.workbook.add_worksheet("IST3")
            self.key_ist3 = ['C', 'E', 'D', 'D', 'D',
                             'B', 'D', 'B', 'E', 'D',
                             'C', 'C', 'C', 'C', 'D',
                             'C', 'C', 'E', 'E', 'E']

            self.res = []
            if ans_ist3!= None:
                self.ans_ist3 = ans_ist3
            for i in range(len(self.ans_ist3)):
                if self.ans_ist3[i] == self.key_ist3[i]:
                    self.res.append(1)
                else:
                    self.res.append(0)
            # print(self.res)
            # print(sum(self.res))

            row = 0
            col = 0
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.desired = self.workbook.add_format({'bg_color': 'cyan'})
            self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
            self.worksheet.write(0, 1, 'Jawaban', self.header)
            self.worksheet.write(0, 2, 'Skor', self.header)
            row += 1
            for x in range(len(self.key_ist3)):
                self.worksheet.write(row, col, self.key_ist3[x], self.desired)
                self.worksheet.write(row, col + 1, self.ans_ist3[x], self.fill)
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

            self.ist3_res = (sum(self.res)*8)+44
            if self.ist3_res <= 78:
                self.ist3_res = 1
            elif self.ist3_res <= 107:
                self.ist3_res = 2
            elif self.ist3_res <=136:
                self.ist3_res = 3
            elif self.ist3_res <= 165:
                self.ist3_res = 4
            else:
                self.ist3_res = 5
        except Exception as e:
            print(e)

        #SAVE SHEET IST5
        try:
            self.worksheet = self.workbook.add_worksheet("IST5")
            if ans_ist5!=None:
                self.ans_ist5 = ans_ist5
            self.key_ist5 = [{'3', '5'}, {'2', '8', '0'}, {'2', '5', '0'}, {'2', '6'}, {'3', '0'},
                         {'7', '0'}, {'4', '5'}, {'5', '0'}, {'4', '8'}, {'7', '8'},
                         {'1', '9'}, {'6'}, {'5', '7'}, {'9', '0'}, {'1', '0', '2'},
                         {'1', '7'}, {'2', '4'}, {'5'}, {'4', '8'}, {'3'}]

            self.res = []
            self.n = 0
            for x in self.ans_ist5:
                if self.key_ist5[self.n] == x:
                    self.res.append(1)
                else:
                    self.res.append(0)
                self.n += 1

            row = 0
            col = 0
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.desired = self.workbook.add_format({'bg_color': 'cyan'})
            self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
            self.worksheet.write(0, 1, 'Jawaban', self.header)
            self.worksheet.write(0, 2, 'Skor', self.header)
            row += 1
            for x in range(len(self.key_ist5)):
                self.worksheet.write(row, col, str(self.key_ist5[x]), self.desired)
                self.worksheet.write(row, col + 1, str(self.ans_ist5[x]), self.fill)
                self.worksheet.write(row, col + 2, self.res[x], self.fill)
                row += 1

            self.worksheet.write(row, 1, "Total", self.header)
            self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)

            self.worksheet.write(1, 5, 'RS', self.header)
            self.worksheet.write(1, 6, '=C' + str(row + 1), self.fill)
            self.worksheet.write(2, 5, 'WS', self.header)
            self.worksheet.write(2, 6, '=(G2*5)+60', self.fill)
            self.worksheet.write(2, 7, "=IF(G3<=82,1,IF(G3<=99,2,IF(G3<=115,3,IF(G3<=132,4,5))))")
            self.worksheet.write(3, 5, 'Scale', self.header)
            self.worksheet.write_formula(3, 6,
                                         '=IF(G3<=82,"(1) Rendah",IF(G3<=99,"(2) Rata-rata Bawah",IF(G3<=115, "(3) Sedang", IF(G3<=132, "(4) Rata-rata atas", "(5) Tinggi"))))',
                                         self.fill)
            self.ist5_res = (sum(self.res)*5)+60
            if self.ist5_res<=82:
                self.ist5_res = 1
            elif self.ist5_res<=99:
                self.ist5_res = 2
            elif self.ist5_res<115:
                self.ist5_res = 3
            elif self.ist5_res<132:
                self.ist5_res = 4
            else:
                self.ist5_res = 5
        except Exception as e:
            print(e)

        #SAVE SHEET ADKUDAG
        try:
            self.worksheet = self.workbook.add_worksheet("ADKUDAG")

            self.key_adkudag = ['s', 's', 'b', 'b', 's', 's', 'b', 's', 'b', 'b',
                             's', 'b', 's', 's', 's', 'b', 's', 's', 's', 's',
                             'b', 'b', 'b', 'b', 's', 's', 's', 'b', 'b', 'b',
                             'b', 'b', 'b', 'b', 'b', 'b', 'b', 'b', 's', 'b',
                             's', 'b', 's', 'b', 'b', 'b', 'b', 's', 's', 'b',
                             'b', 'b', 'b', 'b', 'b', 's', 's', 'b', 'b', 's',
                             's', 'b', 'b', 'b', 'b', 's', 's', 'b', 's', 's',
                             's', 'b', 'b', 's', 'b', 's', 'b', 's', 's', 's',
                             's', 's', 'b', 'b', 'b', 's', 'b', 's', 's', 'b',
                             's', 's', 's', 's', 'b', 's', 's', 's', 's', 's',
                             's', 'b', 's', 'b', 's', 'b', 's', 's', 'b', 'b',
                             'b', 's', 's', 's', 'b', 'b', 's', 's', 'b', 'b',
                             's', 'b', 's', 'b', 'b', 'b', 's', 's', 's', 'b',
                             'b', 'b', 's', 'b', 's', 'b', 's', 'b', 's', 's',
                             'b', 'b', 'b', 's', 's', 'b', 'b', 'b', 'b', 'b']

            if ans_adkudag != None:
                self.ans_adkudag = ans_adkudag

            self.res = []
            for x in range(len(self.ans_adkudag)):
                if self.ans_adkudag[x].lower() == self.key_adkudag[x]:
                    self.res.append(1)
                else:
                    self.res.append(0)
            # print(self.res)

            row = 0
            col = 0
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.desired = self.workbook.add_format({'bg_color': 'cyan'})
            self.percentage = self.workbook.add_format({'num_format': '0.00%', 'bg_color': 'lime'})
            self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
            self.worksheet.write(0, 1, 'Jawaban', self.header)
            self.worksheet.write(0, 2, 'Skor', self.header)
            self.worksheet.write(0, 3, 'Percentile', self.header)
            row += 1
            for x in range(len(self.key_adkudag)):
                self.worksheet.write(row, col, self.key_adkudag[x], self.desired)
                self.worksheet.write(row, col + 1, self.ans_adkudag[x], self.fill)
                self.worksheet.write(row, col + 2, self.res[x], self.fill)
                row += 1

            self.worksheet.write(row, 1, "Total", self.header)
            self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)
            self.worksheet.write(1, 3, "=(C" + str(row + 1) + "/150)", self.percentage)
            self.worksheet.write(1, 5, 'Scale', self.header)
            self.worksheet.write_formula(1, 6,
                                         '=IF(D2<60%,"Kurang",IF(D2<81%,"Cukup",IF(D2<101%,"Baik")))',
                                         self.fill)
            self.adkudag_res = sum(self.res)
        except Exception as e:
            print(e)

        #SAVE SHEET AA
        try:
            self.worksheet = self.workbook.add_worksheet("AA")
            if ans_aa!=None:
                self.ans_aa = ans_aa

            row = 0
            col = 0
            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.worksheet.write(0, 0, 'Kunci Jawaban', self.header)
            self.worksheet.write(0, 1, 'Jawaban', self.header)
            self.worksheet.write(0, 2, 'Skor', self.header)
            row += 1
            for x in range(len(self.ans_aa)):
                self.worksheet.write(row, col + 2, self.ans_aa[x], self.fill)
                row += 1

            self.worksheet.write(row, 1, "Total", self.header)
            self.worksheet.write(row, 2, "=SUM(C2:C" + str(row) + ")", self.fill)
            # save mechanism

            self.worksheet.write(1, 5, 'RS', self.header)
            self.worksheet.write(1, 6, '=C' + str(row + 1), self.fill)
            self.worksheet.write(2, 5, 'Scale', self.header)
            self.worksheet.write_formula(2, 6,
                                         '=IF(G2<6,"Kurang",IF(G2<9,"Cukup",IF(G2<13,"Baik")))',
                                         self.fill)
        except Exception as e:
            print(e)

        #SAVE SHEET DISC
        try:
            if ans_disc != None:
                self.ans_disc = ans_disc

            self.worksheet = self.workbook.add_worksheet("DISC")
            self.res = {
                'D': 0,
                'I': 0,
                'S': 0,
                'C': 0
            }
            row = 0

            self.header = self.workbook.add_format({'bold': True})
            self.fill = self.workbook.add_format({'bg_color': 'lime'})
            self.suitable = self.workbook.add_format({'bg_color': 'lime'})
            self.unsuitable = self.workbook.add_format({'bg_color': 'orange'})
            self.purp = self.workbook.add_format({'bg_color': 'red'})
            self.cyan = self.workbook.add_format({'bg_color': 'cyan'})

            row += 1

            # print("masuk")

            for x in range(len(self.ans_disc)):
                # print(self.choices[x])
                col = 1
                for y in range(len(self.ans_disc[x])):
                    # print(self.choices[x][y].text())
                    if self.discCat[x][y] == 'd':
                        self.bg = self.purp
                    elif self.discCat[x][y] == 'i':
                        self.bg = self.unsuitable
                    elif self.discCat[x][y] == 's':
                        self.bg = self.cyan
                    elif self.discCat[x][y] == 'c':
                        self.bg = self.suitable

                    if self.ans_disc[x][y].text().lower() == 'l':
                        self.worksheet.write(row, col, self.ans_disc[x][y].text().upper(), self.bg)
                        self.calculate(x, y, 'l')
                    elif self.ans_disc[x][y].text().lower() == 'm':
                        self.worksheet.write(row, col, self.ans_disc[x][y].text().upper(), self.bg)
                        self.calculate(x, y, 'm')
                    else:
                        self.worksheet.write(row, col, self.ans_disc[x][y].text().upper(), self.bg)
                    col += 1
                row += 2

            self.sorted_res = sorted(self.res.items(), key=lambda kv: kv[1], reverse=True)
            self.worksheet.write(11, 6, "Result", self.header)

            self.strRes = []
            for x in self.sorted_res:
                self.strRes.append(x[0])
            self.strRes = "".join(self.strRes)
            self.worksheet.write(11, 7, self.strRes)

            self.worksheet.write(0, 7, 'M', self.header)
            self.worksheet.write(0, 8, 'L', self.header)
            self.worksheet.write(0, 9, 'M-L', self.header)

            self.worksheet.write(1, 6, 'D', self.header)
            self.worksheet.write(1, 7,
                                 '=IF(D4="M",1,0)+IF(C6="M",1,0)+IF(B12="M",1,0)+IF(D14="M",1,0)+IF(B16="M",1,0)+IF(D18="M",1,0)+IF(B20="M",1,0)+IF(E22="M",1,0)+IF(C24="M",1,0)+IF(B26="M",1,0)+IF(C28="M",1,0)+IF(E32="M",1,0)+IF(E34="M",1,0)+IF(E36="M",1,0)+IF(C38="M",1,0)+IF(B40="M",1,0)+IF(D42="M",1,0)+IF(D44="M",1,0)+IF(D46="M",1,0)+IF(B48="M",1,0)')
            self.worksheet.write(1, 8,
                                 '=IF(E2="L",1,0)+IF(D4="L",1,0)+IF(C6="L",1,0)+IF(D8="L",1,0)+IF(D10="L",1,0)+IF(B12="L",1,0)+IF(D14="L",1,0)+IF(D18="L",1,0)+IF(B20="L",1,0)+IF(E22="L",1,0)+IF(C24="L",1,0)+IF(E30="L",1,0)+IF(E32="L",1,0)+IF(E34="L",1,0)+IF(E36="L",1,0)+IF(C38="L",1,0)+IF(B40="L",1,0)+IF(D42="L",1,0)+IF(D44="L",1,0)+IF(D46="L",1,0)+IF(B48="L",1,0)')
            self.worksheet.write(1, 9, '=+H2-I2')

            self.worksheet.write(3, 6, 'I', self.header)
            self.worksheet.write(3, 7,
                                 '=IF(C2="M",1,0)+IF(B4="M",1,0)+IF(E6="M",1,0)+IF(E8="M",1,0)+IF(E14="M",1,0)+IF(C16="M",1,0)+IF(B18="M",1,0)+IF(B22="M",1,0)+IF(C26="M",1,0)+IF(D28="M",1,0)+IF(B32="M",1,0)+IF(B36="M",1,0)+IF(D38="M",1,0)+IF(E40="M",1,0)+IF(B44="M",1,0)+IF(B46="M",1,0)+IF(D48="M",1,0)')
            self.worksheet.write(3, 8,
                                 '=IF(B4="L",1,0)+IF(E6="L",1,0)+IF(E8="L",1,0)+IF(B10="L",1,0)+IF(D12="L",1,0)+IF(E14="L",1,0)+IF(B18="L",1,0)+IF(D20="L",1,0)+IF(B22="L",1,0)+IF(B24="L",1,0)+IF(C26="L",1,0)+IF(D28="L",1,0)+IF(D34="L",1,0)+IF(D38="L",1,0)+IF(E40="L",1,0)+IF(B42="L",1,0)+IF(B44="L",1,0)+IF(B46="L",1,0)+IF(D48="L",1,0)')
            self.worksheet.write(3, 9, '=+H4-I4')

            self.worksheet.write(5, 6, 'S', self.header)
            self.worksheet.write(5, 7,
                                 '=IF(B2="M",1,0)+IF(D6="M",1,0)+IF(C8="M",1,0)+IF(E10="M",1,0)+IF(C12="M",1,0)+IF(C14="M",1,0)+IF(C18="M",1,0)+IF(E20="M",1,0)+IF(C22="M",1,0)+IF(E24="M",1,0)+IF(D26="M",1,0)+IF(E28="M",1,0)+IF(B30="M",1,0)+IF(C34="M",1,0)+IF(C36="M",1,0)+IF(E38="M",1,0)+IF(B42="M",1,0)+IF(E44="M",1,0)+IF(C48="M",1,0)')
            self.worksheet.write(5, 8,
                                 '=IF(B2="L",1,0)+IF(E4="L",1,0)+IF(C8="L",1,0)+IF(E10="L",1,0)+IF(C12="L",1,0)+IF(D16="L",1,0)+IF(C18="L",1,0)+IF(E20="L",1,0)+IF(C22="L",1,0)+IF(E24="L",1,0)+IF(D26="L",1,0)+IF(C32="L",1,0)+IF(C34="L",1,0)+IF(E38="L",1,0)+IF(D40="L",1,0)+IF(C42="L",1,0)+IF(E44="L",1,0)+IF(E46="L",1,0)+IF(C48="L",1,0)')
            self.worksheet.write(5, 9, '=+H6-I6')

            self.worksheet.write(7, 6, 'C', self.header)
            self.worksheet.write(7, 7,
                                 '=IF(D2="M",1,0)+IF(C4="M",1,0)+IF(B8="M",1,0)+IF(C10="M",1,0)+IF(E18="M",1,0)+IF(C20="M",1,0)+IF(D24="M",1,0)+IF(B28="M",1,0)+IF(D30="M",1,0)+IF(B34="M",1,0)+IF(B38="M",1,0)+IF(C40="M",1,0)+IF(E42="M",1,0)+IF(C46="M",1,0)+IF(E48="M",1,0)')
            self.worksheet.write(7, 8,
                                 '=IF(D2="L",1,0)+IF(C4="L",1,0)+IF(B6="L",1,0)+IF(C10="L",1,0)+IF(E12="L",1,0)+IF(B14="L",1,0)+IF(E16="L",1,0)+IF(D22="L",1,0)+IF(E26="L",1,0)+IF(B28="L",1,0)+IF(D30="L",1,0)+IF(D32="L",1,0)+IF(D36="L",1,0)+IF(E42="L",1,0)+IF(C44="L",1,0)+IF(E48="L",1,0)')
            self.worksheet.write(7, 9, '=+H8-I8')

            self.worksheet.write(8, 6, 'subtotal', self.header)
            self.worksheet.write(8, 7, '=SUM(H2:H8)')
            self.worksheet.write(8, 8, '=SUM(I2:I8)')
            self.worksheet.write(8, 9, '=SUM(J2:J8)')

            self.worksheet.write(9, 6, 'circle', self.header)
            self.worksheet.write(9, 7,
                                 '=IF(E2="M",1,0)+IF(E4="M",1,0)+IF(B6="M",1,0)+IF(D8="M",1,0)+IF(B10="M",1,0)+IF(D10="M",1,0)+IF(D12="M",1,0)+IF(E12="M",1,0)+IF(B14="M",1,0)+IF(D16="M",1,0)+IF(E16="M",1,0)+IF(D20="M",1,0)+IF(D22="M",1,0)+IF(B24="M",1,0)+IF(E26="M",1,0)+IF(C30="M",1,0)+IF(E30="M",1,0)+IF(C32="M",1,0)+IF(D32="M",1,0)+IF(D34="M",1,0)+IF(D36="M",1,0)+IF(D40="M",1,0)+IF(C42="M",1,0)+IF(C44="M",1,0)+IF(E46="M",1,0)')
            self.worksheet.write(9, 8,
                                 '=IF(C2="L",1,0)+IF(D6="L",1,0)+IF(B8="L",1,0)+IF(C14="L",1,0)+IF(B16="L",1,0)+IF(C16="L",1,0)+IF(E18="L",1,0)+IF(C20="L",1,0)+IF(D24="L",1,0)+IF(B26="L",1,0)+IF(C28="L",1,0)+IF(E28="L",1,0)+IF(B30="L",1,0)+IF(C30="L",1,0)+IF(B32="L",1,0)+IF(B34="L",1,0)+IF(B36="L",1,0)+IF(C36="L",1,0)+IF(B38="L",1,0)+IF(C40="L",1,0)+IF(C46="L",1,0)')
            self.worksheet.write(9, 9, '=SUM(H10:I10)')

            self.worksheet.write(10, 6, 'total', self.header)
            self.worksheet.write(10, 7, '=SUM(H9:H10)')
            self.worksheet.write(10, 8, '=SUM(I9:I10)')
            self.worksheet.write(10, 9, '=SUM(H11:I11)')
        except Exception as e:
            print(e)

        #CREATE REKAP
        try:
            collist = ['No', 'Nama', 'TIU', 'AN', 'RA', 'KETELITIAN',
                       'AA', 'DISC', 'POSISI', 'USIA', 'PENDIDIKAN', 'NO. HP/TLP',
                       'RS', 'WS']
            mergelist = ['KETERANGAN', 'POTENSI KECERDASAN (2,6-3)', 'KETELITIAN (60-80)',
                         'DAYA TANGKAP & KONSENTRASI (6-8)',
                         'JOB-PROFILE MATCH', 'REKOMENDASI', 'POSISI ALTERNATIF']

            self.mergeformat = self.workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': 1
            })
            self.border = self.workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            self.rekap = self.workbook.add_worksheet("REKAP")
            colnum = 0
            for col in collist:
                self.rekap.write(3, colnum, col, self.mergeformat)
                colnum += 1

            self.rekap.merge_range('M2:S2', mergelist[0], self.mergeformat)
            self.rekap.merge_range('M3:N3', mergelist[1], self.mergeformat)
            self.rekap.merge_range('O3:O4', mergelist[2], self.mergeformat)
            self.rekap.merge_range('P3:P4', mergelist[3], self.mergeformat)
            self.rekap.merge_range('Q3:Q4', mergelist[4], self.mergeformat)
            self.rekap.merge_range('S3:S4', mergelist[5], self.mergeformat)
            self.rekap.merge_range('R3:R4', mergelist[6], self.mergeformat)

            self.rekap.write(4, 1, self.name, self.border)
            self.rekap.write(4, 2, self.tiu_res, self.border)
            self.rekap.write(4, 3, self.ist3_res, self.border)
            self.rekap.write(4, 4, self.ist5_res, self.border)
            self.rekap.write(4, 5, floor(self.adkudag_res/150*100)/100, self.border)
            self.rekap.write(4, 6, self.AAresult, self.border)
            self.rekap.write(4, 7, self.strRes, self.border)
            self.rekap.write(4, 8, self.ident[5], self.border)
            self.rekap.write(4, 9, self.ident[1], self.border)
            self.rekap.write(4, 10, self.ident[2], self.border)
            self.rekap.write(4, 11, self.ident[6], self.border)
            self.rekap.write(4, 12, floor((self.tiu_res+self.ist3_res+self.ist5_res)/3*100)/100, self.border)
            if (self.tiu_res+self.ist3_res+self.ist5_res)/3 < 2:
                self.pot_ws = "Sgt Kurang"
            elif (self.tiu_res+self.ist3_res+self.ist5_res)/3 < 2.6:
                self.pot_ws = "Kurang"
            elif (self.tiu_res+self.ist3_res+self.ist5_res)/3 <3.1:
                self.pot_ws = "Cukup"
            else:
                self.pot_ws = "Baik"
            self.rekap.write(4, 13, self.pot_ws, self.border)

            if self.adkudag_res/150 < 0.6:
                self.ketelitian = "Kurang"
            elif self.adkudag_res/150 < 0.81:
                self.ketelitian = "Cukup"
            else:
                self.ketelitian = "Baik"
            self.rekap.write(4, 14, self.ketelitian, self.border)

            if self.AAresult < 6:
                self.konsentrasi = "Kurang"
            elif self.AAresult < 9:
                self.konsentrasi = "Cukup"
            else:
                self.konsentrasi = "Baik"
            self.rekap.write(4, 15, self.konsentrasi, self.border)
            self.rekap.write(4, 16, '', self.border)
            self.rekap.write(4, 17, '', self.border)
            self.rekap.write(4, 18, '', self.border)
        except Exception as e:
            print(e)

        self.workbook.close()

    def calculate(self, num, row, ans):
        if self.discCat[num][row] == 'd':
            if ans == 'm':
                if num != 0 and num != 3 and num != 4 and num != 14:
                    self.res['D'] += 1
                # exc 1,4,5,15
            elif ans == 'l':
                if num != 7 and num != 12 and num != 13:
                    self.res['D'] -= 1
                # exc 8, 13,14
        elif self.discCat[num][row] == 'i':
            if ans == 'm':
                if num != 4 and num != 5 and num != 9 and num != 11 and num != 14 and num != 16 and num != 20:
                    self.res['I'] += 1
                # exc 5,6, 10,12,15,17,21
            elif ans == 'l':
                if num != 0 and num != 7 and num != 14 and num != 15 and num != 17:
                    self.res['I'] -= 1
                # exc 1, 8, 15, 16, 18
        elif self.discCat[num][row] == 's':
            if ans == 'm':
                if num != 1 and num != 7 and num != 15 and num != 19 and num != 22:
                    self.res['S'] += 1
                # exc 2, 8, 16, 20, 23
            elif ans == 'l':
                if num != 2 and num != 6 and num != 13 and num != 14 and num != 17:
                    self.res['S'] -= 1
                # exc 3, 7, 14, 15, 18
        elif self.discCat[num][row] == 'c':
            if ans == 'm':
                if num != 2 and num != 5 and num != 6 and num != 7 and num != 10 and num != 12 and num != 15 and num != 17 and num != 21:
                    self.res['C'] += 1
                # exc 3, 6, 7, 8, 11, 13, 16, 18, 22
            elif ans == 'l':
                if num != 3 and num != 8 and num != 9 and num != 11 and num != 16 and num != 18 and num != 19 and num != 22:
                    self.res['C'] -= 1

    def on_click(self, _num):
        if _num ==1 :
            self.listdat = ["Tes TIU", "8 menit", "Bila A dikecilkan diperoleh B. Bila sekarang dengan C dilakukan hal yang \n"+
                                                  "serupa, jadi C dikecilkan, diperoleh gambar 2. maka dari itu gambar 2\n"+
                                                  "dicoret di bawahnya."] #TIU 8 mnt
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(True)
        elif _num == 2:
            self.listdat = ["Tes IST3", "5 menit", "Ditentukan tiga kata. \n"+
                                                   "Antara kata pertama dan kedua terdapat suatu hubungan tertentu.\n"+
                                                   "Antara kata ketiga dan salah satu kata di antara lima kata pilihan harus pula\n"+
                                                   "terdapat hubungan yang sama itu.\n"+
                                                   "Carilah kata itu.\n\n"+
                                                   "Hubungan antara hutan dan pohon ialah bahwa hutan terdiri atas pohon-pohon, \n"+
                                                   "maka hubungan antara tembok dan salah satu kata pilihan ialah bahwa tembok \n"+
                                                   "terdiri atas batu-batu bata. Oleh karena itu, pada lembar jawaban di belakang\n"+
                                                   "contoh 03, huruf a harus dicoret."] #IST3 5 mnt
            self.pushButton_2.setEnabled(False)
            self.pushButton_3.setEnabled(True)
        elif _num == 3:
            self.listdat = ["Tes IST5", "10 menit", "Kolom ini terdiri atas angka-angka 1 sampai 9 dan 0. \n"+
                                                    "Untuk menunjukkan jawaban suatu soal, maka pilihlah angka-angka yang\n"
                                                    "terdapat"+" di jawaban itu.\n"+
                                                    "Keurutan angka jawaban tidak perlu dihiraukan.\n\n"+
                                                    "Pada contoh, jawaban ialah 75. \n"+
                                                    "Oleh karena itu, pada lembar jawaban, angka 7 dan 5 harus dicoret."] #IST5 10 mnt
            self.pushButton_3.setEnabled(False)
            self.pushButton_4.setEnabled(True)
        elif _num == 4:
            self.listdat = ["Tes B-S", "5 menit", "Anda akan dihadapkan pada serangkaian kombinasi angka, huruf, dan tanda baca\n"+
                                                      "tertentu. Tugas anda adalah menentukan apakah kombinasi di sebelah kiri sama\n"+
                                                      "persis dengan kombinasi di sebelah kanannya.\n\n"+
                                                      "Berilah coretan miring satu kali pada pilihan jawaban Anda.\n"+
                                                      "B = Berbeda\n"+
                                                      "S = Sama"] #ADKUDAG 5 mnt
            self.pushButton_4.setEnabled(False)
            self.pushButton_5.setEnabled(True)
        elif _num == 5:
            self.listdat = ["Tes Army Alpha", "", ""]
            '''"Pada bagian tes ini saudara jumpai dua belas soal. Tiap-tiap soal akan dibacakan\n"+
                                  "sebuah instruksi. Setiap instruksi hanya dibaca satu kali saja. Saudara baru akan\n"+
                                  "diijinkan menulis apabila instruksi selesai dibacakan dan saya memberi tanda YA. \n"+
                                  "Apabila ada tanda STOP, harap segera berhenti menulis dan siap untuk soal \n"+
                                  "berikutnya. Jika saya belum memberi tanda STOP, saudara dapat mengisi atau \n"+
                                  "memperbaiki soal yang salah atau belum diisi."]''' #army alpha - tiap soal
            self.pushButton_5.setEnabled(False)
            self.pushButton_6.setEnabled(True)
        elif _num == 6:
            self.listdat = ["Tes DISC", "", "Dalam persoalan ini terdapat 24 nomor soal. Dimana setiap nomor soal memiliki \n"+
                                            "4 pernyataan. Pada setiap nomor, dari ke empat pernyataan tersebut pilihlah \n"+
                                            "salah satu pernyataan yang paling menggambarkan diri saudara lalu tuliskan \n"+
                                            "huruf M pada kotaknya. Lalu pilih lagi salah satu pernyataan yg paling tidak \n"+
                                            "menggambarkan diri saudara lalu tuliskan huruf L pada kotaknya. Jadi, pada \n"+
                                            "setiap nomor soal saudara hanya diperbolehkan menuliskan satu huruf M dan \n"+
                                            "satu huruf L."]
                                             #DISC - no limit
            self.pushButton_6.setEnabled(False)

        self.listdat.append(self.name)
        self.listdat.append(self.date)
        self.listdat.append(self)

        self.window = QtWidgets.QDialog()
        if _num != 5:
            self.window.ui = Ui_Dialog(self.listdat, self.workbook, self)
        else:
            self.window.ui = Ui_Dialog(self.listdat, self.workbook)
        self.window.ui.setupUi(self.window)
        self.window.showFullScreen()

    def setAAResult(self, _num):
        self.AAresult = _num

    def saveandquit(self):
        try:
            XLSDataInputTest.inst_db(self.workname)
        except Exception as e:
            print(str(e))
        print("Saving and closing...")

        QtWidgets.qApp.quit()
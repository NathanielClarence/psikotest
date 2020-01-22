from PyQt5 import QtWidgets, uic
from datetime import date
import mysql.connector as conn
import xlsxwriter

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('ui/exportUI.ui', self)

        self.initUI()
        self.show()


    def initUI(self):
        f = open("data/targetDB", 'r')

        self.mydb = conn.connect(
            host=f.readline(),
            user="root",
            passwd="root",
            database="psikotest",
            auth_plugin='mysql_native_password'
        )

        self.mycursor = self.mydb.cursor()

        self.setFixedSize(self.width(), self.height())
        self.dt_from.setDate(date.today())
        self.dt_until.setDate(date.today())
        self.dt_from.setMaximumDate(date.today())
        self.dt_until.setMaximumDate(date.today())

        self.btn_export.clicked.connect(self.export)

    def export(self):
        try:
            self.workbook = xlsxwriter.Workbook("data/result/Export_Rekap_"+str(date.today())+".xlsx")
            '''print(self.dt_from.date().toPyDate())
            print(self.dt_until.date().toPyDate())'''

            self.daterange = (str(self.dt_from.date().toPyDate()), str(self.dt_until.date().toPyDate()))
            self.query = "SELECT * FROM rekap_hasil where tanggal_tes between %s and %s;"
            self.mycursor.execute(self.query, self.daterange)
            self.result = self.mycursor.fetchall()
            #print(self.result)

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

                self.currRow = 4
                for x in self.result:
                    self.rekap.write(self.currRow, 0, x[0], self.border)
                    self.rekap.write(self.currRow, 1, x[1], self.border)
                    self.rekap.write(self.currRow, 2, x[2], self.border)
                    self.rekap.write(self.currRow, 3, x[3], self.border)
                    self.rekap.write(self.currRow, 4, x[4], self.border)
                    self.rekap.write(self.currRow, 5, x[5], self.border)
                    self.rekap.write(self.currRow, 6, x[6], self.border)
                    self.rekap.write(self.currRow, 7, x[7], self.border)
                    self.rekap.write(self.currRow, 8, x[8], self.border)
                    self.rekap.write(self.currRow, 9, x[9], self.border)
                    self.rekap.write(self.currRow, 10, x[10], self.border)
                    self.rekap.write(self.currRow, 11, x[11], self.border)
                    self.rekap.write(self.currRow, 12, x[12], self.border)
                    self.rekap.write(self.currRow, 13, x[13], self.border)
                    self.rekap.write(self.currRow, 14, x[14], self.border)
                    self.rekap.write(self.currRow, 15, x[15], self.border)
                    self.rekap.write(self.currRow, 16, '', self.border)
                    self.rekap.write(self.currRow, 17, '', self.border)
                    self.rekap.write(self.currRow, 18, '', self.border)
                    self.rekap.write(self.currRow, 19, str(x[16]), self.border)
                    self.currRow+=1
            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

        self.buttonReply = QtWidgets.QMessageBox
        self.warning = self.buttonReply.question(self, 'WARNING', "Exported to data/result/Export_Rekap_"+str(date.today())+".xlsx",
                                                 QtWidgets.QMessageBox.Ok)
        self.workbook.close()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)

    exportRekap = QtWidgets.QMainWindow()
    exportRekap.ui = Ui()

    sys.exit(app.exec_())
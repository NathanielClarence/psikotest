from PyQt5 import QtWidgets, uic
from PyQt5 import QtCore, QtGui
import sys
from subwindow import Ui_Dialog
import xlsxwriter
import datetime
#import XLSDataInputTest

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

        #print(datetime.date.today())
        self.workname = 'data/result/'+self.name+'_'+str(datetime.date.today())+'.xlsx'
        self.workbook = xlsxwriter.Workbook(self.workname)

        self.worksheet = self.workbook.add_worksheet("Identitas")
        self.header = self.workbook.add_format({'bold': True})
        self.fill = self.workbook.add_format({'bg_color': 'lime'})
        self.worksheet.write(0,0, "Nama", self.header)
        self.worksheet.write(1,0, "Usia",self.header)
        self.worksheet.write(2, 0, "Pendidikan Terakhir", self.header)
        self.worksheet.write(3, 0, "Prodi/Jurusan", self.header)
        self.worksheet.write(4, 0, "Universitas", self.header)
        self.worksheet.write(5, 0, "Posisi Dilamar", self.header)
        self.worksheet.write(6, 0, "No Telepon", self.header)
        self.worksheet.write(7, 0, "Tanggal Tes", self.header)

        self.row = 0
        for x in identity:
            if self.row == 6:
                self.worksheet.write(self.row, 1, "'" + x, self.fill)
            elif self.row ==7:
                self.worksheet.write(self.row, 1, str(x), self.fill)
            else:
                self.worksheet.write(self.row, 1, x, self.fill)

            self.row+=1

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
        self.window.ui = Ui_Dialog(self.listdat, self.workbook)
        self.window.ui.setupUi(self.window)

        '''self.window.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, False)
        self.window.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
        self.window.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)'''
        self.window.showFullScreen()
        #self.window.show()
        #self.window.ui.exButton.clicked.connect(self.closeWin)

    '''def closeWin(self):
        #print("closewin")
        if not self.window.ui.player.isFinished():
            self.window.ui.player.stop()
        self.window.close()'''

    def setAAResult(self, _num):
        self.AAresult = _num
        #print(self.AAresult)

    def saveandquit(self):
        #algo save all
        #no hp di index 6
        collist = ['No','Nama','TIU','AN','RA','KETELITIAN',
                   'AA', 'DISC', 'POSISI', 'USIA', 'PENDIDIKAN', 'NO. HP/TLP',
                   'RS','WS']
        mergelist = ['KETERANGAN', 'POTENSI KECERDASAN (2,6-3)','KETELITIAN (60-80)', 'DAYA TANGKAP & KONSENTRASI (6-8)',
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
            colnum+=1

        self.rekap.merge_range('M2:S2',mergelist[0], self.mergeformat)
        self.rekap.merge_range('M3:N3', mergelist[1], self.mergeformat)
        self.rekap.merge_range('O3:O4', mergelist[2], self.mergeformat)
        self.rekap.merge_range('P3:P4', mergelist[3], self.mergeformat)
        self.rekap.merge_range('Q3:Q4', mergelist[4], self.mergeformat)
        self.rekap.merge_range('S3:S4', mergelist[5], self.mergeformat)
        self.rekap.merge_range('R3:R4', mergelist[6], self.mergeformat)

        self.rekap.write(4, 1, self.name, self.border)
        self.rekap.write(4, 2, "='TIU'!H3", self.border)
        self.rekap.write(4, 3, "='IST3'!H3", self.border)
        self.rekap.write(4, 4, "='IST5'!H3", self.border)
        self.rekap.write(4, 5, "='ADKUDAG'!D2", self.border)
        self.rekap.write(4, 6, self.AAresult, self.border)
        self.rekap.write(4,7, "='DISC'!H12", self.border)
        self.rekap.write(4, 8, self.ident[5], self.border)
        self.rekap.write(4, 9, self.ident[1], self.border)
        self.rekap.write(4, 10, self.ident[2], self.border)
        self.rekap.write(4, 11, self.ident[6], self.border)
        self.rekap.write(4, 12, "=SUM(C5:E5)/3", self.border)
        self.rekap.write(4, 13, '=IF(M5<2,"Sangat Kurang",IF(M5<2.6,"Kurang",IF(M5<3.1,"Cukup",IF(M5<4.1,"Baik"))))', self.border)
        self.rekap.write(4, 14, '=IF(F5<60%,"Kurang",IF(F5<81%,"Cukup",IF(F5<101%,"Baik")))', self.border)
        self.rekap.write(4, 15, '=IF(G5<6,"Kurang",IF(G5<9,"Cukup",IF(G5<13,"Baik")))', self.border)
        self.rekap.write(4, 16,'', self.border)
        self.rekap.write(4, 17, '', self.border)
        self.rekap.write(4, 18, '', self.border)

        self.workbook.close()
        print("Saving and closing...")

        #XLSDataInputTest.inst_db(self.workname)

        QtWidgets.qApp.quit()

#if __name__ == "__main__":
#    app = QtWidgets.QApplication(sys.argv)
#    window = Ui()
#    sys.exit(app.exec_())
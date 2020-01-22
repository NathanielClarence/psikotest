import pandas as pd
import mysql.connector as conn
from datetime import date
#from numpy import isnan

def inst_db(workname):
    #self.workname = workname
    #workname = 'data/result/Teste_2020-01-17.xlsx'
    datas = pd.read_excel(workname, sheet_name='REKAP', header= None)
    insert_val = []
    print(datas)
    for x in datas:
        print(datas[x][4])
        if not pd.isnull(datas[x][4]):
            insert_val.append(str(datas[x][4]))
        elif x==0 or x==16 or x==17 or x==18:
            pass
        else:
            insert_val.append('')
    insert_val.append(str(date.today()))
    print(insert_val)

    f = open("data/targetDB.txt", 'r')

    mydb = conn.connect(
        host = f.readline(),
        user = "root",
        passwd = "root",
        database = "psikotest",
        auth_plugin='mysql_native_password'
    )

    insert_db = "INSERT INTO REKAP_HASIL (nama, tiu, an, ra, ketelitian, aa, disc, posisi, usia, pendidikan, no_hp, rs, ws, " \
                "ketelitian2, daya_tangkap, tanggal_tes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s," \
                "%s);"
    inst_data = tuple(insert_val)
    #print(insert_db)
    mycursor = mydb.cursor()
    mycursor.execute(insert_db, inst_data)
    mycursor.execute('commit;')
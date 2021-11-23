from dbConnection.databaseConnection import ServerConnection as Sc
import os
import pyodbc as pbo
import pandas as pd
import numpy as np


def access_connection(path):
    cxnn = pbo.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path + ';')
    with cxnn:
        cursor = cxnn.cursor()
        cursor.execute('select * from CARDS')
        data = cursor.fetchall()
        cxnn.commit()
        data = np.reshape(data, [len(data), 42])
        data = pd.DataFrame(data)
    return data


def persist(data_frame: pd.DataFrame):
    insert_query = f"insert into [staging_new].[dbo].[CARDS] " \
                   f"values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?," \
                   f"?,?,?,?,?,?,?,?," \
                   f"?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    connect = Sc().connection()
    crsr = connect.cursor()
    crsr.fast_executemany = True
    crsr.executemany(insert_query, data_frame.values.tolist())
    connect.commit()
    print('persisted_data')


def read_file_names(path):
    p = path
    fl_nms = [nm for nm in os.listdir(p) if nm.find('.mdb') > 0]
    sc = Sc()
    reported_files = sc.fetch_data('select distinct file_name from [Staging_new].[dbo].[CARDS]')
    reported_files = [nm[0] for nm in reported_files]
    fl_nms = [nm for nm in fl_nms if nm not in reported_files]
    return fl_nms


if __name__ == '__main__':
    read_directory = "C:\\Users\\M_manteghipoor.SB24\\Desktop\\MOMTAZCARD\\"
    file_names = read_file_names(read_directory)
    i = 0
    for name in file_names:
        print(i)
        print(name)
        pa = read_directory + name
        cx = access_connection(path=pa)
        cx['file_name'] = name
        # cd = cd.append(cx)
        m, n = cx.shape
        print(np.shape(cx))
        persist(cx)
        i += 1
        # s.join([str(elem) for elem in row])
        # query = c

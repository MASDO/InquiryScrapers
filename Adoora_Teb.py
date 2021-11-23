from shutil import copyfile
import datetime as DT
import pandas as pd 
import time
import dbConnection.databaseConnection as dbCxnn

# Open Connection To The DataBase
connection = dbCxnn.server_connection()
connection.set_database("Transaction_Data")

T_now = DT.datetime.today()
H_now = T_now.hour
T_future = (DT.datetime(T_now.year, T_now.month, T_now.day, 15, 17))
while T_now != T_future:
    T_now = DT.datetime.today()
    try:
        if T_now.hour > T_future.hour:
            T_future += DT.timedelta(days=1, seconds=24*60*60)
        sleep_time = int((T_future-T_now).total_seconds())
        time.sleep(sleep_time)
        copyfile("//10.88.37.120//Manteghipoor//AllT.xlsx",
                 "C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.xlsx")
        X = pd.read_excel("C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.xlsx")
        print(X)
        X.to_csv("C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.csv")
        connection.truncate_tbl("Adoora", "1398")
        connection.bulk_insert("C:\\Users\\M_manteghipoor.SB24\\Desktop\\Monthly_Saman\\AllT.csv",
                               "1398", "Adoora", "AllT.csv", X,
                               "Transaction_Data")
        T_future += DT.timedelta(days=1, seconds=24*60*60)
        sleep_time = int((T_future-T_now).total_seconds())
    except Exception as e:
        print(e)
        copyfile("//10.88.37.120//Manteghipoor//AllT.xlsx",
                 "C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.xlsx")
        X = pd.read_excel("C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.xlsx")
        print(X)
        X.to_csv("C://Users//M_manteghipoor.SB24//Desktop//Monthly_Saman//AllT.csv")
        connection.truncate_tbl("Adoora", "1398")
        connection.bulk_insert("C:\\Users\\M_manteghipoor.SB24\\Desktop\\Monthly_Saman\\AllT.csv",
                               "1398", "Adoora", "AllT.csv", X,
                               "Transaction_Data")
        T_future += DT.timedelta(days=1, seconds=24*60*60)
        sleep_time = int((T_future-T_now).total_seconds())
        print('gone...waiting')
        time.sleep(sleep_time)

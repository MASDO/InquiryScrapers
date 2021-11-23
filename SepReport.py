import pandas as pd
from dbConnection import databaseConnection as db
import numpy as np
import os
# max value of consumer scoting ptoblems will solve this issue
class cleansedSepTable:
    def __init__(self, file_name = '', upper_bond = 250000):
        """file name has to have xlsx in it !!! its an excel file name """
        self.column_names = ['شماره ترمینال','شماره پذیرنده','نام پذیرنده','موبایل',
                             'تلفن','کد پستی','شماره حساب','نام صاحب حساب','شعبه',
                             'گروه ترمینالی','شرح صنف','بستر ارتباطی','آدرس','تعداد تراکنش خرید',
                             'مبلغ تراکنش خرید','تعداد تراکنش شارژ','مبلغ تراکنش شارژ',
                             'تعداد تراکنش قبض','مبلغ تراکنش قبض ','تعداد تراکنش مانده گیری',
                             'کد شعبه','کد سپرده','شماره مشتری']

        self.reading_path = 'D:\\SEP_Terminals\\ReadingData\\'
        self.file_name = file_name
        self.file_path = self.reading_path + file_name
        self.related_table = '[1400].[SEP_Monthly_new]'
        self.persistedBefore = True
        self.upper_bond = upper_bond

    def check_persisted(self):
        connect = db.ServerConnection()
        report_date = self.file_name.replace('.xlsx','')
        query = 'select report_date from [posdb].[1400].[pos_report_log] where report_date = {}'.format(report_date)
        connect.set_query(query)
        rep_date = connect.fetch_data(query)
        print(len(rep_date))
        if len(rep_date) == 0:
            self.persistedBefore = False
        return self.persistedBefore

    def expend_file(self):
        os.remove(self.file_path)

    def read_data(self):
        column_names = self.column_names
        p =  'D://SEP_Terminals//ReadingData//{}'.format(self.file_name)
        data = pd.read_excel(open(p, 'rb'))
        df = pd.DataFrame(data)
        new_columns = ['کد شعبه','کد سپرده','شماره مشتری','سریال سپرده']
        df[new_columns] = df['شماره حساب'].str.split('-', expand=True)
        df_needed = df[column_names].copy(deep=True)
        df_needed['تعداد تراکنش خرید'] = df_needed['تعداد تراکنش خرید'].astype(str)
        df_needed['مبلغ تراکنش خرید'] = df_needed['مبلغ تراکنش خرید'].astype(str)
        df_needed['تعداد تراکنش شارژ'] = df_needed['تعداد تراکنش شارژ'].astype(str)
        df_needed['مبلغ تراکنش شارژ'] = df_needed['مبلغ تراکنش شارژ'].astype(str)
        df_needed['تعداد تراکنش قبض'] = df_needed['تعداد تراکنش قبض'].astype(str)
        df_needed['مبلغ تراکنش قبض '] = df_needed['مبلغ تراکنش قبض '].astype(str)
        df_needed['تعداد تراکنش مانده گیری'] = df_needed['تعداد تراکنش مانده گیری'].astype(str)
        # df['نام صاحب حساب'] = df['نام صاحب حساب'].astype(str)
        # df['نام صاحب حساب'] = df['نام صاحب حساب'].map(lambda x: int(x.strip("'")))
        df_needed['تعداد تراکنش خرید'] = df_needed['تعداد تراکنش خرید'].map(lambda x: float(x.strip(",")))
        df_needed['مبلغ تراکنش خرید'] = df_needed['مبلغ تراکنش خرید'].map(lambda x: float(x.strip(",")))
        df_needed['تعداد تراکنش شارژ'] = df_needed['تعداد تراکنش شارژ'].map(lambda x: float(x.strip(",")))
        df_needed['مبلغ تراکنش شارژ'] = df_needed['مبلغ تراکنش شارژ'].map(lambda x: float(x.strip(",")))
        df_needed['تعداد تراکنش قبض'] = df_needed['تعداد تراکنش قبض'].map(lambda x: float(x.strip(",")))
        df_needed['مبلغ تراکنش قبض '] = df_needed['مبلغ تراکنش قبض '].map(lambda x: float(x.strip(",")))
        df_needed['تعداد تراکنش مانده گیری'] = df_needed['تعداد تراکنش مانده گیری'].map(lambda x: float(x.strip(",")))

        return df_needed

    def persist(self):
        X = self.check_persisted()
        if X:
            print("Already Got The Report !!!")
        else:
            data = self.read_data()
            m, n = np.shape(data)
            max_iter = max(m , self.upper_bond)
            df_2 = data.values.tolist()
            connect = db.ServerConnection()
            connect.set_database('Posdb')
            exceptionList = []
            exceptionRecords = []
            for i in range(0, m):
                t = df_2[i]
                s = "','"
                value = s.join([str(elem) for elem in t])
                insert_Query = "insert into [1400].[SEP_Monthly_new] values('{}','{}')". \
                    format(value, int(name.replace('.xlsx', '')))
                log_query = "insert into [1400].[Pos_report_log] values ({},getDate())".\
                    format(int(name.replace('.xlsx', '')))
                print(insert_Query)
                try:
                    connect.set_query(insert_Query)
                    connect.insert_data()
                    # connect.set_query(log_query)
                    # connect.insert_data()
                except Exception as e:
                    exceptionRecords.append(value)
                    exceptionList.append(str(e))
                with open("leftover_logs.txt", 'w') as f:
                    for item in exceptionRecords:
                        f.write(str(item.encode('utf-8')))
                    f.close()
                with open("errorLogs.txt", 'w') as g:
                    for item in exceptionList:
                        g.write(str(item.encode('utf-8')))
                    g.close()
                print("inserted value")
        self.expend_file()
        return



if __name__ == '__main__':
    p = 'D:\\SEP_Terminals\\ReadingData\\'
    file_names = [name for name in os.listdir(p) if name.find('.xls') > 0]
    print(file_names)
    # file_names = ['14000131','14000231']
    for name in file_names:
        # p = '//RD7-119//Users//e_ramezani//Desktop//manteghipoor//{}.xlsx'.format(name)
        sep = cleansedSepTable(file_name=name)
        data = sep.read_data()
        print(data)
        sep.persist()
        print("done the services")

import pyodbc as pbo
import pandas as pd
import numpy as np
import sqlalchemy


class ServerConnection:
    def __init__(self, query: str = '', database:str = '', **kwargs):
        self.database = database
        self.query = query
        self.driver = '{ODBC Driver 17 for Sql Server}'
        self.server = r'RD-125\MASOUD_ADMIN'

    def connection(self):   
        cxnn = pbo.connect(Trusted_connection='yes',
                           driver=self.driver,
                           server=self.server,
                           database=self.database,
                           p_str="")
        return cxnn

    def set_driver(self, driver: str):
        self.driver = driver

    def set_server(self, server: str):
        self.server = server

    def set_query(self, query):
        self.query = query

    def set_database(self, db_name: str):
        self.database = db_name

    def insert_data(self):
        query = self.query
        cxnn = self.connection()
        cursor = cxnn.cursor()
        cursor.execute(query)
        cxnn.commit()
        cxnn.close()

    def fetch_data(self, query):
        cxnn_f = self.connection()
        cursor = cxnn_f.cursor()
        cursor.execute(query)
        self.data = cursor.fetchall()
        cxnn_f.commit()
        cxnn_f.close()
        return self.data

    def delete(self, query):
        cxnn_f = self.connection()
        cursor = cxnn_f.cursor()
        cursor.execute(query)
        cxnn_f.commit()
        cxnn_f.close()
        return 'Deleted Data'

    def truncate_tbl(self, tbl_name , schema_name):
        cxnn_f = self.connection()
        query = "truncate table [{0}].[{1}]".format(schema_name,tbl_name) 
        cursor = cxnn_f.cursor()
        cursor.execute(query)
        cxnn_f.commit()
        cxnn_f.close()

    def bulk_insert(self, path, schema, tbl, csvName, dataframe, dbname='attrition_Db'):
        self.set_database(dbname)
        path_t = path + csvName
        dataframe.to_csv(path_or_buf=path,index = False) 
        str_directory = "'" + path + "'"
        insert_into_temp_table = ("bulk insert " + dbname + ".[" + schema + "].[" + tbl + "] from " + str_directory +
                                  " with (firstrow = 2 , fieldterminator = ',', rowterminator = '0x0a')")
        self.query = insert_into_temp_table
        self.insert_data()




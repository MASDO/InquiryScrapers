import pandas as pd
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import getpass
import dbConnection.databaseConnection as d
import numpy as np
import pyautogui as pye
import os
# this file is not for sale and customer scoring problems will be solved

url = 'http://biservice:2002/BatchInquiry/'
print(getpass.getuser())
# \\RD7-115\Users\p_mohebmaleki\Desktop\share
# C:\Users\M_manteghipoor.SB24\Desktop\Inquiry
path = 'C:\\Users\\M_manteghipoor.SB24\\Desktop\\Inquiry'
# '//RD7-115//Users//p_mohebmaleki//Desktop//share'
# os.path.join(os.path.join(os.environ['USERPROFILE'], 'Desktop'))
# print(path)
Inquiry_path_load = path + '\\ChequeInquiry.xls'
Inquiry_path_med = path + '\\Tokenized\\ChequeInquiry_save.xls'
ehsan_custno_path_custno = '//RD7-119//Users//e_ramezani//Desktop//manteghipoor//ChequeInquiry_custno.xlsx'
ehsan_custno_path_national_code = '//RD-125//Users//M_manteghipoor.SB24//' \
                                  'Desktop//Inquiry//master//ChequeInquiry_save.xls'
    # '//RD7-119//Users//e_ramezani//Desktop//manteghipoor//' \
    #                               'ChequeInquiry_national_code.xlsx'
national_code = pd.read_excel(Inquiry_path_load)
# print(national_code)
# print(Inquiry_path_load)

style_text_align_vert_center_horiz_center = xlwt.easyxf("font: name Calibry ;align: vert centre, horiz centre")
style_text_align_vert_bottom_horiz_left = xlwt.easyxf("font: name Calibry ; align: vert bottom, horiz left")
style_text_align_vert_top_horiz_right = xlwt.easyxf("font: name Calibry ;align:vert top, horiz right")
style_text_wrap_font_bold_red_color = xlwt.easyxf("font: name Calibry;align:wrap on; font: bold on, color-index red")
# worksheet.write(1, 0, "Vert Center Horiz Center", style_text_align_vert_center_horiz_center)
# worksheet.write(2, 0, "Vert Bottom Horiz Left", style_text_align_vert_bottom_horiz_left)
# worksheet.write(3, 0, "Vert Top Horiz Right", style_text_align_vert_top_horiz_right)
# worksheet.write(4, 0, "Wrapped Font, Bold, Red in Color", style_text_wrap_font_bold_red_color)
# workbook.save("C:\\YourDirectory\\FileName.xls")


rb = open_workbook(filename=Inquiry_path_load, formatting_info=True, on_demand=True)
rb_1 = open_workbook(filename=Inquiry_path_med, formatting_info=True, on_demand=True)
ws = copy(rb_1)
sh = ws.get_sheet(0)

# book = Workbook(Inquiry_path_load)
# writer = pd.ExcelWriter(Inquiry_path_load, engine='xlwt')


def write_with_style(ws , row, col, value):
    # if ws.rows[row]._Row__cells[col]:
    #     old_xf_idx = ws.rows[row]._Row__cells[col].xf_idx
    #     ws.write(row, col, value, style_text_align_vert_center_horiz_center)
    #     ws.rows[row]._Row__cells[col].xf_idx = old_xf_idx
    #  else:
    ws.write(row, col, value, style_text_align_vert_center_horiz_center)


def right(s, num):
    return s[-num:]


def extract_custno(is_national_code=True):
    if is_national_code:
        f_data = pd.read_excel(open(ehsan_custno_path_national_code, 'rb'))
        n_data = f_data['National'].to_list()
    else:
        f_data = pd.read_excel(open(ehsan_custno_path_custno, 'rb'))
        f_data = f_data['custno'].to_list()
        cxnn = d.ServerConnection()
        s = "','"
        value = s.join([str(elem) for elem in f_data])
        value = "'"+value + "'"
        query = "SELECT distinct National_code FROM [DEMOGRAPHIC_INFO].[permanent].[Demographics] where custno in " \
                "({})".format(value)
        n_data = cxnn.fetch_data(query)
    return n_data


if __name__ == '__main__':
    batch_size = 100
    data = extract_custno(is_national_code=True)
    # print(data)
    data_list = [(count, d) for count, d in enumerate(data)]
    DF = pd.DataFrame(data_list, columns=['Row', 'NationalCode'])
    print(np.shape(DF))
    id_num = 1
    # DF = pd.DataFrame(columns=['Row', 'NationalCode'])
    t = 1
    for index, n in DF.iterrows():
        Inquiry_path_save = path + '\\Tokenized\\ChequeInquiry_save_0{}.xls'.format(t)
        nationalCode_str = right('000000' + str(n[1]), 10)
        print('String code = {}'.format(nationalCode_str))
        if id_num <= batch_size:
            new_list = {'Row': id_num, 'NationalCode': n[1]}
            write_with_style(sh, id_num, 0, id_num)
            write_with_style(sh, id_num, 1, nationalCode_str)
            # n must be converted to a string
            print(id_num)
            id_num += 1
        else:
            id_num = 1
            t += 1
        ws.save(Inquiry_path_save)
    exit('Done sir')


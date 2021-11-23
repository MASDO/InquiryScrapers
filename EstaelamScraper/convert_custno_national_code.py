import os
import pyautogui as pye
import pydirectinput as pyd

if __name__ == '__main__':
    print('ready for test')
    pye.sleep(3)
    pyd.press('down')
    # making files usable for getting records
    path = 'C:\\Users\\M_manteghipoor.SB24\\Desktop\\Inquiry'
    Inquiry_path_files = path + '\\Tokenized\\'
    files_names = [name for name in os.listdir(Inquiry_path_files) if name.find('.xls') > 0]
    number_of_itr = len(files_names)
    remote_position = (174, 1066)
    a = 186
    b = 172
    step = \
        24
    print(pye.position())
    pye.click(remote_position)
    for i in range(0, number_of_itr):
        excel_position = (a, b)
        if i == 0:
            pye.click(excel_position)
        else:
            pyd.press('down')
        pye.press('enter')
        pye.sleep(7)
        pye.click(1160, 719)
        pye.click(49, 7)
        pye.sleep(3)
        pye.click(1905, 7)
        pye.press('enter')
        pye.sleep(7)
        b += step

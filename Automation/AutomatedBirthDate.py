import pandas as pd
import os
import pyautogui as pye
import pydirectinput as pyd
import pyperclip
import csv


def insert_addess(site_address='https://portal.pbn.net/tabid/136/view/EstlmShenase/p1/tas/Default.aspx'):
    address_bar_position = (442, 51)
    pye.click(address_bar_position)
    pye.typewrite(site_address)
    pye.press('enter')
    pass


def select_real_iranian():
    personality_position = (1331, 446)
    real_iranian_position = (1360, 488)
    pye.click(personality_position)
    pye.click(real_iranian_position)


def insert_data(national_code):
    insert_position = (1359, 494)
    pye.click(insert_position)
    pye.sleep(0.5)
    pye.click(insert_position)
    pye.sleep(0.5)
    pye.click(insert_position)
    pye.sleep(0.5)
    pye.typewrite(national_code)
    Button_position = (1440, 526)
    pye.click(Button_position)


def click_button():
    Button_position = (1374, 540)
    pye.click(Button_position)


def copy_data(n, total_list):
    birth_date_position = (495, 484)
    pye.sleep(1)
    pye.tripleClick(birth_date_position)
    pye.hotkey('ctrl', 'c')
    X = pyperclip.paste()
    X = X[16:26]
    national_code_position = (765, 450)
    pye.doubleClick(national_code_position)
    pye.hotkey('ctrl', 'c')
    Y = pyperclip.paste()
    # Y = X[16:26]
    total_list.append([n, X, Y])
    return


def log_out():
    log_out_position = (41, 118)
    pye.click(log_out_position)
    pye.sleep(5)


def click_login_position():
    login_page = "https://portal.pbn.net/tabid/38/Default.aspx"
    address_bar_position = (442, 51)
    pye.click(address_bar_position)
    pye.typewrite(login_page)
    pye.press('enter')
    login_position = (442, 51)
    pye.click(login_position)
    pye.sleep(5)


def login_masoud():
    user = 'saman1-99999'
    passwd = '123456789'
    login_page = "https://portal.pbn.net/tabid/38/Default.aspx"
    address_bar_position = (442, 51)
    pye.click(address_bar_position)
    pye.typewrite(login_page)
    login_position = (177, 410)
    pye.tripleClick(login_position)
    pye.typewrite(user)
    passwd_position = (219, 438)
    pye.tripleClick(passwd_position)
    pye.typewrite(passwd)
    button_position = (294, 467)
    pye.click(button_position)


def create_person_url(national_code):
    url = "https://portal.pbn.net/tabid/136/view/EstelamOut/p1/idtas/p2/{}/p3/1/Default.aspx".format(national_code)
    address_bar_position = (442, 51)
    pye.click(address_bar_position)
    pye.typewrite(url)
    pye.press('enter')
    pye.click(address_bar_position)
    pye.press('enter')


if __name__ == '__main__':
    wait_time_seconds = 3
    n_counter = 10
    main_data = pd.read_excel('National_code_data.xlsx')
    # main_data = main_data.parse()
    data = main_data['NationalCode'].to_list()
    site_address = "https://portal.pbn.net/tabid/136/view/EstlmShenase/p1/tas/Default.aspx"
    pye.sleep(6)
    personality_position = (1383, 526)
    real_position = (1360, 488)
    address_bar_position = (442, 51)
    print(pye.position())
    explorer_position = (363, 1062)
    counter = 0
    total_list = []
    # pye.click(explorer_position)
    # click_login_position()
    for counter, n in enumerate(data):
        if counter == 0:
            pye.click(explorer_position)
        create_person_url(n)
        pye.sleep(2)
        copy_data(n, total_list)
        if (counter/n_counter) == (counter//n_counter):
            pye.sleep(wait_time_seconds)
            total_df = pd.DataFrame.from_records(total_list, columns=['report_national_code',
                                                                     'Birth_date',
                                                                     'Exact_national_code'])
            print(total_df)
            total_df.to_csv('Result_{}.csv'.format(21 + int(counter//n_counter)))
            # login_masoud()
            print(total_df)

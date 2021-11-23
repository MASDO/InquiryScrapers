from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from dbConnection.databaseConnection import ServerConnection as server_connection
from os import listdir
from multiprocessing import Pool
import multiprocessing as mp
import time


# this might solve the issue of customer scoring problems
# thats the way we do the standing problems
# the main problem will stand by far the deepest issues of all time I am done

def open_file_uploader(filename='ChequeInquiry_save_1.xls',loan_cheque = 1):
    if loan_cheque == 1:
        push_button  = 'MainContent_btnFacilityInquiry'
    elif loan_cheque == 0:
        push_button = 'MainContent_btnChequeInqury'
    delay = 5
    path = 'C:\\Users\\M_manteghipoor.SB24\\Desktop\\Inquiry\\Tokenized\\{}'.format(filename)
    url = 'http://biservice:2002/BatchInquiry/'
    chrome_options = webdriver.ChromeOptions() # creating chrome options class
    # chrome_options.add_argument('headless') # keeps chrome from starting
    chrome_options.add_experimental_option("excludeSwitches", ['enable-automation']) # casting machine to a human
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(
        executable_path='C:\\Users\\M_manteghipoor.SB24\\PycharmProjects\\pythonProject2\\chromedriver.exe',
        options=chrome_options)
    #driver = webdriver.Edge(executable_path=
    #                        'C:\\Users\\M_manteghipoor.SB24\\PycharmProjects\\pythonProject2\\msedgedriver.exe')

    # seconds
    driver.get(url)
    driver.find_element_by_id('FileUploader').send_keys(path)
    driver.find_element_by_id(push_button).click()
    empty_dict = {}
    try:
        WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'MainContent_ResultGrid')))
        # driver.set_window_size()
        # driver.minimize_window()
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("window.scrollTo(document.body.scrollHeight,0);")
        table = driver.find_element_by_id('MainContent_ResultGrid')
        tbl = table.find_elements_by_tag_name('tr')
        id = 1
        for row in tbl:
            ro = row.find_elements_by_tag_name('td')
            customer_list = [r.text for r in ro]
            empty_dict[id] = customer_list
            id += 1
    except Exception as e:
        print("took too long or other errors")
        print(e)
    driver.quit()
    return empty_dict

def get_inquiry_results(f_name:str , report_type:int):
    """... gets the reports for customer ..."""
    data = []
    result_code = 0
    if report_type == 1:
       tbl_name = 'permanent.loan_inquiry'
    elif report_type == 0:
       tbl_name = 'permanent.chequeInquiry'
    # data = open_file_uploader(f_name)
    try:
        data = open_file_uploader(f_name,report_type)
        del data[1]
        with open('successful_inquiry_log.txt', 'a') as f:
            f.write('\n')
            f.write(f_name + ',')
        # print(data)
        for key, value in data.items():
            print(value)
            s = "',N'"
            v = s.join([str(elem) for elem in value])
            insert_query = "insert into {} values('{}',getdate(),'{}') ".format(tbl_name,v,f_name)
            print(insert_query)
            # print(filename)
            connection = server_connection()
            connection.set_database('Loan_DB')
            connection.set_query(insert_query)
            connection.insert_data()
            # with open('inquiry_log.txt', 'w') as f:
            #     for item in inquiry_log:
            #         f.write(item)
            result_code = 1
    except Exception as e:
        print(e)
        # inquiry_log.append(f_name)
        with open('failed_inquiry_log.txt', 'a') as f:
            f.write('\n')
            f.write(f_name + ',')
    return result_code

def make_it_concurrent(f_name,down):
    # 'ChequeInquiry_save_{}.xls'.format(i)
    print(f_name)
    # file_name
    try:
        get_inquiry_results(f_name=f_name, report_type=1)
        get_inquiry_results(f_name=f_name, report_type=0)
        print('got {} down ....!!!'.format(down))
    except Exception as e:
        print(e)

def get_reported_list():
    cx = server_connection()
    read_query = 'SELECT distinct ReportFileName FROM [Loan_DB].[permanent].[loan_Inquiry]'
    # cx.set_query(read_query)
    reported_list = cx.fetch_data(read_query)
    reported_list = [r[0] for r in reported_list]
    return reported_list



if __name__ == '__main__':
    # print('waiting')
    # ime.sleep(7200)
    r = get_reported_list()
    path = 'C:\\Users\\M_manteghipoor.SB24\\Desktop\\Inquiry'
    Inquiry_path_files = path + '\\Tokenized\\'
    dir_list = [name for name in listdir(Inquiry_path_files) if name.find('.xls') > 0]
    dir_list = [name for name in dir_list if name not in r]
    r = get_reported_list()
    dir_list_1 = set(dir_list)
    print(dir_list_1)
    inquiry_log = []
    pool = Pool(int(5) , maxtasksperchild=1)
    results = [pool.apply_async(make_it_concurrent,(filename,id)) for id,filename in enumerate(dir_list)]
    for r in results:
        r.get()
    path = 'C:\\Users\\M_manteghipoor.SB24\\Desktop\\Inquiry'
    Inquiry_path_files = path + '\\Tokenized\\'
    dir_list = [name for name in listdir(Inquiry_path_files) if name.find('.xls') > 0]
    dir_list = [name for name in dir_list if name not in r]
    print(dir_list)
    print("Done the Report completely")

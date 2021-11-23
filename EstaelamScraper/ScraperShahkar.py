import sys
import requests
import os
import requests.utils
import requests.adapters
import certifi.core
import http.client as hc
import pyautogui as pye
import time
from dbConnection import databaseConnection
import pyperclip




def override_where():
    """ overrides certifi.core.where to return actual location of cacert.pem"""
    # change this to match the location of cacert.pem
    return os.path.abspath("SBCA1")


# is the program compiled ?
def call_shahkar_api(identification_number: str, mobile: str):
    if hasattr(sys, "frozen"):
        os.environ["SBCA1"] = override_where()
        certifi.core.where = override_where
        # delay importing until after where() has been replaced
        # replace these variables in case these modules were
        # imported before we replaced certifi.core.where
        requests.utils.DEFAULT_CA_BUNDLE_PATH = override_where()
        requests.adapters.DEFAULT_CA_BUNDLE_PATH = override_where()
    # xt = os.environ.get("C:/Users/M_manteghipoor.SB24/Desktop/Customer360_Certificate_Ver2/")
    # print(xt)
    v_path = ""

    url = "https://biservice:8099/InternetShahkar/ShahkarSrvSvc.svc?wsdl"
    parameters = {"IdentificationNumber": identification_number, "ServiceNumber": mobile}
    resp = requests.get(url, parameters)
    print(resp)
    print(resp.status_code)
    return resp


class ShahkarReport:
    def __init__(self):
        self.customer = 'masoud'
    # todo:this is a class for fetching data


def get_result(result_pos: tuple = (1850, 789)):
    """this function gets the results of the api call"""
    pye.click(result_pos)
    pye.hotkey('ctrl', 'a')
    pye.hotkey('ctrl', 'c')
    txt = pyperclip.paste()
    if txt.find('<a:Response>200</a:Response>') > 0:
        r = 'تطابق'
    elif txt.find('<a:Response>600</a:Response>') > 0:
        r = 'عدم تطابق'
    elif txt.find('<a:Response>311</a:Response>') > 0:
        r = 'کد ملی نامعتبر'
    else:
        r = 'خطای سرویس'
    print(r)
    return r


def clear_sessions(cl_pos: tuple = (260, 689)):
    pye.click(cl_pos)


def type_value(pos: tuple, value: str):
    pye.moveTo(pos)
    pye.doubleClick(pos)
    pye.typewrite(value)


if __name__ == '__main__':
    db = databaseConnection.ServerConnection()
    db.set_database('staging_new')
    read_data_query = ("select * from adooraEstelam "
                       "where National_code in (select right('000000' + nationalCode , 10) collate Arabic_CI_AS "
                       "from [DEMOGRAPHIC_INFO].[permanent].[shahkar_Inquiry]"
                       " where [shahkarResult] = N'{}')".format('خطای سرویس'))
    print(read_data_query)
    db.set_query(read_data_query)
    data = db.fetch_data(read_data_query)
    pye.click(669, 1061)

    for row in data:
        national_code = row[0]
        mobile = row[1]

        nationalCode_Pos = (700, 193)
        type_value(nationalCode_Pos, national_code)
    #
        mobile_no_pos = (685, 242)
        type_value(mobile_no_pos, mobile)
    #
        clear_sessions((260, 689))
    #
        inquiry_key_pos = (899, 311)
        pye.click(899, 311)
        time.sleep(8)
        inquiry_result_pos = (1607, 836)

        result = get_result()
        result_list = [national_code, mobile, result]
        delete_query = "delete from [DEMOGRAPHIC_INFO].[permanent].[shahkar_Inquiry] " \
                       "where NationalCode = right('000000'+{},10)".format(national_code)

        db.delete(delete_query)
        query = "insert into " \
                "DEMOGRAPHIC_INFO.[permanent].[shahkar_Inquiry] " \
                "values (right('000000'+{},10), {}, N'{}',getdate())".\
                format(national_code, mobile, result)
        db.set_query(query)
        db.insert_data()
        time.sleep(0.5)

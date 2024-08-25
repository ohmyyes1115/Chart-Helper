import ast
import base64
import json
import os
import pickle
import platform
import random
import threading
import time
from enum import Enum
from io import StringIO, BytesIO

import imutils as imutils
import pyperclip
import wx

import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageGrab
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import constants
import home_therapy_web_manager
# from home_therapy_web_manager import HomeTherapyWebManager
import UI
# from home_therapy_web_manager import BrowserTab
from UI import MainWindow
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# pytesseract.pytesseract.tesseract_cmd = r'<C:\Users\yy\Desktop\HomeTherapy\tesseract-4.0.0-alpha>\tesseract.exe'
# img = Image.open('sample.png')
# captcha_str = pytesseract.image_to_string(img, config="--psm 8 --oem 0 -c tessedit_char_whitelist"
#                                                                "=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")
# print("captcha=" + captcha_str)

# 衛服部首頁
from constants import HomeTherapyCompany

HOME_THERAPY_GOVERN_HOME_PAGE_URL = r'https://csms2.sfaa.gov.tw/lcms/'
# 照顧平台 - 首頁
TAKE_CARE_HOME_PAGE_URL = r'https://csms2.sfaa.gov.tw/lcms/saTree/treeIndex?saPortalId=278266377'
# 照顧平台 - 照會列表
CLIENT_SEARCH_PAGE_URL = r'https://csms2.sfaa.gov.tw/lcms/qd/listQd111'

# CHROME_DRIVER_EXE_PATH = r'C:\Users\yy\Desktop\HomeTherapy\chromedriver.exe'
CHROME_DRIVER_EXE_PATH = r'chromedriver.exe'
FIREFOX_DRIVER_EXE_PATH = r'geckodriver.exe'
PHANTOM_JS_EXE_PATH = r'phantomjs.exe'  # deprecated

# LC035 - 個案名稱3 - [CA01] / CA03
UPDATE_DB_URL1 = r'https://csms2.sfaa.gov.tw/lcms/qd/saveQd120Batch?' \
                 r'qd120.id=' \
                 r'&qd120.ca100=919606075' \
                 r'&qd120ServdtNum=1' \
                 r'&qd120dt.1.servDt=108%2F10%2F10' \
                 r'&qd120ServtimeNum=1' \
                 r'&qd120time.1.startHh=11' \
                 r'&qd120time.1.startMm=12' \
                 r'&qd120time.1.endHh=13' \
                 r'&qd120time.1.endMm=14' \
                 r'&qd120.servUsers=882319126' \
                 r'&qd120.noAa09=' \
                 r'&qd120.commt=記錄錄' \
                 r'&qd120Num=1' \
                 r'&qd120.1.qd100=285889258' \
                 r'&qd120.1.addr1=' \
                 r'&qd120.1.addr2=' \
                 r'&qd120.1.carNum=' \
                 r'&qd120.1.driver=' \
                 r'&qd120.1.amount=1' \
                 r'&qd120.1.price=4500.0' \
                 r'&qd120.1.stype=1' \
                 r'&qd120.1.lastServ=false' \
                 r'&qd120.1.missedVisit=false' + \
 \
                 r'&qd120Num=2' \
                 r'&qd120.2.qd100=285889260' \
                 r'&qd120.2.addr1=' \
                 r'&qd120.2.addr2=' \
                 r'&qd120.2.carNum=' \
                 r'&qd120.2.driver=' \
                 r'&qd120.2.amount=0' \
                 r'&qd120.2.price=4500.0' \
                 r'&qd120.2.stype=1' \
                 r'&qd120.2.lastServ=false' \
                 r'&qd120.2.missedVisit=false'

# CCO0168 - 個案名稱1 - CA03 only
UPDATE_DB_URL1 = r'https://csms2.sfaa.gov.tw/lcms/qd/saveQd120Batch?' \
                 r'qd120.id=' \
                 r'&qd120.ca100=362191551' \
                 r'&qd120ServdtNum=1' \
                 r'&qd120dt.1.servDt=108%2F10%2F09' \
                 r'&qd120ServtimeNum=1' \
                 r'&qd120time.1.startHh=10' \
                 r'&qd120time.1.startMm=11' \
                 r'&qd120time.1.endHh=12' \
                 r'&qd120time.1.endMm=13' \
                 r'&qd120.servUsers=864119605' \
                 r'&qd120.noAa09=' \
                 r'&qd120.commt=%E8%A8%98%E9%8C%841%0D%0A%E8%A8%98%E9%8C%842' \
                 r'&qd120Num=1' \
                 r'&qd120.1.qd100=285889260' \
                 r'&qd120.1.addr1=' \
                 r'&qd120.1.addr2=' \
                 r'&qd120.1.carNum=' \
                 r'&qd120.1.driver=' \
                 r'&qd120.1.amount=1' \
                 r'&qd120.1.price=4500.0' \
                 r'&qd120.1.stype=1' \
                 r'&qd120.1.lastServ=false' \
                 r'&qd120.1.missedVisit=false'


class DbQuery_CreateClientRecord:
    # common field
    qd120_id = ''
    qd120_ca100 = ''
    qd120ServdtNum = '1'
    qd120dt_1_servDt = '108/10/12'
    qd120ServtimeNum = '1'
    qd120time_1_startHh = ''
    qd120time_1_startMm = ''
    qd120time_1_endHh = ''
    qd120time_1_endMm = ''
    qd120_servUsers = ''
    qd120_noAa09 = ''
    qd120_commt = ''
    qd120Num = '1'

    # CA-01
    qd120_1_qd100 = '285889258'  # 同一人，不同次記錄，也都相同。[實驗]治療師沒填也會有，猜是QA01的ID
    qd120_1_addr1 = ''
    qd120_1_addr2 = ''
    qd120_1_carNum = ''
    qd120_1_driver = ''
    qd120_1_amount = ''
    qd120_1_price = '4500.0'  # TODO: 從 DB 撈
    qd120_1_stype = '1'
    qd120_1_lastServ = 'false'
    qd120_1_missedVisit = 'false'

    # CA-03
    qd120Num = '2'
    qd120_2_qd100 = ''
    qd120_2_addr1 = ''
    qd120_2_addr2 = ''
    qd120_2_carNum = ''
    qd120_2_driver = ''
    qd120_2_amount = ''
    qd120_2_price = '4500.0'  # TODO: 從 DB 撈
    qd120_2_stype = '1'
    qd120_2_lastServ = 'false'
    qd120_2_missedVisit = 'false'


def get_captcha(login_page_screenshot):
    cropped_captcha = crop_image(login_page_screenshot)
    stroke_width = stroke_height = 6
    img_with_white_bg = Image.new('RGB',
                                  (cropped_captcha.width + (stroke_width << 1),
                                   cropped_captcha.height + (stroke_height << 1)), "white")
    img_with_white_bg.paste(cropped_captcha,
                            (stroke_width,
                             stroke_height,
                             cropped_captcha.width + stroke_width,
                             cropped_captcha.height + stroke_height))
    img_with_white_bg.save("cropped_extend.png")

    img_denoised = cv2.imread("cropped_extend.png", cv2.IMREAD_GRAYSCALE)
    img_denoised = cv2.resize(img_denoised, None, fx=10, fy=10, interpolation=cv2.INTER_LINEAR)
    cv2.imwrite("cropped_extend_1_resized.png", img_denoised)
    img_denoised = cv2.medianBlur(img_denoised, 9)
    cv2.imwrite("cropped_extend_2_medianBlur.png", img_denoised)

    # to binary image
    th, img_denoised = cv2.threshold(img_denoised, 175, 255, cv2.THRESH_BINARY)
    cv2.imwrite("cropped_extend_mediaBlur.png", img_denoised)

    # fill black noise
    kernel = np.ones((5, 5), np.uint8)
    img_denoised = cv2.morphologyEx(img_denoised, cv2.MORPH_CLOSE, kernel)
    cv2.imwrite("cropped_extend_morphologyEx.png", img_denoised)

    os.remove("cropped_extend.png")
    os.remove("cropped_extend_1_resized.png")
    os.remove("cropped_extend_2_medianBlur.png")
    os.remove("cropped_extend_mediaBlur.png")
    os.remove("cropped_extend_morphologyEx.png")

    # 1/I 1/J 1/L 3/S 3/B 5/S 6/G 7/Y 8/S 8/B 9/Q
    # F/P G/Q
    # captcha_str = pytesseract.image_to_string(img_denoised, config="--psm 8 --oem 0 -c tessedit_char_whitelist"
    #                                                            "=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    # 1/l 5/> 7// 8/S 9/O 9/$ 9/Y
    captcha_str = pytesseract.image_to_string(img_denoised, config="--psm 10")
    print("raw-captcha=" + captcha_str)
    captcha_str = captcha_str.replace("/", "7").replace("l", "1").replace("]", "J")
    captcha_str = strip_invalid_captcha_char(captcha_str)
    print("new-captcha=" + captcha_str)

    return captcha_str


def strip_invalid_captcha_char(captcha):
    stripped_captcha_str = ''

    for i in range(len(captcha)):
        if is_valid_captcha_char(captcha[i]):
            stripped_captcha_str += captcha[i]

    return stripped_captcha_str


def is_valid_captcha_char(char):
    return 'A' <= char <= 'Z' or '0' <= char <= '9'


def crop_image(png_image_name):
    image = Image.open(png_image_name)
    # box = (335, 406, 398, 447)  # whole captcha area
    if platform.node() == "LAPTOP-GAGMOCQA":
        box = (396, 407, 459, 443)
    elif platform.node() == "yy-i7":
        box = (335, 416, 403, 435)
    else:
        box = (335, 416, 403, 435)

    # print(platform.node())

    return image.crop(box)
    region.save("cropped.png")


def press_key(browser, key_code):
    ActionChains(browser).send_keys(key_code).perform()


class HomeTherapyLoginProfile:
    def __init__(self, account: str, pw: str, home_therapy_company: HomeTherapyCompany):
        self.account = account
        self.pw = pw
        self.home_therapy_company = home_therapy_company


def create_theropy_cc_profile():
    return HomeTherapyLoginProfile(constants.CC_ACCOUNT, constants.CC_PASSWORD, HomeTherapyCompany.CC)


def create_theropy_lc_profile():
    return HomeTherapyLoginProfile(constants.LC_ACCOUNT, constants.LC_PASSWORD, HomeTherapyCompany.LC)


CC = True


def load_web_page(ui, browser_display_ratio):
    # lease = 14 * 24 * 60 * 60  # 14 days in seconds
    # lease = 1572224444  # cookie expiry
    # end = time.gmtime(lease)
    # expires = time.strftime("%a, %d-%b-%Y %T GMT", end)
    # print(expires)
    # return

    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('--headless')  # 啟動無頭模式, 規避google bug
    chrome_options.add_argument('--disable-gpu')
    # (maximize, 1680_1.25): 0.8=(2124, 1274) 1=(1698, 1018) no_set=(1360, 816)
    # (maximize, 1680_1.0) : 0.8=(2120, 1283) 1=(1680, 1026) no_set=(1696, 1026)
    # (1280, 720, 1680_1.25): 0.8=(1280, 720) 1=(1280, 720) no_set=(1282, 722)
    # (1280, 720, 1680_1.0) : 0.8=(1280, 720) 1 =(1280, 720) no_set=(1280, 720)
    # (1400, 1000, 1680_1.25): 0.8=(1400, 1000) 1=(1400, 1000) no_set=(1363, 859)
    # (1400, 1000, 1680_1.0) : 0.8=(1400, 1000) 1=(1400, 1000) no_set=(1400, 1000)
    chrome_options.add_argument(f'--force-device-scale-factor={browser_display_ratio}')

    # chrome_options.add_argument("--window-size=1200x990")

    # DeprecationWarning: use options instead of chrome_options chrome_options=chrome_options
    browser = webdriver.Chrome(
        executable_path=CHROME_DRIVER_EXE_PATH,
        options=chrome_options)
    # browser = webdriver.PhantomJS(executable_path=PHANTOM_JS_EXE_PATH)  # deprecated...
    browser.set_window_size(1920, 1080)
    browser.set_window_position(900, 150)
    # browser.maximize_window()
    # print('window size = ' + str(browser.get_window_size()))

    # browser.get('chrome://settings/')
    # browser.execute_script('chrome.settingsPrivate.setDefaultZoom(0.7);')  # 字變小，但可用範圍一樣小

    # browser.execute_script("document.body.style.zoom='90%'")

    # fire_fox_options = webdriver.FirefoxOptions()
    # fire_fox_options.headless = True
    # browser = webdriver.Firefox(
    #     executable_path=FIREFOX_DRIVER_EXE_PATH,
    #     # options=fire_fox_options
    # )
    # browser.get(r'https://csms2.sfaa.gov.tw/lcms/simpleCaptcha/captcha')
    # print(browser.page_source)  # JLZX  3b161dde395196200ff8e09c22ed4812
    # print(browser.get_cookies())  # {'name': 'captcha', 'value': '3b161dde395196200ff8e09c22ed4812', 'path': '/', 'domain': 'csms2.sfaa.gov.tw', 'expiry': 1572193308, 'secure': False, 'httpOnly': False}
    # return

    # ================= 第 1 網頁 =====================

    browser_tab_list = []

    lc_tab_idx, lc_cookies = login_home_therapy_page(browser, create_theropy_lc_profile())
    browser_tab_list.append(home_therapy_web_manager.BrowserTab(HomeTherapyCompany.LC, lc_tab_idx, lc_cookies))

    # browser.execute_script("document.body.style.zoom='75%'")  # 只有字變小，可用範圍沒有變大
    # zoom_out(browser)  # ctrl-+ not work
    # print(browser.get_window_size())

    # Formal login flow. But now login is too easy, skip for now
    # is_need_re_login = True
    #
    # if os.path.exists('cookie.json'):
    #     f1 = open('cookie.json')
    #     print("cookie.json loaded")
    #     cookie = f1.read()
    #     cookie = json.loads(cookie)
    #     for c in cookie:
    #         # print(c)
    #         if 'expiry' in c:
    #             del c['expiry']
    #         browser.add_cookie(c)
    #
    #     browser.get(TAKE_CARE_HOME_PAGE_URL)
    #     if "captcha" in browser.page_source:
    #         print('cookie expired.  Need re-login')
    #     else:
    #         is_need_re_login = False
    #
    # else:
    #     print("cookie.json not exist")
    #
    # # ActionChains(browser).move_by_offset(104, 40).click().perform()  # 家裡電腦：x=103可關,104不行 y=29可關,30不行
    #
    # # y: 24 可關公告  25不行
    # # 從網址列最下緣開始算Y後，要扣8才當作標下去 click
    # # ActionChains(browser).move_by_offset(300, 24).click().perform()  # 公司電腦：x: 163 可以關閉公告  164不行
    #
    # is_account_filled = False
    #
    # if is_need_re_login:
    #     while True:
    #         # ESC
    #         ActionChains(browser).key_down(Keys.ESCAPE).key_up(Keys.ESCAPE).perform()
    #
    #         if not is_account_filled:
    #             user_name_input = browser.find_element_by_id('username')
    #             user_name_input.send_keys(CC_ACCOUNT)
    #             is_account_filled = True
    #
    #         password_input = browser.find_element_by_id('password')
    #         password_input.send_keys(PASSWORD)
    #
    #         browser.save_screenshot("login_page.png")
    #         captcha = get_captcha("login_page.png")
    #         os.remove("login_page.png")
    #
    #         captcha_input = browser.find_element_by_id('captcha')
    #         captcha_input.send_keys(captcha)
    #         print(browser.get_cookies())
    #         print("========================")
    #
    #         time.sleep(5)
    #
    #         if True or len(captcha) == 4:
    #             print('ready login')
    #
    #             login_btn = browser.find_element_by_name('_action_login')
    #             login_btn.click()
    #
    #             try:
    #                 alert = browser.switch_to.alert
    #                 print(alert.text)  # '驗証碼錯誤'
    #                 # time.sleep(5)
    #                 alert.accept()
    #                 # time.sleep(5)
    #                 # print('============== retry login')
    #                 continue
    #             except NoAlertPresentException:
    #                 browser.get(TAKE_CARE_HOME_PAGE_URL)
    #                 break
    #         else:
    #             print("captcha len != 4, refresh!!")
    #             browser.get(TAKE_CARE_HOME_PAGE_URL)
    #             is_account_filled = False
    #             continue
    # print(browser.page_source)

    # ui.set_cc_browser(browser)

    browser.delete_all_cookies()

    # Open a new window for LC company
    browser.execute_script(f"window.open('{TAKE_CARE_HOME_PAGE_URL}');")
    browser.switch_to.window(browser.window_handles[1])

    cc_tab_idx, cc_cookies = login_home_therapy_page(browser, create_theropy_cc_profile())
    browser_tab_list.append(home_therapy_web_manager.BrowserTab(HomeTherapyCompany.CC, cc_tab_idx, cc_cookies))

    if False:
        # test start
        browser.switch_to.window(browser.window_handles[0])

        browser.delete_all_cookies()

        for cookie in cc_cookies:
            if 'expiry' in cookie:
                del cookie['expiry']
            browser.add_cookie(cookie)
            # print(cookie)
        print("added cookies")
        print("=================")

        tmp_cookies = browser.get_cookies()
        for cookie in tmp_cookies:
            print(cookie)

        browser.get(
            'https://csms2.sfaa.gov.tw/lcms/qd/filterQd111/?doQuery=true&qdList1=yes&caseno=&serialno=&name=個案名稱1&idno=&noteDt1=&noteDt2=&processDt1=&processDt2=&qcntcode=&twnspcode=&vilgcode=&qinformCntcode=&informTwnspcode=&informVilgcode=&_qd111Resp=&_qd111adjHint=&_qd111pi400Hint=&_qlistQd111Export=&_qsortNotNull=&qsort=&qorder=1&limit=10&offset=0&order=asc&_=1570458820317');

        # print("=================")
        # tmp_cookies = browser.get_cookies()
        # for cookie in tmp_cookies:
        #     print(cookie)

        # return
        time.sleep(3)
        #
        browser.delete_all_cookies()

        for cookie in cc_cookies:
            if 'expiry' in cookie:
                del cookie['expiry']
            browser.add_cookie(cookie)

        browser.get('https://csms2.sfaa.gov.tw/lcms/qd/filterQd120/919606075?iframe_tab_params=cascade+previous.tabid%3DROOT&iframe_tab_params=cascade%28nofix%3Dtrue%29+entrance%3D384520077&%3Fiframe_tab_params=cascade+tabid%3DTREE-384520077&qd111id=972470871&ca100id=362191551&qdperms=true&perms=true&limit=10&offset=0&order=asc&_=1571579327737')

        print("=================")
        tmp_cookies = browser.get_cookies()
        for cookie in tmp_cookies:
            print(cookie)
        print("get url again...")
        return

    ui.set_home_therapy_web_manager(home_therapy_web_manager.HomeTherapyWebManager(browser, browser_tab_list))

    print("LC & CC tabs are ready")

    return

    # =========== 直接搜尋資料庫 =================
    # SEARCH_CLIENT_URL = r'https://csms2.sfaa.gov.tw/lcms/qd/filterQd111/?doQuery=true&qdList1=yes&caseno=&serialno=&name=個案名稱1&idno=&noteDt1=&noteDt2=&processDt1=&processDt2=&qcntcode=&twnspcode=&vilgcode=&qinformCntcode=&informTwnspcode=&informVilgcode=&_qd111Resp=&_qd111adjHint=&_qd111pi400Hint=&_qlistQd111Export=&_qsortNotNull=&qsort=&qorder=1&limit=10&offset=0&order=asc&_=1570458820317'
    # browser.get(SEARCH_CLIENT_URL)
    # print(browser.page_source)
    # time.sleep(5000)

    # 打完紀錄送出
    UPDATE_DB_URL1 = r'https://csms2.sfaa.gov.tw/lcms/qd/saveQd120Batch?qd120.id=&qd120.ca100=919606075&qd120ServdtNum=1&qd120dt.1.servDt=108%2F10%2F10&qd120ServtimeNum=1&qd120time.1.startHh=11&qd120time.1.startMm=12&qd120time.1.endHh=13&qd120time.1.endMm=14&qd120.servUsers=882319126&qd120.noAa09=&qd120.commt=%E8%A8%98%E9%8C%84%E9%8C%84&qd120Num=1&qd120.1.qd100=285889258&qd120.1.addr1=&qd120.1.addr2=&qd120.1.carNum=&qd120.1.driver=&qd120.1.amount=1&qd120.1.price=4500.0&qd120.1.stype=1&qd120.1.lastServ=false&qd120.1.missedVisit=false&qd120Num=2&qd120.2.qd100=285889260&qd120.2.addr1=&qd120.2.addr2=&qd120.2.carNum=&qd120.2.driver=&qd120.2.amount=0&qd120.2.price=4500.0&qd120.2.stype=1&qd120.2.lastServ=false&qd120.2.missedVisit=false'
    UPDATE_DB_URL2 = r'https://csms2.sfaa.gov.tw/lcms/qd/filterQd120/919606075?iframe_tab_params=cascade+previous.tabid%3DROOT&amp;iframe_tab_params=cascade%28nofix%3Dtrue%29+entrance%3D384520077&amp;%3Fiframe_tab_params=cascade+tabid%3DTREE-384520077&amp;qd111id=961129015&amp;ca100id=919606075&amp;qdperms=true/?&ca100id=919606075&perms=true&stype=&status=&servDt1=&servDt2=&qd120Pi400=&servUserName=&aa10Status=&limit=10&offset=0&order=asc&_=1570545378711'
    # saveQd120Batch = iframe_tab_params=cascade+previous.tabid%3DROOT&amp;
    # iframe_tab_params=cascade%28nofix%3Dtrue%29+entrance%3D384520077&amp;
    # %3Fiframe_tab_params=cascade+tabid%3DTREE-384520077&amp;
    # qd111id=961129015&amp;
    # ca100id=919606075&amp;
    # qdperms=true/?&ca100id=919606075&perms=true&stype=&status=&servDt1=&servDt2=&qd120Pi400=&servUserName=&aa10Status=&limit=10&offset=0&order=asc&_=1570545378711

    # DB_query_result = {"total":2,"rows":
    # [{"qd111id":961129015,
    #   "ca100id":919606075,
    #   "noteDt":"108/09/25",
    #   "processDt":"108/09/26",
    #   "accept":"可提供服務",
    #   "servStatus":"服務中",
    #   "idno":"A232782829",
    #   "name":"個案名稱3",
    #   "caseno":"108X17105",
    #   "caseStatusCh":"<div title='開案服務中' style='color:#af6832;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin:auto;'><b>開案服務中<\u002fb><\u002fdiv>",
    #   "mainGb200":"新北市長期照顧管理中心-總站",
    #   "pi400aOrg":"[新北特約](A)財團法人台北市立心慈善基金會",
    #   "mainUser":"使用者1",
    #   "assign":"[新北特約](A)財團法人台北市立心慈善基金會",
    #   "servCode":"CA01[IADLs復能照護--居家],CA03[ADLs復能照護--居家]",
    #   "twnspname":"臺北市信義區安康里",
    #   "informTwnspname":"新北市三重區碧華里",
    #   "processCommt":"9/27(五) 17:30 使用者1CA03:1初次訪視\r\n10/2(三) 17:30 職能治療師名稱 CA01:2初次訪視《系統訊息》因照專於計畫異動後，貴單位仍可提供個案原照顧計畫的服務項目。為維持服務不中斷，系統於原額度下自動延續服務。若計畫異動後，可用額度較原計畫高，則需請照專手動照會新增的額度。"},

    #  {"qd111id":931302802,
    #  "ca100id":919606075,
    #  "noteDt":"108/09/25",
    #  "processDt":"108/09/26",
    #  "accept":"可提供服務",
    #  "servStatus":"服務中",
    #  "idno":"A232782829",
    #  "name":"個案名稱3",
    #  "caseno":"108X17105",
    #  "caseStatusCh":"<div title='開案服務中' style='color:#af6832;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin:auto;'><b>開案服務中<\u002fb><\u002fdiv>",
    #  "mainGb200":"新北市長期照顧管理中心-總站",
    #  "pi400aOrg":"[新北特約](A)財團法人台北市立心慈善基金會",
    #  "mainUser":"使用者",
    #  "assign":"[新北特約](A)財團法人台北市立心慈善基金會",
    #  "servCode":"CA01[IADLs復能照護--居家],CA03[ADLs復能照護--居家]",
    #  "twnspname":"臺北市信義區安康里",
    #  "informTwnspname":"新北市三重區碧華里",
    #  "processCommt":"9/27(五) 17:30 孫紀鈺0922968317CA03:1初次訪視\r\n10/2(三) 17:30 職能治療師 CA01:2初次訪視"}]}

    # ================= 第 2 網頁 =====================

    # print(browser.get_cookies())

    # Save already login cookie. But now login is too easy, skip for now
    # json_cookies = json.dumps(browser.get_cookies())
    # with open('cookie.json', 'w') as f:
    #     f.write(json_cookies)

    if CC:
        tree_107 = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_11_switch']"))
        )
    else:
        tree_107 = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_3_switch']"))
        )
    tree_107.click()

    if CC:
        btn_101 = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_13_ico']"))
        )
    else:
        btn_101 = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_5_ico']"))
        )
    btn_101.click()

    iframe_search_engine = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.ID, "TREE-384520077"))
    )

    iframe_offset = (iframe_search_engine.location['x'], iframe_search_engine.location['y'])
    browser.switch_to.frame(iframe_search_engine)  # (主分頁) 101-照會列表

    if CC:
        client_name = '個案名稱1'
    else:
        client_name = '個案名稱2'

    client_name_field = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.NAME, "name"))
    )
    client_name_field.send_keys(client_name)

    # 查詢
    search_btn = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.NAME, "search"))
    )
    search_btn.click()

    id_btn = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//a[@href='javascript: void(0)' and contains(text(), '{client_name}')]"))
    )
    id_btn.click()

    browser.switch_to.default_content()  # back to root level

    level_1_iframe = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id, 'QD-LIST1-')]"))
    )
    browser.switch_to.frame(level_1_iframe)  # Level-1-iframe: 「個案服務: XXX」

    WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '服務項目明細')]"))
    )

    client_btn_in_level_1 = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '個案服務:')]"))
    )
    client_btn_in_level_1.click()

    level_2_iframe = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id, 'QD-LIST1-2-')]"))
    )
    browser.switch_to.frame(level_2_iframe)  # Level-2-iframe: 個案服務: XXX

    service_record_tab = WebDriverWait(browser, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, "//*[contains(text(), '服務紀錄')]"))
    )[1]
    service_record_tab.click()

    service_record_table = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.ID, "searchQd120Result"))
    )
    table_rows = service_record_table.find_elements_by_tag_name("tr")  # get all of the rows in the table

    # 找到『服務人員』 和 『編輯/檢視』 在哪一行
    serv_person_col_idx = None
    header_row = table_rows[0]
    header_title_fields = header_row.find_elements_by_tag_name("th")
    for idx, table_header_col in enumerate(header_title_fields):
        if "服務人員" in table_header_col.text:
            serv_person_col_idx = idx
            break

    if serv_person_col_idx is None:
        return
    action_btn_col_idx = len(header_title_fields) - 1

    # 用這兩行來找第一列符合『治療師名稱』 和 『可編輯狀態』
    table_body_rows = table_rows[1:]
    for body_row in table_body_rows:
        body_fields = body_row.find_elements_by_tag_name("td")
        service_person_filed = body_fields[serv_person_col_idx]
        show_serv_record_detail_btn = body_fields[action_btn_col_idx]
        if service_person_filed.text == "治療師名稱" and show_serv_record_detail_btn.text == "編輯":
            # TODO: make sure the service datetime is matched
            # found last added record
            show_serv_record_detail_btn.click()

            edit_serv_record_popup = WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "modal-footer"))
            )
            time.sleep(0.5)  # the popup is showing animation...
            size = edit_serv_record_popup.size
            location = edit_serv_record_popup.location

            html_crop_right_bottom_pos = (iframe_offset[0] + location['x'] + size['width'], iframe_offset[1] + location['y'] + size['height'])
            screenshot_crop_right_bottom_pos = from_html_coor_to_screenshot_coor(html_crop_right_bottom_pos)
            tmp_img = get_screenshot(browser)
            browser.save_screenshot("ttt.png")
            final_client_record_screenshot = tmp_img.crop((iframe_offset[0], iframe_offset[1], screenshot_crop_right_bottom_pos[0], screenshot_crop_right_bottom_pos[1]))
            final_client_record_screenshot.save("final_client_record_screenshot.png")
            press_key(browser, Keys.ESCAPE)
            return

    print("No matched record found...")
    # print(browser.page_source)
    return

    global screenshot_top, screenshot_bottom, screenshot_left, screenshot_right

    # No need opencv-template-matching to do this...
    tree_pos_107 = wait_until_template_shown(browser, r"template\template_107.png", "QD-107新制服務區", 1)[1]
    click_screen(browser, from_screenshot_coor_to_click_coor(tree_pos_107), y_offset=7)  # open QD-107
    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_101_2.png", "101-照會列表", 1)[1]), y_offset=7)

    click_screen(browser, from_screenshot_coor_to_click_coor(pos_107), y_offset=7)  # close QD-107
    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_client_name.png", "案主姓名", 1.5)[1]), x_offset=100, y_offset=7)

    pyperclip.copy('個案名稱1')
    ActionChains(browser).key_down(Keys.SHIFT).send_keys(Keys.INSERT).key_up(Keys.SHIFT).perform()

    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_search.png", "查詢")[1]))

    # press_key(browser, Keys.PAGE_DOWN)

    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_name_id.png", "姓名-身分證號", 0.5)[1]), y_offset=72)

    wait_until_template_shown(browser, r"template\template_service_item_detail.png", "服務項目明細")  # ??? need this ???

    pos = wait_until_template_shown(browser, r"template\template_client_tab.png", "個案服務:", 2)[1]
    screenshot_left, screenshot_top = int(pos[0] - 5), int(pos[1] - 8)
    click_screen(browser, from_screenshot_coor_to_click_coor(pos))

    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_service_record.png", "服務紀錄")[1]))

    press_key(browser, Keys.PAGE_DOWN)

    click_screen(browser, from_screenshot_coor_to_click_coor(wait_until_template_shown(browser, r"template\template_edit.png", "編輯", 0.5)[1]))

    img, pos = wait_until_template_shown(browser, r"template\template_delete_2.png", "刪除", 2)
    screenshot_right, screenshot_bottom = int(pos[0] + 58), int(pos[1] + 27)

    # img, pos = wait_until_template_shown(browser, r"template\template_save.png", "儲存", 2)
    # screenshot_right, screenshot_bottom = int(pos[0] + 100), int(pos[1] - 51)

    final_client_record_screenshot = img[screenshot_top:screenshot_bottom, screenshot_left:screenshot_right]
    cv2.imwrite("final_client_record_screenshot.png", final_client_record_screenshot)

    press_key(browser, Keys.ESCAPE)

    os.remove("tmp.png")
    return

    # 107新制服務區
    ActionChains(browser).move_by_offset(84, 208).click().perform()
    # ActionChains(browser).move_by_offset(20, 62).click().perform()  # 62=TopTitle, 63=empty, (20, 80)=HideLeftMenu

    # 101-照會列表
    ActionChains(browser).move_by_offset(16, 40).click().perform()
    time.sleep(0.5)  # depends on network

    # 案主姓名
    ActionChains(browser).move_by_offset(277, 38).click().perform()
    print("click(20, 20)")
    browser.save_screenshot("ttttt.png")
    time.sleep(10000)

    # btn = browser.find_element_by_id('west-ztree_12_span')
    # print(btn.location)
    # while True:
    #     action = ActionChains(browser)
    #     x = random.randrange(0, 500)
    #     y = random.randrange(0, 500)
    #     action.move_by_offset(x, y)
    #     # action.move_to_element(btn)  #.move_by_offset(100, 400)
    #     action.click()
    #     action.perform()
    #     time.sleep(0.5)
    #
    # time.sleep(10000)
    # browser.find_element_by_id('west-ztree_10').click()

    # TODO: survey how use wait.until
    # WebElement el = wait.until(ExpectedConditions.visibilityOfElementLocated(
    #     By.xpath(".//span[@class = 'middle' and contains(text(), 'Next')]")));
    # el.click();

    # from selenium.webdriver.support.ui import WebDriverWait as wait
    # wait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='No']"))).click()

    browser.find_element_by_id('west-ztree_12_span').click()
    time.sleep(5)
    # print(browser.page_source)

    browser.find_element(By.XPATH, "//span[contains(text(), '首頁')]").click()
    time.sleep(5)

    browser.find_elements(By.XPATH, "//span[contains(text(), '101-照會列表')]")[1].click()
    time.sleep(5000)
    # for it in ll:
    #     print('id=' + it.id)
    # print('len=' + str(len(ll)))

    item = browser.find_element_by_class_name(r'ui-state-default ui-corner-top ui-tabs-selected ui-state-active')
    item.click()
    print('found 101-照會列表')
    time.sleep(5)

    client_name = browser.find_element_by_name('name')  # Dynamic generated content not found...
    client_name.send_keys(r'個案名稱1')

    time.sleep(5)

    search_btn = browser.find_element_by_name('search')
    search_btn.click()

    time.sleep(15)

    time.sleep(30)
    # except NoSuchElementException:
    #     print('NoSuchElementException')

    browser.close()
    browser.quit()
    print("end")


def login_home_therapy_page(browser, home_therapy_login_profile: HomeTherapyLoginProfile):
    print('login_home_therapy_page for ' + str(home_therapy_login_profile.home_therapy_company))

    cur_url = browser.current_url
    if cur_url != HOME_THERAPY_GOVERN_HOME_PAGE_URL:
        browser.delete_all_cookies()
        browser.get(HOME_THERAPY_GOVERN_HOME_PAGE_URL)

    account = home_therapy_login_profile.account
    pw = home_therapy_login_profile.pw
    captcha = None

    cookie = browser.get_cookie('captcha')
    browser.delete_cookie('captcha')
    if home_therapy_login_profile.home_therapy_company == HomeTherapyCompany.CC:
        captcha = "JLZX"
        cookie['value'] = '3b161dde395196200ff8e09c22ed4812'  # JLZX : CC
    elif home_therapy_login_profile.home_therapy_company == HomeTherapyCompany.LC:
        captcha = "9BN6"
        cookie['value'] = 'f3608d3d7d22dd2589e9dadaf91b8011'  # 9BN6 : LC
    if 'expiry' in cookie:
        del cookie['expiry']
    browser.add_cookie(cookie)

    browser.get(rf'https://csms2.sfaa.gov.tw/lcms?loginAttempt=true&username={account}&password={pw}&captcha={captcha}&_action_login=%E7%99%BB+%E5%85%A5')
    browser.get(TAKE_CARE_HOME_PAGE_URL)

    window_idx = None
    for idx, window_handle in enumerate(browser.window_handles):
        if browser.current_window_handle == window_handle:
            window_idx = idx
            break

    return window_idx, browser.get_cookies()


def zoom_out(browser, step=1):
    ActionChains(browser).key_down(Keys.LEFT_CONTROL).key_down(Keys.SUBTRACT).key_up(Keys.SUBTRACT).key_up(Keys.LEFT_CONTROL).perform()
    # ActionChains(browser).key_down(Keys.LEFT_CONTROL).send_keys(Keys.SUBTRACT).key_up(Keys.LEFT_CONTROL).perform()


def click_screen(browser, pos, x_offset=0, y_offset=0):
    x, y = pos[0] + x_offset, pos[1] + y_offset
    ActionChains(browser).move_by_offset(x, y).click().move_by_offset(-x, -y).perform()


def wait_until_template_shown2(browser, template_file_name, template_name=None, delay_time_before_start=0):
    while True:
        browser.save_screenshot("tmp.png")
        img_rgb = cv2.imread("tmp1.png")
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        template = cv2.imread(template_file_name, 0)
        w, h = template.shape[::-1]
        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
        threshold = 0.7
        loc = np.where(res >= threshold)

        # TODO: How to find max index with pythonic way?
        largest_pt = ()
        largest_value = -1
        for pt in zip(*loc[::-1]):
            candidate = res[pt[::-1]]
            if candidate > largest_value:
                largest_value = candidate
                largest_pt = pt
        if largest_value != -1:
            print(f'{template_name} = ({largest_pt[0]}, {largest_pt[1]})')
            return [img_rgb, largest_pt]
        # cv2.rectangle(img_rgb, pt, (largest_pt[0] + w, largest_pt[1] + h), (7, 249, 151), 2)
            # return [img_rgb, pt]
        # cv2.imshow('Detected', img_rgb)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()
        # break


global x_ratio, y_ratio


def wait_until_template_shown(browser, template_file_name, template_name=None, delay_time_before_start=0):
    if delay_time_before_start != 0:
        time.sleep(delay_time_before_start)

    while True:
        # efficient way ? Image.open(StringIO(base64.decodestring(driver.get_screenshot_as_base64())))
        browser.save_screenshot("tmp.png")
        img_rgb = cv2.imread("tmp.png")
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        template = cv2.imread(template_file_name, cv2.IMREAD_GRAYSCALE)

        found = None

        global browser_display_ratio

        # TODO: 可能有多個編輯  找Y最高，或是名字是治療師名稱

        for scale in np.linspace(1 / browser_display_ratio, 0.2 / browser_display_ratio, 20):
            resized_image = imutils.resize(img_gray, width=int(img_gray.shape[1] * scale))
            ratio = img_gray.shape[1] / float(resized_image.shape[1])

            if resized_image.shape[0] < template.shape[0] or resized_image.shape[1] < template.shape[1]:
                break

            result = cv2.matchTemplate(resized_image, template, cv2.TM_CCOEFF_NORMED)
            (_, maxVal, _, maxLoc) = cv2.minMaxLoc(result)

            if found is None or maxVal > found[0]:
                found = (maxVal, maxLoc, ratio)

        (maxValue, maxLoc, ratio_max) = found
        if maxValue > 0.5:
            pos = (int(maxLoc[0] * ratio_max), int(maxLoc[1] * ratio_max))

            # draw a bounding box around the detected result and display the image
            (startX, startY) = pos
            (endX, endY) = (int((maxLoc[0] + template.shape[1]) * ratio_max), int((maxLoc[1] + template.shape[0]) * ratio_max))
            cv2.rectangle(img_rgb, (startX, startY), (endX, endY), (0, 0, 255), 2)
            cv2.imwrite("tmp_locate.png", img_rgb)
            # cv2.imshow("Image", img_rgb)
            # cv2.waitKey(0)

            # resized_pos = int(pos[0] / browser_display_ratio), int(pos[1] / browser_display_ratio)

            # print(f'pos={pos} x_ratio={x_ratio} y_ratio={y_ratio}  / browser_display_ratio={browser_display_ratio}')
            # print(f'template found: ({resized_pos[0]}, {resized_pos[1]}) {template_name}')
            print(f'template found: ({pos[0]}, {pos[1]}) {template_name}')

            # return [img_rgb, resized_pos]
            return [img_rgb, pos]

    return None


def from_screenshot_coor_to_click_coor(screenshot_pos):
    global browser_display_ratio
    return int(screenshot_pos[0] / browser_display_ratio), int(screenshot_pos[1] / browser_display_ratio)


def from_html_coor_to_screenshot_coor(html_pos) -> (int, int):
    global browser_display_ratio
    return int(html_pos[0] * browser_display_ratio), int(html_pos[1] * browser_display_ratio)


def get_screenshot(browser):
    return Image.open(BytesIO(browser.get_screenshot_as_png()))


def prepare_ui():
    app = wx.App()
    ui = MainWindow(None, "Home Therapy Machine")
    app.MainLoop()


class ServicerProfile:
    # service_info is dict['account': str / 'pw': str / 'clients': list]
    def __init__(self, servicer_name, servicer_info_list: list):
        self.servicer_name = servicer_name
        self.servicer_info_list = servicer_info_list


class LoginProfile:
    def __init__(self, account: str, pw: str):
        self.account = account
        self.pw = pw


def load_servicer_profile() -> ServicerProfile:
    cc_client_list = ["個案名稱1"]
    cc_info = {'company': HomeTherapyCompany.CC, 'account': constants.CC_ACCOUNT, 'pw': constants.CC_PASSWORD, 'clients': cc_client_list}

    lc_client_list = ["個案名稱2"]
    lc_info = {'company': HomeTherapyCompany.LC, 'account': constants.LC_ACCOUNT, 'pw': constants.LC_PASSWORD, 'clients': lc_client_list}

    # servicer_info_list = [cc_info, lc_info]
    servicer_info_list = [lc_info, cc_info]

    return ServicerProfile("治療師名稱", servicer_info_list)


global browser_display_ratio


def main():
    # img = ImageGrab.grab()
    # print(img.size)

    app = wx.App()
    resolution = wx.GetDisplaySize()
    # print(resolution)
    # print(wx.GetDisplayPPI())

    global x_ratio, y_ratio
    if platform.node() == "yy-i7":
        x_ratio, y_ratio = float(resolution[0]) / 1680, float(resolution[1]) / 1050
    else:
        x_ratio, y_ratio = float(resolution[0]) / 1920, float(resolution[1]) / 1080

    global browser_display_ratio
    system_display_ratio = [x_ratio, y_ratio]
    browser_display_ratio = 0.9
    # browser_display_ratio = 1
    # print(f'system_ratio={system_display_ratio}  browser_display_ratio={browser_display_ratio}')

    servicer_profile = load_servicer_profile()

    client_tuple_list = []
    for servicer_info in servicer_profile.servicer_info_list:
        client_tuple_list.append((servicer_info['company'], servicer_info['clients']))

    ui = MainWindow(None, client_tuple_list, system_display_ratio=system_display_ratio, browser_display_ratio=browser_display_ratio)

    web_page_loading_thread = threading.Thread(target=load_web_page, args=[ui, browser_display_ratio])
    web_page_loading_thread.start()

    app.MainLoop()

    # prepare_ui()


if __name__ == '__main__':
    main()

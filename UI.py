import json
import os
import re
import sys
import time
import threading
from io import BytesIO
import win32clipboard

import cv2 as cv
import numpy as np
import imutils
import pyperclip

import wx
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from wx import BoxSizer
from wx import Button
from wx import StaticBox
from wx import StaticBoxSizer
from wx import StaticLine
from wx import StaticText
from wx import TextCtrl

from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from PIL import Image

import HomeTherapy
import constants
from constants import HomeTherapyCompany
from constants import TAKE_CARE_HOME_PAGE_URL
import home_therapy_web_manager
# from home_therapy_web_manager import HomeTherapyWebManager
# from HomeTherapy import HomeTherapyCompany

ID_ABOUT = 101
ID_OPEN = 102
ID_SAVE = 103
ID_BUTTON1 = 300
ID_EXIT = 200

# qd120_servUsers
DB_ID_CO0168 = r'864119605'
DB_ID_LC035 = r'882319126'

TSAI_ID = r'362191551'  # CO00168個案名稱1
HUANG_ID = r'919606075'  # LC035個案名稱3
LI_ID = r'874056375'  # LC035個案名稱2

OT_COMPANY_JIA_JIN = 1  #
OT_COMPANY_LE_CHIN = 2  #

CLIENT_LIST = {
    #         [CA01, CA03] = {-1=不存在, 0=不是, 1=是}
    u"個案名稱1": ([-1, 1], TSAI_ID, OT_COMPANY_JIA_JIN),
    u"個案名稱3": ([1, 0], HUANG_ID, OT_COMPANY_LE_CHIN),
    u"個案名稱2": ([1, -1], LI_ID, OT_COMPANY_LE_CHIN)
}


class BrowserTab:
    def __init__(self, home_therapy_company, tab_idx: int, cookies: dict):
        self.home_therapy_company = home_therapy_company
        self.tab_idx = tab_idx
        self.cookies = cookies


class Test:
    def set_func(self, func):
        self.func = func


class MainWindow(wx.Frame):

    def take_func1(self, func):
        self.func1 = func

    def take_func2(self, func):
        self.func2 = func

    def call_func(self):
        self.func1()

    # client_tuple_list: [(HomeTherapyCompany, ["個案名字1", "個案名字2"]), ...]
    def __init__(self, parent, client_tuple_list, system_display_ratio=1, browser_display_ratio=1):
        # based on a frame, so set up the frame
        wx.Frame.__init__(self, parent, title="Home Therapy Machine", size=(1200, 800))
        panel = wx.Panel(self)
        self.client_tuple_list = client_tuple_list
        self.system_display_ratio = system_display_ratio
        self.browser_display_ratio = browser_display_ratio

        self.CreateStatusBar()  # A Statusbar in the bottom of the window
        menu_bar = self.createMenuBar()
        self.SetMenuBar(menu_bar)  # Adding the MenuBar to the Frame content.
        # Note - previous line stores the whole of the menu into the current object

        # HomeTherapy ++
        add_client_text_ctrl = TextCtrl(panel, size=(89, 25), style=wx.TE_PROCESS_ENTER | wx.TE_CENTER)
        add_client_text_ctrl.Bind(wx.EVT_CHAR, self.on_add_new_client)
        client_list_sizer = BoxSizer(wx.VERTICAL)
        client_list_sizer.Add(add_client_text_ctrl, 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)
        client_list_sizer.AddSpacer(10)
        # client_list_sizer.Add(Button(panel, label=u"個案名稱1"), 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)
        # client_list_sizer.Add(Button(panel, label=u"個案名稱3"), 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)
        # client_list_sizer.Add(Button(panel, label=u"個案名稱2"), 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)
        self.add_client_text_ctrl = add_client_text_ctrl
        self.client_list_sizer = client_list_sizer

        # TODO: may beautify UI with ObjectListView?
        # http://www.blog.pythonlibrary.org/2009/12/23/wxpython-using-objectlistview-instead-of-a-listctrl/

        # {HomeTherapyCompany: {1: "個案名字1"，"個案名字2"}，...}
        self.dict_of_client_row_dict = {}

        lc_on_client_item_selected = lambda event: self.on_client_item_selected(event, HomeTherapyCompany.LC)
        cc_on_client_item_selected = lambda event: self.on_client_item_selected(event, HomeTherapyCompany.CC)

        lc_on_client_item_double_clicked = lambda event: self.on_client_item_double_clicked(event, HomeTherapyCompany.LC)
        cc_on_client_item_double_clicked = lambda event: self.on_client_item_double_clicked(event, HomeTherapyCompany.CC)

        # self.tests = []

        # (HomeTherapyCompany, ["個案名字1", "個案名字2"])
        for ot_company, client_list in client_tuple_list:

            client_list_ctrl = wx.ListCtrl(panel, -1, style=wx.LC_REPORT)
            client_list_ctrl.InsertColumn(0, u"姓名", width=100)
            client_list_ctrl.InsertColumn(1, u"狀態1", wx.LIST_FORMAT_RIGHT, 40)
            client_list_ctrl.InsertColumn(2, u"狀態2", wx.LIST_FORMAT_RIGHT, 40)

            # [Fucking Bug] variable name 'ot_company' will be overwritten by the last assignment, while the lambda is called back
            # on_client_item_selected_tmp = lambda event: self.on_client_item_double_clicked(event, ot_company)
            # client_list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, on_client_item_selected_tmp, client_list_ctrl)
            # on_client_item_double_clicked_tmp = lambda event: self.on_client_item_double_clicked(event, ot_company)
            # client_list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, on_client_item_double_clicked_tmp, client_list_ctrl)

            if ot_company == HomeTherapyCompany.LC:
                client_list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, lc_on_client_item_selected, client_list_ctrl)
                client_list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, lc_on_client_item_double_clicked, client_list_ctrl)
            else:
                client_list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, cc_on_client_item_selected, client_list_ctrl)
                client_list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, cc_on_client_item_double_clicked, client_list_ctrl)

            # [Fucking Bug] simulating this bug...
            # test = Test()
            # test.set_func(on_client_item_selected_tmp)
            # self.tests.append(test)

            client_row_dict = {}
            for client_name in client_list:
                index = client_list_ctrl.InsertItem(client_list_ctrl.GetItemCount(), client_name)
                client_list_ctrl.SetItem(index, 1, u'O')
                client_list_ctrl.SetItem(index, 2, u'X')
                client_row_dict[index] = client_name

            self.dict_of_client_row_dict[ot_company] = client_row_dict

            # index = client_list_ctrl.InsertItem(client_list_ctrl.GetItemCount(), u"個案名稱1")
            # client_list_ctrl.SetItem(index, 1, u'O')
            # client_list_ctrl.SetItem(index, 2, u'X')
            # index = client_list_ctrl.InsertItem(client_list_ctrl.GetItemCount(), u"個案名稱3")
            # client_list_ctrl.SetItem(index, 1, u'O')
            # client_list_ctrl.SetItem(index, 2, u'X')
            # index = client_list_ctrl.InsertItem(client_list_ctrl.GetItemCount(), u"個案名稱2")
            # client_list_ctrl.SetItem(index, 1, u'O')
            # client_list_ctrl.SetItem(index, 2, u'X')

            client_list_sizer.Add(client_list_ctrl, 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)

        evaluation_comment_text_ctrl = TextCtrl(panel, size=(600, 300), style=wx.TE_MULTILINE)
        evaluation_comment_text_ctrl.Bind(wx.EVT_CHAR, self.on_enter)

        text_ctrl_font = wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Consolas')
        evaluation_comment_text_ctrl.SetFont(text_ctrl_font)

        self.evaluation_comment_text_ctrl = evaluation_comment_text_ctrl

        # +++++++++++++++++++++++++ Date ++++++++++++++++++++++++++++++
        date_month_text_ctrl = TextCtrl(panel, size=(25, 18), style=wx.TE_PROCESS_ENTER | wx.TE_CENTER)
        date_splitter = StaticText(panel, label=u"／")
        date_day_text_ctrl = TextCtrl(panel, size=(25, 18), style=wx.TE_PROCESS_ENTER | wx.TE_CENTER)
        date_month_text_ctrl.Bind(wx.EVT_CHAR, self.on_add_new_client)
        date_day_text_ctrl.Bind(wx.EVT_CHAR, self.on_add_new_client)

        date_box_sizer = BoxSizer()
        date_box_sizer.Add(date_month_text_ctrl, 0, wx.RIGHT, border=5)
        date_box_sizer.Add(date_splitter, 0)
        date_box_sizer.Add(date_day_text_ctrl, 0, wx.LEFT, border=5)

        self.date_month_text_ctrl = date_month_text_ctrl
        self.date_day_text_ctrl = date_day_text_ctrl
        # -------------------------- Date ----------------------------

        # +++++++++++++++++++++++++ Time ++++++++++++++++++++++++++++++
        time_hour_text_ctrl = TextCtrl(panel, size=(25, 18), style=wx.TE_CENTER)
        time_splitter = StaticText(panel, label=u"：")
        time_minute_text_ctrl = TextCtrl(panel, size=(25, 18), style=wx.TE_CENTER)

        time_sbox_sizer = BoxSizer()
        time_sbox_sizer.Add(time_hour_text_ctrl, 0, wx.RIGHT, border=5)
        time_sbox_sizer.Add(time_splitter, 0)
        time_sbox_sizer.Add(time_minute_text_ctrl, 0, wx.LEFT, border=5)

        self.time_hour_text_ctrl = time_hour_text_ctrl
        self.time_minute_text_ctrl = time_minute_text_ctrl
        # --------------------------- Time ------------------------------

        # ++++++++++++++++++++++++++ DateTime +++++++++++++++++++++++++++
        datetime_sbox = StaticBox(panel, -1, u'時間:')
        datetime_sbox_sizer = StaticBoxSizer(datetime_sbox, wx.VERTICAL)
        datetime_sbox_sizer.Add(date_box_sizer, 0, wx.ALL, 5)
        datetime_sbox_sizer.Add(time_sbox_sizer, 0, wx.ALL, 5)
        # --------------------------- DateTime ---------------------------

        # +++++++++++++++++++++++++++ CA01 / CA03 +++++++++++++++++++++++++++
        ca01_radio_btn = wx.RadioButton(panel, label="CA01")
        ca03_radio_btn = wx.RadioButton(panel, label="CA03")

        client_group_sbox = StaticBox(panel, -1, u'類型:')
        client_group_sbox_sizer = StaticBoxSizer(client_group_sbox, wx.VERTICAL)
        client_group_sbox_sizer.Add(ca01_radio_btn, 0, wx.ALL | wx.LEFT, 5)
        client_group_sbox_sizer.Add(ca03_radio_btn, 0, wx.ALL | wx.LEFT, 5)
        # --------------------------- CA01 / CA03 ----------------------------------------

        date_time_group_sizer = BoxSizer()
        date_time_group_sizer.Add(datetime_sbox_sizer, 0)
        date_time_group_sizer.AddSpacer(400)
        date_time_group_sizer.Add(client_group_sbox_sizer)

        right_client_evaluation_detail_sizer = BoxSizer(wx.VERTICAL)
        right_client_evaluation_detail_sizer.Add(evaluation_comment_text_ctrl, 5, wx.EXPAND)
        right_client_evaluation_detail_sizer.AddSpacer(20)
        right_client_evaluation_detail_sizer.Add(date_time_group_sizer , 2, wx.EXPAND)

        root_sizer = BoxSizer()
        root_sizer.Add(client_list_sizer, 0, wx.SHAPED)
        # root_sizer.Add(StaticLine(panel, style=wx.LI_VERTICAL), flag=wx.EXPAND)
        # root_sizer.AddSpacer(20)
        root_sizer.Add(right_client_evaluation_detail_sizer, 1, wx.EXPAND, border=10)
        root_sizer.AddSpacer(10)

        self.SetSizer(root_sizer)
        self.SetAutoLayout(1)
        self.Fit()
        # root_sizer.Fit(self)
        self.Show(True)

        self.panel = panel

        # HomeTherapy --

        # Define widgets early even if they're not going to be seen
        # so that they can come up FAST when someone clicks for them!
        self.aboutme = wx.MessageDialog(self, " A sample editor \n"
                                              " in wxPython", "About Sample Editor", wx.OK)
        self.doiexit = wx.MessageDialog(self, " Exit - R U Sure? \n",
                                        "GOING away ...", wx.YES_NO)

        # dirname is an APPLICATION variable that we're choosing to store
        # in with the frame - it's the parent directory for any file we
        # choose to edit in this frame
        # self.dirname = ''

        self._current_home_therapy_company = None
        self._current_client_idx = None

        self.browser = None
        self.lc_browser = None
        self.cc_client_profile_dict = {}
        self.lc_client_profile_dict = {}
        self.global_client_profile_dict = {}
        self.home_therapy_web_manager = None

    def create_client_profiles(self, home_therapy_company: HomeTherapyCompany, client_name_list: list) -> dict:
        client_profile_dict = {}

        for client_name in client_name_list:
            client_profile = self.home_therapy_web_manager.create_client_profile(home_therapy_company, client_name, "治療師名稱")
            client_profile_dict[client_name] = client_profile

        return client_profile_dict

    def on_add_new_client(self, event):
        key_code = event.GetKeyCode()
        if key_code == 13:
            print("ENTER = " + self.add_client_text_ctrl.GetValue())
            # TODO: why wrong layout
            # self.client_list_sizer.Add(Button(self.panel, label=self.add_client_text_ctrl.GetValue()), 0, wx.TOP | wx.LEFT | wx.RIGHT, border=8)
            # # self.client_list_sizer.Layout()  # ????????
            # self.client_list_sizer.Fit(self.panel)  # ???????
            # self.Fit()  # ???????

            x, y = int(self.date_month_text_ctrl.GetValue()), int(self.date_day_text_ctrl.GetValue())
            # "<<" = (20x20)=(1, 75, 20, 94)  (小畫家上只有 14x14)  (screen_shot=13x13)
            # pin = (206, 75, 225, 94)
            ActionChains(self.browser).move_by_offset(x, y).click().move_by_offset(-x, -y).perform()  # 家裡電腦 [(0, 51)標題 (0, 52)不行] [(14, 52)標題 (13, 52)不行

        else:
            event.Skip()

    def on_client_item_selected(self, event, home_therapy_company: HomeTherapyCompany):
        # print("[on_client_item_selected] " + str(home_therapy_company))
        # self.tests[0].func(None)
        # return
        idx = event.GetIndex()
        client_row_dict = self.dict_of_client_row_dict[home_therapy_company]
        client_name = client_row_dict[idx]
        print(f"{home_therapy_company}[{idx}] = {client_name}")
        self._current_home_therapy_company = home_therapy_company
        self._current_client_idx = idx

    # TODO: refactor to support dynamic num of therapy companies
    def cc_on_client_item_double_clicked(self, event):
        self.on_client_item_double_clicked(event, HomeTherapyCompany.CC)

    def lc_on_client_item_double_clicked(self, event):
        self.on_client_item_double_clicked(event, HomeTherapyCompany.CC)

    def on_client_item_double_clicked(self, event, home_therapy_company: HomeTherapyCompany):
        # print("double-clicked " + str(home_therapy_company))
        # return
        idx = event.GetIndex()
        client_row_dict = self.dict_of_client_row_dict[home_therapy_company]
        client_name = client_row_dict[idx]
        print(client_name + " double-clicked")
        if self.evaluation_comment_text_ctrl.GetValue() != "":
            self.show_info_message(f'評估欄沒有清空，不允許覆蓋!')
        else:
            print("Start load last record for " + client_name)
            # self.getClientEvaluationRecords(client_name)
            # return

            evaluation_record_page_url = r'https://csms2.sfaa.gov.tw/lcms/qd/showEditQd120?qd120id=' + self.get_last_client_evaluation_record_id(client_name, home_therapy_company, "治療師名稱")
            evaluation_record_page_source = self.home_therapy_web_manager.get_url_page_source(evaluation_record_page_url, home_therapy_company)
            evaluation_comment = self.pull_up_evaluation_comment(evaluation_record_page_source)
            print(evaluation_comment)
            self.evaluation_comment_text_ctrl.SetValue(evaluation_comment)

    def get_last_client_evaluation_record_id(self, client_name, home_therapy_company: HomeTherapyCompany, servicer_name):
        records = self.get_client_evaluation_records(client_name, home_therapy_company)
        for record in records:
            if servicer_name in record:
                last_record_id = record[:record.find(",")]
                print("Last record for " + client_name + " is " + last_record_id)
                return last_record_id

    def get_client_evaluation_records(self, client_name, home_therapy_company: HomeTherapyCompany):
        client_id = self.get_client_id(client_name)
        url = r'https://csms2.sfaa.gov.tw/lcms/qd/filterQd120/' + client_id + r'?iframe_tab_params=cascade+previous.tabid%3DROOT&iframe_tab_params=cascade%28nofix%3Dtrue%29+entrance%3D384520077&%3Fiframe_tab_params=cascade+tabid%3DTREE-384520077&qd111id=972470871&ca100id=' + client_id + '&qdperms=true&perms=true&limit=10&offset=0&order=asc&_=1571575842161'
        html = self.home_therapy_web_manager.get_url_page_source(url, home_therapy_company)
        split_array = re.split('"id":', html)
        return split_array[1:]

    @staticmethod
    def pull_up_evaluation_comment(html_content):
        # str = r'<textarea name="qd120.commt" id="F4PvkPo3tJXMLaf1" rows="4" style="width: 98%" class="form-control">1. </textarea>'
        # reg = re.findall(r"")
        # htmlContent.
        rm_left_part = re.split('<textarea[^<>]+>', html_content)[1]
        evaluation_comment = re.split('</textarea', rm_left_part)[0]
        return evaluation_comment

    def on_enter(self, event):
        key_code = event.GetKeyCode()
        # print("key_code=" + str(key_code) + " Ctrl=" + str(event.ControlDown()))
        if key_code == 10:
            print("CTRL ENTER")

            idx = self._current_client_idx
            client_row_dict = self.dict_of_client_row_dict[self._current_home_therapy_company]
            client_name = client_row_dict[idx]
            print('client_name=' + client_name)
            self.add_service_record_to_db()
            # print(self.evaluation_comment_text_ctrl.GetValue())
        # elif key_code == 13:
        #     print("ENTER")
        else:
            event.Skip()

    def add_service_record_to_db(self):
        if self.home_therapy_web_manager is None:
            print("HomeTherapyWebManager is not ready...")
            return

        idx = self._current_client_idx
        client_row_dict = self.dict_of_client_row_dict[self._current_home_therapy_company]
        client_name = client_row_dict[idx]

        client_profiles = self.global_client_profile_dict[self._current_home_therapy_company]
        client_profile = client_profiles[client_name]

        service_record = self.create_service_record(client_profile)
        url = self.createUrl_AddServiceRecord(service_record)
        print(url)

        if url is None:
            print("add_service_record_to_db, generated url is None")
            return

        page_source = self.home_therapy_web_manager.get_url_page_source(url, self._current_home_therapy_company)

        if u"成功" in page_source:
            print("[Success] add_service_record_to_db = " + client_name)

            if client_profile.home_therapy_company == HomeTherapyCompany.CC:
                save_record_screenshot = threading.Thread(target=self.save_service_record_screenshot, args=[service_record])
                save_record_screenshot.start()

            self.show_info_message(f'[{client_name}] 新增紀錄成功')
        else:
            print("[Failed] add_service_record_to_db" + client_name)
            print(page_source)
            self.show_err_message(f'[{client_name}] 新增紀錄失敗')

    CC = True

    def save_service_record_screenshot(self, service_record):
        self.home_therapy_web_manager.get_service_record_screenshot(service_record, self.browser_display_ratio)

        return

        if False:
            pos_107 = self.wait_until_template_shown(browser, r"template\template_107.png", "QD-107新制服務區", 1)[1]
            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(pos_107),
                              y_offset=7)  # open QD-107
            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_101_2.png", "101-照會列表", 1)[1]),
                              y_offset=7)  # open 101
            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(pos_107),
                              y_offset=7)  # close QD-107
            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_client_name.png", "案主姓名", 1.5)[1]),
                              x_offset=100, y_offset=7)

            pyperclip.copy(client_name)
            ActionChains(browser).key_down(Keys.SHIFT).send_keys(Keys.INSERT).key_up(Keys.SHIFT).perform()

            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_search.png", "查詢")[1]))

            print("page-down")
            WebDriverUtil.press_key(browser, Keys.PAGE_DOWN)

            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_name_id.png", "姓名-身分證號", 0.5)[1]),
                              y_offset=72)

            self.wait_until_template_shown(browser, r"template\template_service_item_detail.png", "服務項目明細", 2)

            pos = self.wait_until_template_shown(browser, r"template\template_client_tab.png", "個案服務:", 2)[1]
            screenshot_left, screenshot_top = int(pos[0] - 5), int(pos[1] - 8)
            self.click_screen(browser, self.from_screenshot_coor_to_click_coor(pos))

            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_service_record.png", "服務紀錄")[1]))

            print("page-down")
            WebDriverUtil.press_key(browser, Keys.PAGE_DOWN)

            self.click_screen(browser,
                              self.from_screenshot_coor_to_click_coor(self.wait_until_template_shown(browser, r"template\template_edit.png", "編輯", 0.5)[1]))

            img, pos = self.wait_until_template_shown(browser, r"template\template_delete_2.png", "刪除", 2)
            screenshot_right, screenshot_bottom = int(pos[0] + 58), int(pos[1] + 27)

            final_client_record_screenshot = img[screenshot_top:screenshot_bottom, screenshot_left:screenshot_right]
            cv.imwrite("final_client_record_screenshot.png", final_client_record_screenshot)
            image = Image.open('final_client_record_screenshot.png')
            image.show()

            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()

            SystemUtil.send_to_clipboard(win32clipboard.CF_DIB, data)

            WebDriverUtil.press_key(browser, Keys.ESCAPE)

            os.remove("tmp.png")

    def from_html_coor_to_screenshot_coor(self, html_pos) -> (int, int):
        return int(html_pos[0] * self.browser_display_ratio), int(html_pos[1] * self.browser_display_ratio)

    def wait_until_template_shown(self, browser, template_file_name, template_name=None, delay_time_before_start=0):
        if delay_time_before_start != 0:
            time.sleep(delay_time_before_start)

        while True:
            browser.save_screenshot("tmp.png")
            img_rgb = cv.imread("tmp.png")
            img_gray = cv.cvtColor(img_rgb, cv.COLOR_BGR2GRAY)
            template = cv.imread(template_file_name, cv.IMREAD_GRAYSCALE)

            found = None

            # TODO: 可能有多個編輯  找Y最高，或是名字是治療師名稱

            for scale in np.linspace(1 / self.browser_display_ratio, 0.2 / self.browser_display_ratio, 20):
                resized_image = imutils.resize(img_gray, width=int(img_gray.shape[1] * scale))
                ratio = img_gray.shape[1] / float(resized_image.shape[1])

                if resized_image.shape[0] < template.shape[0] or resized_image.shape[1] < template.shape[1]:
                    break

                result = cv.matchTemplate(resized_image, template, cv.TM_CCOEFF_NORMED)
                (_, maxVal, _, maxLoc) = cv.minMaxLoc(result)

                if found is None or maxVal > found[0]:
                    found = (maxVal, maxLoc, ratio)

            (maxValue, maxLoc, ratio_max) = found
            if maxValue > 0.5:
                pos = (int(maxLoc[0] * ratio_max), int(maxLoc[1] * ratio_max))

                # draw a bounding box around the detected result and display the image
                (startX, startY) = pos
                (endX, endY) = (int((maxLoc[0] + template.shape[1]) * ratio_max), int((maxLoc[1] + template.shape[0]) * ratio_max))
                cv.rectangle(img_rgb, (startX, startY), (endX, endY), (0, 0, 255), 2)
                cv.imwrite("tmp_locate.png", img_rgb)
                # cv.imshow("Image", img_rgb)
                # cv.waitKey

                # resized_pos = int(pos[0] / self.browser_display_ratio), int(pos[1] / self.browser_display_ratio)

                # print(f'pos={pos} system_display_ratio={self.system_display_ratio}  browser_display_ratio={self.browser_display_ratio}')
                # print(f'template found: ({resized_pos[0]}, {resized_pos[1]}) {template_name}')
                print(f'template found: ({pos[0]}, {pos[1]}) {template_name}')

                # return [img_rgb, resized_pos]
                return [img_rgb, pos]

        return None

    def from_screenshot_coor_to_click_coor(self, screenshot_pos):
        return int(screenshot_pos[0] / self.browser_display_ratio), int(screenshot_pos[1] / self.browser_display_ratio)

    @staticmethod
    def click_screen(browser, pos, x_offset=0, y_offset=0):
        x, y = pos[0] + x_offset, pos[1] + y_offset
        ActionChains(browser).move_by_offset(x, y).click().move_by_offset(-x, -y).perform()

    @staticmethod
    def press_key(browser, key_code):
        ActionChains(browser).send_keys(key_code).perform()

    def show_info_message(self, msg):
        wx.MessageBox(msg, '訊息', wx.OK)

    def show_err_message(self, msg):
        wx.MessageBox(msg, '錯誤', wx.OK | wx.ICON_ERROR)

    def get_current_client_name(self):
        return self.client_row_dict[self._current_client_idx]

    def create_service_record(self, client_profile):

        if client_profile is None:
            print("create_service_record, create_service_recordNo client profile provided.")
            return None

        # idx = self._current_client_idx
        # ot_company = self._current_home_therapy_company

        # client_row_dict = self.dict_of_client_row_dict[ot_company]
        # client_name = client_row_dict[idx]

        client_name = client_profile.client_name
        home_therapy_company = client_profile.home_therapy_company
        client_id = client_profile.client_id
        client_idno = client_profile.client_idno
        evaluation_group_id = client_profile.client_group_id

        # client_profile_dict = self.global_client_profile_dict[ot_company]
        # client_idno = client_profile_dict[client_name]

        year = 2019
        month = self.date_month_text_ctrl.GetValue()
        day = self.date_day_text_ctrl.GetValue()

        start_hour = self.time_hour_text_ctrl.GetValue()
        start_min = self.time_minute_text_ctrl.GetValue()

        evaluation_comment = self.evaluation_comment_text_ctrl.GetValue().replace("\n", "%0D%0A")  # replace "\n"
        # client_profile = self.cc_client_profile_dict[client_name]

        return self.EvaluationRecord(home_therapy_company, client_name, client_id, client_idno, year, month, day, start_hour, start_min, evaluation_group_id, evaluation_comment)

    # TODO: 從 DB 撈
    def get_client_id(self, client_name):
        client_attr = self.get_client_attr(client_name)
        return client_attr[1]

    # TODO: 從 DB 撈
    def get_client_group_id(self, client_name):
        for name, attr_descriptor in CLIENT_LIST.items():
            if client_name == name:
                group_descriptor = attr_descriptor[0]
                return constants.CLIENT_TYPE_ID_CA01 if group_descriptor[0] == 1 else constants.CLIENT_TYPE_ID_CA03

        return None

    def get_client_attr(self, clientName):
        for name, attr_descriptor in CLIENT_LIST.items():
            if clientName == name:
                return attr_descriptor

        return None

    # TODO: 用個人 profile 資訊，從 DB 撈
    def get_therapy_id_from_client(self, client_name):
        client_attr = self.get_client_attr(client_name)
        return DB_ID_CO0168 if client_attr[2] == OT_COMPANY_JIA_JIN else DB_ID_LC035

    # class ClientEvaluationGroup(Enum):
    #     CA01 = 'CA01'
    #     CA03 = 'CA03'

    class EvaluationRecord:

        def __init__(self, home_therapy_company, name, client_id, client_idno, year, month, day, start_hour, start_min, evaluation_group_id, evaluation_comment):
            self.home_therapy_company = home_therapy_company
            self.name = name
            self.client_id = client_id
            self.client_idno = client_idno
            self.year = year
            self.month = int(month)
            self.day = int(day)
            self.start_hour = int(start_hour)
            self.start_min = int(start_min)
            self.end_hour, self.end_min = self.calculate_end_time(self.start_hour, self.start_min)
            self.evaluation_group_id = evaluation_group_id
            self.evaluation_comment = evaluation_comment

        def calculate_end_time(self, start_hour, start_min):
            if start_min < 10:
                return start_hour, start_min + 50
            else:
                return start_hour + 1, start_min - 10

    def createUrl_AddServiceRecord(self, evaluation_record):
        day_str_tmp = str(evaluation_record.day)
        day = '0' + day_str_tmp if evaluation_record.day < 10 else day_str_tmp

        month_str_tmp = str(evaluation_record.month)
        month = '0' + month_str_tmp if evaluation_record.month < 10 else month_str_tmp

        year = str(evaluation_record.year - 1911)

        client_name = evaluation_record.name
        client_group_id = evaluation_record.evaluation_group_id

        if client_group_id is None:
            print("createUrl_AddEvaluationRecord, client_group_id is None, skip...")
            return None

        res = r'https://csms2.sfaa.gov.tw/lcms/qd/saveQd120Batch?' \
               r'qd120.id=' \
               r'&qd120.ca100=' + str(evaluation_record.client_id) + \
               r'&qd120ServdtNum=1' \
               r'&qd120dt.1.servDt=' + year + '%2F' + month + '%2F' + day + \
               r'&qd120ServtimeNum=1' \
               r'&qd120time.1.startHh=' + str(evaluation_record.start_hour) + \
               r'&qd120time.1.startMm=' + str(evaluation_record.start_min) + \
               r'&qd120time.1.endHh=' + str(evaluation_record.end_hour) + \
               r'&qd120time.1.endMm=' + str(evaluation_record.end_min)
        res += \
               r'&qd120.servUsers=' + self.get_therapy_id_from_client(client_name) + \
               r'&qd120.noAa09='
        res += \
               r'&qd120.commt=' + evaluation_record.evaluation_comment
        res += \
               r'&qd120Num=1' \
               r'&qd120.1.qd100=' + str(client_group_id) + \
               r'&qd120.1.addr1=' \
               r'&qd120.1.addr2=' \
               r'&qd120.1.carNum=' \
               r'&qd120.1.driver=' \
               r'&qd120.1.amount=1' \
               r'&qd120.1.price=4500.0' \
               r'&qd120.1.stype=1' \
               r'&qd120.1.lastServ=false' \
               r'&qd120.1.missedVisit=false'

        return res

    # 個案名稱1 qd120.id=&qd120.ca100=362191551&qd120ServdtNum=1&qd120dt.1.servDt=108%2F10%2F18&qd120ServtimeNum=1&qd120time.1.startHh=11&qd120time.1.startMm=11&qd120time.1.endHh=11&qd120time.1.endMm=11&qd120.servUsers=864119605&qd120.noAa09=&qd120.commt=&qd120Num=1&qd120.1.qd100=285889260&qd120.1.addr1=&qd120.1.addr2=&qd120.1.carNum=&qd120.1.driver=&qd120.1.amount=1&qd120.1.price=4500.0&qd120.1.stype=1&qd120.1.lastServ=false&qd120.1.missedVisit=false
    # 個案名稱1 qd120.id=&qd120.ca100=362191551&qd120ServdtNum=1&qd120dt.1.servDt=108%2F05%2F05&qd120ServtimeNum=1&qd120time.1.startHh=7&qd120time.1.startMm=8&qd120time.1.endHh=7&qd120time.1.endMm=58&qd120.servUsers=864119605&qd120.noAa09=&qd120.commt=123&qd120.1.qd100=285889260&qd120Num=1&qd120.1.addr1=&qd120.1.addr2=&qd120.1.carNum=&qd120.1.driver=&qd120.1.amount=1&qd120.1.price=4500.0&qd120.1.stype=1&qd120.1.lastServ=false&qd120.1.missedVisit=false
    # 個案名稱2 qd120.id=&qd120.ca100=874056375&qd120ServdtNum=1&qd120dt.1.servDt=108%2F10%2F18&qd120ServtimeNum=1&qd120time.1.startHh=7&qd120time.1.startMm=8&qd120time.1.endHh=9&qd120time.1.endMm=10&qd120.servUsers=882319126&qd120.noAa09=&qd120.commt=123&qd120Num=1&qd120.1.qd100=285889258&qd120.1.addr1=&qd120.1.addr2=&qd120.1.carNum=&qd120.1.driver=&qd120.1.amount=1&qd120.1.price=4500.0&qd120.1.stype=1&qd120.1.lastServ=false&qd120.1.missedVisit=false
    # 個案名稱3 qd120.id=&qd120.ca100=919606075&qd120ServdtNum=1&qd120dt.1.servDt=108%2F10%2F10&qd120ServtimeNum=1&qd120time.1.startHh=11&qd120time.1.startMm=12&qd120time.1.endHh=13&qd120time.1.endMm=14&qd120.servUsers=882319126&qd120.noAa09=&qd120.commt=&qd120Num=1&qd120.1.qd100=285889258&qd120.1.addr1=&qd120.1.addr2=&qd120.1.carNum=&qd120.1.driver=&qd120.1.amount=1&qd120.1.price=4500.0&qd120.1.stype=1&qd120.1.lastServ=false&qd120.1.missedVisit=false&qd120Num=2&qd120.2.qd100=285889260&qd120.2.addr1=&qd120.2.addr2=&qd120.2.carNum=&qd120.2.driver=&qd120.2.amount=0&qd120.2.price=4500.0&qd120.2.stype=1&qd120.2.lastServ=false&qd120.2.missedVisit=false

    def set_home_therapy_web_manager(self, home_therapy_web_manager):
        self.home_therapy_web_manager = home_therapy_web_manager

        # (HomeTherapyCompany, ["個案名字1", "個案名字2"])
        for ot_company, client_list in self.client_tuple_list:
            # {{LC: {個案名稱1: ClientProfile1, ...}}, CC: {個案名稱1: ClientProfile1, ...}}
            self.global_client_profile_dict[ot_company] = self.create_client_profiles(ot_company, client_list)

    def createMenuBar(self):
        # Setting up the menu. filemenu is a local variable at this stage.
        filemenu = wx.Menu()
        # use ID_ for future easy reference - much better that "48", "404" etc
        # The & character indicates the short cut key
        menuOpen = filemenu.Append(wx.ID_OPEN, "&Open", "Open a file")
        menuAbout = filemenu.Append(wx.ID_ABOUT, "&About", "Information about this program")  # (ID, 專案名稱, 狀態列資訊)
        filemenu.AppendSeparator()  # 分割線
        menuExit = filemenu.Append(wx.ID_EXIT, "&Exit", "Terminate the program")  # (ID, 專案名稱, 狀態列資訊)

        # Creating the menubar.
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "&File")  # Adding the "filemenu" to the MenuBar

        # 設定 events
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)

        return menuBar

    def OnAbout(self, e):
        # A modal show will lock out the other windows until it has
        # been dealt with. Very useful in some programming tasks to
        # ensure that things happen in an order that  the programmer
        # expects, but can be very frustrating to the user if it is
        # used to excess!
        self.aboutme.ShowModal()  # Shows it
        # widget / frame defined earlier so it can come up fast when needed

    def OnExit(self, e):
        # A modal with an "are you sure" check - we don't want to exit
        # unless the user confirms the selection in this case ;-)
        igot = self.doiexit.ShowModal()  # Shows it
        if igot == wx.ID_YES:
            self.Close(True)  # Closes out this simple application

    def OnOpen(self, e):
        # In this case, the dialog is created within the method because
        # the directory name, etc, may be changed during the running of the
        # application. In theory, you could create one earlier, store it in
        # your frame object and change it when it was called to reflect
        # current parameters / values
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()

            # Open the file, read the contents and set them into
            # the text edit window
            filehandle = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(filehandle.read())
            filehandle.close()

            # Report on name of latest file read
            self.SetTitle("Editing ... " + self.filename)
            # Later - could be enhanced to include a "changed" flag whenever
            # the text is actually changed, could also be altered on "save" ...
        dlg.Destroy()

    def OnSave(self, e):
        # Save away the edited text
        # Open the file, do an RU sure check for an overwrite!
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*", \
                            wx.SAVE | wx.OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            # Grab the content to be saved
            itcontains = self.control.GetValue()

            # Open the file for write, write, close
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            filehandle = open(os.path.join(self.dirname, self.filename), 'w')
            filehandle.write(itcontains)
            filehandle.close()
        # Get rid of the dialog to keep things tidy
        dlg.Destroy()

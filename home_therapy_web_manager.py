import json
import time
from io import BytesIO
from selenium.webdriver.support import expected_conditions as EC

import win32clipboard
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait

import SystemUtil
import WebDriverUtil
import constants
from constants import TAKE_CARE_HOME_PAGE_URL
from HomeTherapy import HomeTherapyCompany


class BrowserTab:
    def __init__(self, home_therapy_company: HomeTherapyCompany, tab_idx: int, cookies: dict):
        self.home_therapy_company = home_therapy_company
        self.tab_idx = tab_idx
        self.cookies = cookies


class ClientProfile:
    def __init__(self, client_name, home_therapy_company, client_id, client_idno, client_group_id):
        self.client_name = client_name
        self.home_therapy_company = home_therapy_company
        self.client_id = client_id
        self.client_idno = client_idno
        self.client_group_id = client_group_id


class HomeTherapyWebManager:
    def __init__(self, browser, browser_tab_list: list):
        self.do_not_directly_use_this_browser = browser
        self.browser_tab_list = browser_tab_list

    def switch_to(self, home_therapy_company: HomeTherapyCompany):
        browser = self.do_not_directly_use_this_browser

        for browser_tab in self.browser_tab_list:
            if browser_tab.home_therapy_company == home_therapy_company:

                # switch focus to this tab
                browser.switch_to.window(browser.window_handles[browser_tab.tab_idx])

                # set corresponding cookies to this tab
                self.replace_cookies_to_window(browser, browser_tab.cookies)
                return browser

        return None

    @staticmethod
    def replace_cookies_to_window(browser, cookies: dict):
        browser.delete_all_cookies()
        for cookie in cookies:
            if 'expiry' in cookie:
                del cookie['expiry']
            browser.add_cookie(cookie)

    def get_url_page_source(self, url: str, home_therapy_company: HomeTherapyCompany = None):
        if home_therapy_company is not None:
            browser = self.switch_to(home_therapy_company)
        else:
            browser = self.do_not_directly_use_this_browser
        browser.get(url)
        return browser.page_source

    def create_client_profile(self, home_therapy_company: HomeTherapyCompany, client_name, servicer_name: str) -> ClientProfile:
        browser = self.switch_to(home_therapy_company)

        """ get client id """
        browser.get(
            'https://csms2.sfaa.gov.tw/lcms/qd/filterQd111/?doQuery=true&qdList1=yes&caseno=&serialno=&name=' + client_name +
            '&idno=&noteDt1=&noteDt2=&processDt1=&processDt2=&qcntcode=&twnspcode=&vilgcode=&qinformCntcode=&informTwnspcode=&informVilgcode=&_qd111Resp=&_qd111adjHint=&_qd111pi400Hint=&_qlistQd111Export=&_qsortNotNull=&qsort=&qorder=1&limit=10&offset=0&order=asc&_=1570458820317');
        client_profile_json = json.loads(browser.find_element_by_tag_name('body').text)
        num_of_profile = client_profile_json['total']
        client_profile = client_profile_json['rows'][0]

        client_id = client_profile['ca100id']
        client_idno = client_profile['idno']  # 身份證字號

        # for client_attr in client_profile:
        #     print(client_attr + '-' + str(client_profile[client_attr]))

        # TODO: don't know why cookies failed to keep log in
        browser = self.switch_to(home_therapy_company)

        """ get client group id """
        url_getting_client_group_id = 'https://csms2.sfaa.gov.tw/lcms/qd/filterQd120/919606075?iframe_tab_params=cascade+previous.tabid%3DROOT&iframe_tab_params=cascade%28nofix%3Dtrue%29+entrance%3D384520077&%3Fiframe_tab_params=cascade+tabid%3DTREE-384520077&qd111id=972470871&ca100id=' + str(client_id) + '&qdperms=true&perms=true&limit=10&offset=0&order=asc&_=1571579327737'
        browser.get(url_getting_client_group_id)

        client_service_records_json = json.loads(browser.find_element_by_tag_name('body').text)
        num_of_service_record = client_service_records_json['total']
        client_service_records = client_service_records_json['rows']

        for service_record in client_service_records:
            if servicer_name in service_record['servUser']:
                serv_record_title = service_record['title']
                client_group_id = constants.CLIENT_TYPE_ID_CA01 if "CA01" in serv_record_title else constants.CLIENT_TYPE_ID_CA03
                break

        return ClientProfile(client_name, home_therapy_company, client_id, client_idno, client_group_id)

    def get_service_record_screenshot(self, service_record, browser_display_ratio):

        ht_company = service_record.home_therapy_company
        browser = self.switch_to(ht_company)
        browser.get(TAKE_CARE_HOME_PAGE_URL)

        client_name = service_record.name

        if ht_company == HomeTherapyCompany.CC:
            tree_107 = WebDriverWait(browser, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_11_switch']"))
            )
        else:
            tree_107 = WebDriverWait(browser, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='west-ztree_3_switch']"))
            )
        tree_107.click()

        if ht_company == HomeTherapyCompany.CC:
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

        client_name_field = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.NAME, "name"))
        )
        client_name_field.send_keys(client_name)

        while True:
            try:
                # 查詢
                search_btn = WebDriverWait(browser, 20).until(
                    EC.element_to_be_clickable((By.NAME, "search"))
                )
                search_btn.click()  # 為何有時沒有按到按鈕???

                id_btn = WebDriverWait(browser, 20).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//a[@href='javascript: void(0)' and contains(text(), '{client_name}')]"))
                )
                id_btn.click()
                break

            except TimeoutException:
                print(f"Timeout: Button({client_name}) not shown...")

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
        serv_date_col_idx = None
        header_row = table_rows[0]
        header_title_fields = header_row.find_elements_by_tag_name("th")
        for idx, table_header_col in enumerate(header_title_fields):
            if "服務人員" in table_header_col.text:
                serv_person_col_idx = idx
            elif "服務日期" in table_header_col.text:
                serv_date_col_idx = idx

        if serv_person_col_idx is None or serv_date_col_idx is None:
            return

        action_btn_col_idx = len(header_title_fields) - 1

        # 用這兩行來找第一列符合『治療師名稱』 和 『可編輯狀態』
        table_body_rows = table_rows[1:]
        for body_row in table_body_rows:
            body_fields = body_row.find_elements_by_tag_name("td")
            service_person_field = body_fields[serv_person_col_idx]
            service_date_field = body_fields[serv_date_col_idx]
            show_serv_record_detail_btn = body_fields[action_btn_col_idx]
            serv_date_str = str(service_record.year - 1911) + "/" + str(service_record.month) + "/" + str(service_record.day)
            # make sure the service record is matched
            if service_person_field.text == "治療師名稱" and service_date_field.text == serv_date_str and show_serv_record_detail_btn.text == "編輯":
                # found last added record
                show_serv_record_detail_btn.click()

                edit_serv_record_popup = WebDriverWait(browser, 20).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "modal-footer"))
                )
                time.sleep(0.5)  # the popup is showing animation...
                size = edit_serv_record_popup.size
                location = edit_serv_record_popup.location

                html_crop_right_bottom_pos = (
                    iframe_offset[0] + location['x'] + size['width'],
                    iframe_offset[1] + location['y'] + size['height'])
                screenshot_crop_right_bottom_pos = int(html_crop_right_bottom_pos[0] * browser_display_ratio), int(html_crop_right_bottom_pos[1] * browser_display_ratio)
                tmp_img = WebDriverUtil.get_screenshot(browser)
                browser.save_screenshot("ttt.png")
                final_client_record_screenshot = tmp_img.crop((iframe_offset[0], iframe_offset[1],
                                                               screenshot_crop_right_bottom_pos[0],
                                                               screenshot_crop_right_bottom_pos[1]))
                final_client_record_screenshot.save("final_client_record_screenshot.png")
                final_client_record_screenshot.show()

                output = BytesIO()
                final_client_record_screenshot.convert("RGB").save(output, "BMP")
                data = output.getvalue()[14:]
                output.close()

                SystemUtil.send_to_clipboard(win32clipboard.CF_DIB, data)

                WebDriverUtil.press_key(browser, Keys.ESCAPE)
                return

        print("No matched record found...")

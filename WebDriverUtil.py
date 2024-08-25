from io import BytesIO

from PIL import Image
from selenium.webdriver import ActionChains


def get_screenshot(browser):
    return Image.open(BytesIO(browser.get_screenshot_as_png()))


def press_key(browser, key_code):
    ActionChains(browser).send_keys(key_code).perform()

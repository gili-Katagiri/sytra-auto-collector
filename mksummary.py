from appium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os

msapp = '...............'
pafrase='...............'

# get Desktop
desired_caps = {}
desired_caps['app'] = "Root"

time.sleep(10)

driver= webdriver.Remote(
    command_executor = 'http://127.0.0.1:4723',
    desired_capabilities = desired_caps
)

# find MS Window
driver_ms = driver.find_element_by_name(msapp)

# move to comp-summary
driver_ms.send_keys(Keys.ALT + Keys.F5)
time.sleep(1)
# move to Login Dialog
driver_ms.send_keys(Keys.ENTER)
time.sleep(1)
# login
driver_ms.find_element_by_class_name('#32770').find_element_by_accessibility_id('10002').send_keys(pafrase + Keys.ENTER)
time.sleep(3)

# find Excel Window
driver_xl = driver.find_element_by_class_name("XLMAIN")
time.sleep(10)
# read summary_base.csv and auto trans
driver_xl.send_keys(Keys.CONTROL + "i")
time.sleep(15)
# save as summary.csv
driver_xl.send_keys(Keys.CONTROL + "q")
time.sleep(10)

# quit apps
# close Excel
driver_xl.find_element_by_name("閉じる").click()
time.sleep(1)
# get Context and close RSS
driver.find_element_by_class_name("TrayNotifyWnd").find_element_by_class_name('SysPager').find_element_by_name("RSS  接続中").click()
time.sleep(1)
driver.find_element_by_name('コンテキスト').find_element_by_name('終了').click()
time.sleep(1)
# close MS
driver_ms.find_element_by_accessibility_id('Close').click()

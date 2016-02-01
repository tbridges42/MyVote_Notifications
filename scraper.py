from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import  ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
import time


def do_and_switch_to_new_window(browser, action):
    old_windows = browser.window_handles
    action()
    new_windows = browser.window_handles
    diff = set(new_windows) - set(old_windows)
    browser.switch_to_window(diff.pop())


def get_webdriver():
    profile = FirefoxProfile()
    profile.set_preference("browser.download.panel.shown", False)
    profile.set_preference("browser.helperApps.neverAsk.openFile","text/csv,application/vnd.ms-excel")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/vnd.ms-excel")
    profile.set_preference("browser.download.folderList", 2);
    profile.set_preference("browser.download.dir", "c:\\firefox_downloads\\")
    return webdriver.Firefox(firefox_profile=profile)


def main():
    browser = get_webdriver()
    browser.get("http://wisvote.wi.gov")
    browser.maximize_window()
    main_window = browser.current_window_handle
    assert "Sign In" in browser.title
    username = browser.find_element_by_id("userNameInput")
    username.click()
    username.clear()
    username.send_keys("svrs\\bridgt")
    browser.find_element_by_id("passwordInput").send_keys("" + Keys.ENTER)
    time.sleep(4)
    do_and_switch_to_new_window(browser, browser.find_element_by_id("advancedFindImage").click)
    time.sleep(4)
    assert "Advanced Find" in browser.title
    browser.switch_to.frame(browser.find_element_by_id("contentIFrame0"))
    entity = Select(browser.find_element_by_id("slctPrimaryEntity"))
    entity.select_by_visible_text("Absentee Applications")
    view = Select(browser.find_element_by_id("savedQuerySelector"))
    view.select_by_visible_text("Absentee mailing")
    masked_date = browser.find_element_by_id("advFindEFGRP0FFLD2CCVALLBL")
    hover = ActionChains(browser).move_to_element(masked_date)
    hover.perform()
    date = browser.find_element_by_id("DateInput")
    date.click()
    date.clear()
    date.send_keys("1/29/2016")
    browser.switch_to.default_content()
    browser.find_element_by_id("Mscrm.AdvancedFind.Groups.Show.Results-Large").click()
    time.sleep(3)
    browser.switch_to.alert.accept()
    time.sleep(.2)
    browser.switch_to.alert.accept()
    time.sleep(.2)
    browser.find_element_by_id("edm_absenteeapplication|NoRelationship|SubGridStandard|Mscrm.SubGrid.edm_absenteeapplication.ExportToExcel-Large").click()
    time.sleep(.5)
    browser.switch_to.frame(browser.find_element_by_id("InlineDialog_Iframe"))
    browser.find_element_by_id("dialogOkButton").click()



    print("Done")


if __name__ == "__main__":
    main()

from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import  ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from threading import Thread
from datetime import date, datetime, timedelta
import time
import excel
import send_email


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


def wait_for_id(driver, id):
    WebDriverWait(driver, 10).until(
        expected_conditions.presence_of_element_located((By.ID, id))
    )


def get_creds():
    creds = {}
    with open("config.txt") as file:
        for line in file:
            (key, val) = [x.strip() for x in line.split('=')]
            creds[key] = val
    return creds


def sign_in(browser):
    creds = get_creds()
    username = browser.find_element_by_id("userNameInput")
    username.click()
    username.clear()
    username.send_keys("svrs\\" + creds['username'])
    browser.find_element_by_id("passwordInput").send_keys(creds['password'] + Keys.ENTER)


def get_yesterday():
    yesterday = date.today() - timedelta(1)
    return yesterday.strftime('%m/%d/%Y')


def get_days_ago(delta):
    day = date.today() - timedelta(delta)
    return day.strftime('%m/%d/%Y')


def click_if_available(browser, element_id, retry=False, timeout=5.0):
    time_elapsed = 0.0
    while time_elapsed < timeout:
        try:
            browser.find_element_by_id(element_id).click()
            break
        except NoSuchElementException:
            if retry:
                time_elapsed += 0.1
                time.sleep(0.1)
            else:
                break


def dismiss_alerts(browser):
    while True:
        try:
            time.sleep(0.1)
            browser.switch_to.alert.accept()
        except NoAlertPresentException:
            break


def construct_export_id(entity_name):
    entity_name = "".join(entity_name.split()).lower()
    return "edm_" + entity_name + "|NoRelationship|SubGridStandard|Mscrm.SubGrid.edm_" + \
           entity_name + ".ExportToExcel-Large"


def download_excel(browser, entity_name):
    browser.find_element_by_id("Mscrm.AdvancedFind.Groups.Show.Results-Large").click()
    dismiss_alerts(browser)
    browser.find_element_by_id(construct_export_id(entity_name)).click()
    time.sleep(.5)
    browser.switch_to.frame(browser.find_element_by_id("InlineDialog_Iframe"))
    click_if_available(browser, "printAll")
    browser.find_element_by_id("dialogOkButton").click()
    browser.switch_to.default_content()


def select_view(browser, entity_name, view_name):
    entity = Select(browser.find_element_by_id("slctPrimaryEntity"))
    entity.select_by_visible_text(entity_name)
    dismiss_alerts(browser)
    view = Select(browser.find_element_by_id("savedQuerySelector"))
    view.select_by_visible_text(view_name)


def change_date(browser, absentee_date):
    wait_for_id(browser, "advFindEFGRP0FFLD2CCVALLBL")
    masked_date = browser.find_element_by_id("advFindEFGRP0FFLD2CCVALLBL")
    hover = ActionChains(browser).move_to_element(masked_date)
    hover.perform()
    date = browser.find_element_by_id("DateInput")
    date.click()
    date.clear()
    date.send_keys(absentee_date)


def switch_to_iframe(browser):
    browser.switch_to.frame(browser.find_element_by_id("contentIFrame0"))


def switch_out_of_iframe(browser):
    browser.switch_to.default_content()


def download_absentee_list(browser, absentee_date):
    switch_to_iframe(browser)
    select_view(browser, "Absentee Applications", "MyVote Mailing 2")
    change_date(browser, absentee_date)
    switch_out_of_iframe(browser)
    download_excel(browser, "Absentee Application")


def download_email_list(browser):
    switch_to_iframe(browser)
    select_view(browser, "Jurisdictions", "Jurisdiction Emails w Provider")
    switch_out_of_iframe(browser)
    download_excel(browser, "Jurisdictions")


def return_to_query(browser):
    time.sleep(0.1)
    browser.find_element_by_class_name("ms-cui-tt-a").click()


def get_absentee_list(absentee_date):
    browser = get_webdriver()
    browser.get("http://wisvote.wi.gov")
    browser.maximize_window()
    main_window = browser.current_window_handle
    if "Sign In" in browser.title:
        sign_in(browser)
    else:
        assert "Easy Navigate" in browser.title
    wait_for_id(browser, "contentIFrame0")
    time.sleep(0.75)
    do_and_switch_to_new_window(browser, browser.find_element_by_id("advancedFindImage").click)
    wait_for_id(browser, "contentIFrame0")
    download_absentee_list(browser, absentee_date)
    time.sleep(1)
    browser.close()
    browser.switch_to.window(main_window)
    browser.close()


def get_email_list():
    browser = get_webdriver()
    browser.get("http://wisvote.wi.gov")
    browser.maximize_window()
    main_window = browser.current_window_handle
    if "Sign In" in browser.title:
        sign_in(browser)
    else:
        assert "Easy Navigate" in browser.title
    wait_for_id(browser, "contentIFrame0")
    time.sleep(1)
    do_and_switch_to_new_window(browser, browser.find_element_by_id("advancedFindImage").click)
    wait_for_id(browser, "contentIFrame0")
    download_email_list(browser)
    time.sleep(1)
    browser.close()
    browser.switch_to.window(main_window)
    browser.close()


def create_emails(data, emails):
    with open('text_template.txt', encoding="utf8") as text:
        body = text.readlines()
    with open('email_template.html', encoding="utf8") as html:
        html_body = html.readlines()
    for key in data:
        html_table = "<tr><td><table>"
        text_table = ""
        for record in data[key]:
            html_table += "<tr>"
            for datum in record:
                html_table += "<td>" + datum + "</td>"
                text_table += datum + "\t\t"
            html_table += "</tr>"
            text_table += "\n"
        html_table += "</table></td></tr>"
        html_body[34] = html_table
        body[34] = text_table
        body_string = ""
        for string in body:
            body_string += string
        html_string = ""
        for string in html_body:
            html_string += string
        send_email.send_mail(get_creds(), '', "Notification from MyVote", body_string, html_string)
        break


def main():
    count = 1
    get_absentee_list(get_yesterday())
    if datetime.today().weekday() == 0:
        get_absentee_list(get_days_ago(2))
        get_absentee_list(get_days_ago(3))
        count = 3
    get_email_list()
    data, emails = excel.main(count)
    #create_emails(data, emails)
    #for key in data:
    #    send_email.send_mail(get_creds(), emails[key], 'Test', str(data[key]))
    print("Done")


if __name__ == "__main__":
    main()

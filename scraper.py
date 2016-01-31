from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time


def main():
    mydriver = webdriver.Firefox()
    mydriver.get("http://wisvote.wi.gov")
    mydriver.maximize_window()
    time.sleep(4)
    assert "Sign In" in mydriver.title
    mydriver.find_element_by_id("userNameArea").send_keys("svrs\\bridgt")
    mydriver.find_element_by_id("passwordArea").send_keys("67942Klyel1")
    print("Done")


if __name__ == "__main__":
    main()

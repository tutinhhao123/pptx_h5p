import webbrowser
url='https://docs.python.org/'
# webbrowser.open_new_tab(url)
#
# # Open URL in new window, raising the window if possible.
# webbrowser.open_new(url)
from selenium import webdriver
driver = webdriver.Chrome(r"E:\download\chromedriver.exe")
driver.get("https://www.verywellmind.com/what-is-personality-testing-2795420")
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])
driver.get('https://www.indeed.com/career-advice/career-development/types-of-personality-test')
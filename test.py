from selenium import webdriver
import pandas as pd


# https://g.co/gemini/share/ada92acb29a0
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

driver = webdriver.Edge(options=options)
driver.get("https://allapp.vhcy.gov.tw/webreg/frmOpdSchedule_wc")


table = pd.read_html(driver.page_source)[0]

print(table)



from selenium import webdriver

# https://g.co/gemini/share/ada92acb29a0
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

driver = webdriver.Edge(options=options)
sss = driver.get("https://www.google.com.tw")

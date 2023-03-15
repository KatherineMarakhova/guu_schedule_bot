from selenium import webdriver
import time

from selenium.webdriver.common.by import By


def get_file(course):
    PATH = "/Users/katherine.marakhova/PycharmProjects/exampleBot/files"
    # Google Chrome
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": PATH}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)

    driver.get('https://guu.ru/студентам/расписание-сессий/schedule/')
    time.sleep(1)

    if (course == 4):
        driver.get("https://guu.ru/wp-content/uploads/4-курс-бакалавриат-ОФО-42.xlsx")
    if (course == 3):
        driver.get('https://guu.ru/wp-content/uploads/3-курс-бакалавриат-ОФО-50.xlsx')
    if (course == 2):
        driver.get('https://guu.ru/wp-content/uploads/2-курс-бакалавриат-ОФО-48.xlsx')
    if (course == 1):
        driver.get('https://guu.ru/wp-content/uploads/1-курс-бакалавриат-ОФО-48.xlsx')

    time.sleep(1)
    driver.quit()

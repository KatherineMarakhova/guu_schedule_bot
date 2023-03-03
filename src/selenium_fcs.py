from selenium import webdriver
import time


def get_file(course):
    PATH = "/Users/katherine.marakhova/PycharmProjects/exampleBot/files"
    # Google Chrome
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": PATH}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)

    driver.get('https://guu.ru/студентам/расписание-сессий/schedule/')
    time.sleep(3)
    # селекторы надо обязательно переписать!!!!!!!!!!!
    if(course == 4):
        driver.get("https://guu.ru/wp-content/uploads/4-курс-бакалавриат-ОФО-42.xlsx")
    if (course == 3):
        driver.get('https://guu.ru/wp-content/uploads/3-курс-бакалавриат-ОФО-50.xlsx') #файл третьего курса
    time.sleep(2)
    driver.quit()

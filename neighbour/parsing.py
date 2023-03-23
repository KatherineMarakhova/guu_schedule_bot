import requests
import schedule
import time
import wget
from bs4 import BeautifulSoup
import os
import shutil
from pathlib import Path
from src.bot import message_to_makareshka

# НЕЗАБЫВАТЬ ЗАПУСКАТЬ ЭТОТ ФАЙЛ ЧЕРЕЗ nohup python MyScheduledProgram.py &

def get_file(course):
    url = 'https://guu.ru/студентам/расписание-сессий/schedule/'
    response = requests.get(url)

    bs = BeautifulSoup(response.text,"lxml")

    rows = bs.find_all('a', class_ = "doc-unit odd")

    for row in rows:
        link = row.attrs["href"]
        if link == 'None': continue
        # print(link)
        if str(link).find(f'{course}-курс-бакалавриат') != -1:
            # response = requests.get(link, '../files')
            wget.download(link, '../files/')
            course += 1
            break


def update_docs(self):
    path = '../files'
    try:
        shutil.rmtree(path)
        print("Папка удалена.")
    except OSError as error:
        print("Возникла ошибка.")

    os.mkdir(path)
    print("Папка создана.")

    with Path(r"../files") as direction:

        for i in range(1, 5):
            s = str(self.course) + "-курс-бакалавриат*.xlsx"
            get_file(i)  # скачиваем новый файл
            path = self.get_file_path(s)
            self.unmerge_all_cells(path)
            self.unmerge_institutes(path)
        sec = time.time()
        struct = time.localtime(sec)
        t = time.strftime('%d.%m.%Y %H:%M', struct)
        message_to_makareshka(t)

schedule.every().day.at("03:00").do(update_docs)
while True:
    schedule.run_pending()
    time.sleep(1) # wait one minute


##TODO
##сгенерировать отчет из какой-нибудь папки
##переименовать отчет - для удобства работы. Можно переименовать в их код из ReportService
##DONE получить длину таблицы
##DONE из этой таблицы рандомно почекать ссылки из разных ячеек
##DONE попытаться перейти по этим ссылкам
import random

import pandas as pd
import openpyxl
import webbrowser
import sys

wb = openpyxl.load_workbook('issue.xlsx')
ws = wb.active
max_col = ws.max_column
max_row = ws.max_row
print(f'Max row in list: {max_row}')


def get_issue_screen_link():
    screen_link_list = []
    for row in range(5, max_row):
        screen_link = ws[f'X{row + 1}'].value
        screen_link_list.append(screen_link)

    random_url = random.sample(screen_link_list, 3)

    for row in range(len(random_url)):
        webbrowser.open(random_url[row]) ##проверить можно ли оптимизировать этот код?


def get_issue_screen_check_link():
    screen_check_list = []
    for row in range(5, max_row):
        screen_check_link = ws[f'Y{row + 1}'].value
        if screen_check_link is not None:
            screen_check_list.append(screen_check_link)
        else:
            pass

    random_url = random.sample(screen_check_list, 3)

    for row in range(len(random_url)):
        webbrowser.open(random_url[row]) ##проверить можно ли оптимизировать этот код?


if __name__ == '__main__':
    get_issue_screen_link()
    ##get_issue_screen_check_link()

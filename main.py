##TODO
##сгенерировать отчет из какой-нибудь папки
##переименовать отчет - для удобства работы. Можно переименовать в их код из ReportService
##DONE получить длину таблицы
##из этой таблицы рандомно почекать ссылки из разных ячеек
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
    screen_list = []
    for row in range(5, max_row):
        screen_link = ws[f'X{row + 1}'].value
        screen_list.append(screen_link)
    random_url = random.sample(screen_list, 3)

    for row in range(len(random_url)):
        webbrowser.open(random_url[row]) ##проверить можно ли оптимизировать этот код?


def get_issue_screen_check_link():
    for row in range(5, max_row):
        screen_link = ws[f'Y{row + 1}'].value
        if screen_link is not None:
            print(screen_link)
            webbrowser.open(screen_link)
        else:
            pass


if __name__ == '__main__':
    get_issue_screen_link()
    ##get_issue_screen_check_link()

import platform
import pandas as pd
import numpy as np
import re
import requests
from bs4 import BeautifulSoup as bs

import docx
import os

import time
import datetime
from calendar import monthrange

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


def reformat_date(date: str, year):
    """
    Функция переформатирует даты
    """
    date = date.strip()
    flag = True if ((year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)) else False
    if date == 'январь':
        date = '31 january'
    elif date == 'февраль' and flag:
        date = '29 february'
    elif date == 'февраль':
        date = '28 february'
    elif date == 'март':
        date = '31 march'
    elif date == 'апрель':
        date = '30 April'
    elif date == 'май':
        date = '31 may'
    elif date == 'июнь':
        date = '30 june'
    elif date == 'июль':
        date = '31 july'
    elif date == 'август':
        date = '31 august'
    elif date == 'сентябрь':
        date = '30 september'
    elif date == 'октябрь':
        date = '31 october'
    elif date == 'ноябрь':
        date = '30 november'
    elif date == 'декабрь':
        date = '31 december'
    return date


def pars_year_by_months():
    """
    Функция для получения ссылок на документы по месяцам.
    Для ВВП реализовано возвращение названия последнего доступного месяца в конкретном году
    и ссылки на соответствующий раздел.
    """
    header = {
        'user-agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:86.0) Gecko/20100101 Firefox/86.0'
    }
    time.sleep(15)
    url = f'https://www.cbr.ru/statistics/macro_itm/svs/'
    response = requests.get(url, headers=header)
    soup = bs(response.content, "html.parser")
    for i in soup.find_all('a'):
        if i.text.replace('\n', '').strip() == 'Внешняя торговля Российской Федерации товарами (по методологии ' \
                                               'платежного баланса)':
            link_to_download = f'https://www.cbr.ru' + str(i['href'])
            print(link_to_download)
            time.sleep(15)
            dok_name_to_download = 'file.xls'
            folder = os.getcwd()
            response = requests.get(link_to_download, headers=header)
            folder = os.path.join(folder, 'word_data', dok_name_to_download)
            if response.status_code == 200:
                with open(folder, 'wb') as f:
                    f.write(response.content)
                print(f'Document was downloaded.')
            else:
                print('FAILED:', link_to_download)
            return 'word_data/' + dok_name_to_download


def parse_docx_document(path):
    """
    Функция осуществляет парсинг документа.
    path - путь к документу (обязательно в формате .docx)
    year - текущий год
    """
    data_xlsx = pd.read_excel(path)
    data_xlsx = data_xlsx.iloc[:, [1, 2, 3, 8, 9]]
    count = 0
    flag = False
    data = pd.DataFrame()
    data_xlsx.columns = ['Дата', 'Экспорт, млн долл. США', 'Экспорт,  % накопленным итогом год к году', 'Импорт, млн '
                                                                                                        'долл. США',
                         'Импорт,  % накопленным итогом год к году']
    for i in data_xlsx.iloc[::-1, [1]].values:
        if not np.isnan(i[0]):
            if count != 1:
                string = data_xlsx.loc[data_xlsx['Экспорт, млн долл. США'] == i[0]].iloc[0]
                data = data._append(string)
            if flag is False:
                flag = True
        elif flag is True and np.isnan(i[0]):
            flag = False
            count += 1
        if count == 3 and np.isnan(i[0]):
            break
    year = datetime.datetime.now().year
    data = data.iloc[::-1].reset_index(drop=True)
    data.iloc[12:, 0] = data.iloc[12:, 0].apply(lambda x: reformat_date(x, year))
    data.iloc[:12, 0] = data.iloc[:12, 0].apply(lambda x: reformat_date(x, year - 1))

    for i in range(len(data)):
        if i <= 11:
            data.iloc[i, 0] = pd.to_datetime(data.iloc[i, 0] + str(year - 1))
        else:
            data.iloc[i, 0] = pd.to_datetime(data.iloc[i, 0] + str(year))

    return data


def create_new_date(last_date_in_file_year, last_date_in_file_month):
    now = datetime.datetime.now()
    lst_date = []
    _, last_day = monthrange(now.year, now.month)
    last_date = datetime.datetime.strptime(f"{now.year}-{now.month}-{last_day}", "%Y-%m-%d").date()

    for i in range((last_date.year - last_date_in_file_year) * 12 + last_date.month - last_date_in_file_month - 1):
        if last_date.month - 1 != 0:
            _, last_day = monthrange(last_date.year, last_date.month - 1)
            last_date = datetime.datetime.strptime(f"{last_date.year}-{last_date.month - 1}-{last_day}",
                                                   "%Y-%m-%d").date()
        else:
            _, last_day = monthrange(last_date.year - 1, 12)
            last_date = datetime.datetime.strptime(f"{last_date.year - 1}-{12}-{last_day}", "%Y-%m-%d").date()
        lst_date.append(last_date)
    return sorted(lst_date)


def append_date_rez_file_Y(xlsx_path='rez_file_Y_v2.xlsx'):
    """
        Функция осуществляет дабавление месяцев, если их нет в файле.
    """
    data_xlsx = pd.read_excel(xlsx_path)
    year = pd.to_datetime(pd.read_excel('rez_file_Y_v2.xlsx')['Целевой показатель'].iloc[-1]).year
    month = pd.to_datetime(pd.read_excel('rez_file_Y_v2.xlsx')['Целевой показатель'].iloc[-1]).month
    date_lst = create_new_date(year, month)
    for date in date_lst:
        new_string = {'Целевой показатель': [date]}
        new_string.update({c: [None] for c in data_xlsx.columns[1:]})
        new_string = pd.DataFrame(new_string)
        if not data_xlsx.empty and not new_string.empty:
            data_xlsx = pd.concat([data_xlsx, new_string])
    data_xlsx.to_excel(xlsx_path, index=False)


def update_rez_file_y(data, xlsx_path='rez_file_Y_v2.xlsx'):
    """
        Функция осуществляет обновление файла со всеми данными rez_file_Y_v2.xlsx
    """
    data_xlsx = pd.read_excel(xlsx_path)
    if data.values[-1][0] not in list(data_xlsx['Целевой показатель']):
        append_date_rez_file_Y()
        data_xlsx = pd.read_excel(xlsx_path)
    for c in data.columns[1:]:
        index = list(data.columns).index(c)
        for j in data.values:
            data_xlsx.loc[data_xlsx['Целевой показатель'] == j[0], c] = float(str(j[index]).replace(',', '.'))

    data_xlsx.to_excel(xlsx_path, index=False)


def main():
    """
    Основная функция. Выполняет проверку данных на полноту. Скачивет недостающие
    данные и дополняет ими файл с данными.
    """
    path = pars_year_by_months()
    data = parse_docx_document(path)
    if not data.empty:
        update_rez_file_y(data, xlsx_path='rez_file_Y_v2.xlsx')


if __name__ == '__main__':
    main()

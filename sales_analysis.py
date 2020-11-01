import pandas as pd
from openpyxl import load_workbook
import datetime
import collections
import supportfunc
logs_data = pd.read_excel('logs.xlsx', sheet_name='log')
logs_data_dict = logs_data.to_dict(orient='records')

# Собираем лист из словарей с информацией о браузере и дате посещения
browser_list = []
ru = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
      'октябрь', 'ноябрь', 'декабрь']
for element in logs_data_dict:
    for k, v in element.items():
        if k == 'Браузер':
            name = v
        elif k == 'Дата посещения':
            name1 = datetime.datetime.strptime(str(v), "%Y-%m-%d %H:%M:%S")
            date_m = name1.month
            date_month = ru[date_m-1]
    el_append = {name: date_month}
    browser_list.append(el_append)

# Собираем словарь с ключем браузер и  значенимем - словарь с месяцами
# в которых встретился браузер
date_comb = {}
for el in browser_list:
    for k, v in el.items():
        if k in date_comb:
            new_el = date_comb[k]
            new_el.append(v)
            date_comb.update({k: new_el})
        else:
            date_comb[k] = [v]


# Словарь с браузером и счетчиком по месяцам, сколько раз месяца повторились
browser_dict = {}
for k, v in date_comb.items():
    counter_date = collections.Counter(v)
    browser_dict.update({k: counter_date})

# Определяем наиболее популярные браузеры
dict_most_common = collections.defaultdict(int)
for el in browser_list:
    for element in el.keys():
        dict_most_common[element] += 1
popular = collections.Counter(dict_most_common).most_common(7)

# отработка функции из модуля supportfunc 
# Заполняем эксель файл с 7 популярными браузерами и количеством посещений по месяцам
# и вставляем в эксель
supportfunc.insert_data_task(5,4,browser_dict,popular,'report.xlsx','Лист1','report.xlsx')


# отработка функции из модуля supportfunc 
# функция для подсеча итогов по месяцам и заполнения их в файл эксель
supportfunc.total_count(3, 5, 10, 12, 'report.xlsx', 'Лист1', 'report.xlsx')

# Собираем лист из словарей с информацией о продаже и дате продажи
goods_list = []
ru = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
      'октябрь', 'ноябрь', 'декабрь']
for element in logs_data_dict:
    for goods, items in element.items():
        if goods == 'Купленные товары':
            name = items
        elif goods == 'Дата посещения':
            name1 = datetime.datetime.strptime(str(items), "%Y-%m-%d %H:%M:%S")
            date_m = name1.month
            date_month = ru[date_m-1]
    el_append = {name: date_month}
    goods_list.append(el_append)

# Распарсиваем товары и присваиваем им дату покупки
list_product_date = []
for element in goods_list:
    for k, v in element.items():
        new_fotmat = k.split(',')
        for i in new_fotmat:
            new_el = {i: v}
            list_product_date.append(new_el)

# Собираем словарь из товаров и всех месяцев покупки
dict_product_date = {}
for product_date in list_product_date:
    for k, v in product_date.items():
        if k in dict_product_date:
            dict_el = dict_product_date[k]
            dict_el.append(v)
            dict_product_date.update({k: dict_el})
        else:
            dict_product_date[k] = [v]

# Собираем словарь с подсчитаным кол-вом повторений месяцев
new_dict_product_date = {}
for k, v in dict_product_date.items():
    counter_data = collections.Counter(v)
    new_dict_product_date.update({k:counter_data})


# Собираем лист с наиболее популярным товаром
all_products = collections.defaultdict(int)
for item in goods_list:
    for part in item.keys():
        for i in part.split(','):
            all_products[i] += 1
popular_product = collections.Counter(all_products).most_common(7)


# отработка функции из модуля supportfunc 
# Заполняем эксель файл с 7 популярными товарами и количеством повторений по месяцам
# и вставляем в эксель
supportfunc.insert_data_task(19,18,new_dict_product_date,popular_product,'report.xlsx','Лист1','report.xlsx')


# отработка функции из модуля supportfunc 
# функция для подсеча итогов по месяцам и заполнения их в файл эксель
supportfunc.total_count(3, 19, 10, 26, 'report.xlsx', 'Лист1', 'report.xlsx')

# Определяем список популярных товаров с признаком пол купившего
preferece_list = []
for element in logs_data_dict:
    for goods, items in element.items():
        if goods == 'Купленные товары':
            name = items.split(',')
        elif goods == 'Пол':
            name1 = items
    el_append = {name1: name}
    preferece_list .append(el_append)

# Функция определения самого популярного или не популярного товара
# среди мужчин или женщин и заполнение этих данных в эксель
supportfunc.preferences_customers(preferece_list,'м', '+', 0, 'B31', 'report.xlsx', 'Лист1', 'report.xlsx')
supportfunc.preferences_customers(preferece_list,'ж', '+', 0, 'B32', 'report.xlsx', 'Лист1', 'report.xlsx')
supportfunc.preferences_customers(preferece_list,'м', '-', 0, 'B33', 'report.xlsx', 'Лист1', 'report.xlsx')
supportfunc.preferences_customers(preferece_list,'ж', '-', 0, 'B34', 'report.xlsx', 'Лист1', 'report.xlsx')
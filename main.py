import pandas as pd
import numpy as np
import os
import xlwings as xw
import win32com.client
import matplotlib.pyplot as plt
from matplotlib import ticker
from function import intersect, min_lenght
# from tqdm import tqdm
import tkinter as tk
from time import sleep
import logging
import sys

data_file = "База перспективного ПФ (15.02).xlsx" # Файл для расчета ' — копия

# distance = 150 # Минимальное расстояние между скважинами
# diff_depth = 30

# Названия столбцов в Excel
team = 'Команда'
field = 'Месторождение'
cluster_pad = 'КП'
object_name = 'Объект'
work_marker = 'Назначение'
well_number = 'Скважина - Забой'
x1 = "X1"
y1 = "Y1"
z1 = "Z1"
x3 = "X3"
y3 = "Y3"
z3 = "Z3"
x11 = "X11"
y11 = "Y11"
z11 = "Z11"
x33 = "X33"
y33 = "Y33"
z33 = "Z33"


def target(df, data_file, distance, diff_depth, win, pb, label_4):

    # df = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file), header=2)  # Открытие экселя
    name_columns = pd.MultiIndex.from_tuples([('', '', 'Команда'), ('', '', 'Месторождение'), ('', '', 'КП'), ('', '', 'Объект'), ('', '', 'Назначение'), ('', '', 'Скважина - Забой'),
                                         ('1й ствол', 'T1', 'X'), ('', '', 'Y'), ('', '', 'Z'), ('', 'T3', 'X'), ('', '', 'Y'), ('', '', 'Z'), ('', '', 'Комментарий'),
                                              ('2й ствол', 'T1', 'X'), ('', '', 'Y'), ('', '', 'Z'), ('', 'T3', 'X'), ('', '', 'Y'), ('', '', 'Z'),
                                              ('', 'Пересечения', 'скв (команда/куст/объект)')]) #, ('', 'Скважины рядом', 'скв (команда/куст/объект)')])

    if list(df)[-1] == 'Unnamed: 18': # последний столбец не имел заголовка
        df['Unnamed: 18'] = np.nan
        df = df.rename(columns={'X': x1, 'Y': y1, 'Z': z1, 'X.1': x3, 'Y.1': y3, 'Z.1': z3, 'X.2': x11, 'Y.2': y11, 'Z.2': z11, 'X.3': x33, 'Y.3': y33, 'Unnamed: 18': z33}) # Переименование столбца
    else:
        df = df.rename(columns={'X': x1, 'Y': y1, 'Z': z1, 'X.1': x3, 'Y.1': y3, 'Z.1': z3, 'X.2': x11, 'Y.2': y11, 'Z.2': z11, 'X.3': x33, 'Y.3': y33, 'Z.3': z33})  # Переименование столбца

    if df.shape[1] > 19:
        df.drop(df.columns[20:], axis=1, inplace=True)

    df_result = pd.DataFrame()
    statistics = pd.DataFrame(columns=["Месторождение", "Количество пересечений", "Пересекающиеся команды"])

    list_coord = [x1, y1, z1, x3, y3, z3, x11, y11, z11, x33, y33, z33]

    for coord in list_coord:
        df[coord] = np.where((df[coord] == '-'), np.nan, df[coord]) # замена '-' на np.nan
        print(df[coord].dtype)
        # костыль - на случай, если числа вставлены как текст
        if df[coord].dtype == "object":
            df_cells_str = df[pd.to_numeric(df[coord], errors='coerce').isnull()].dropna(subset=[coord])
            list_cells_str = df_cells_str[coord].tolist()
            df_cells_str[coord] = df_cells_str[coord].str.replace(',', ".")
            new_list_cells_float = df_cells_str[coord].tolist()
            df[coord] = df[coord].replace(list_cells_str, new_list_cells_float)
        # try:
            # df[coord] = df[coord].str.replace(',', ".")
        # except AttributeError:
        #     df[coord] = df[coord]
        df[coord] = df[coord].astype(float)
        df[coord] = np.where((df[coord] < 0), df[coord] * (-1), df[coord]) # замена отрицательных на положительные координаты

    list_fields = df[field].unique() # всего месторождений

    app1 = xw.App(visible=False) # доступ к файлу?
    new_wb = xw.Book() # открываем Excel файл

    # создание папки для результатов
    dir_result = str(os.path.dirname(data_file)) + "/Result"
    if not os.path.isdir(dir_result):
        os.mkdir(dir_result)

    # перебор месторождений
    for field_name in list_fields:
    # for field_name in tqdm(list_fields, desc='fields'): #, tk_parent=win): # для консоли (без -w при упаковке в exe)
    # for field_name in tqdm_gui(list_fields): #, desc='fields'):
        print(field_name)
        # logging.basicConfig(filename="log.txt", level=logging.INFO)
        # logging.info(field_name)

        df_field = df[df[field] == field_name]
        df_field['Пересечения'] = np.nan
        # df_field['Скважины рядом'] = np.nan

        num_intersections = 0 # счетчик пересечений/скважин рядом
        list_intersection_teams = set() # список пересекающихся команд
        # df_intersection = pd.DataFrame()

        # создание графика по месторождению matplotlib
        fig, ax = plt.subplots()
        fig.set_size_inches(9, 9)
        ax.set_xlabel("X", fontsize=8, fontweight='bold')
        ax.set_ylabel("Y", fontsize=8, fontweight='bold')
        ax.grid(True)
        ax.yaxis.set_major_formatter(ticker.StrMethodFormatter("{x:.0f}"))
        plt.title(field_name, fontsize=16)
        plt.xticks(fontsize=7)
        plt.yticks(fontsize=7)

        num_wells = df.shape[0]

        # перебор текущих скважин
        for i in range(len(df_field)):  # for each row:
            # print(f'текущая скважина - {df_field.iloc[i, 5]}')

            # проверка на пилот pl
            if df_field.iloc[i, 4] == 'pl':
                continue

            color = 'k' # цвет не пересекающихся скважин
            start_well_hole_current = [6] # столбец X1 первого ствола

            list_wells_intersect = []
            list_wells_near = []

            if not np.isnan(df_field.iloc[i, 13]): # проверка есть ли данные по второму стволу
                start_well_hole_current.append(13)

            # перебор стволов скважины
            # координаты
            for n in start_well_hole_current:
                x1_current = df_field.iloc[i, n]
                y1_current = df_field.iloc[i, n + 1]
                z1_current = df_field.iloc[i, n + 2]
                x3_current = df_field.iloc[i, n + 3]
                y3_current = df_field.iloc[i, n + 4]
                z3_current = df_field.iloc[i, n + 5]

                # перебор окружающих скважин
                for j in range(len(df_field)):
                    Flag = 0 # флаг пересечения, 0 - нет, 1 - да
                    # print(f'проверяемая скважина - {df_field.iloc[j, 5]}')

                    # проверка на пилот pl
                    if df_field.iloc[j, 4] == 'pl':
                        continue

                    start_well_hole_round = [6] # столбец X1 первого ствола

                    if not np.isnan(df_field.iloc[j, 13]): # проверка есть ли данные по второму стволу
                        start_well_hole_round.append(13)

                    # перебор стволов скважины
                    # координаты
                    for k in start_well_hole_round:
                        x1_round = df_field.iloc[j, k]
                        y1_round = df_field.iloc[j, k + 1]
                        z1_round = df_field.iloc[j, k + 2]
                        x3_round = df_field.iloc[j, k + 3]
                        y3_round = df_field.iloc[j, k + 4]
                        z3_round = df_field.iloc[j, k + 5]

                        # проверка принадлежности к одному объекту по абс.отметкам
                        if abs(z1_current - z1_round) >= diff_depth:
                            continue

                        # проверка на пересечение отрезков скважин
                        if intersect(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round) and i != j:
                            Flag = 1 # отрезки пересекаются, Flag = 0 - отрезки не пересекаются
                            num_intersections += 1
                            # print(f'скважина текущая - {df_field.iloc[i, 5]}')
                            # print(f'пересекающаяся скважина - {df_field.iloc[j, 5]}')
                            list_wells_intersect.append(str(df_field.iloc[j, 5]) + ' (' + str(df_field.iloc[j, 0]) + '/ ' + str(df_field.iloc[j, 2]) + '/ ' + str(df_field.iloc[j, 3]) + ')')
                            list_intersection_teams.add(df_field.iloc[i, 0])
                            list_intersection_teams.add(df_field.iloc[j, 0])

                        # поиск расстояния между отрезками и проверка на нахожение отрезков рядом
                        elif min_lenght(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round) < distance and i != j:
                            num_intersections += 1
                            # print(f'скважина текущая - {df_field.iloc[i, 5]}')
                            # print(f'скважина рядом - {df_field.iloc[j, 5]}')
                            list_wells_near.append(str(df_field.iloc[j, 5]) + ' (' + str(df_field.iloc[j, 0]) + '/ ' + str(df_field.iloc[j, 2]) + '/ ' + str(df_field.iloc[j, 3]) + ')')
                            list_intersection_teams.add(df_field.iloc[i, 0])
                            list_intersection_teams.add(df_field.iloc[j, 0])

                list_wells_intersect_near = list_wells_near + list_wells_intersect # пересечения и ближайшие скважины
                if len(list_wells_intersect_near) > 0: # изменение цвета скважины на наличии пересечения/скважин рядом
                    color = 'r'

                # добавление траектории скважины
                # plt.plot([list of Xs], [list of Ys])
                plt.plot([df_field.iloc[i, n], df_field.iloc[i, n + 3]], [df_field.iloc[i, n + 1], df_field.iloc[i, n + 4]], c=color)
                plt.plot(df_field.iloc[i, n], df_field.iloc[i, n + 1], marker=".", c=color, markersize=4)
                ax.annotate(df_field.iloc[i, 5], (df_field.iloc[i, n] + 40, df_field.iloc[i, n + 1] - 30), size=5)

                # добавление данных по пересечению/скважин рядом
                df_field.iloc[i, 19] = ", ".join(map(str, list_wells_intersect_near))
                # df_field.iloc[i, 19] = ", ".join(map(str, list_wells_intersect))
                # df_field.iloc[i, 20] = ", ".join(map(str, list_wells_near))

                # проверка на пересение сразу всего столбца
                # df_field['result'] = df_field.apply(intersect, args=(x1_current, y1_current, x3_current, y3_current), axis=1)
                # df_field_intersect = df_field[df_field['result'] == True]
                # if len(df_field_intersect) > 0:
                #     print(df_field.iloc[i, 5])
                #     print(df_field_intersect.iloc[:, 5])

        # df только пересечениями/скважинами рядом для страниц по месторождению
        df_intersection = df_field.copy()
        df_intersection.iloc[:, 19].replace('', np.nan, inplace=True)
        df_intersection = df_intersection.dropna(subset=df_intersection.columns[[19]])

        statistics = statistics._append({'Месторождение': field_name, 'Количество пересечений': num_intersections / 2,
                                         'Пересекающиеся команды': ", ".join(map(str, list(list_intersection_teams)))}, ignore_index=True)

        df_result = df_result._append(df_field, ignore_index=False)

        # Запись в эксель - каждое месторождение с пересечениями на свой лист
        if len(df_intersection) > 0:
            new_wb.sheets.add(str(field_name), before='Лист1')
            sht = new_wb.sheets(str(field_name))
            sht.api.Tab.ColorIndex = 16
            sht.range('A1').options(index=False).value = df_intersection
            sht.pictures.add(fig, name='Plot', update=True, left=sht.range("W1").left, top=sht.range("W1").top) # добавление графика matplotlib в эксель

        # plt.show()
        # fig.savefig(str(field_name))

        # str(os.path.dirname(os.path.abspath(data_file))) + "/Result"

        fig.savefig(dir_result + "/" + field_name)

        # fig.savefig(str(os.path.join(os.path.dirname(__file__), data_file)).replace(str(os.path.basename(data_file)), "") + field_name)

        #для progressbara
        pb['value'] += 100 / len(list_fields)
        label_4['text'] = round(pb['value']), '%'
        win.update_idletasks()
        sleep(0.05)

    df_result.columns = name_columns

    # result = pd.concat([df, df_result], axis=1)
    # result = pd.merge(df, df_result, left_index=True, right_index=True)
    sht = new_wb.sheets[0].name

    if "Статистика" in new_wb.sheets:
        xw.Sheet["Статистика"].delete()
    new_wb.sheets.add('Статистика', before=new_wb.sheets[0].name)
    sht = new_wb.sheets('Статистика')
    sht.range('A1').options(index=False).value = statistics

    if "Цели" in new_wb.sheets:
        xw.Sheet["Цели"].delete()
    new_wb.sheets.add("Цели") # добавление страницы "История"
    sht = new_wb.sheets("Цели") # получение данных с листа
    sht.range('A1').options(index=False).value = df_result # сохранение данных из Pandas в Excel

    # print(str(os.path.join(os.path.dirname(__file__), data_file)).replace(".xlsx", "") + "_out.xlsx")
    # print(str(os.path.basename(data_file)))
    # print(str(os.path.dirname(os.path.abspath(__file__))))
    # print(str(os.path.join(os.path.dirname(__file__))))

    new_wb.save(str(os.path.join((dir_result + "/"), os.path.basename(data_file))).replace(".xlsx", "") + "_out.xlsx")  # сохранение нового эксель в той же диреектории, что и исходный файл
    # new_wb.save(str(os.path.join(os.path.dirname(__file__), data_file)).replace(".xlsx", "") + "_out.xlsx") # сохранение нового эксель в той же диреектории, что и исходный файл
    # app1.kill()
    app1.quit()
    del app1

    label_4['text'] = '' # при повторном нажатии на "расчет" почему-то процент расчета накладывается на 100%
    win.update_idletasks()






import numpy as np
import tkinter as tk
import os
import sys

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


# Определение ориентации отрезков
def check_orientation(x1, y1, x2, y2, x3, y3):
    res = (y2 - y1) * (x3 - x2) - (y3 - y2) * (x2 - x1)

    if res > 0:  # по часовой стрелке
        return 1
    elif res < 0:  # против часовой стрелки
        return 2
    else: # коллинеарны
        return 0


# Проверка пересечений отрезков
def intersect(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round):
    if (check_orientation(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round) !=
            check_orientation(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round) and
            check_orientation(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current) !=
            check_orientation(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current)):
        return True


# Расстояние между двумя точками
def distance_between_points(x1, y1, x2, y2):
    lenght = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    return lenght


# Минимальное расстояние между отрезком и точкой
def distance_between_point_segment(x1, y1, x3, y3, x_round, y_round):
    L1 = distance_between_points(x_round, y_round, x1, y1)
    L2 = distance_between_points(x_round, y_round, x3, y3)
    L = distance_between_points(x1, y1, x3, y3)  # длина ГС
    if (L1 * L1 > L2 * L2 + L * L) or (L2 * L2 > L1 * L1 + L * L):
        P = min(L1, L2)
    else:
        if x1 == x3 and y1 != y3:
            x_base = x1
            y_base = y_round
        elif y1 == y3 and x1 != x3:
            x_base = x_round
            y_base = y1
        elif (x1 == x3 and y1 == y3) or (np.isnan(x3) and np.isnan(y3)):
            x_base = x1
            y_base = y1
        else:
            A = y3 - y1
            B = x1 - x3
            C = -1 * x1 * (y3 - y1) + y1 * (x3 - x1)

            x_base = (x1 * ((y3 - y1)**2) + x_round * ((x3 - x1)**2) + (x3 - x1) * (y3 - y1) * (y_round - y1)) / ((y3 - y1)**2 + (x3 - x1)**2)
            y_base = (x3 - x1) * (x_round - x_base) / (y3 - y1) + y_round

            x_base = (B * x_round / A - C / B - y_round) * A * B / (A * A + B * B)
            y_base = B * x_base / A + y_round - B * x_round / A

        P = distance_between_points(x_round, y_round, x_base, y_base)

    return P


# Минимальное расстояние между двумя непересекающимися отрезками
def min_lenght(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round):
    # if x1_current == x3_current and y1_current == y3_current
    min_len = min(distance_between_point_segment(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round),
                  distance_between_point_segment(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round),
                  distance_between_point_segment(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current),
                  distance_between_point_segment(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current))
    return min_len


def errors(Flag):
    root = tk.Toplevel()
    root.title('Error')
    root.attributes('-toolwindow', True)
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(application_path, 'point.png')
    print(config_path)
    logo = tk.PhotoImage(file=config_path)
    root.iconphoto(False, logo)
    root.config(bg='light grey')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
    root.geometry("200x80+200+200")  # размер окна и расположение (его можно не указывать)
    root.resizable(width=False, height=False)  # Нельзя изменять размер окна
    if Flag == 1:
        text = "Исходные данные \n ""не выбраны!"
    elif Flag == 2:
        text = "Ошибка в расчёте!"
    else:
        text = "Расчёт не выполнен!"
    label_1 = tk.Label(root, text=text,
                       bg='light grey',
                       fg='black',
                       font=("Arial", 9, "bold"),
                       # padx=20, # отступ по x
                       # pady=30, # отступ по y
                       # width=20, # ширина
                       # height=10, # высота
                       # anchor='n', # расположение текста в лейбле
                       justify=tk.CENTER)
    label_1.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
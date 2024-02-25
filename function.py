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

    if res > 0: # по часовой стрелке
        return 1
    elif res < 0: # против часовой стрелки
        return 2
    else: # коллинеарны
        return 0

# Проверка пересечений отрезков
def intersect(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round):
    if check_orientation(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round) != check_orientation(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round) and \
            check_orientation(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current) != check_orientation(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current):
        return True

# Расстояние между двумя точками
def distance_between_points(x1, y1, x2, y2):
    lenght = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    return lenght

# Минимальное расстояние между отрезком и точкой
def distance_between_point_segment(x1, y1, x3, y3, x_round, y_round):
    L1 = distance_between_points(x_round, y_round, x1, y1)
    L2 = distance_between_points(x_round, y_round, x3, y3)
    L = distance_between_points(x1, y1, x3, y3) # длина ГС
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


    # print(f'расстояние - {P}')

    return P

# Минимальное расстояние между двумя непересекающимися отрезками
def min_lenght(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round):
    # if x1_current == x3_current and y1_current == y3_current
    min_len = min(distance_between_point_segment(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round), distance_between_point_segment(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round), \
                  distance_between_point_segment(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current), distance_between_point_segment(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current))
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



class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

# Given three collinear points p, q, r, the function checks if
# point q lies on line segment 'pr'
def onSegment(p, q, r):
    if ((q.x <= max(p.x, r.x)) and (q.x >= min(p.x, r.x)) and
            (q.y <= max(p.y, r.y)) and (q.y >= min(p.y, r.y))):
        return True
    return False

def orientation(p, q, r):
    # to find the orientation of an ordered triplet (p,q,r)
    # function returns the following values:
    # 0 : Collinear points
    # 1 : Clockwise points - по часовой
    # 2 : Counterclockwise - против часовой

    val = (float(q.y - p.y) * (r.x - q.x)) - (float(q.x - p.x) * (r.y - q.y))
    if (val > 0):
        # Clockwise orientation
        return 1
    elif (val < 0):
        # Counterclockwise orientation
        return 2
    else:
        # Collinear orientation
        return 0

# The main function that returns true if
# the line segment 'p1q1' and 'p2q2' intersect.
def doIntersect(p1, q1, p2, q2):
    # Find the 4 orientations required for
    # the general and special cases
    o1 = orientation(p1, q1, p2)
    o2 = orientation(p1, q1, q2)
    o3 = orientation(p2, q2, p1)
    o4 = orientation(p2, q2, q1)

    # General case
    if ((o1 != o2) and (o3 != o4)):
        return True

    # Special Cases

    # p1 , q1 and p2 are collinear and p2 lies on segment p1q1
    if ((o1 == 0) and onSegment(p1, p2, q1)):
        return True

    # p1 , q1 and q2 are collinear and q2 lies on segment p1q1
    if ((o2 == 0) and onSegment(p1, q2, q1)):
        return True

    # p2 , q2 and p1 are collinear and p1 lies on segment p2q2
    if ((o3 == 0) and onSegment(p2, p1, q2)):
        return True

    # p2 , q2 and q1 are collinear and q1 lies on segment p2q2
    if ((o4 == 0) and onSegment(p2, q1, q2)):
        return True

    # If none of the cases
    return False


# # Driver program to test above functions:
# p1 = Point(1, 1)
# q1 = Point(10, 1)
# p2 = Point(1, 2)
# q2 = Point(10, 2)
#
# if doIntersect(p1, q1, p2, q2):
#     print("Yes")
# else:
#     print("No")
#
# p1 = Point(10, 0)
# q1 = Point(0, 10)
# p2 = Point(0, 0)
# q2 = Point(10, 10)
#
# if doIntersect(p1, q1, p2, q2):
#     print("Yes")
# else:
#     print("No")
#
# p1 = Point(-5, -5)
# q1 = Point(0, 0)
# p2 = Point(1, 1)
# q2 = Point(10, 10)
#
# if doIntersect(p1, q1, p2, q2):
#     print("Yes")
# else:
#     print("No")




# Return true if line segments AB and CD intersect
# def intersect(x1, y1, x2, y2, x3, y3, x4, y4):
#     if check_orientation(x1, y1, x2, y2, x3, y3) != check_orientation(x1, y1, x2, y2, x4, y4) and \
#             check_orientation(x3, y3, x4, y4, x1, y1) != check_orientation(x3, y3, x4, y4, x2, y2):
#         return True

# def intersect(row, x1_current, y1_current, x3_current, y3_current):
#     if check_orientation(x1_current, y1_current, x3_current, y3_current, row[x1], row[y1]) != check_orientation(x1_current, y1_current, x3_current, y3_current, row[x3], row[y3]) and \
#             check_orientation(row[x1], row[y1], row[x3], row[y3], x1_current, y1_current) != check_orientation(row[x1], row[y1], row[x3], row[y3], x3_current, y3_current):
#         return True
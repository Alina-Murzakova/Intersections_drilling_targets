import tkinter as tk
from tkinter import filedialog as fd
import customtkinter as ctk
import subprocess
from tkinter.ttk import Progressbar
import os
import sys
import pandas as pd
from main import target
from function import errors
# from tqdm import tqdm

def load_file(button, entry):
    # global data_file
    # global file_name
    file_name = fd.askopenfilename(defaultextension=('.xlsx', '.xls'))
    print(file_name)
    entry.configure(state='normal')
    entry.delete(0, tk.END) # очищение виджета
    if file_name: # файл выбран
        entry.insert(0, file_name) # вставка текста в виджет
        entry.configure(state='disabled')
    else: # файл не выбран
        entry.insert(0, 'Выберите файл')
    # return

def get_file(event):
    data_init = entry_1.get() # получение исходного файлы
    distance = int(entry_2.get()) # получение минимального расстояния
    diff_depth = int(entry_3.get()) # получение минимальной разницы абс отметок (для понимания один пласт или нет)

    if data_init == 'Выберите файл':
        Flag = 1
        errors(Flag)

    else:
        pb['value'] = 0
        win.update_idletasks()
        label_4['text'] = round(pb['value']), '%'
        win.update_idletasks()

        df_initial = pd.read_excel(os.path.join(os.path.dirname(__file__), data_init), header=2)  # Открытие экселя
        try:
            target(df_initial, data_init, distance, diff_depth, win, pb, label_4)
        except:
            Flag = 2
            errors(Flag)

def open_report():
    if entry_1.get() == 'Выберите файл':
        Flag = 1
        errors(Flag)
    elif pb['value'] > 99.9:
        subprocess.Popen(os.path.join(str(os.path.dirname(entry_1.get()) + "/Result"), os.path.basename(entry_1.get())).replace(".xlsx", "") + "_out.xlsx", shell=True)
        # print(str(os.path.dirname(entry_1.get())))
        # print(str(os.path.abspath(entry_1.get())))
        # = str(os.path.dirname(os.path.abspath(data_file))) + "/Result"
        # str(os.path.join((dir_result + "/"), os.path.basename(data_file))).replace(".xlsx", "") + "_out.xlsx")
    else:
        Flag = 3
        errors(Flag)


# Начало
win = tk.Tk()  # главное окно, еще называют root
win.title('Поиск пересечений целей бурения')

# метод получения пути приложения
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

config_path = os.path.join(application_path, 'logo.png')
logo = tk.PhotoImage(file=config_path)
win.iconphoto(False, logo)
win.config(bg='#DADADA')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
win.geometry("400x400+100+100")  # размер окна и расположение (его можно не указывать)
# win.minsize(400, 300) # минимальный возможный размер, если True в resizable
# win.minsize(800, 700) # максимальный возможный размер, если True в resizable
win.resizable(width=False, height=False)  # Нельзя изменять размер окна


# вставляем виджеты
frame_1 = ctk.CTkFrame(master=win,
                       width=255,
                       height=170,
                       corner_radius=20,
                       fg_color='#EDEDED'
                       )

frame_2 = ctk.CTkFrame(master=win,
                       width=255,
                       height=55,
                       corner_radius=20,
                       fg_color='#EDEDED'
                       )

label_1 = tk.Label(master=frame_1, text="Исходные данные",
                   bg='#EDEDED',
                   fg='black',
                   font="Arial 11 bold",
                   # padx=20, # отступ по x
                   # pady=30, # отступ по y
                   # width=50, # ширина
                   # height=10, # высота
                   # anchor='n', # расположение текста в лейбле
                   justify=tk.CENTER
                   )

label_2 = tk.Label(master=frame_1, text="Минимальное расстояние, м:",
                   bg='#EDEDED',
                   fg='black',
                   font="Arial 9",
                   justify=tk.CENTER)

label_3 = tk.Label(master=frame_1, text="Разница абс.отметок, м:",
                   bg='#EDEDED',
                   fg='black',
                   font="Arial 9",
                   justify=tk.CENTER)

label_4 = tk.Label(master=win, text="",
                   bg='#DADADA',
                   fg='black',
                   font="Arial 9",
                   justify=tk.CENTER)

btn_1 = tk.Button(master=frame_1, text='Цели бурения', width=33, bg='#909090', fg='white')
btn_2 = ctk.CTkButton(win, text='Расчёт', corner_radius=10, width=150, fg_color='#0070BA', font=('Arial', 12,  'bold'))
btn_3 = tk.Button(master=frame_2, text='Открыть отчёт', command=open_report, width=33, bg='#909090', fg='white')

entry_1 = tk.Entry(master=frame_1, width=34)
entry_1.insert(0, 'Выберите файл')

entry_2 = tk.Entry(master=frame_1, width=8, justify='center')
entry_2.insert(0, '150')

entry_3 = tk.Entry(master=frame_1, width=8, justify='center')
entry_3.insert(0, '30')

btn_1.bind('<ButtonRelease-1>', lambda event, button=btn_1, entry=entry_1: load_file(button, entry))
btn_2.bind('<ButtonRelease-1>', get_file)

pb = Progressbar(win, orient="horizontal", mode="determinate", length=150, maximum=100, value=0)

frame_1.place(relx=0.5, rely=0.3, anchor=tk.CENTER)
frame_2.place(relx=0.5, rely=0.86, anchor=tk.CENTER)
label_1.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
label_2.place(relx=0.07, rely=0.6)
label_3.place(relx=0.07, rely=0.76)
label_4.place(relx=0.7, rely=0.68)
btn_1.place(relx=0.5, rely=0.26, anchor=tk.CENTER)
btn_2.place(relx=0.5, rely=0.61, anchor=tk.CENTER)
btn_3.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
entry_1.place(relx=0.5, rely=0.44, anchor=tk.CENTER)
entry_2.place(relx=0.82, rely=0.66, anchor=tk.CENTER)
entry_3.place(relx=0.82, rely=0.82, anchor=tk.CENTER)
pb.place(relx=0.5, rely=0.71, anchor=tk.CENTER)


# чтобы exe не работал в фоновом режиме после закрытия
def on_closing():
    win.quit
    win.destroy()
    sys.exit()

win.protocol("WM_DELETE_WINDOW", on_closing)


win.mainloop()  # запускает цикл обработки событий; пока мы не вызовем эту функцию, окно не откроется





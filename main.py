from Tkinter import *  # Импорт всех классов и функций из библиотеки Tkinter для создания графического интерфейса
from decimal import *  # Импорт Decimal для работы с десятичными числами (не используется в коде)
import cx_Oracle  # Импорт cx_Oracle для подключения к базе данных Oracle
import time  # Импорт модуля time для работы с временем (не используется в коде)
from threading import Timer  # Импорт Timer для создания таймеров
from ctypes import *  # Импорт ctypes для взаимодействия с C-библиотеками
import ConfigParser  # Импорт ConfigParser для чтения конфигурационных файлов


def oracle_connect():
    # Создание объекта ConfigParser для чтения конфигурационного файла 'meteo.ini'
    config = ConfigParser.RawConfigParser()
    config.read('meteo.ini')

    # Получение данных пользователя и пароля, а также алиаса для подключения к базе данных из конфигурационного файла
    user = config.get('xxxxx', 'user')
    cpwd = config.get('xxxxx', 'xxxxx')
    alias = config.get('oracle', 'alias')

    # Загрузка и использование DLL для дешифрования пароля
    cfu = cdll.LoadLibrary("chippercrypt.dll").DeCryptString
    cfu.restype = c_char_p
    pwd = cfu(cpwd, user)

    # Подключение к базе данных Oracle и возвращение курсора для выполнения запросов
    return cx_Oracle.connect(user, pwd, alias).cursor()


def update_all():  # Функция для обновления данных в графическом интерфейсе
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id = (SELECT MAX (m_id)
                 FROM refbus.meteo)""")

    # Обновление текста меток в GUI с последними данными из базы данных
    for d, t, v in cursor:
        vd["text"] = d
        vt["text"] = t
        vw["text"] = v


def export_excel_slk():
    # Экспорт данных в файл .slk (Spreadsheet Link)
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id + 24 > (SELECT MAX (m_id)
                 FROM refbus.meteo)
         ORDER BY 1 desc""")

    # Открытие шаблона .slk и создание нового файла отчета
    templ = open('meteo.slk', 'r')
    ft = open('meteo_report.slk', 'w')

    # Чтение шаблона и замена меток на актуальные данные из базы данных
    for d, t, w in cursor:
        while 1:
            line = templ.readline()
            if line == '':
                break
            dok = tok = wok = False
            if line.find("t_date") != -1 and not dok:
                ft.write(line.replace("t_date", str(d)))
                dok = True
                break
            elif line.find("t_temperature") != -1 and not tok:
                ft.write(line.replace("\"t_temperature\"", str(t)))
                tok = True
                break
            elif line.find("t_wind_speed") != -1 and not wok:
                ft.write(line.replace("\"t_wind_speed\"", str(w)))
                wok = True
                break
            else:
                ft.write(line)
            if dok and tok and wok:
                break
    while 1:
        line = templ.readline()
        if line == '':
            break
        ft.write(line)
    ft.close()
    templ.close()


def export_excel():
    # Экспорт данных в CSV файл
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id + 24 > (SELECT MAX (m_id)
                 FROM refbus.meteo)
         ORDER BY 1 desc""")

    # Открытие нового CSV файла и запись заголовков
    f = open('meteo_report.csv', 'w')
    f.write("Date;Temperature;Wind Speed\n")

    # Запись данных из базы в CSV файл, замена точек на запятые для десятичных чисел
    for d, t, w in cursor:
        f.write(str(d) + ";" +
                str(t).replace(".", ",") + ";" +
                str(w).replace(".", ",") + "\n")
    f.close()


# Основная программа

# Создание графического интерфейса
tk = Tk()
tk.title("Meteo")  # Установка заголовка окна
tk.geometry('280x125')  # Установка размеров окна

# Создание виджетов (меток и кнопок) для отображения данных и управления программой
hd = Label(tk, text="Date")
ht = Label(tk, text="Temperature")
hw = Label(tk, text="Wind")
vd = Label(tk, text="01.01.2011")
vt = Label(tk, text="0")
vw = Label(tk, text="0")

bu = Button(tk, text="Update")
be = Button(tk, text="Excel export")
bes = Button(tk, text="Excel slk export")

# Настройка макета виджетов с использованием сетки
top = tk.winfo_toplevel()
top.rowconfigure(6, weight=1)
top.columnconfigure(1, weight=1)
tk.rowconfigure(6, weight=1)
tk.columnconfigure(1, weight=1)

hd.grid(row=0, column=0)
ht.grid(row=1, column=0)
hw.grid(row=2, column=0)
vd.grid(row=0, column=1)
vt.grid(row=1, column=1)
vw.grid(row=2, column=1)

bu.grid(row=3, column=1, sticky=N + S + E + W)
be.grid(row=4, column=1, sticky=N + S + E + W)
bes.grid(row=5, column=1, sticky=N + S + E + W)

# Подключение к базе данных, создание таймера, установка команд для кнопок и запуск основного цикла событий
cursor = oracle_connect()

t = Timer(10 * 60, update_all)  # Создание таймера для обновления данных каждые 10 минут
t.start()

bu["command"] = update_all
be["command"] = export_excel
bes["command"] = export_excel_slk

update_all()  # Начальная загрузка данных
tk.mainloop()  # Запуск основного цикла обработки событий

t.cancel()  # Остановка таймера при завершении программы
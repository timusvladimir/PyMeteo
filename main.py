from Tkinter import *
from decimal import * # for py2exe
import cx_Oracle
import time
from threading import Timer
from ctypes import *
import ConfigParser


def oracle_connect():
    config = ConfigParser.RawConfigParser()
    config.read('meteo.ini')
    user = config.get('xxxxx', 'user')
    cpwd = config.get('xxxxx', 'xxxxx')
    alias = config.get('oracle', 'alias')

    cfu = cdll.LoadLibrary("chippercrypt.dll").DeCryptString
    cfu.restype = c_char_p
    pwd = cfu(cpwd, user)
    return cx_Oracle.connect(user, pwd, alias).cursor()


def update_all(): #(vd, vt, vw):
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id = (SELECT MAX (m_id)
                 FROM refbus.meteo)""")
    for d, t, v in cursor:
        vd["text"] = d
        vt["text"] = t
        vw["text"] = v


def export_excel_slk():
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id + 24 > (SELECT MAX (m_id)
                 FROM refbus.meteo)
         ORDER BY 1 desc""")
    templ = open('meteo.slk', 'r')
    ft = open('meteo_report.slk', 'w')


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
                break;
    while 1:
        line = templ.readline()
        if line == '':
            break
        ft.write(line)
    ft.close()
    templ.close()


def export_excel():
    cursor.execute("""
        SELECT m_datetime, m_temperature, m_wind_speed
          FROM refbus.meteo
         WHERE m_id + 24 > (SELECT MAX (m_id)
                 FROM refbus.meteo)
         ORDER BY 1 desc""")
    f = open('meteo_report.csv', 'w')
    f.write("Date;Temperature;Wind Speed\n")

    for d, t, w in cursor:
        f.write(str(d) + ";" + 
            str(t).replace(".", ",") + ";" + 
            str(w).replace(".", ",") + "\n")
    f.close()

# Main Program

# Create widgets
tk = Tk()
tk.title("Meteo")
tk.geometry('280x125')
hd = Label(tk, text="Date")
ht = Label(tk, text="Temperature")
hw = Label(tk, text="Wind")
vd = Label(tk, text="01.01.2011")
vt = Label(tk, text="0")
vw = Label(tk, text="0")

bu = Button(tk, text="Update")
be = Button(tk, text="Excel export")
bes = Button(tk, text="Excel slk export")

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

bu.grid(row=3, column=1, sticky=N+S+E+W)
be.grid(row=4, column=1, sticky=N+S+E+W)
bes.grid(row=5, column=1, sticky=N+S+E+W)

# Ok, let's connect to oracle,
# create timer, set buttons and go event loop
cursor = oracle_connect()

t = Timer(10 * 60, update_all) # 10 min
t.start()

bu["command"] = update_all
be["command"] = export_excel
bes["command"] = export_excel_slk

update_all()
tk.mainloop()

t.cancel()
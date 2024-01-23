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
    user = config.get('oracle', 'user')
    cpwd = config.get('oracle', 'password')
    alias = config.get('oracle', 'alias')
    return cx_Oracle.connect(user, cpwd, alias).cursor()

def update_all(): #(vd, vt, vw):
    cursor.execute("""_______""")
    for d, t, v in cursor:
        vd["text"] = d
        vt["text"] = t
        vw["text"] = v

# Main Program

# Create widgets
tk = Tk()
tk.title("Test")
tk.geometry('280x125')
hd = Label(tk, text="Date")
ht = Label(tk, text="Temperature")
hw = Label(tk, text="AVNR")
vd = Label(tk, text="01.01.2011")
vt = Label(tk, text="0")
vw = Label(tk, text="0")

bu = Button(tk, text="Update")

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

# Ok, let's connect to oracle,
# create timer, set buttons and go event loop
cursor = oracle_connect()

t = Timer(5 * 60, update_all) # 5 min
t.start()

bu["command"] = update_all

update_all()
tk.mainloop()

t.cancel()

import psutil
import tkinter as tk
from tkinter import ttk
import subprocess
import sys
from threading import Thread
import configparser
import os.path
import ctypes
import time

LOG = False

config = configparser.ConfigParser()
config.read("settings.ini")
name_path = {}
# Разрешить\запретить индивидуальные настройки РДП
run_local_rdp = config.get("CONNECT", "Local_RDP")
ts = config.get("CONNECT", "TS")
if run_local_rdp == "True":
    run_local_rdp = True
    if LOG:
        print("лольный рдп активен")
else:
    run_local_rdp = False
# Разрешить\запретить запуск локальных приложений
allow_local_progs = config.get("CONNECT", "Local_Programs")
if allow_local_progs == "True":
    allow_local_progs = True
    if LOG:
        print("локальные программы активированы")
else:
    allow_local_progs = False

mstsc = 'C:/Windows/system32/mstsc.exe'

if run_local_rdp:
    rdp_file = r"C:/RDP Control/Local.rdp"
else:
    rdp_file = (r"\\run.contabo.de\rdp" + "\\" + ts)

if allow_local_progs:
    w = '270'
    h = '200'
else:
    w = '270'
    h = '160'


# При наведении меняет цвет кнопок
class HoverButton(tk.Button):
    def __init__(self, master, **kw):
        tk.Button.__init__(self, master=master, **kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self['background'] = self['activebackground']

    def on_leave(self, e):
        self['background'] = self.defaultBackground


class Window(tk.Tk):
    def __init__(self):
        super().__init__()

        def start_program(p):
            subprocess.call(name_path[p])

        def run_program():
            p = cmb.get()
            Thread(target=start_program, args=(p,), daemon=True).start()

        # Размер
        self.minsize(width=w, height=h)
        self.maxsize(width=w, height=h)
        # Убирает обрамление
        self.overrideredirect(1)
        # Центрирование на экране
        x = (self.winfo_screenwidth() / 2 - (int(w) / 2))
        y = (self.winfo_screenheight() / 2 - (int(h) / 2))
        self.wm_geometry("+%d+%d" % (x, y))
        # Поверх всех окон
        '''if not allow_local_progs:
            self.call('wm', 'attributes', '.', '-topmost', '1')'''
        self["bg"] = "black"
        # Кнопки
        remote = HoverButton(self, text='НАЧАТЬ РАБОТУ', background='#A9F5A9', activebackground='#2EFE2E',
                             font=("Verdana", 8, "bold"), command=self.start_mstsc)
        remote.place(anchor='nw', height='40', width='250', x='10', y='10')
        change = HoverButton(self, text='СМЕНИТЬ ПОЛЬЗОВАТЕЛЯ', background='#F2F5A9', activebackground='#FFFF00',
                             font=("Verdana", 8, "bold"), command=change_user)
        change.place(anchor='nw', height='40', width='250', x='10', y='60')
        shutdown = HoverButton(self, text='ВЫКЛЮЧИТЬ КОМПЬЮТЕР', background='#F5A9A9', activebackground='#FE2E2E',
                               font=("Verdana", 8, "bold"), command=power_off)
        shutdown.place(anchor='nw', height='40', width='250', x='10', y='110')
        run = HoverButton(self, text='ЗАПУСТИТЬ', background='#FFA500', activebackground='#FFD700',
                          font=("Verdana", 8, "bold"), command=run_program)
        run.place(anchor='nw', height='30', width='90', x='170', y='160')
        # Комбобокс
        cmb = ttk.Combobox(self, values=list_prog, state="readonly")
        cmb.place(anchor='nw', height='30', width='150', x='10', y='160')
        if not list_prog:
            cmb.current()
        else:
            cmb.current(0)

    def run(self):
        if LOG:
            print("показываю gui")
        self.mainloop()

    def start_mstsc(self):
        check_ver()
        on_start_kill_proc()
        if not allow_local_progs:
            # Сокрытие окна
            Thread(target=self.start_rdp, daemon=True).start()
        else:
            Thread(target=self.start_rdp, daemon=True).start()

    def start_rdp(self):
        if LOG:
            print("запускаю рдп")
        result = os.path.exists(rdp_file)
        if result:
            subprocess.call(['C:/Windows/system32/mstsc.exe', rdp_file])
        else:
            warning = "Нет связи"
            mb = ctypes.windll.user32.MessageBoxW
            mb(None, warning, 'Ошибка 0x01', 0)


def show_window():
    root = Window()
    root.run()


def search():
    rdp = 'mstsc.exe'
    for proc in psutil.process_iter():
        if proc.name() == rdp:
            return True


def power_off():
    subprocess.check_output('shutdown /s /t 0')
    sys.exit()


def change_user():
    subprocess.check_output('shutdown /L')
    sys.exit()


def check_ver():
    if LOG:
        print("проверяю версию программы")
    def start_update():
        subprocess.call(r'C:/RDP Control/Updater.exe')

    config.read("ver.ini")
    current = config.get("VERSION", "ver")
    config.read(r"\\run.contabo.de\rdp\Update\version.ini")
    latest = config.get("VERSION", "ver")
    if current != latest:
        if LOG:
            print("обновляю")
        Thread(target=start_update, daemon=True).start()


def on_start_kill_proc():
    if LOG:
        print("убиваю процесс рдп")
    rdp = 'mstsc.exe'
    for proc in psutil.process_iter():
        if proc.name() == rdp:
            proc.kill()


def as_dict(cfg):
    dictionary = {}
    for section in cfg.sections():
        dictionary[section] = []
        for option in cfg.options(section):
            dictionary[section].append(cfg.get(section, option))
    return dictionary


def check_progs():
    global list_prog, name_path
    list_prog = []
    cp = configparser.ConfigParser()
    if LOG:
        print("читаю файл программ")
    cp.read(r"\\run.contabo.de\rdp\progs.ini")
    temp_dict = as_dict(cp)
    i = 1
    while i <= len(temp_dict):
        for val in temp_dict[str(i)]:
            result = os.path.exists(val)
            if result:
                list_prog.append(temp_dict[str(i)][0])
                name_path[temp_dict[str(i)][0]] = val
                break
        i += 1


def mainloop():
    if LOG:
        print("старт программы")
    global list_prog
    if allow_local_progs:
        check_progs()
    else:
        list_prog = "Closed"
    on_start_kill_proc()
    while True:
        time.sleep(0.5)
        rdp = search()
        if rdp:
            continue
        else:
            show_window()
            #Thread(target=show_window, daemon=True).start()
    # root = Window()
    # ninja = root.run()
    # Thread(target=root.run(), daemon=True).start()
    # while True:
    #    time.sleep(1)
    #    print("ok")
    # root.focus_force()


if __name__ == '__main__':
    mainloop()

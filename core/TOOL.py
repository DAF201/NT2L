from tkinter import Tk
from tkinter.filedialog import askopenfilename, askopenfilenames, askdirectory
import datetime
import winsound
import gc
import sys
import config.config
import glob
import os
import itertools
import win32gui
import win32con
import win32console
import subprocess


tk_root = Tk()
tk_root.overrideredirect(True)
tk_root.geometry('0x0+0+0')
tk_root.attributes('-topmost', True)
tk_root.withdraw()
tk_root.update()


RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
BLUE = "\033[34m"
RESET = "\033[0m"


def get_runtime() -> list[int, int, int]:
    """get how long has the system beed running"""
    system_operated_time = (
        datetime.datetime.now() - config.config.system_start_time).total_seconds()
    operate_hour = system_operated_time // 3600
    operate_min = (system_operated_time % 3600) // 60
    operate_sec = system_operated_time % 60
    return int(operate_hour), int(operate_min), int(operate_sec)


def get_today_date() -> str:
    """today's month/day, both can be 1 digit"""
    now = datetime.datetime.now()
    return f"{now.month}/{now.day}"


def get_today_date_pad2() -> str:
    """today's month/day, with month being at least 2 digits"""
    now = datetime.datetime.now()
    return now.strftime(r"%m%d")


def message(*args, **kwargs) -> None:
    """regular message"""
    operate_hour, operate_min, operate_sec = get_runtime()
    caller_name = sys._getframe(1).f_code.co_name
    time_str = f"{str(operate_hour).rjust(2, '0')}:{str(operate_min).rjust(2, '0')}:{str(operate_sec).rjust(2, '0')}"
    print(f"[{time_str}] {caller_name}:\t", *args, **kwargs)
    gc.collect()


def alert(*args, **kwargs) -> None:
    """alert message in RED"""
    operate_hour, operate_min, operate_sec = get_runtime()
    caller_frame = sys._getframe(1)
    caller_name = caller_frame.f_code.co_name
    caller_line = caller_frame.f_lineno
    time_str = f"{str(operate_hour).rjust(2, '0')}:{str(operate_min).rjust(2, '0')}:{str(operate_sec).rjust(2, '0')}"
    print(
        f"[{time_str}] {caller_name} {RED}@LINE:{caller_line}{RESET}\t", *args, **kwargs)
    gc.collect()


def alert_and_beep(msg) -> None:
    """alert with noise"""
    alert(msg)
    winsound.Beep(1000, 1000)


def select_file(title: str, format=None) -> str:
    """format is in form of [('image','*.jpg;*.png;*.bmp')] or something like that"""
    if format is None:
        format = [("all file", "*.*")]
    return askopenfilename(parent=tk_root, title=title, filetypes=format)


def select_files(title: str, format=None) -> list[str]:
    """format is in form of [('images','*.jpg;*.png;*.bmp')] or something like that"""
    if format is None:
        format = [("all files", "*.*")]
    return askopenfilenames(tk_root, title=title, filetypes=format)


def select_directory(title: str) -> str:
    """select a directory"""
    return askdirectory(title=title)


def file_search_in_directory(dir: str, key: str, num_of_files=1) -> list[str]:
    """search for file start with this key in this directory"""
    files = glob.iglob(os.path.join(dir, key))
    return list(itertools.islice(files, num_of_files))


def get_invoice() -> str:
    """get the current invoice number for GR, will update counter each call"""
    invoice = f"{config.config.global_config['GR']['invoice_header']}{str(config.config.global_config['GR']['invoice_record']).rjust(4, '0')}"
    config.config.global_config["GR"]["invoice_record"] = config.config.global_config["GR"]["invoice_record"]+1
    config.config.config_update()
    return invoice


def focus_console() -> None:
    """pull the console to top to input"""
    try:
        hwnd = win32console.GetConsoleWindow()
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
    except:
        pass


def OneDrive_Sync():
    subprocess.run(
        ["cmd", "/c", f'attrib +U "{config.config.global_config["Excel"]["target_table"]}"'])
    subprocess.Popen(
        [
            "cmd", "/c",
            f'{config.config.global_config["OneDrive"]["path"]} /background"'
        ])

import os
import json
import datetime

global_config: dict
system_start_time = datetime.datetime.now()

config_path = os.path.join(os.path.dirname(
    __file__), '..', 'config', 'config.json')
config_path = os.path.abspath(config_path)


def config_init() -> None:
    """kill all the excel process to avoid those random conflicts between Excel app and Excel COM"""
    os.system("taskkill /f /im excel.exe >nul 2>&1")
    global global_config
    with open(config_path, "r") as config:
        global_config = json.load(config)


def config_update() -> None:
    """mostly for update invoice number"""
    with open(f"{os.path.dirname(__file__)}\\config.json", "w") as config:
        json.dump(global_config, config)


config_init()

# import pandas as pd 
# import xlwings as xw
# import re

# path  = r"C:\Users\mnguy\Downloads\Report Scan Verify Shiftly (RCV) (4).xlsm"
# wb = xw.Book(path,update_links=False)
# wb.save()
# user_input = input('Please input the sheet name you want detele form n to n: ').strip()
# chosse = re.split('\D+', user_input)

# for i in range(int(chosse[0]), int(chosse[1]) + 1):
#     sheet_name = f'Sheet{i}'
#     wb.sheets[sheet_name].delete()

# from pywinauto import Application 
import random as rd
from pywinauto import Application
from pywinauto import WindowSpecification
from pywinauto import ElementNotFoundError
import keyboard as kb
import psutil as ps
import time 
import pandas as pd
import pyperclip
import win32gui
import win32api
import my_thread as mt
import threading
import sys 
import os
import ctypes
from ctypes import wintypes
import pythoncom
from my_thread import wait_printing, is_printing,is_alt_l
from pynput import keyboard

global second
second :int = 60
inanttenion = 10

def pressed_alt_l(key)->None:
        if key ==keyboard.Key.alt_l:
                is_alt_l.clear()
        if key ==keyboard.Key.home:
                time.sleep(5)
                is_alt_l.set()
        if key == keyboard.Key.esc:
            os._exit(0)
        if key == keyboard.Key.down and inanttenion > 0:
            inanttenion=-1
            print(inanttenion)
        if key == keyboard.Key.up and inanttenion <101:
            inanttenion=+1
            print(inanttenion)
                





def get_data(path)->pd.DataFrame:
    data  = pd.read_excel(path).values
    return data


def hook_process()-> WindowSpecification:
    try:
        app = Application(backend="uia").connect(title_re=".*HCM Goods Receipt Verification.*")
        window = app.top_window()
        print('Đã lấy được cửa sổ')
        return window
    except Exception as e:
        print(f'Error {e}')
        print(f'Vui lòng bật file HCM GOOD RECEIPT VERIFICATION : Error -> ')
    
def get_form(window:WindowSpecification)->WindowSpecification:
        try:
            form = window.child_window(title='HCM GOOD RECEIPT VERIFICATION',control_type="Window")
            return form
        except Exception as e:
            print(f'Error {e}')
            print('Kiểm tra xem đã bật file HCM GOOD RECEIPT VERIFICATION chưa ')


def get_group_ui(form:WindowSpecification)->WindowSpecification:
    try:
        return{
                'group_jabil' : form.child_window(title='Jabil Information',control_type='Group'),
                'group_vendor': form.child_window(title="Vendor Information", control_type="Group"),
                'good_receipt': form.child_window(title='GoodReceipt Information',control_type="Group")
        }
    except:
        print('Kiểm tra xem đã bật GRN Verify lên hay chưa')


try:
    group_jabil = get_group_ui(get_form(hook_process()))['group_jabil']
    group_vendor = get_group_ui(get_form(hook_process()))['group_vendor']
    group_goodreceipt = get_group_ui(get_form(hook_process()))['good_receipt']

except Exception as e:
    print(f'Error {e}')
    print('Không thể lấy children do cửa sổ GRN Verify chưa được bật vui lòng kiểm tra lại')
    os._exit(0)

is_printing.set()
group_goodreceipt.print_control_identifiers()
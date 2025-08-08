# from pywinauto import Application 
import random as rd
from pywinauto import Application
from pywinauto import WindowSpecification
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
from math import modf
global count 
count  = 0
# def random_phi()->float:
#     phi = (1+5**0.5 )/2
#     frac = modf(time.time())
def pressed_alt_l(key)->None:
        if key ==keyboard.Key.alt_l:
                is_alt_l.clear()
        if key ==keyboard.Key.home:
                time.sleep(5)
                is_alt_l.set()
        if key == keyboard.Key.esc:
            os._exit(0)
                





def get_data(path)->pd.DataFrame:
    data  = pd.read_excel(path).values
    for i in data:
        i[0] = '9K'+str(i[0])
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
                'group_vendor': form.child_window(title="Vendor Information", control_type="Group")
        }
    except:
        print('Kiểm tra xem đã bật GRN Verify lên hay chưa')


try:
    group_jabil = get_group_ui(get_form(hook_process()))['group_jabil']
    group_vendor = get_group_ui(get_form(hook_process()))['group_vendor']

except Exception as e:
    print(f'Error {e}')
    print('Không thể lấy children do cửa sổ GRN Verify chưa được bật vui lòng kiểm tra lại')
    os._exit(0)

def create_UI_mapping(three_ui) -> dict:
    mapping_UI = []
    all_edits= three_ui.descendants(control_type='Edit')
    for  edit in all_edits:
        mapping_UI.append(edit)
    return mapping_UI


element_UI = create_UI_mapping(group_vendor)

def custom_kb(s:str)->None:
            kb.press_and_release(s)
    
EVENT_SYSTEM_FOREGROUND= 0x0003
HOOK_OUT_OF_CONTEXT = 0x0000

WIN_WRAPER_FUNC = ctypes.WINFUNCTYPE(
        None,
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.HWND,
        wintypes.LONG,
        wintypes.LONG,
        wintypes.DWORD,
        wintypes.DWORD
)


def print_waiting()->None:
    icon : list[str] = ['.','..','...']
    while not is_printing.is_set(): 
        wait_printing.wait()
        for i in icon:
            
            sys.stdout.write(f'\rWaiting{i}')
            sys.stdout.flush()
            time.sleep(0.2)
    return


def write_hooking(data:pd.DataFrame,group_vendor,group_jabil:WindowSpecification)->None:
    try:
        text_grn =group_jabil.child_window(title="GRN", control_type="Edit")
        mt.foreground.wait()
        is_alt_l.wait()
        for i in range(0,len(data)):
            text_grn.set_focus()
            custom_kb('ctrl+a')
            mt.foreground.wait()
            is_alt_l.wait()
            custom_kb('backspace')
            mt.foreground.wait()
            is_alt_l.wait()
            pyperclip.copy(data[i][0])
            mt.foreground.wait()
            is_alt_l.wait()
            custom_kb('ctrl+v')
            mt.foreground.wait()
            is_alt_l.wait()
            custom_kb('enter')
            mt.foreground.wait()
            is_alt_l.wait()
            time.sleep(.4)
            mt.foreground.wait()
            is_alt_l.wait()


            for j in range(1,len(data[i])):
                mt.foreground.wait()
                is_alt_l.wait()
                custom_kb('ctrl+a')
                mt.foreground.wait()
                is_alt_l.wait()
                custom_kb('backspace')
                mt.foreground.wait()
                is_alt_l.wait()
                pyperclip.copy(data[i][j])
                mt.foreground.wait()
                is_alt_l.wait()
                custom_kb('ctrl+v')
                mt.foreground.wait()
                is_alt_l.wait()
                custom_kb('enter')
                time.sleep(rd.uniform(0.4,0.6))
                mt.foreground.wait()
                is_alt_l.wait()
                
            kb.press_and_release('enter')
        mt.is_done.set()
        os._exit(0)
        is_printing.set()
        return
    except Exception as e :
        mt.is_done.set()
        print(f'\nKhông thể hooking : error {e}')
        os._exit(0)
        is_printing.set()
        return


def callback(hwWinEvent,event,hdwn,idObject,idChild,dwEventThread,dwmsEventTime):
        
        global count
        if win32gui.GetWindowText(hdwn) =='HCM GOOD RECEIPT VERIFICATION':
            mt.foreground.set()
            wait_printing.clear()
            print(f'\nWINDOW : {win32gui.GetWindowText(hdwn)} [FOCUS]')
        else:
            print(f'\nWINDOW : {win32gui.GetWindowText(hdwn)} [FOCUS]')
            wait_printing.set()
            mt.foreground.clear()
        
            


event_hook = WIN_WRAPER_FUNC(callback)
hook = ctypes.windll.user32.SetWinEventHook(
        EVENT_SYSTEM_FOREGROUND,
        EVENT_SYSTEM_FOREGROUND,
        0,
        event_hook,
        0,
        0,
        HOOK_OUT_OF_CONTEXT
)
def time_return(event:threading.Event,event_child:threading.Event)->None:
        while True:
            second:int = 50
            for i in range(0,second+1):
                event_child.wait()
                if second <1:
                    os._exit(0)
                elif event.is_set() == False:
                    time.sleep(1)
                    print(f'\n[  Count down :{second}  ]')
                    second-=1


def main()->None:
        listener = keyboard.Listener(on_press=pressed_alt_l)
        data = get_data(r"C:\Users\3601183\Desktop\data_input.xlsx")
        task_1 = threading.Thread(target=print_waiting)
        task_2 = threading.Thread(target=write_hooking,args=(data,group_vendor,group_jabil))
        task_3 = threading.Thread(target=time_return,args=(is_printing,wait_printing))
        task_1.start()
        mt.foreground.set()
        task_2.start()
        task_3.start()
        listener.start()
        pythoncom.PumpMessages()
        task_1.join()
        task_2.join()
        task_3.join()
        listener.stop()

main()








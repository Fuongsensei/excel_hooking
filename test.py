import pandas as pd 
import xlwings as xw
import re

path  = r"C:\Users\mnguy\Downloads\Report Scan Verify Shiftly (RCV) (4).xlsm"
wb = xw.Book(path,update_links=False)
wb.save()
user_input = input('Please input the sheet name you want detele form n to n: ').strip()
chosse = re.split('\D+', user_input)

for i in range(int(chosse[0]), int(chosse[1]) + 1):
    sheet_name = f'Sheet{i}'
    wb.sheets[sheet_name].delete()


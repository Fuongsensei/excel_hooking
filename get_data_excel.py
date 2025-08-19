
import os
from posixpath import supports_unicode_filenames
from webbrowser import get
import pandas as pd
import xlwings as xw
import re
import getpass
import random as rd
from openpyxl import Workbook
from ui import print_user_table_clean as putc
from xlwings import Book


# Define the path to the Excel file and the data input file
path = r"C:\Users\mnguy\Downloads\Report Scan Verify Shiftly (RCV) (4).xlsm"
# Open the Excel file with xlwings
wb = xw.Book(path,update_links=False)
ws = wb.sheets['Summary']



def get_data_path() -> str:
    is_exist = os.path.exists(rf"C:\Users\{getpass.getuser()}\Documents\data\data.xlsx")
    if is_exist:
        return rf"C:\Users\{getpass.getuser()}\Documents\data\data.xlsx"
    else:
        os.makedirs(rf"C:\Users\{getpass.getuser()}\Documents\data", exist_ok=True)
        wb  = Workbook()
        wb.active.title = 'data'
        wb.save(rf"C:\Users\{getpass.getuser()}\Documents\data\data.xlsx")
        wb.close()
        return rf"C:\Users\{getpass.getuser()}\Documents\data\data.xlsx"

wb_data = xw.Book(get_data_path(),update_links=False)
ws_data = wb_data.sheets['data']


def get_location_and_df(path:str ,name : str)-> dict:

     df = pd.read_excel(path,sheet_name=name,keep_default_na=True,usecols=[0,1,2,3,4,5,6,7,8,9])
     col_name = ['Site','Shift','Keyin Type','Status','Check Wrong Batch','Check Date code','Check Lot No','Media Code','Total','Percent']
     df.columns = col_name
     filter_na = df[~df['Status'].isna()]

     col_index  = filter_na.columns.tolist().index('Total')+1
     print(f'col_index : {col_index}')
     filter_df = filter_na[filter_na['Status'] == 'Verification scan incomplete']
     row_index = filter_df.index.tolist()
     return{'col_index': col_index,
             'row_index': row_index}




def get_before_sheets(wb:xw.Book)-> list[str]:
    list_sheets = [sheet.name for sheet in wb.sheets ]
    return list_sheets

def click_deital(col:int,row:int|list):
    if isinstance(row,list):
        for i in row:
            ws.range((i+2,col)).api.ShowDetail = True
            
    else:
        ws.range((int(row+2),col)).api.ShowDetail = True
    wb.save()


before_sheets = get_before_sheets(wb)
click_deital(get_location_and_df(path,'Summary')['col_index'],get_location_and_df(path,'Summary')['row_index'])



def get_detail_sheet_name(wb:Book,sheets:list[str])-> list[str]:
    detail_sheets = [sheet.name for sheet in wb.sheets if sheet.name not in sheets]
    return detail_sheets


def get_detail_data(sheets:list[str])-> pd.DataFrame:
    df_detail = pd.read_excel(path,sheet_name=sheets,usecols=[5,12,6,7,19,8])
    df_detail = pd.concat(df_detail.values(),ignore_index=True,axis=0)
    df_detail = df_detail[['GRN Number','Masked MPN','Vendor Date Code','Lot number','Quantity','Name 1']]
    return df_detail


df_detail = get_detail_data(get_detail_sheet_name(wb,before_sheets))
name_vendor = df_detail['Name 1'].unique()


for i,name in enumerate(name_vendor):
    print(f'{i} : {name} ')
user_input = input("Vui lòng chọn nhà cung cấp bạn muốn chạy auto bằng số thứ tự:    ").strip()
vendor_list  = [name_vendor[int(i)] for i in re.split('\D+',user_input)]

after_df = []
def before_data_process(vendor:list)-> pd.DataFrame:
    left_random = rd.randint(1,10)
    right_random = rd.randint(1,10)
    head_index = 0
    for name in vendor:
        temp_df = df_detail[df_detail['Name 1'] == name]
        mpn_list = temp_df[temp_df.columns[1]].unique().tolist()
        for mpn in mpn_list:
            mpn_df = temp_df[temp_df[temp_df.columns[1]]== mpn]
            if left_random <= right_random:
                    
                    after_df.append(mpn_df.head(left_random))
                    mpn_df.drop(mpn_df.iloc[head_index:left_random])

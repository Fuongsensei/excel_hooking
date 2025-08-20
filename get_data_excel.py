
import os
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




def choose_vendor(vendor:list) -> list[str]:
    for i,name in enumerate(vendor):
        print(f'{i} : {name} ')
    user_input = input("Vui lòng chọn nhà cung cấp bạn muốn chạy auto bằng số thứ tự:    ").strip()
    vendor_list  = [vendor[int(i)] for i in re.split(r'\D+',user_input)]
    return vendor_list



def get_data_by_mpn(data:pd.DataFrame)->list[pd.DataFrame]:
    df_mpn  = []
    mpn_list = data['Masked MPN'].unique()
    for i in mpn_list:
            data_mpn = data[data['Masked MPN'] == i].reset_index(drop=True)
            df_mpn.append(data_mpn)
    return df_mpn
        




def random_dataframe(data_list : list[pd.DataFrame]) -> list:
    random_data = []
    user_input = int(input('Nhập số lượng dòng dữ liệu mà bạn muốn lấy ngẫu nhiên từ mỗi MPN nhỏ nhất là 1 lớn nhất là 10 : ').strip())
    os.system('cls')
    if user_input in range(1,11):
        for data in data_list:
            temp_df = []
            if len(data) > user_input:
                while len(data) > 0:
                    left = rd.randint(0, len(data)-1)%11
                    right = rd.randint(0, len(data)-1)%11
                    if left < right:
                        temp_df.append(data.iloc[left:right])
                        data.drop(index=range(left,right), inplace=True)
                    elif left > right:
                        temp_df.append(data.iloc[right:left])
                        data.drop(index=range(right,left), inplace=True)
                        
                    else:
                            temp_df.append(data.iloc[left:left+1])
                            data.drop(index=left, inplace=True)
                    data.reset_index(drop=True, inplace=True)


            else:
                temp_df.append(data)
            random_data.append(pd.concat(temp_df, ignore_index=True))
        return random_data
        

    else:
        print('Số lượng dòng dữ liệu bạn nhập không hợp lệ vui lòng nhập lại từ 1 đến 10')
        return random_dataframe(data_list)
    



def get_data_vendor(vendor_list:list,data:pd.DataFrame)->pd.DataFrame:
    df_vendor = []
    for vendor in vendor_list:
        df_temp = data[data['Name 1'] == vendor]
        df_vendor.append(df_temp)
    return pd.concat(df_vendor, ignore_index=True)

def get_success_data(data:list[pd.DataFrame],vendor,callback) -> pd.DataFrame:
    success_data = pd.concat(data,ignore_index=True)
    print(f'Dữ liệu đã được lấy ngẫu nhiên từ mỗi MPN {success_data}')
    success_data = success_data.astype(str)
    success_data['GRN Number'] = success_data['GRN Number'].apply(lambda x: '9K' + x if not x.startswith('9K') else x)
    success_data = callback(vendor,success_data)
    print(f'Dữ liệu đã được lọc theo nhà cung cấp {success_data['Name 1'].unique()} : {success_data}')
    success_data = success_data[['GRN Number','Masked MPN','Vendor Date Code','Lot number','Quantity']]
    return success_data

success_data = get_success_data(random_dataframe(get_data_by_mpn(df_detail)), choose_vendor(name_vendor), get_data_vendor)
# Save the success data to an Excel file
def write_to_excel(data: pd.DataFrame):
    for i in ['A', 'B', 'C', 'D', 'E']:
        ws_data.range(f'{i}:{i}').number_format = '@'
    ws_data.range('A1').value = [success_data.columns.tolist()] + success_data.values.tolist()
    wb_data.save()
    wb_data.close()

def clear_sheet()->None:
    ws_data.range('A2').expand().clear_contents()
def delete_detail_sheet(sheets_name:list[str])->None:
    for sheet in sheets_name:
        if sheet in wb.sheets:
            wb.sheets[sheet].delete()
    wb.save()

delete_detail_sheet(get_detail_sheet_name(wb,before_sheets))   
clear_sheet()
write_to_excel(success_data)
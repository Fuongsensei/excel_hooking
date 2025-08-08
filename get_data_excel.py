import pandas as pd
import xlwings as xw
import re
import random as rd
path = r"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm"
data_path = r"C:\Users\3601183\Desktop\data_input.xlsx"
wb_data = xw.Book(data_path,update_links=False)
ws_data = wb_data.sheets['data']

df = pd.read_excel(path,sheet_name='Summary',keep_default_na=True,usecols=[0,1,2,3,4,5,6,7,8,9])

col_name = ['Site','Shift','Keyin Type','Status','Check Wrong Batch','Check Date code','Check Lot No','Media Code','Total','Percent']

df.columns = col_name


print(df['Status'].isna())

filter_na = df[~df['Status'].isna()]

print(filter_na)

col_index  = filter_na.columns.tolist().index('Total')+1
print(f'col_index : {col_index}')
filter_df = filter_na[filter_na['Status'] == 'Verification scan incomplete']
print(filter_df)
row_index = filter_df.index.tolist()
print(f'row_index : {row_index}')
wb = xw.Book(path,update_links=False)
ws = wb.sheets['Summary']
list_sheets = [sheet.name for sheet in wb.sheets ]

print(list_sheets)

def click_deital(col:int,row:int|list):
    if isinstance(row,list):
        for i in row:
            ws.range((i+2,col)).api.ShowDetail = True
            
    else:
        ws.range((int(row+2),col)).api.ShowDetail = True
    wb.save()
click_deital(col_index,row_index)

detail_sheet_name = [sheet.name for sheet in wb.sheets if sheet.name  not in list_sheets]
print(f'Sheets detail is : {detail_sheet_name}')


df_detail = pd.read_excel(path,sheet_name=detail_sheet_name,usecols=[5,12,6,7,19,8])
df_detail = pd.concat(df_detail.values(),ignore_index=True,axis=0)
name_vendor = df_detail['Name 1'].unique()

df_detail = df_detail[['GRN Number','Masked MPN','Vendor Date Code','Lot number','Quantity','Name 1']]

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

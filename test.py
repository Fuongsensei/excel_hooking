# import pandas as pd 
# import xlwings as xw


# path  = r"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm"
# des_path  = r"C:\Users\3601183\Desktop\data_input.xlsx"
# wb = xw.Book(des_path,update_links=False)
# ws = wb.sheets['Sheet1']
# df = pd.read_excel(path,sheet_name='dsa',usecols=[5,12,6,7,19])

# df = df[['Unnamed: 5','Unnamed: 12','Unnamed: 7','Unnamed: 6','Unnamed: 19']]
# for i in ['A','B','C','D']:
#     ws.range(f'{i}:{i}').number_format ='@'

# ws.range('A1').value = df.values
# print(df)
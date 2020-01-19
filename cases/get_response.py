from data.get_data import ReadBaseData
from common.read_excel import ReadExcel
import requests
from data.write_data import WriteExcel
import time


# readexcel = ReadExcel(0)
# max_row = readexcel.get_max_row()
host = 'http://auto-api.mobimedical.cn'
excel = ReadBaseData()

# data = excel.get_all_datas()
# for i in range(max_row):
#     time.sleep(5)
#     readexcel = ReadExcel(0)
#     params = excel.get_all_data(i)
#     url = host + params[1]
#     method = params[0]
#     data = params[2]
#     expect = params[3]
#     is_need = params[4]
#     cell = params[5]
#     print(params)
#     print(i)
#     header = {'number': 'P2320190329'}
#     if is_need == 1:
#         if method == 'get':
#             r = requests.get(url, params=data, headers=header)
#             writeexcel = WriteExcel()
#             writeexcel.write_excel_xls(cell,str(r.json()))
#             print(r.json())
#             time.sleep(5)
#             writeexcel1 = WriteExcel()
#             print(i)
#             writeexcel1.get_key(i)
#             time.sleep(5)
#         else:
#             r = requests.post(url, data=data, headers=header)
#             print(r.json())

readexcel = ReadExcel(0)
params = excel.get_all_data(3)

url = host + params[1]
method = params[0]
data = params[2]
expect = params[3]
is_need = params[4]
cell = params[5]
print(data)
# header = {'number': 'P2320190329'}
# if is_need == 1:
#     if method == 'get':
#         r = requests.get(url, params=data, headers=header)
#         writeexcel = WriteExcel()
#         writeexcel.write_excel_xls(cell,str(r.json()))
#         print(r.json())
#         time.sleep(5)
#         writeexcel1 = WriteExcel()
#         print(2)
#         writeexcel1.get_key(3)
#         time.sleep(5)
#     else:
#         r = requests.post(url, data=data, headers=header)
#         print(r.json())

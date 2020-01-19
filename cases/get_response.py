from data.get_data import ReadBaseData
from common.read_excel import ReadExcel
import requests
from data.write_data import WriteExcel
from data.open_excel import just_open
import time


readexcel = ReadExcel(0)
max_row = readexcel.get_max_row()

excel = ReadBaseData()

data = excel.get_all_datas()
for i in range(5):
    readexcel = ReadBaseData()
    params = excel.get_all_data(i)
    url = host + params[1]
    method = params[0]
    data = params[2]
    expect = params[3]
    is_need = params[4]
    cell = params[5]
    print(cell)

    if is_need == 1:
        if method == 'get':
            r = requests.get(url, params=data, headers=header)
            writeexcel = WriteExcel()
            writeexcel.write_excel_xls(cell,str(r.json()))
            writeexcel1 = WriteExcel()
            print(i)
            writeexcel1.get_key(i)
            just_open()
        else:
            r = requests.post(url, data=data, headers=header)
            print(r.json())

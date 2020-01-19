from data.get_data import ReadBaseData
from common.read_excel import ReadExcel
import requests
from data.write_data import WriteExcel
import time


excel = ReadBaseData()
params = excel.get_all_data(3)
host = 'http://auto-api.mobimedical.cn'

url = host + params[1]
method = params[0]
data = params[2]
expect = params[3]
is_need = params[4]
cell = params[5]
print(data)
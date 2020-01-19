import requests
from data.write_data import WriteExcel

excel = WriteExcel()


def test_cases(url, cell):
    header = {'number': 'P2320190329'}
    r = requests.get(url, headers=header)
    data = str(r.json())
    excel.write_excel_xls(cell, data)



test_cases('D2')

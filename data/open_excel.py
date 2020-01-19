from win32com.client import Dispatch
from data.datapath import file_name


def just_open(filename=file_name):
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()

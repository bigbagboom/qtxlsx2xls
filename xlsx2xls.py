from win32com.client import Dispatch
import os
import time

prgpath = os.path.dirname(os.path.abspath(__file__))

def xlsx2xls(filepath):
    xl=Dispatch('Excel.Application')
    wb=xl.Workbooks.Add(filepath)
    wb.Application.DisplayAlerts = False
    wb.SaveAs(filepath[:-1], FileFormat=56)
    xl.Quit()
    
def xls2xlsx(filepath):
    xl=Dispatch('Excel.Application')
    wb=xl.Workbooks.Add(filepath)
    wb.Application.DisplayAlerts = False
    wb.SaveAs(filepath + 'x', FileFormat=51)
    xl.Quit()
    
if __name__ == '__main__':

    start=time.time()
    
    filename='CCS余额报表20211123.xlsx'
    
    filepath = os.path.join(prgpath, filename)
    
    xlsx2xls(filepath)

    end=time.time()

    print(end-start)

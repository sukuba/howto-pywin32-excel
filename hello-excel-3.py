#!python3

import pythoncom
import win32com.client
import datetime

class excel_app(object):
    """
    Excel.Application as a context
    """
    def __init__(self, visible=True):
        self.visible = visible
        self.app = None
        
    def __enter__(self):
        pythoncom.CoInitialize()
        #pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        # these two are almost equivalent.
        # http://docs.activestate.com/activepython/3.4/pywin32/pythoncom__CoInitialize_meth.html
        self.app = win32com.client.DispatchEx('Excel.Application')
        #self.app = win32com.client.Dispatch('Excel.Application')
        # DispatchEx always creates new instance
        # http://stackoverflow.com/questions/18648933/using-pywin32-what-is-the-difference-between-dispatch-and-dispatchex
        self.app.Application.Visible = self.visible
        self.app.Application.AskToUpdateLinks = False
        return self.app
        
    def __exit__(self, type, value, traceback):
        self.app.Quit()
        del self.app
        pythoncom.CoUninitialize()

class excel_book(object):
    """
    Excel.Application.Workbook as a context
    """
    def __init__(self, app, filename):
        self.app = app
        self.filename = filename
        self.book = None
        
    def __enter__(self):
        self.book = self.app.Workbooks.Open(self.filename)
        return self.book
        
    def __exit__(self, type, value, traceback):
        self.book.Saved = True
        self.book.Close()
        del self.book

def main():
    """
    start Excel.
    open an existing book.
    read a cell value at A1.
    write to a cell A2.
    save the book.
    terminate Excel.
    """
    with excel_app() as app:
        print(app)
        input("press enter to continue ...")
        with excel_book(app, r'E:\scratch\Hello1.xls') as book:
             print(book.Worksheets[0].Range("A1").Value)
             book.Worksheets[0].Range("A2").Value = str(datetime.date.today())
             input("press enter to continue ...")
             book.Save()

if __name__ == '__main__':
    main()

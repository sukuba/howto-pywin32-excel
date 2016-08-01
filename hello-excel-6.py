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
        self.app = win32com.client.DispatchEx('Excel.Application')
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
    read cells for the first worksheet.
    terminate Excel.
    
    Result:
    Hello2
    $A$1/年度/ＭＳ Ｐゴシック
    $B$1/分類/ＭＳ Ｐゴシック
    $C$1/タイトル/ＭＳ Ｐゴシック
    $D$1/URL/ＭＳ Ｐゴシック
    $A$1/$C$2
    $B$1/$D$2
    $C$1/$E$2
    $D$1/$F$2
    $A$1/$B$2
    $B$1/$C$2
    $C$1/$D$2
    $D$1/$E$2
    Traceback (most recent call last):
      File "hello-excel-6.py", line 87, in <module>
        main()
      File "hello-excel-6.py", line 81, in main
        print(cell.Address(False, False))
    TypeError: 'str' object is not callable
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheet = book.Worksheets[0]
            print(sheet.Name)
            for row in sheet.UsedRange.Rows:
                for cell in row.Cells:
                    print('/'.join((cell.Address, cell.Value, cell.Font.Name)))
                    # work as expected
                for cell in row.Cells:
                    print('/'.join((cell.Address, cell.Offset(2,3).Address)))
                    # work, but the result is not as expected.
                    # expect "D3", but return "C2".
                    # so, Offset(0,0) in VBA is equivalent to Offset(1,1) in Python.
                for cell in row.Cells:
                    print('/'.join((cell.Address, cell.Range("B2").Address)))
                    # work as expected
                for cell in row.Cells:
                    print(cell.Address(False, False))
                    # ERROR
                    # Address is not a method. Not receive arguments.
                    # So, we can't get the relative address like "A1".

if __name__ == '__main__':
    main()

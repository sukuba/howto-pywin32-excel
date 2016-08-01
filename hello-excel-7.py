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
    read bulk cells for each worksheet.
    terminate Excel.
    
    Result:
    <COMObject <unknown>>
    1
    Hello2
    (('年度', '分類', 'タイトル', 'URL'), ('平成27年', 'ダイズ', 'ダイズの低温特性',
     'soybean/low-temp'), ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen'),
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators'), ('平成28年', 'ナタ
    ネ', 'ナタネの埋土種子', 'oilseed/seedbank'))
    (('年度', '分類', 'タイトル', 'URL'), ('平成27年', 'ダイズ', 'ダイズの低温特性',
     'soybean/low-temp'), ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen'),
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators'), ('平成28年', 'ナタ
    ネ', 'ナタネの埋土種子', 'oilseed/seedbank'))
    (('年度', '分類', 'タイトル', 'URL'), ('平成27年', 'ダイズ', 'ダイズの低温特性',
     'soybean/low-temp'), ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen'),
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators'), ('平成28年', 'ナタ
    ネ', 'ナタネの埋土種子', 'oilseed/seedbank'))
    $A$1:$D$5
    ＭＳ Ｐゴシック
    -4142
    None
    Traceback (most recent call last):
      File "hello-excel-7.py", line 89, in <module>
        main()
      File "hello-excel-7.py", line 84, in main
        print(app.xlColorIndex.xlColorIndexNone)
      File "E:\bin\GRA-dev\bin\WinPython-64bit-3.4.4.3Qt5\python-3.4.4.amd64\lib\sit
    e-packages\win32com\client\dynamic.py", line 527, in __getattr__
        raise AttributeError("%s.%s" % (self._username_, attr))
    AttributeError: Excel.Application.xlColorIndex
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheets = book.Worksheets
            print(sheets)
            print(len(sheets))
            for sheet in sheets:
                print(sheet.Name)
                print(sheet.Range("A1").CurrentRegion())
                # tuple of Cell.Value strings
                print(sheet.Range("A1").CurrentRegion)
                # tuple of Cell.Value strings
                print(sheet.Range("A1").CurrentRegion.Cells)
                # tuple of Cell.Value strings
                print(sheet.Range("A1").CurrentRegion.Address)
                # string $A$1:$D$5
                print(sheet.Range("A1").CurrentRegion.Font.Name)
                # string ＭＳ Ｐゴシック
                print(sheet.Range("A1").CurrentRegion.Interior.ColorIndex)
                # numeric -4142
                sheet.Range("A2").Interior.ColorIndex = 3 # red
                print(sheet.Range("A1").CurrentRegion.Interior.ColorIndex)
                # None
                print(app.xlColorIndex.xlColorIndexNone)
                # ERROR AttributeError: Excel.Application.xlColorIndex
                # so, we must use explicit number -4142 to represent default color

if __name__ == '__main__':
    main()

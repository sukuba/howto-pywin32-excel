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
    ('年度', '分類', 'タイトル', 'URL')
    ('平成27年', 'ダイズ', 'ダイズの低温特性', 'soybean/low-temp')
    ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen')
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators')
    ('平成28年', 'ナタネ', 'ナタネの埋土種子', 'oilseed/seedbank')
    (('年度', '分類', 'タイトル', 'URL'),)
    ('年度', '分類', 'タイトル', 'URL')
    年度
    分類
    タイトル
    URL
    (('平成27年', 'ダイズ', 'ダイズの低温特性', 'soybean/low-temp'),)
    ('平成27年', 'ダイズ', 'ダイズの低温特性', 'soybean/low-temp')
    平成27年
    ダイズ
    ダイズの低温特性
    soybean/low-temp
    (('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen'),)
    ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen')
    平成28年
    ダイズ
    ダイズの花粉
    soybean/pollen
    (('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators'),)
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators')
    平成27年
    ナタネ
    ナタネの食害
    oilseed/predators
    (('平成28年', 'ナタネ', 'ナタネの埋土種子', 'oilseed/seedbank'),)
    ('平成28年', 'ナタネ', 'ナタネの埋土種子', 'oilseed/seedbank')
    平成28年
    ナタネ
    ナタネの埋土種子
    oilseed/seedbank
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheet = book.Worksheets[0]
            print(sheet.Name)
            for row in sheet.UsedRange.Rows():
                print(row)
                # represented as a row tuple
                # so, the following fails, because a tuple doesn't have Cells
                #for cell in row.Cells():
                #    print(cell)
            for row in sheet.UsedRange.Rows:
                print(row)
                # represented as a tuple of a row tuple
                for cell in row.Cells():
                    print(cell)
                # represented as a row tuple
                # doesn't seem what is expected
                for cell in row.Cells:
                    print(cell)
                # represented as a cell value string
                # as expected

if __name__ == '__main__':
    main()

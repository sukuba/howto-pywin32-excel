#!python3

import pythoncom
import win32com.client
import json
import codecs

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
    r"""
    start Excel.
    open an existing book.
    read cells for each worksheet.
    write json files with the 1st row as dictionary key.
    terminate Excel.
    
    Result:
    (('年度', '分類', 'タイトル', 'URL'), ('平成27年', 'ダイズ', 'ダイズの低温特性',
     'soybean/low-temp'), ('平成28年', 'ダイズ', 'ダイズの花粉', 'soybean/pollen'),
    ('平成27年', 'ナタネ', 'ナタネの食害', 'oilseed/predators'), ('平成28年', 'ナタ
    ネ', 'ナタネの埋土種子', 'oilseed/seedbank'))
    5
    ('年度', '分類', 'タイトル', 'URL')
    4
    E:\scratch\hello-excel-9-Hello2.json
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheets = book.Worksheets
            for sheet in sheets:
                #nrow = sheet.Range("A1").CurrentRegion.Rows.Count
                data = sheet.Range("A1").CurrentRegion.Value
                #data = sheet.Range("A1").CurrentRegion.FormulaR1C1
                # convert bulk cell values at once.
                # after that, everything is python object.
                # this is good when only values are needed,
                # because the com automation costs will be minimized.
                # Formula family can also be used.
                
                nrow = len(data)
                print(data)
                print(nrow)
                
                keys = data[0]
                ncol = len(keys)
                print(keys)
                print(ncol)
                data_dict = []
                
                for r in range(1, nrow):
                    row = {}
                    for c in range(ncol):
                        row[keys[c]] = data[r][c]
                    data_dict.append(row)
                
                file_name = r'E:\scratch\hello-excel-9-%s.json' % sheet.Name
                print(file_name)
                with codecs.open(file_name, 'w', 'utf-8') as outfile:
                    json.dump(data_dict, outfile, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()

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
    read bulk cells for each worksheet.
    terminate Excel.
    
    Result:
    E:\scratch\hello-excel12-v-Hello2.json
    E:\scratch\hello-excel12-t-Hello2.json
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello4.xls') as book:
            sheets = book.Worksheets
            for sheet in sheets:
                region = sheet.Range("A1").CurrentRegion.Address
                data_v = sheet.Range(region).Value
                
                nrow = sheet.Range(region).Rows.Count
                ncol = sheet.Range(region).Columns.Count
                
                data_t = []
                for r in range(nrow):
                    row = []
                    for c in range(ncol):
                        row.append(sheet.Cells(r+1, c+1).Text)
                    data_t.append(row)
                
                print(data_v)
                print(data_t)
                
                file_name_v = r'E:\scratch\hello-excel12-v-%s.json' % sheet.Name
                file_name_t = r'E:\scratch\hello-excel12-t-%s.json' % sheet.Name
                
                print(file_name_v)
                print(file_name_t)
                
                with codecs.open(file_name_v, 'w', 'utf-8') as outfile:
                    json.dump(data_v, outfile, ensure_ascii=False, indent=4)
                with codecs.open(file_name_t, 'w', 'utf-8') as outfile:
                    json.dump(data_t, outfile, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()

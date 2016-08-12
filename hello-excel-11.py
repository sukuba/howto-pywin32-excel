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
    E:\scratch\hello-excel11-v-Hello2.json
    E:\scratch\hello-excel11-t-Hello2.json
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello4.xls') as book:
            sheets = book.Worksheets
            for sheet in sheets:
                data_v = sheet.Range("A1").CurrentRegion.Value
                data_t = sheet.Range("A1").CurrentRegion.Text
                # this returns a None.
                print(sheet.Range("A2").Text)
                # each cell can return "Text"
                print(sheet.Range("A2:D2").Text)
                # this returns a None.
                print(sheet.Range("A2").Value) # 27.0
                print(str(sheet.Range("A2").Value)) # 27.0
                
                file_name_v = r'E:\scratch\hello-excel11-v-%s.json' % sheet.Name
                file_name_t = r'E:\scratch\hello-excel11-t-%s.json' % sheet.Name
                
                print(file_name_v)
                print(file_name_t)
                
                with codecs.open(file_name_v, 'w', 'utf-8') as outfile:
                    json.dump(data_v, outfile, ensure_ascii=False, indent=4)
                with codecs.open(file_name_t, 'w', 'utf-8') as outfile:
                    json.dump(data_t, outfile, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()

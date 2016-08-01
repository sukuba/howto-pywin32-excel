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
    E:\scratch\ascii-Hello2.json
    E:\scratch\sjis-Hello2.json
    E:\scratch\utf8-Hello2.json
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheets = book.Worksheets
            for sheet in sheets:
                data = sheet.Range("A1").CurrentRegion.Value
                
                file_name = r'E:\scratch\ascii-%s.json' % sheet.Name
                print(file_name)
                with open(file_name, 'w') as outfile:
                    json.dump(data, outfile, indent=4)
                    # output result is escaped as utf-8 to be ascii encoded.
                
                file_name = r'E:\scratch\sjis-%s.json' % sheet.Name
                print(file_name)
                with open(file_name, 'w') as outfile:
                    json.dump(data, outfile, ensure_ascii=False, indent=4)
                    # output result is shift_jis encoded.
                
                file_name = r'E:\scratch\utf8-%s.json' % sheet.Name
                print(file_name)
                with codecs.open(file_name, 'w', 'utf-8') as outfile:
                    json.dump(data, outfile, ensure_ascii=False, indent=4)
                    # output result is utf-8 encoded.

if __name__ == '__main__':
    main()

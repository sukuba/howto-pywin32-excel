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
    5
    4
    ('年度', '分類', 'タイトル', 'URL')
    E:\scratch\hello-excel-10-Hello2.json
    """
    with excel_app() as app:
        with excel_book(app, r'E:\scratch\Hello2.xls') as book:
            sheets = book.Worksheets
            for sheet in sheets:
                nrow = sheet.Range("A1").CurrentRegion.Rows.Count
                ncol = sheet.Range("A1").CurrentRegion.Columns.Count
                # walk through each row.
                # at every rown, excel sheet object must be called, 
                # because the Rows object cannnot be kept in python.
                # com costs grow large by cell number.
                # this is needed when you are interested in 
                # other cell attributes than Value and Formula.
                # color, font, or hyperlinks.
                
                print(nrow)
                print(ncol)
                
                #print(sheet.Rows(1)) # returns the 1st row values
                #print(sheet.Rows[1]) # returns the 2nd row values
                # index base differs between excel and python.
                # both return a tuple of a row tuple.
                # don't know which is efficient,
                # but I have a bad feeling to call sheet.Cells[1234,5678] or something.
                
                keys = sheet.Rows(1).Value[0][:ncol]
                print(keys)
                
                data_dict = []
                
                for r in range(1, nrow):
                    row = {}
                    for c in range(ncol):
                        row[keys[c]] = sheet.Cells(r+1, c+1).Value
                    data_dict.append(row)
                
                file_name = r'E:\scratch\hello-excel-10-%s.json' % sheet.Name
                print(file_name)
                with codecs.open(file_name, 'w', 'utf-8') as outfile:
                    json.dump(data_dict, outfile, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    main()

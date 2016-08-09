#!python3

import argparse
import os.path
import codecs
import re
import json
import win32com.client
import pythoncom

class ExcelApp(object):
    """
    Excel.Application as a context
    """
    def __init__(self, visible=True, coinitialize=False):
        self.visible = visible
        self.coinitialize = coinitialize
        self.app = None
        
    def __enter__(self):
        if self.coinitialize:
            pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx('Excel.Application')
        self.app.Application.Visible = self.visible
        self.app.Application.AskToUpdateLinks = False
        return self.app
        
    def __exit__(self, type, value, traceback):
        self.app.Quit()
        del self.app
        if self.coinitialize:
            pythoncom.CoUninitialize()

class ExcelBook(object):
    """
    Excel.Application.Workbook as a context
    """
    def __init__(self, app, filename):
        self.app = app
        self.filename = filename
        self.book = None
        
    def __enter__(self):
        print('Opening %s ...' % self.filename)
        self.book = self.app.Workbooks.Open(self.filename)
        return self.book
        
    def __exit__(self, type, value, traceback):
        self.book.Saved = True
        self.book.Close()
        del self.book
        print('%s closed.' % self.filename)

class AddressHelper:
    """
    manipulates cell adress string.
    """
    p1 = re.compile(r'^([$\w]+)(?::([$\w]+))?$')
    p2 = re.compile(r'\$')
    p3 = re.compile(r'(\$?[a-zA-Z]+)(\$?[0-9]+)')
    p4 = re.compile(r'\$?([a-zA-Z]+)\$?([0-9]+)')
    p5 = re.compile(r'(?<!\$)((?<![a-zA-Z])[a-zA-Z]+|(?<![0-9])[0-9]+)')
    
    @classmethod
    def join(cls, address1, address2):
        return ':'.join((address1, address2))
    
    @classmethod
    def split(cls, address):
        return cls.p1.match(address).groups()
    
    @classmethod
    def begin_at(cls, address):
        return cls.split(address)[0]
    
    @classmethod
    def end_at(cls, address):
        range = cls.split(address)
        return range[1] if range[1] else range[0]
    
    @classmethod
    def no_absolute(cls, address):
        return re.sub(cls.p2, '', address)
    
    @classmethod
    def col_row(cls, address, absolute=None):
        if absolute == None:
            cr = cls.p3.match(address).groups()
        elif absolute == False:
            cr = cls.p4.match(address).groups()
        else:
            cr = cls.p3.match(re.sub(cls.p5, r'$\1', address)).groups()
        return cr
    
    @classmethod
    def col(cls, address):
        return cls.col_row(address, False)[0]
    
    @classmethod
    def row(cls, address):
        return int(cls.col_row(address, False)[1])

def make_params(args, sheet):
    """
    returns params defined by both of args and the 1st sheet.
    """
    params = {}
    (col_begin, col_end, row_begin, has_header, headers) = get_header(args, sheet)
    params['col_begin'] = col_begin
    params['col_end'] = col_end
    params['row_begin'] = row_begin
    params['has_header'] = has_header
    params['headers'] = headers if not args.url else headers + ('URL',)
    
    return params

def get_header(args, sheet):
    """
    determines the column range and gets header content.
    (columnBegin, columnEnd, rowBegin, hasHeader, Headers)
    """
    # continuous region from the origin
    end = end_of_current_region(sheet, args.origin)
    begin = args.origin
    
    # filter the columns
    if args.columns:
        col_begin = AddressHelper.begin_at(args.columns)
        col_end = AddressHelper.end_at(args.columns)
    else:
        col_begin = AddressHelper.col(begin)
        col_end = AddressHelper.col(end)
    
    # work around headers
    has_header = not args.noheader
    row_begin = AddressHelper.row(begin)
    if has_header:
        header_address = '%s%d:%s%d' % (col_begin, row_begin, col_end, row_begin)
        headers = sheet.Range(header_address).Value[0]
        row_begin += 1
    else:
        headers = None
    
    return (col_begin, col_end, row_begin, has_header, headers)

def get_region(params, args, sheet):
    """
    get table addresses to convert with processing command arguments.
    (whole table, url table)
    """
    # continuous region from the origin
    end = end_of_current_region(sheet, args.origin)
    
    # data rows
    row_end = AddressHelper.row(end)
    row_begin = params['row_begin']
    if row_begin > row_end:
        # no data rows
        return (None, None)
    
    # data columns
    col_begin = params['col_begin']
    col_end = params['col_end']
    
    whole_table = '%s%d:%s%d' % (col_begin, row_begin, col_end, row_end)
    
    # url
    if args.url:
        url_table = '%s%d:%s%d' % (args.url, row_begin, args.url, row_end)
    else:
        url_table = None
    
    return (whole_table, url_table)
    

def end_of_current_region(sheet, origin):
    """
    get an address of the last cell of current region.
    """
    current_region = sheet.Range(origin).CurrentRegion.Address
    return AddressHelper.end_at(current_region)

def append_url(sheet, whole_table, url_table):
    """
    append url of hyperlinks found at url_table 
    after the last column of each row of whole_table.
    """
    whole = sheet.Range(whole_table).Value
    data = []
    n = len(whole)
    
    for i in range(n):
        hyperlinks = sheet.Range(url_table).Rows(i+1).Hyperlinks
        url = hyperlinks(1).Address if hyperlinks.Count > 0 else None
        data.append(whole[i] + (url,))
    
    return data

def make_json_files(params, args, sheets):
    """
    converts Worksheets into json files.
    returns index of files generated.
    """
    index = {}
    counter = 0
    
    for sheet in sheets:
        if sheet.Range(args.origin).Value == None:
            print('Skipping blank worksheet at %s' % sheet.Name)
            continue
            
        whole_table, url_table = get_region(params, args, sheet)
        if not whole_table:
            print('Skipping empty line worksheet at %s' % sheet.Name)
            continue
        
        if url_table:
            data = append_url(sheet, whole_table, url_table)
        else:
            data = sheet.Range(whole_table).Value
            
        counter += 1
        file_name = os.path.join(args.dest, r'sheet%d.json' % counter)
        sheet_name = sheet.name
        index[sheet_name] = file_name
        print('%d: %s: %s' % (counter, file_name, sheet_name))
            
        with codecs.open(file_name, 'w', 'utf-8') as outfile:
            json.dump(data, outfile, ensure_ascii=False, indent=4)
            
    if params['has_header']:
        column_file_name = os.path.join(args.dest, 'columns.json')
        with codecs.open(column_file_name, 'w', 'utf-8') as outfile:
            json.dump(params['headers'], outfile, ensure_ascii=False, indent=4)
    
    index_file_name = os.path.join(args.dest, 'index.json')
    with codecs.open(index_file_name, 'w', 'utf-8') as outfile:
        json.dump(index, outfile, ensure_ascii=False, indent=4)
    
    return index

def excel_to_json(args):
    """
    converts Excel book into json files per worksheet.
    expects args; filename, dest, columns, url, noheader, verbose
    """
    with ExcelApp(visible=not args.invisible) as app:
        with ExcelBook(app, args.filename) as book:
            sheets = book.Worksheets
            params = make_params(args, sheets[0])
            index = make_json_files(params, args, sheets)
            
            if args.verbose:
                print(params)
                print(index)

def main():
    r"""
    エクセル表をjsonファイルに変換する。
    
    excel-to-json-2.py E:\scratch\Hello2.xls E:\scratch
    excel-to-json-2.py Hello2.xls E:\scratch --noheader
    excel-to-json-2.py Hello2.xls E:\scratch --columns B:C
    excel-to-json-2.py Hello2.xls E:\scratch --origin B3 --columns B:C
    excel-to-json-2.py Hello3.xls E:\scratch --url C
    
    第1引数: 変換元エクセルブック（フルパスまたは出力先パス）
    第2引数: 出力先ディレクトリ（フルパス）
    --origin: 基準セル（表範囲の左上アドレス、省略時はA1）
    --columns: 有効カラム（絶対列アドレス）
    --url: Hyperlink情報を持つ列（絶対列アドレス）
    --noheader: 1行目からデータとして扱う
    --verbose: 冗長な情報を出力する
    
    生成するjsonは、ArrayのArray。行優先マトリックス。
    A1がブランクでないシートを変換対象とする。
    シート単位でjsonファイルを作る。
    すべてのシートが同じ形式だと想定している。
    origin (A1)からの連続領域を有効データの範囲とする。
    columns指定があれば、それを優先し、なければ、先頭シートの形式を優先する。
    各シートを、sheet0.json等の連番ファイルに変換する。
    ファイルとシート名の関連情報を、index.jsonに書き出す。
    ヘッダー情報があれば、columns.jsonに書き出す。
    url指定した列にHyperlink情報があれば、url情報としてカラム末尾に追加する。
    
    エクセル本体が必要（インストール済みであること）。
    
    このプログラムはエクセルを起動し、
    指定されたブックを開き、
    変換したjsonを保存し、
    エクセルを閉じる。
    
    """
    parser = argparse.ArgumentParser(description='エクセル表をjsonファイルに変換する')
    parser.add_argument('filename', help='エクセルファイル名(読み取り)')
    parser.add_argument('dest', help='出力先ディレクトリ')
    parser.add_argument('-o', '--origin', default="A1", help='基準セル')
    parser.add_argument('-c', '--columns', help='有効カラム（列アドレス）')
    parser.add_argument('-u', '--url', help='Hyperlink情報のあるカラム（列アドレス）')
    parser.add_argument('-n', '--noheader', action='store_true', help='先頭行からデータとする')
    parser.add_argument('-i', '--invisible', action='store_true', help='エクセルを非表示にする')
    parser.add_argument('-v', '--verbose', action='store_true', help='冗長な情報を出力する')
    args = parser.parse_args()
    
    args.filename = os.path.join(args.dest, args.filename)
    
    if args.verbose:
        print(args)
    excel_to_json(args)

if __name__ == '__main__':
    main()

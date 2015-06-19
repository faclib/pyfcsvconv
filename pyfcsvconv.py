# -*- coding: UTF-8 -*-
# Copyright (C) 2015 Dmitriy Tyurin

__author__ = 'Dmitriy Tyurin <fobia3d@gmail.com>'
__license__ = "MIT"
__version__ = '1.5'

import chardet
import codecs
import cStringIO
import csv
import os
import os.path
import re
import subprocess
import sys
import tempfile
import types
import xlrd
import xlwt
import HTMLParser
from time import *

# Кодировка по умолчанию
# reload(sys)
# sys.setdefaultencoding('utf-8')
# ------------------------------


def unicode_filename(filename, encoding='utf-8'):
    """
    Преобразует имя файла в правельный формат
    """

    filename = filename.decode(encoding)
    return filename
# -----------------------------------------------


def utf8_encode(text, encoding='utf-8'):
    u"""
    Конвертирует строку в кодировку utf-8
    """

    enc = chardet.detect(text).get("encoding")
    if enc and enc.lower() != encoding:
        try:
            text = text.decode(enc)
            text = text.encode(encoding)
        except:
            pass
    return text
# -----------------------------------------------


def file_utf8_encode(filename, output = None):
    filename = unicode_filename(filename)
    f = open(filename, 'rb')
    text = f.read()
    f.close()

    text = utf8_encode(text)
    if type(output) == file:
        output.write(text)
    elif type(output) == str or type(output) == unicode:
        output = unicode_filename(output)
        f = open(output, 'wb')
        f.write(text)
        f.close()
# -----------------------------------------------


htmlCodes = (
    (" ", "&nbsp;"),
    ("'", '&#39;'),
    ('"', '&quot;'),
    ('>', '&gt;'),
    ('<', '&lt;')
    # ,('&', '&amp;')
)
def html_decode(s):
    """
    Returns the ASCII decoded version of the given HTML string. This does
    NOT remove normal HTML tags like <p>.
    """

    for code in htmlCodes:
        s = s.replace(code[1], code[0])
    s = s.replace('&amp;', '&')
    return s

def html_entities(s):
    """
    Преобразует все возможные символы в соответствующие HTML-сущности
    """

    for code in htmlCodes:
        s = s.replace(code[0], code[1])
    return s
# -----------------------------------------------


def detect_dialect(f):
    u"""
    Определить формат разделителей CSV
    Input:
        f - <type 'file'> дискриптор открытого файла
    Output:
        dialect
    """

    try:
        f.seek(0)
        dialect = csv.Sniffer().sniff(f.read(), delimiters=';,|\t')
    except BaseException, e:
        f.seek(0)
        dialect=csv.excel
        dialect.lineterminator = '\n'
        row = f.read(1024)
        if len(re.compile(';').findall(row)) > len(re.compile(',').findall(row)):
            dialect.delimiter = ';'
    f.seek(0)
    return dialect
# -----------------------------------------------


def get_type_sheet(filename):
    u"""
    Тип таблици (CSV, XLS, XLSX)

    $ file --brief --mime test/data/file.csv
        text/plain; charset=utf-8
        text/plain; charset=iso-8859-1
        text/plain; charset=us-ascii
    """

    filename = unicode_filename(filename)
    p = re.compile('html|plain|csv|xml|office|document|msword|excel|zip', re.IGNORECASE)
    c = re.compile('utf-8|us-ascii', re.IGNORECASE)

    res = {
        'type': '',
        'utf8': True
    }
    try:
        proc = subprocess.Popen("/usr/bin/file --mime '%s'" % (filename), shell=True, stdout=subprocess.PIPE)
        out = proc.stdout.readlines()
        S = " ".join(out).split(":")[1]
        S = S.lower()

        m = p.search(S).group()
    except:
        return res


    if m == 'plain' or m == 'csv':
        m = 'csv'
    elif m == 'xml' or m == 'msword' or m == 'office' or m == 'document':
        m = 'xls'
    elif m == 'zip' or m == 'excel':
        m = 'xlsx'
    elif not m:
        m = ''

    if m == 'csv' or m == 'html':
        if not c.search(S):
            res['utf8'] = False

    res['type'] = m.lower()
    return res
# -----------------------------------------------


def convert_xlsx(input_filename, output, delimiter=','):
    import xlsx2csv
    kwargs = {
        'delimiter' : delimiter,
        'sheetdelimiter' : '--------',
        'dateformat' : '%%Y-%%m-%%d',
        'skip_empty_lines' : False,
        'escape_strings' : False,
        'hyperlinks' : False,
        'cmd' : False,
        'include_sheet_pattern' : "^.*$",
        'exclude_sheet_pattern' : "",
        'merge_cells' : False
    }
    x = xlsx2csv.Xlsx2csv(input_filename, **kwargs)
    # return x
    # for s in self.workbook.sheets:
    # print x.workbook.sheets
    x.convert(output)
# -----------------------------------------------


class UTF8Recoder:
    u"""
    Итератор, который читает кодированный поток и перекодирует вход для UTF-8
    """

    def __init__(self, f, encoding="utf-8"):
        self.reader = codecs.getreader(encoding)(f)

    def __iter__(self):
        return self

    def read(self):
        return self.reader.read().encode("utf-8")

    def next(self):
        return self.reader.next().encode("utf-8")
# -----------------------------------------------


class FitSheetWrapper(object):
    """Try to fit columns to max size of any entry.
    To use, wrap this around a worksheet returned from the
    workbook's add_sheet method, like follows:

        sheet = FitSheetWrapper(book.add_sheet(sheet_name))

    The worksheet interface remains the same: this is a drop-in wrapper
    for auto-sizing columns.
    """
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()

    def write(self, r, c, label='', *args, **kwargs):
        self.sheet.write(r, c, label, *args, **kwargs)
        slen = str(label).__len__()
        if r > 0:
            slen = slen + 3
        else:
            slen = (slen + 2) * 1.25

        if slen > 100:
            slen = 100
        width = int(slen * 200)

        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = width

    def get_sheet(self):
        return self.sheet

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)
# -----------------------------------------------


class CSVUnicodeReader:
    """
    CSV Reader, который будет перебирать строки в CSV файл "f",
    который кодируется в данной кодировке.
    """

    def __init__(self, f, encoding="utf-8", **kwds):
        dialect = detect_dialect(f)
        rec = UTF8Recoder(f, encoding)
        self.reader = csv.reader(rec, dialect=dialect, **kwds)

    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]

    def __iter__(self):
        return self
# -----------------------------------------------


class XLSReader:
    def __init__(self, filename, **kwds):
        filename = unicode_filename(filename)
        book = xlrd.open_workbook(filename)
        self.sh = book.sheet_by_index(0)
        self.reader = self.get_reader()

    def get_reader(self):
        for rx in xrange(self.sh.nrows):
            row = []
            for cell in self.sh.row(rx):
                data = str(cell.value)
                row.append(data)
            yield row

    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]

    def __iter__(self):
        return self
# -----------------------------------------------


class HTMLReader:
    u"""
    HTML Reader
    """

    def __init__(self, f, encoding="utf-8", **kwds):
        text = f.read().encode("utf-8")
        text = re.compile('<(\w+)([^">]*(="[^"]*")?)*>', re.DOTALL).sub(r'<\1>', text)
        text = re.compile('\s+', re.DOTALL).sub(' ', text)
        text = re.compile('\r|( *<(?!/?(td|tr|th)>).+?> *)', re.DOTALL).sub('', text)

        text = re.compile('<th>(.*?)</th>', re.DOTALL).sub(r'<td>\1</td>', text)

        ptr = re.compile('<tr>.*?</tr>', re.DOTALL)
        ptd = re.compile('<td>(.*?)</td>', re.DOTALL)

        rows = list()
        mtr = ptr.findall(text)
        for tr in mtr:
            row = ptd.findall(tr)
            if len(row):
                rows.append(row)

        self.reader = enumerate(rows)

    def parser(self, text):
        text = html_decode(unicode(text, "utf-8")).strip(' ')
        return text

    def next(self):
        idx, row = self.reader.next()
        return [self.parser(s) for s in row]

    def __iter__(self):
        return self
# -----------------------------------------------


class CSVUnicodeWriter:
    u"""
    CSV Writer, который напишет строки в CSV файл "f",
    который закодирован в данной кодировке.
    """

    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        self.queue = cStringIO.StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.encoder = codecs.getincrementalencoder(encoding)()

    def writerow(self, row):
        self.writer.writerow([s.encode("utf-8") for s in row])
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        data = self.encoder.encode(data)
        self.stream.write(data)
        self.queue.truncate(0)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)

    def write_reader(self, reader):
        for rowi, row in enumerate(reader):
            self.writerow(row)

    def get_file(self, seek=False):
        F = self.stream
        if type(seek) == int:
            F.seek(seek)
        return F
# -----------------------------------------------


class XLSWriter:
    def __init__(self, sheetname=None, **kwds):
        if not sheetname:
            sheetname = "Sheet1"
        sheetname = utf8_encode(sheetname)

        self.book = xlwt.Workbook(encoding='utf-8')
        self.sheet = FitSheetWrapper(self.book.add_sheet(sheetname))

        # добавить новый цвет в палитре и установить RGB
        xlwt.add_palette_colour("custom_colour", 0x21)
        self.book.set_colour_RGB(0x21, 244, 236, 197) # FFF4ECC5 # border FFCCC085

        # Создаем новые стили
        # self.general_style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')

        # По умолчанию
        self.general_style = xlwt.XFStyle()
        self.general_style.num_format_str = 'general'
        # self.general_style.alignment.wrap = 1

        alignment = xlwt.Alignment() # Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_LEFT
        alignment.vert = xlwt.Alignment.VERT_TOP
        alignment.wrap = 1
        self.general_style.alignment = alignment

        # Стиль шапки
        self.head_style = self.general_style


    def set_head(self, hex='#F4ECC5'):
        _color = {
            'yellow': 'F4ECC5',
            'red':    'FFC7CE',
            'blue':   'C5D9F1'
        }
        if hex in _color.keys():
            value = _color[hex]
        else:
            value = hex.lstrip('#')

        lv = len(value)
        rgb = tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))
        self.book.set_colour_RGB(0x21, rgb[0], rgb[1], rgb[2])

        # Обнулим стиль
        self.head_style = self._get_style(True)

    def write_reader(self, reader):
        for rowi, row in enumerate(reader):
            if rowi == 0:
                self._firs_row(row)
                continue
            for coli, value in enumerate(row):
                value = value.decode('utf-8')
                self.sheet.write(rowi, coli, value, self.general_style)
                # print(value)

    def _firs_row(self, row):
        for coli, value in enumerate(row):
            value = value.decode('utf-8')
            self.sheet.write(0, coli, value, self._get_style())

    def _get_style(self, reset = False):
        if self.head_style and not reset:
            return self.head_style

        # Шрифт первой строчки
        fnt = xlwt.Font()
        fnt.name = 'Arial'
        fnt.bold = True

        borders = xlwt.Borders()
        borders.right = 0x1
        borders.right_colour = 0x13

        style = xlwt.easyxf('pattern: pattern solid, fore_colour custom_colour')
        # style = xlwt.XFStyle()
        style.font = fnt
        style.borders = borders

        self.head_style = style
        return self.head_style

    def save(self, filename):
        try:
            filename = unicode_filename(filename)
            self.book.save(filename)
        except:
            raise RuntimeError("Не удалось сохранить файл.")

    def frozen(self):
        sheet = self.sheet.get_sheet()
        sheet.panes_frozen = True
        sheet.horz_split_pos = 1
# -----------------------------------------------


class ConvertCSV:
    def __init__(self, filename):
        self.filename = unicode_filename(filename)
        self.tp = get_type_sheet(filename)

    def convert(self, output, delimiter=','):
        t = self.tp['type']

        if t == 'xlsx':
            try:
                convert_xlsx(self.filename, output, delimiter=delimiter)
                return
            except:
                t = 'xls'

        if t == 'xls' or not t:
            try:
                reader = XLSReader(self.filename)
            except:
                raise ValueError("Не верный формат файла таблицы")
        else:
            F=open(self.filename, 'rb')
            if not self.tp['utf8']:
                W = tempfile.TemporaryFile()
                W.write(utf8_encode(F.read()))
                F.close()
                F = W
                F.seek(0)
            if t == 'csv':
                reader = CSVUnicodeReader(F)
            elif t == 'html':
                reader = HTMLReader(F)

        if type(output) == str or type(output) == unicode:
            output = open(unicode_filename(output), 'wb')
        elif type(output) == file:
            pass
        else:
            raise ValueError("Не верный аргумент файла записи")

        writer = self._get_writer(output, delimiter=delimiter)
        writer.write_reader(reader)

    def _get_writer(self, output, delimiter=','):
        return CSVUnicodeWriter(output, delimiter=delimiter, encoding='utf-8', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
# -----------------------------------------------



if __name__ == "__main__":
    try:
        import argparse
        import sys

        reload(sys)
        sys.setdefaultencoding('utf-8')

        parser = argparse.ArgumentParser(description='конвертация таблиц')
        parser.add_argument('-i', action='store_true', help='редактирование файлов на месте (создает копию, если указано расширение)')
        parser.add_argument('-o', '--output', dest='output', default=None, help="сохранить в файл")
        parser.add_argument("-d", "--delimiter", dest="delimiter", default=",", help="delimiter - columns delimiter in csv (default: ',')")
        parser.add_argument('file',  help='входной файл таблицы') #nargs='+',

        args = parser.parse_args()

        if not args.output:
            args.output = sys.stdout
        else:
            args.output = open(args.output, 'wb')
        if args.i == True:
            args.i = tempfile.TemporaryFile()

        convertCsv=ConvertCSV(args.file)
        convertCsv.convert(args.output, args.delimiter)

        if args.i == True:
            import shutil
            del convertCsv

            args.output.seek(0)
            with open(args.file, 'wb') as fdest:
                shutil.copyfileobj(args.output, fdest)

        # params = {
        #     'i': args.i,
        #     'output': args.output, # None,
        #     'delimiter': args.delimiter, #  ',',
        #     'file': args.file # '../TableSheets/test/data/fake.xlsx'
        # }
        # params = {
        #     'i': False,
        #     'output': None,
        #     'delimiter': ',',
        #     'file': '../TableSheets/test/data/fake.xlsx'
        # }
        #
        # if not params['output']:
        #     params['output'] = sys.stdout
        # else:
        #     params['output'] = open(params['output'], 'w')
        #
        # if params['i'] == True:
        #     params['output'] = tempfile.TemporaryFile()
        #
        # convertCsv = ConvertCSV(params['file'])
        # convertCsv.convert(params['output'], params['delimiter'])
        #
        # if params['i'] == True:
        #     import shutil
        #     del convertCsv
        #
        #     params['output'].seek(0)
        #     with open(params['file'], 'wb') as fdest:
        #         shutil.copyfileobj(params['output'], fdest)
    except BaseException as e:
        print "Error parse: {0}".format(e)
        sys.exit(1)

    # prime_numbers = get_prime_numbers(count)
    # print("Answer: %d" % prime_numbers[-1])
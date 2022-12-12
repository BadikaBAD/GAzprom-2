# !/usr/local/bin/python3.8
# -*- coding: utf8 -*-

import os
import sqlite3
from docx import Document
import re
import csv


class doc_analyzer(object):
    def __init__(self, _dir):
        self.files = []
        for _dir in os.walk(_dir):
            #            _record = [_file[0], _file[1], []]
            self.files.append([_dir[0], _dir[1], []])
            for _filename in _dir[2]:
                _name, _ext = os.path.splitext(_filename)
                if _ext.lower() == '.pdf':
                    self.files[-1][2].append(_filename)

        self.current_dir = ''
        self.current_file = ''
        self.data_fullname = ''
        self.id = ''
        self.current_page = -1
        self.current_line = -1

    def next_dir(self):
        while (len(self.files) > 0) and (len(self.files[-1][2]) == 0):
            del self.files[-1]

        if len(self.files) > 0:
            self.current_dir = self.files[-1][0]
            return self.files[-1]
        else:
            self.current_dir = ''
            return ['', [], []]

    def next_path(self):
        _dir = self.next_dir()
        if len(self.current_dir) > 0:
            self.current_file = _dir[2][-1]
            del _dir[2][-1]
            self.data_fullname = os.path.join(self.current_dir, self.current_file)
            _match = re.search(r'(?:СТО|Р) Газпром (?:РД ){0,1}[\.\d]{1,7}(?:-[\.\d]{1,4}){0,1}(?:-\d{1,4}){0,1}-\d{4}',
                               self.current_file)
            self.id = _match[0] if _match else ''
        else:
            self.current_file = ''
            self.data_fullname = ''
            self.id = ''
        return self.current_file

    def set_document(self, _document):
        self.document = _document
        print(self.current_file)
        self.pages = []
        for page in self.document:
            self.pages.append(page.lines)
        self.current_page = 0
        self.current_line = 0

    def next_line(self):
        self.current_line = self.current_line + 1
        while ((len(self.pages) - 1) > self.current_page) and (self.current_line == len(self.pages[self.current_page])):
            self.current_page = self.current_page + 1
            self.current_line = 0
        try:
            return self.pages[self.current_page][self.current_line].strip()
        except Exception:
            print(f"len(pages): {len(self.pages)}")
            print(f"current_page: {self.current_page}")
            print(f"len(lines): {len(self.pages[self.current_page])}")
            print(f"current_line: {self.current_line}")

    def get_toc(self):
        _toc = []
        s = self.pages[self.current_page][self.current_line].strip()
        while (self.current_page < (len(self.pages) - 1)) and (s.upper() != 'СОДЕРЖАНИЕ'):
            s = self.next_line()

        while (self.current_page < (len(self.pages) - 1)) and (s.strip().upper() != 'ВВЕДЕНИЕ'):
            s = self.next_line().replace("'", '')

            if (len(_toc) > 0) and (len(s) > 2) and (s[0] == '.') and (s[1] == '.'):
                _toc[-1] = _toc[-1] + s
            #            elif s.count(id)==0:
            else:
                _toc.append(s)

        _term_chapters = []
        for _line in _toc:
            _line = re.sub(r'(\. )\1+\.', r'|', _line)
            _line = re.sub(r'(\.)\1+', r'|', _line)
            if _line.count('|') > 0:
                chapter, page = _line.split('|')
                if (len(_term_chapters) > 0) and (len(_term_chapters[-1]) < 4):
                    _term_chapters[-1].append(chapter.strip())
                    _term_chapters[-1].append(page.strip())

                if chapter.upper().find('ТЕРМИН') > 0:
                    _term_chapters.append([chapter.strip(), page.strip()])

        return _term_chapters

    def get_chapter(self, _toc):
        _chapter = []
        try:
            s = self.pages[self.current_page][self.current_line].strip()
        except Exception:
            print(f"len(pages): {len(self.pages)}")
            print(f"current_page: {self.current_page}")
            print(f"len(lines): {len(self.pages[self.current_page])}")
            print(f"current_line: {self.current_line}")
        for _item in _toc:
            while (self.current_page < (len(self.pages) - 1)) and (s.upper() != _item[0].upper()):
                s = self.next_line()

            _newline = True
            while (self.current_page < (len(self.pages) - 1)) and (s.strip().upper() != _item[2].upper()):
                if _newline:
                    _chapter.append(s)
                else:
                    _chapter[-1] = _chapter[-1] + s
                #                try:
                _newline = (len(s) == 0) or (s[-1] == '.') or (s[-1] == ':') or (s.upper() == _item[0].upper())
                #                except Exception:
                #                    print(_newline)
                s = self.next_line()

        return _chapter

    def get_terms(self, _chapter):
        _colon = 0
        _dash = 0
        _terms = []
        for _line in _chapter:
            _ = _line.split(':', 1)
            if (len(_) > 1) and (len(_[0]) < len(_[1])): _colon = _colon + 1
            _ = _line.split('–', 1)
            if (len(_) > 1) and (len(_[0]) < len(_[1])): _dash = _dash + 1

        _sep = ':' if (_colon > _dash) else '–'
        for _line in _chapter:
            _ = _line.split(_sep, 1)
            if (len(_) > 1) and (len(_[0]) < len(_[1])):
                _0 = re.sub(r'\d\.[\d]{1,3}', '', _[0]).strip()
                if len(_0) > 0: _0 = _0[0].upper() + _0[1:]
                _1 = _[1].strip()
                if len(_1) > 0: _1 = _1[0].upper() + _1[1:]
                _terms.append([_0, _1])
                db.cursor.execute(
                    f"INSERT OR IGNORE INTO terms (filename, file_id, term, term_def) VALUES ('{analyzer.data_fullname}', '{analyzer.id}', '{_0}', '{_1}')")
            elif len(_terms) > 0:
                _terms[-1] = f"{_terms[-1]}\n{_line}"

        return _terms


class db_writer(object):
    def __init__(self):
        self.dir_path = os.path.dirname(os.path.abspath(__file__))
        self.db_name = 'regulatory.sqlite'
        self.db_fullname = f'{self.dir_path}/{self.db_name}'

        if os.path.exists(self.db_fullname): os.remove(self.db_fullname)
        os.system(f'sqlite3 {self.db_fullname} < {self.dir_path}/DB/DDL/01#create_tables#sqlite.sql')

        self.db = sqlite3.connect(self.db_fullname)
        self.cursor = self.db.cursor()


db = db_writer()
analyzer = doc_analyzer(_dir='/home/a.shalamov@adm.ggr.gazprom.ru/SOURCES/Фонд/2008/')

sto_files = []
while (analyzer.next_path() != ''):

    try:
        doc = Document(analyzer.data_fullname)
    except Exception:
        db.cursor.execute(f"INSERT OR IGNORE INTO not_process (filename, kind) VALUES ('{analyzer.data_fullname}',1)")
    else:
        analyzer.set_document(doc)

    term_chapters = analyzer.get_toc()
    if len(term_chapters) == 0:
        db.cursor.execute(f"INSERT OR IGNORE INTO not_process (filename, kind) VALUES ('{analyzer.data_fullname}', 2)")
    else:
        chapter = analyzer.get_chapter(term_chapters)
        if len(chapter) == 0:
            db.cursor.execute(
                f"INSERT OR IGNORE INTO not_process (filename, kind) VALUES ('{analyzer.data_fullname}', 3)")
        else:
            terms = analyzer.get_terms(chapter)

    #    print(analyzer.get_terms(analyzer.get_chapter(analyzer.get_toc())))
    if analyzer.id == '':
        sto_files.append([analyzer.data_fullname, '', analyzer.id, '', '', ''])

csvfile = 'source_file.csv'
with open(csvfile, "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerows(sto_files)

db_writer.execute(f"COMMIT")

con = sqlite3.connect("C:\Users\Никита\OneDrive\Рабочий стол\sqlite-tools-win32-x86-3400000\sqlite-tools-win32-x86-3400000\Chinook")  # change to 'sqlite:///your_filename.db'
cur = con.cursor()
cur.execute("CREATE TABLE document (file_id	, filepath, filename, document_id, terms_start_page, terms_end_page );")

with open('source_file.csv', 'r') as fin:
    dr = csv.DictReader(fin)
    to_db = [(i['col1'], i['col2']) for i in dr]

cur.executemany(
    "INSERT INTO document (file_id	, filepath, filename, document_id, terms_start_page, terms_end_page ) Values (?, ?);", to_db)
con.commit()
con.close()

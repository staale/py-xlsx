# -*- coding: utf-8 -*-
__author__="Ståle Undheim <staale@staale.org>"

import re
import zipfile

from xml.dom import minidom

class DomZip(object):
    def __init__(self, filename):
        self.filename = filename

    def __getitem__(self, key):
        # @type ziphandle ZipFile
        ziphandle = zipfile.ZipFile(self.filename)
        dom = minidom.parseString(ziphandle.read(key))
        ziphandle.close()
        return dom

class Workbook(object):

    def __init__(self, filename):
        self.__sheetsByIndex = []
        self.__sheetsByName = {}
        self.filename = filename
        self.domzip = DomZip(filename)
        self.sharedStrings = SharedStrings(self.domzip["xl/sharedStrings.xml"])
        workbookDoc = self.domzip["xl/workbook.xml"]
        sheets = workbookDoc.firstChild.getElementsByTagName("sheets")[0]
        for sheetNode in sheets.childNodes:
            name = sheetNode._attrs["name"].value
            id = int(sheetNode._attrs["r:id"].value[3:])

            sheet = Sheet(self, id, name)
            self.__sheetsByIndex.append(sheet)
            self.__sheetsByName[name] = sheet
            assert sheet.name in self.__sheetsByName

    def keys(self):
        return self.__sheetsByName.keys()

    def __len__(self):
        return len(self.__sheetsByName)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self.__sheetsByIndex[key]
        else:
            return self.__sheetsByName[key]

class SharedStrings(list):
    def __init__(self, sharedStringsDom):
        nodes = sharedStringsDom.firstChild.childNodes
        for text in [n.firstChild.firstChild for n in nodes]:
            self.append(text.nodeValue if text and text.nodeValue else self.__getIfInline(text))
            
    def __getIfInline(self, text):
        if text.hasChildNodes():
            nodes = text.parentNode.parentNode.childNodes
            return "".join([node.getElementsByTagName("t")[0].firstChild.nodeValue for node in nodes])
        else:
            return ""

class Sheet(object):

    def __init__(self, workbook, id, name):
        self.workbook = workbook
        self.id = id
        self.name = name
        self.loaded = False
        self.addrPattern = re.compile("([a-zA-Z]*)(\d*)")
        self.__cells = {}
        self.__cols = {}
        self.__rows = {}

    def __load(self):
        sheetDoc = self.workbook.domzip["xl/worksheets/sheet%d.xml"%self.id]
        sheetData = sheetDoc.firstChild.getElementsByTagName("sheetData")[0]
        # @type sheetData Element
        rows = {}
        columns = {}
        for rowNode in sheetData.childNodes:
            rowNum = rowNode.getAttribute("r")
            for columnNode in rowNode.childNodes:
                colType = columnNode.getAttribute("t")
                cellId = columnNode.getAttribute("r")
                colNum = cellId[:len(cellId)-len(rowNum)]
                formula = None
                if colType == "s":
                    stringIndex = columnNode.firstChild.firstChild.nodeValue
                    data = self.workbook.sharedStrings[int(stringIndex)]
                elif columnNode.firstChild:
                    data = getattr(columnNode.getElementsByTagName("v")[0].firstChild, "nodeValue", None)
                else:
                    data = ""
                if columnNode.getElementsByTagName("f"):
                    formula = getattr(columnNode.getElementsByTagName("f")[0].firstChild, "nodeValue", None)
                if not rowNum in rows:
                    rows[rowNum] = []
                if not colNum in columns:
                    columns[colNum] = []
                cell = Cell(rowNum, colNum, data,formula=formula)
                rows[rowNum].append(cell)
                columns[colNum].append(cell)
                self.__cells[cellId] = cell
        for rowNum in rows.keys():
            self.__rows[rowNum] = sorted(rows[rowNum])
        self.__cols = columns
        self.loaded=True

    def rows(self):
        if not self.loaded:
            self.__load()
        return self.__rows
    
    def cols(self):
        if not self.loaded:
            self.load()
        return self.__cols

    def __getitem__(self, key):
        if not self.loaded:
            self.__load()
        (column, row) = self.addrPattern.match(key).groups()
        if column and row:
            if not key in self.__cells:
                return None
            return self.__cells[key]
        if column:
            return self.__cols[key]
        if row:
            return self.__rows[key]

    def __iter__(self):
        if not self.loaded:
            self.__load()
        return self.__cells.__iter__()

class Cell(object):
    def __init__(self, row, column, value, formula=None):
        self.row = int(row)
        self.column = column
        self.value = value
        self.formula = formula
        self.id = "%s%s"%(column, row)

    def __cmp__(self, other):
        if other.column == self.column:
            return self.row - other.row
        else:
            if self.column < other.column:
                return -1
            elif self.column > other.column:
                return 1
            else:
                return 0

    def __str__(self):
        return "<Cell [%s] : \"%s\" (%s)>"%(self.id, self.value, self.formula)

# -*- coding: utf-8 -*-
""" Small footprint xlsx reader """
__author__="Ståle Undheim <staale@staale.org>"

import re
import zipfile
from xldate import xldate_as_tuple

try:
    from xml.etree import cElementTree as ET
except:
    import cElementTree as ET

class DomZip(object):
    """ Excel xlsx files are zip files containing xml documents.
    This class handles parsing those xml documents into dom objects

    """

    def __init__(self, filename):
        """ Open up the xlsx document.
        Arguments::

            filename -- can be a filepath or a file-like object

        """

        self.ziphandle = zipfile.ZipFile(filename, 'r')

    def __getitem__(self, key):
        """ Get a domtree from a document in the zip file
        Arguments::

            key -- path inside the zip file (xml document)

        """

        return ET.fromstring(self.ziphandle.read(key))

    def __del__(self):
        """Close the zip file when finished"""

        self.ziphandle.close()

class Workbook(object):
    """Main class that contains sheets organized by name or by id.
    Id being the order number of the sheet starting from 1

    """
    def __init__(self, filename):
        self.__sheetsById = {}
        self.__sheetsByName = {}
        self.filename = filename
        self.domzip = DomZip(filename)
        try : # Not all xlsx documents contain Shared Strings
            self.sharedStrings = SharedStrings(
                self.domzip["xl/sharedStrings.xml"])
        except KeyError :
            self.sharedStrings = None

        # Extract the last modification date; based upon an answer at:
        #  http://superuser.com/questions/195548/excel-2007-modify-creation-date-statistics
        self.dcterms_modified = None
        modified_date_elements = self.domzip["docProps/core.xml"].findtext("{http://purl.org/dc/terms/}modified")
        if modified_date_elements:
            self.dcterms_modified = modified_date_elements

        workbookDoc = self.domzip["xl/workbook.xml"]
        sheets = workbookDoc.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        id = 1
        for sheetNode in sheets:
            name = sheetNode.get("name")
            sheet = Sheet(self, id, name)
            self.__sheetsById[id] = sheet
            self.__sheetsByName[name] = sheet
            assert sheet.name in self.__sheetsByName
            id += 1

    def keys(self):
        return self.__sheetsByName.keys()

    def close(self):
        self.domzip.__del__()

    def __len__(self):
        return len(self.__sheetsByName)

    def __iter__(self):
        for sheet in self.__sheetsByName.values():
            yield sheet

    def __getitem__(self, key):
        if isinstance(key, int):
            return self.__sheetsById[key]
        else:
            return self.__sheetsByName[key]

class SharedStrings(list):

    def __init__(self, sharedStringsDom):
        nodes = [x for x in sharedStringsDom]
        for text in [n[0] for n in nodes]:
            self.append(text.text)

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
        sheetDoc = self.workbook.domzip["xl/worksheets/sheet%d.xml" % self.id]
        sheetData = sheetDoc.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData")
        # @type sheetData Element
        rows = {}
        columns = {}
        for rowNode in sheetData:
            rowNum = int(rowNode.get("r"))
            for columnNode in rowNode:
                colType = columnNode.get("t")
                cellId = columnNode.get("r")
                cellS = columnNode.get("s")
                colNum = cellId[:len(cellId)-len(str(rowNum))]
                formula = None
                data = ''
                try:
                    if colType == "s":
                        stringIndex = columnNode[0].text
                        data = self.workbook.sharedStrings[int(stringIndex)]
                    #Date field
                    elif cellS in ('1', '2', '3', '4') and colType == "n":
                        data = xldate_as_tuple(
                            int(columnNode[0].text),
                            datemode=0)
                    elif len(columnNode)>0 and columnNode[0] is not None:
                        data = columnNode.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v").text

                    if columnNode.find("f"):
                        formula = columnNode.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f").text
                except Exception:
                    raise #pass
                if not rowNum in rows:
                    rows[rowNum] = []
                if not colNum in columns:
                    columns[colNum] = []
                cell = Cell(rowNum, colNum, data, formula=formula)
                rows[rowNum].append(cell)
                columns[colNum].append(cell)
                self.__cells[cellId] = cell
        self.__rows = rows
        self.__cols = columns
        self.loaded=True

    def rows(self):
        if not self.loaded:
            self.__load()
        return self.__rows

    def cols(self):
        if not self.loaded:
            self.__load()
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

    def __lt__(self, other):
        return self.__cmp__(other) == -1

    def __gt__(self, other):
        return self.__cmp__(other) == 1

    def __eq__(self, other):
        return self.__cmp__(other) == 0

    def __ne__(self, other):
        return self.__cmp__(other) != 0

    def __le__(self, other):
        return self.__cmp__(other) != 1

    def __ge__(self, other):
        return self.__cmp__(other) != -1

    def __unicode__(self):
        return u"<Cell [%s] : \"%s\" (%s)>" % (self.id, self.value,
                                               self.formula, )

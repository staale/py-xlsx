from pyxlsx import Workbook
import os
import os.path

import unittest

TEST_PATH = os.path.realpath(os.path.join(__file__, '..','..','..','..','..','..','testdata'))
TEST_FILE = os.path.join(TEST_PATH, os.listdir(TEST_PATH)[0])

class  WorkbookTestCase(unittest.TestCase):
    def setUp(self):
        self.filename = TEST_FILE
    
    #def tearDown(self):
    #    self.foo.dispose()
    #    self.foo = None

    def testFilename(self):
        assert os.path.exists(self.filename), "Could not find file %s"%self.filename
        assert os.path.isfile(self.filename), "%s is not a file"%self.filename
        assert os.path.splitext(self.filename)[1] == ".xlsx", "Wrong extension for %s"%self.filename

    def testConstruction(self):
        Workbook(self.filename)

    def testSheets(self):
        workbook = Workbook(self.filename)
        assert len(workbook) > 0, "No worksheets found in %s"%self.filename

        # assert each index and name
        for index in range(len(workbook)):
            assert workbook[index] != None, "Missing worksheet at index %d"%index
            sheet = workbook[index]
            assert sheet.id == workbook[sheet.name].id, "No name reference for sheet %s at %d"%(sheet.name, index)
            assert sheet.name in workbook.keys()

    def testSheetData(self):
        workbook = Workbook(self.filename)
        sheet = workbook[0]
        assert sheet["A2"].value != None, "Missing A2 cell from worksheet"
        assert type(sheet["A"]) == list, "A column is not a list"
        assert type(sheet["1"]) == list, "1 row is not a list"

    def testSpecificData(self):
        workbook = Workbook(self.filename)
        sheet = workbook[0]
        assert sheet["A2"].value == "Level A", "Column [%s] does not match expected [%s], is instead [%s]"%("A2", "Level A", sheet["A1"])

if __name__ == '__main__':
    unittest.main()


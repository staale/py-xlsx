python-xlsx
===========

A small footprint xslx reader that understands shared strings and can process
excel dates.

Usage
+++++++

::

    book = Workbook('filename or filedescriptor') #Open xlsx file
    for sheet in book:
        print sheet.name
        for row, cells in sheet.rows().iteritems(): # or sheet.cols()
            print row # prints row number
            for cell in cells:
                print cell.id, cell.value, cell.formula

    # or you can access the sheets by their name:

    some_sheet = book['some sheet name']
    ...

Alternatives
------------

To my knowledge there are other python alternatives:

 * https://bitbucket.org/ericgazoni/openpyxl/
 * https://github.com/leegao/pyXLSX

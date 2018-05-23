#!/usr/bin/env python
"""
Your task is as follows:
- read the provided Excel file
- find and return the min and max values for the COAST region
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

# To run this code you'll need xlrd
# pip install xlrd

import os
import errno
import xlrd
from zipfile import ZipFile

datafile = "2013-ercot-hourly-load-data"


def open_zip(datafile):
    remove_file('{0}.xls'.format(datafile))  # removes xls file if already exists
    print 'Extracting zip:   ' + '{0}.zip'.format(datafile)
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:  # opens file to extract it in local directory
        myzip.extractall()


def remove_file(filename):
    try:
        os.remove(filename)
    except OSError as e:  # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
            raise  # re-raise exception if a different error occurred


def parse_file(datafile):
    wbfilename = '{0}.xls'.format(datafile)
    print 'Opening workbook: ' + wbfilename
    workbook = xlrd.open_workbook(wbfilename)
    sheet = workbook.sheet_by_index(0)

    # we just want the region COAST, that is column[1] in the excel sheet
    sheet_data = [[sheet.cell_value(r, 1)] for r in range(1, sheet.nrows)]  # 1 to skip header

    print(sheet_data)
    ### other useful methods:
    # print "\nROWS, COLUMNS, and CELLS:"
    # print "Number of rows in the sheet:", 
    # print sheet.nrows
    # print "Type of data in cell (row 3, col 2):", 
    # print sheet.cell_type(3, 2)
    # print "Value in cell (row 3, col 2):", 
    # print sheet.cell_value(3, 2)
    # print "Get a slice of values in column 3, from rows 1-3:"
    # print sheet.col_values(3, start_rowx=1, end_rowx=4)

    # print "\nDATES:"
    # print "Type of data in cell (row 1, col 0):", 
    # print sheet.cell_type(1, 0)
    # exceltime = sheet.cell_value(1, 0)
    # print "Time in Excel format:",
    # print exceltime
    # print "Convert time to a Python datetime tuple, from the Excel float:",
    # print xlrd.xldate_as_tuple(exceltime, 0)

    data = {
        'maxtime': (0, 0, 0, 0, 0, 0),
        'maxvalue': 0,
        'mintime': (0, 0, 0, 0, 0, 0),
        'minvalue': 0,
        'avgcoast': 0
    }
    return data


def test():
    open_zip(datafile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)

    remove_file(datafile)


test()

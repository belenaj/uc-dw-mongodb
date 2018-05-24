#!/usr/bin/env python
"""
Your task is as follows:
- read the provided Excel file
- find and return the min, max and average values for the COAST region
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

    print('sheet_data:    ' + str(type(sheet_data)))        # sheet_data has 'list' Type
    print('sheet data[0]: ' + str(type(sheet_data[0])))     # ! each element of sheet_data is also a 'list' of size 1 !

    # print(sheet_data)  # uncomment to display data

    # to calculate the average, let's create a list of floats
    float_ls = []
    for x in sheet_data:
        float_ls.append(x[0])  # all the lists within the list have only one element (index = 0)

    # getting the average value
    avgcoast = sum(float_ls) / float(len(float_ls))
    print 'Average value for COAST: ' + '%.6f' % avgcoast

    # getting min and max values (as float)
    minvalue = min(sheet_data)[0]
    maxvalue = max(sheet_data)[0]
    print 'Min. value: ' + '%.6f' % minvalue
    print 'Max. value: ' + '%.6f' % maxvalue

    # getting index of min and max values (+1 because we skipped the header)
    minvalue_idx = sheet_data.index([minvalue]) + 1
    maxvalue_idx = sheet_data.index([maxvalue]) + 1

    # extract time at index ( ! as float)
    mintime_cell = sheet.cell_value(minvalue_idx, 0)
    maxtime_cell = sheet.cell_value(maxvalue_idx, 0)

    print('mintime_cell: ' + str(type(mintime_cell)))  # ! it returns the time as float !

    # print "Convert time to a Python datetime tuple,
    mintime = xlrd.xldate_as_tuple(mintime_cell, 0)
    maxtime = xlrd.xldate_as_tuple(maxtime_cell, 0)

    print mintime
    print maxtime

    data = {
            'maxtime': maxtime,
            'maxvalue': maxvalue,
            'mintime': mintime,
            'minvalue': minvalue,
            'avgcoast': avgcoast
    }
    return data


def test():
    open_zip(datafile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)

    remove_file('{0}.xls'.format(datafile))


test()

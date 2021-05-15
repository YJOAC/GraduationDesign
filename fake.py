import xlrd
import xlwt
import os
import sys
import random
import string
from formExtraction_ln import *

def getBoundry(sheet, nrows, col):
    _min = 10000
    _max = -1
    for i in range(1, nrows):
        value = sheet.cell_value(i, col)
        if value:
            _min = _min if float(value) > _min else float(value)
            _max = _max if float(value) < _max else float(value)
    return _min * 0.9, _max * 1.1

def createFakeData(filePath, excel_rd, excel_wt, sheetType, num=300):
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], excel_rd)

    sheet_rd = excel_rd.sheet_by_index(sheetType.value) # 获取对应操作页
    sheet_wt = excel_wt.get_sheet(sheetType.value)

    for i in range(1, ncols):
        _min, _max = getBoundry(sheet_rd, nrows, i)
        _range = _max - _min
        for n in range(num):
            value = random.random()*_range + _min
            sheet_wt.write(nrows + n, i)
    excel_wt.save(filePath)
    
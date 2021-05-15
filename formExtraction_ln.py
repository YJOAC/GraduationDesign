import sys
import os
import pickle
import re
import codecs
import string
import shutil
import enum
import xlwt
import xlrd
import time
from xlutils.copy import copy
from datetime import datetime
import docx
from docx import Document
from win32com import client as wc

# 确保sheets[TableType.value]=sheetName, eg.sheets[1]='参建单位'
# 确保sheetsHeads[TableType.value]对应该Sheet的表头 eg.sheetsHeads[1]为参建单位sheet表头
sheets = ['参建单位', '施工任务', '工程完成情况', '施工进度目标','施工质量控制', '工程质量验收', '合同管理及投资控制', '安全文明施工', '问题和措施建议'] # Excel中的sheet列表名称
sheetsHeads = [
    [['监理月报日期', '标段名称', '承包人', '备注', '标段简称']],
    [['监理月报日期', '标段', '施工任务']],
    [['监理月报日期', '标段', '施工项目', '单位', '设计量', '累计完成', '累计完成比例']],
    [['监理月报日期', '标段', '本月进度分析', '原因', '下月计划']],
    [['监理月报日期', '质量管理体系文件', '施工文件', '设计图纸', '设计通知单', '工程技术联系单及备忘录']],
    [['监理月报日期', '标段', '本期验收单元数量', '总计', '合格数', '优良数', '合格率', '优良率']],
    [['监理月报日期', '标段', '本期结算', '累计结算', '合同结算率', '变更累计', '状态及评价']],
    [['监理月报日期', '安全文函', '监理工作', '安全周例会', '违规整改处罚', '危险源管理情况', '安全及文明施工监理工作']],
    [['监理月报日期', '序号', '问题', '措施及建议']]
]

# 原表格中数据项之前的无用行
sheetsHeadRows = [0, 0, 1, 0, 0, 0, 0, 0, 0]
# 原表格中数据项之前的无用列
sheetsHeadCols = [0, 0, 0, 0, 0, 0, 0, 0, 0]
# 总表相比原表在数据项之前增加的列数
newLines       = [0, 0, 1, 0, 0, 0, 0, 0, 0]

# 表格的类型，不同类型对应最终Excel中不同的sheet
class TableType(enum.Enum):
    default = -1
    contractor = 0 # 参建单位
    constructionTask = 1 # 施工任务
    constructionComplete = 2 # 工程完成情况
    constructionSchedule = 3 # 施工进度目标
    constructionQuality = 4 # 施工质量控制
    qualityAcception = 5 # 工程质量验收
    contractManagement = 6 # 合同管理
    secureConstruction = 7 # 安全文明施工
    problems = 8 # 问题及措施建议

# 获取fileName所有文件中的所有表格
def getTableData(fileName, pathDir, excelFile, date):
    document = Document(pathDir + fileName) # 读入文件
    tables = document.tables  # 获取文件中的表格集
    for table in tables:
        # printTableData(table)
        processTableData(table, pathDir, excelFile, date)

# 通过对表头的分析，判断输入的表属于哪个项目，并做出对应处理
def processTableData(table, pathDir, excelFile, date):
    head = table.rows[0] # 表头
    if head.cells[:]:
        cells = head.cells[:] # 表头的每个单元格（合并单元格重复显示）
        head_content = []
        for cell in cells:
            head_content.append(cell.text)
        # print(head_content)
        numOfColumns = len(cells) # 表的列数
        determinTableType(table, head_content, numOfColumns, pathDir, excelFile, date)
        # print('--------------------------')

# 判定传入head对应的表的类型
def determinTableType(table, head, num, pathDir, excelFile, date, flag=True):
    table_type = TableType.default
    if num == 6:
        if(_match(head[0], "施工") and _match(head[1], "施工") and _match(head[2], "单位") and _match(head[4], "完成")):
            # 工程完成情况标
            table_type = TableType.constructionComplete
            data = xlrd.open_workbook(pathDir+excelFile)
            nrows, ncols = getSheetRowsCols(sheets[table_type.value], data)
            writeSheetContent(table, table_type, nrows, ncols, pathDir, excelFile, date)
    elif num == 7:
        if(_match(head[1], "施工") and _match(head[2], "施工") and _match(head[3], "单位") and _match(head[5], "完成")):
            # 工程完成情况标
            table_type = TableType.constructionComplete
            data = xlrd.open_workbook(pathDir+excelFile)
            nrows, ncols = getSheetRowsCols(sheets[table_type.value], data)
            writeSheetContent(table, table_type, nrows, ncols, pathDir, excelFile, date, 1)
    elif num == 8:
        if(_match(head[1], "施工") and _match(head[2], "施工") and _match(head[3], "单位") and _match(head[6], "完成")):
            # 工程完成情况标
            table_type = TableType.constructionComplete
            data = xlrd.open_workbook(pathDir+excelFile)
            nrows, ncols = getSheetRowsCols(sheets[table_type.value], data)
            writeSheetContent(table, table_type, nrows, ncols, pathDir, excelFile, date, 1)
    elif num == 4:
        if _match(head[0], "序号") and _match(head[1], "文件名称") and _match(head[2], "文件编号"):
            # 安全文函
            table_type = TableType.secureConstruction
            
            data = xlrd.open_workbook(pathDir+excelFile)
            nrows, ncols = getSheetRowsCols(sheets[table_type.value], data)
            
            excel = copy(data) # 完成xlrd对象向xlwt的转换
            sheet = excel.get_sheet(table_type.value) # 获得需要的操作页
            # if flag:
            #     sheet.write(nrows, 1, len(table.rows)-1)
            # else:
            sheet.write(nrows, 4, len(table.rows)-1)
            excel.save(pathDir+excelFile)


# 获取将要写入的sheet的现有行数列数
def getSheetRowsCols(sheetName, data):
    sheet = data.sheet_by_name(sheetName)
    return sheet.nrows, sheet.ncols

# 写入sheet的具体内容
def writeSheetContent(table, sheetType, nrows, ncols, pathDir, excelFile, date, _r=0):
    data = xlrd.open_workbook(pathDir+excelFile, formatting_info=True)
    excel = copy(data) # 完成xlrd对象向xlwt的转换
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    headerRows = getTableHeaderRows(table, sheetType.value) # 获取数据项前的无用行数
    headerCols = getTableHeaderCols(table, sheetType.value) # 获取数据项前的无用列数

    # 写入
    for i in range(headerRows, len(table.rows)):
        sheet.write(nrows, 0, date)

        row = table.rows[i]
        cells = row.cells[:]
        for k in range(headerCols+_r, len(cells)):
            n = k+newLines[sheetType.value]
            sheet.write(nrows, n, cells[k].text)
        nrows += 1
    
    # mergeAllCells(sheet, table, sheetType.value, nrows) # 合并单元格
    excel.save(pathDir+excelFile)

# 获取该表格的表头的行数，即返回值n表示从table的第n行起为数据项
def getTableHeaderRows(table, n):
    return sheetsHeadRows[n]

# 获取该表格的表头的列数，即返回值n表示从table的第n列起为数据项
def getTableHeaderCols(table, n):
    return sheetsHeadCols[n]

# 判断两个字符串是否相同，前者是从表单元格读取的文本可能含有\n，\r，空格等
def _match(s1, s2):
    if(s2 in s1.replace("\n", "").replace("\r", "").replace(" ", "").replace("（","(").replace("）",")")):
        return True
    else: return False

# 将getTableData中所有表格结果输出
def printTableData(table):
    for i, row in enumerate(table.rows[:]): # 读每行
        # print(row)
        row_content = []
        for cell in row.cells[:]: # 读一行中的所有单元格
            c = cell.text
            row_content.append(c)
        print (row_content) # 以列表形式输出每一行数据

# 将pathDir目录下.doc形式的文件转化为.docx形式
def saveAsDocx(pathDir):
    word = wc.Dispatch('Word.Application')
    files = os.listdir(pathDir) # 路径下所有文件
    for _f in files:
        name = _f.split('.')
        # 文件是.doc形式且未被另存为.docx
        if(name[1] == 'doc' and not os.path.exists(pathDir+name[0]+'.docx')):
            doc = word.Documents.Open(pathDir+_f)
            doc.SaveAs(pathDir+name[0]+'.docx', 12, False, "", True, "", False, False, False, False)
            doc.Close()
    word.Quit()

# 将Excel初始化，新建sheets
# reCreate=True表示需要重建该Excel
def initExcel(pathDir, excelFile,reCreate=False):
    if(reCreate and os.path.exists(pathDir+excelFile)):
        os.remove(pathDir+excelFile)
    if(not os.path.exists(pathDir+excelFile)):
        workbook = xlwt.Workbook(encoding='utf-8')
        for i in range(len(sheets)):
            sheet = workbook.add_sheet(sheets[i], cell_overwrite_ok=True)
            writeSheetHead(sheet, i)
        workbook.save(pathDir+excelFile)

# 为不同的sheet添加表头
def writeSheetHead(sheet, sheetNum):
    style = xlwt.XFStyle() # 新建样式
    font = xlwt.Font() # 新建字体
    font.bold = True # 加粗
    style.font = font

    headers = sheetsHeads[sheetNum]
    for i in range(len(headers)):
        head = headers[i]
        for k in range(len(head)):
            sheet.write(i, k, head[k], style)
    mergeHeadCells(sheet, sheetNum, style)

# 整体合并单元格
def mergeAllCells(sheet, table, sheetNum, nrows):
    rows = table.rows
    m = len(rows)
    n = len(rows[0].cells)
    for i in range(sheetsHeadRows[sheetNum], m):
        for j in range(sheetsHeadCols[sheetNum], n):
            if len(rows[m-i-1].cells[n-j-1].text) == 0:
                continue
            elif m-i-1 > 0 and rows[m-i-1].cells[n-j-1].text == rows[m-i-2].cells[n-j-1].text:
                sheet.write_merge(m-i-2+nrows, m-i-1+nrows, n-j-1, n-j-1, rows[m-i-1].cells[n-j-1].text)
            elif n-j-1 > 0 and rows[m-i-1].cells[n-j-1].text == rows[m-i-1].cells[n-j-2].text:
                sheet.write_merge(m-i-1+nrows, m-i-1+nrows, n-j-2, n-j-1, rows[m-i-1].cells[n-j-1].text)

# 合并表头的单元格
def mergeHeadCells(sheet, sheetNum, style):
    headers = sheetsHeads[sheetNum]
    m = len(headers)
    n = len(headers[0])
    for i in range(m):
        for j in range(n):
            if m-i-1 > 0 and headers[m-i-1][n-j-1] == headers[m-i-2][n-j-1]:
                sheet.write_merge(m-i-2, m-i-1, n-j-1, n-j-1, headers[m-i-2][n-j-1], style)
            if n-j-1 > 0 and headers[m-i-1][n-j-1] == headers[m-i-1][n-j-2]:
                sheet.write_merge(m-i-1, m-i-1, n-j-2, n-j-1, headers[m-i-1][n-j-2], style)
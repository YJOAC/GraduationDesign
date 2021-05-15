from pyhanlp import *
import re
import docx
import time
from docx import Document
from datetime import datetime

from formExtraction_ln import *
from HpParser import *

pathDir = u'D:\study\works\BE\BE\doc\ln\\' # 文件路径
excelFile = 'result_ln.xls' # 结果保存在excelFile中

def getAllFile(pathDir):
    fileName = []
    files = os.listdir(pathDir) # 路径下所有文件
    for _f in files:
        name = _f.split('.')
        # 文件是.docx形式
        if name[1] == 'docx' and '~' not in name[0]:
            fileName.append(_f)
    return fileName

def getAllTextData(fileName):
    for f in fileName:
        document = Document(pathDir + f) # 读入文件
        paragraphs = document.paragraphs     # 获取所有段落
        pgsText = preProcessing([p.text for p in paragraphs]) # 获取文本
        # print(pgsText)
        date = getDate(pgsText).strftime("%Y-%m")
        getTableData(f, pathDir, excelFile, date) # 获取所有的表格数据
        
        data = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格
        excel = copy(data) # xlrd对象转换为xlwt对象（可写入）
        # printAllText(pgsText) # 输出所有文本及其所属段落
        findSpecificText(data, excel, pgsText, date) # 在文本中寻找特定的信息
        # break

def printAllText(text):
    for i in range(len(text)):
        print(i, text[i])

def preProcessing(text):
    for s in text:
        if s == "" or '\t' in s:
            text.remove(s) # 删除空的段落或包含表格的段落
    return text

def findSpecificText(data, excel, text, date):
    constructionSchedule(data, excel, text, date)
    constructionQuality(data, excel, text, date)
    qualityAcception(data, excel, text, date)
    contractManagement(data, excel, text, date)
    secureConstruction(data, excel, text, date)

def constructionSchedule(data, excel, text, date):
    sheetType = TableType.constructionSchedule
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期

    # 本月工程进度主要问题分析及处理情况
    for i in range(len(text)):
        if re.match(r'.*施工进度', text[i]):
            # 逐标段分析
            pos = i+1
            while(True):
                if len(text[pos]) <= 20 and re.match(r'.*[A-Z][0-9]标', text[pos]):
                    sheet.write(nrows, 0, date) # 写入监理月报日期
                    # 标段名称
                    name = re.findall(r'.*([A-Z][0-9]标)', text[pos])
                    if name: sheet.write(nrows, 1, name[0])
                    
                    while(True):
                        pos += 1
                        if re.match(r'.*?本月进度', text[pos]):
                            sheet.write(nrows, 2, text[pos+1])
                        elif re.match(r'.*?进度计划', text[pos]):
                            sheet.write(nrows, 4, text[pos+1])
                            break
                        if pos >= i + 40: break
                    nrows += 1
                        
                pos += 1
                if pos >= i + 40: break
            break

    excel.save(pathDir+excelFile)

def constructionQuality(data, excel, text, date):
    sheetType = TableType.constructionQuality
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期

    for i in range(len(text)):
        if re.match(r'.*质量管理体系文件', text[i]):
            if "无" in text[i+1]:
                sheet.write(nrows, 1, 0)
            else:
                _f = re.findall(r'(《.*?》)', text[i+1])
                sheet.write(nrows, 1, len(_f))
        elif re.match(r'.*施工措施.*?文件', text[i]):
            if "无" in text[i+1]:
                sheet.write(nrows, 2, 0)
            else:
                pos = i + 1
                num = 0
                while(pos < i+5):
                    if re.match(r'.*图纸', text[pos]): break
                    else:
                        _n = re.findall(r'\d', text[pos])
                        for number in _n:
                            num += int(number)
                        pos += 1
                sheet.write(nrows, 2, num)
        elif re.match(r'.*设计图纸', text[i]):
            if "无" in text[i+1]:
                sheet.write(nrows, 3, 0)
            else:
                pos = i + 1
                num = 0
                while(pos < i+5):
                    if re.match(r'.*通知', text[pos]): break
                    else:
                        _n = re.findall(r'\d', text[pos])
                        for number in _n:
                            num += int(number)
                        pos += 1
                sheet.write(nrows, 3, num)
        elif re.match(r'.*设计通知单', text[i]):
            if "无" in text[i+1]:
                sheet.write(nrows, 4, 0)
            else:
                pos = i + 1
                num = 0
                while(pos < i+5):
                    if re.match(r'.*备忘录', text[pos]): break
                    else:
                        _n = re.findall(r'\d', text[pos])
                        for number in _n:
                            num += int(number)
                        pos += 1
                sheet.write(nrows, 4, num)
        elif re.match(r'.*备忘录', text[i]):
            if "无" in text[i+1]:
                sheet.write(nrows, 5, 0)
            else:
                pos = i + 1
                num = 0
                while(pos < i+5):
                    if re.match(r'.*质量', text[pos]) or re.match(r'4.3', text[pos]): break
                    else:
                        _n = re.findall(r'\d', text[pos])
                        for number in _n:
                            num += int(number)
                        pos += 1
                sheet.write(nrows, 5, num)
            break
    excel.save(pathDir+excelFile)

def qualityAcception(data, excel, text, date):
    sheetType = TableType.qualityAcception
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期

    for i in range(len(text)):
        if re.match(r'.*工程质量验收评定情况', text[i]):
            # 逐标段分析
            pos = i+1
            while(pos < i + 5):
                if re.match(r'.*?总体评价', text[pos]): break
                else:
                    sheet.write(nrows, 0, date) # 写入监理月报日期
                    # 标段名称
                    name = re.findall(r'.*([A-Z][0-9]标)', text[pos])
                    if name: sheet.write(nrows, 1, name[0])

                    n1 = re.findall(r'验收.*?(\d+)[个]', text[pos])
                    if n1: sheet.write(nrows, 3, n1[0])
                    else: sheet.write(nrows, 3, 0)

                    n2 = re.findall(r'合格(\d+)[个]', text[pos])
                    if n2: sheet.write(nrows, 4, n2[-1])
                    else: sheet.write(nrows, 4, 0)

                    n3 = re.findall(r'优良(\d+)[个]', text[pos])
                    if n3: sheet.write(nrows, 5, n3[0])
                    else: sheet.write(nrows, 5, 0)

                    n4 = re.findall(r'合格率([\d\.]+\%)', text[pos])
                    if n4: sheet.write(nrows, 6, n4[-1])
                    else: sheet.write(nrows, 6, 0)

                    n5 = re.findall(r'优良率([\d\.]+\%)', text[pos])
                    if n5: sheet.write(nrows, 7, n5[-1])
                    else: sheet.write(nrows, 7, 0)

                    nrows += 1
                pos += 1
            break
    excel.save(pathDir+excelFile)

def contractManagement(data, excel, text, date):
    sheetType = TableType.contractManagement
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期

    for i in range(len(text)):
        if re.match(r'.*进度款审核', text[i]):
            # 逐标段分析
            pos = i+1
            while(pos < i + 10):
                if re.match(r'.*?投资控制', text[pos]): break
                else:
                    sheet.write(nrows, 0, date) # 写入监理月报日期
                    # 标段名称
                    name = re.findall(r'[\(\d\)]*?(.*?标)', text[pos])
                    if name: sheet.write(nrows, 1, name[0])

                    n1 = re.findall(r'本月.*?([\d\.]+[万]*?)元', text[pos])
                    if n1: sheet.write(nrows, 2, n1[0])
                    else: sheet.write(nrows, 2, 0)

                    n2 = re.findall(r'累计完成.*?([\d\.]+[万]*)元', text[pos])
                    if n2: sheet.write(nrows, 3, n2[-1])
                    else: sheet.write(nrows, 3, 0)

                    n3 = re.findall(r'占.*?([\d\.]+\%)', text[pos])
                    if n3: sheet.write(nrows, 4, n3[-1])
                    else: sheet.write(nrows, 4, 0)
                    nrows += 1
                pos += 1
            break
    excel.save(pathDir+excelFile)

def secureConstruction(data, excel, text, date):
    sheetType = TableType.secureConstruction
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期

    for i in range(len(text)):
        # 安全监理工作
        if re.match(r'.*安全监理工作', text[i]):
            pos = i+1
            num = 0
            while(pos < i + 20):
                if re.match(r'.*?安全违规', text[pos]): break
                else:
                    n = re.findall(r'（(\d+)）', text[pos])
                    if n: num = int(n[-1]) if int(n[-1]) > num else num
                pos += 1
            sheet.write(nrows, 2, num)
        
        # 安全违规整改处罚
        elif re.match(r'.*?危险源监管', text[i]):
            if re.match(r'失控', text[i+1]) or re.match(r'未受控', text[i+1]):
                sheet.write(nrows, 5, 0)
            else: sheet.write(nrows, 5, 1)

    # 安全周例会
    for i in range(len(text)):
        if re.match(r'.*?周例会', text[i]):
            _n = re.findall(r'安全.*?周例会(\d+)次', text[i])
            if _n: sheet.write(nrows, 3, _n[-1])
        
    excel.save(pathDir+excelFile)


def getDate(text):
    for i in range(20):
        if re.match(r'.*?(\d{4}年\d{0,2}月)', text[i]):
            _d = re.findall(r'.*?(\d{4}年\d{0,2}月)', text[i])
            if _d:
                print('report date:', _d[0])
                date = datetime.strptime(_d[0], '%Y年%m月').date()
                return date

if __name__ == '__main__':
    # saveAsDocx(pathDir) # 将doc形式的文件另存为docx
    fileName = getAllFile(pathDir)
    # fileName = ["[HZJL][LNGS]-(综合）-2019-015关于报送“2019年5月份河南洛宁抽水蓄能电站建设工程施工监理月报”的函.docx"]
    initExcel(pathDir, excelFile, True) # 初始化结果表格
    getAllTextData(fileName)
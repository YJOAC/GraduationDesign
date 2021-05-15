from pyhanlp import *
import re
import docx
import time
from docx import Document
from datetime import datetime

from formExtraction import *
from HpParser import *

pathDir = u'D:\study\works\BE\BE\doc\wd\\' # 文件路径
excelFile = 'result.xls' # 结果保存在excelFile中

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

def printAllText(text):
    for i in range(len(text)):
        print(i, text[i])

def preProcessing(text):
    for s in text:
        if s == "" or '\t' in s:
            text.remove(s) # 删除空的段落或包含表格的段落
    return text

def findSpecificText(data, excel, text, date):
    secureTrainingData(data, excel, text, date)
    scheduleCompleted(data, excel, text, date)
    scheduleDescription(data, excel, text, date)
    rawMaterialTest(data, excel, text, date)
    qualityProblemDescription(data, excel, text, date)
    standardProcess(data, excel, text, date)

def secureTrainingData(data, excel, text, date):
    sheetType = TableType.secureTrainingData
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    sheet.write(nrows, 0, date) # 写入监理月报日期
    sheet.write(nrows, 1, date) # 写入数据周期

    # 监理安全周例会（次数）
    for i in range(len(text)):
        if re.match(r'.*监理安全周例会[0-9]+', text[i]):
            n = re.findall(r'(\d)次', text[i])[0]
            sheet.write(nrows, 2, n)
            break
    # 安全教育培训情况
    for i in range(len(text)):
        if re.match(r'安全教育培训情况', text[i]):
            sheet.write(nrows, 3, '是')
            sheet.write(nrows, 4, text[i+1])
            break
        else:
            sheet.write(nrows, 3, '否')
    # 职业健康
    for i in range(len(text)):
        if re.match(r'职业健康采取的措施', text[i]):
            sheet.write(nrows, 6, '是')
            sheet.write(nrows, 7, text[i+1])
            break
    excel.save(pathDir+excelFile)

def scheduleCompleted(data, excel, text, date):
    sheetType = TableType.scheduleCompleted
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页

    # 本月工程进度主要问题分析及处理情况
    for i in range(len(text)):
        if re.match(r'.*本月工程进度[\u4e00-\u9fa5]*问题[\u4e00-\u9fa5]*', text[i]):
            # 逐标段分析
            pos = i+1
            while(True):
                if len(text[pos]) <= 20 and re.match(r'.*[A-Z][0-9]标', text[pos]):
                    sheet.write(nrows, 0, date) # 写入监理月报日期
                    sheet.write(nrows, 1, date) # 写入数据周期
                    # 标段名称
                    name = re.findall(r'(.+?标)', text[pos])
                    _name = re.findall(r'([A-Z]\d+标)', text[pos])
                    if name: sheet.write(nrows, 2, name[0])
                    if _name: sheet.write(nrows, 3, _name[0])
                    # 没有实际完成工作，跳过该标段
                    pos += 1
                    if not re.match(r'.*完成.*', text[pos]):
                        pos += 1
                        nrows += 1
                        continue
                    # 计划、实际、完成率
                    n1 = re.findall(r'计划.*?([\d\.]+)项', text[pos])
                    if n1: sheet.write(nrows, 4, n1[0])
                    else: sheet.write(nrows, 4, 0)

                    n2 = re.findall(r'完成.*?([\d\.]+)项', text[pos])
                    if n2: sheet.write(nrows, 5, n2[-1])
                    else: sheet.write(nrows, 5, 0)
                    
                    while(True):
                        # 原因
                        pos += 1
                        # print(text[pos])
                        if re.match(r'.*原因', text[pos]):
                            sheet.write(nrows, 7, text[pos])
                            break
                        elif pos >= i + 50:
                            break
                        else: continue
                    while(True):
                        # 处理情况
                        pos += 1
                        if re.match(r'.*处理情况', text[pos]):
                            sheet.write(nrows, 8, text[pos])
                            break
                        elif pos >= i + 50:
                            break
                        else: continue
                    nrows += 1
                pos += 1
                if pos >= i+50: break
            break

    excel.save(pathDir+excelFile)

def scheduleDescription(data, excel, text, date):
    sheetType = TableType.scheduleDescription
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页

    sheet.write(nrows, 0, date) # 写入监理月报日期
    sheet.write(nrows, 1, date) # 写入数据周期

    for i in range(len(text)):
        if re.match(r'.*施工周进度计划', text[i]):
            n1 = re.findall(r'施工周进度计划.*?(\d+)次', text[i])
            n2 = re.findall(r'专项检查.*?(\d+)次', text[i])
            n3 = re.findall(r'进度分析例会.*?(\d+)次', text[i])
            n4 = re.findall(r'生产协调例会.*?(\d+)次', text[i])
            n5 = re.findall(r'施工专题会.*?(\d+)次', text[i])
            # print(n1, n2, n3, n4, n5)
            sheet.write(nrows, 2, n1[0] if n1 else 0)
            sheet.write(nrows, 3, n2[0] if n2 else 0)
            sheet.write(nrows, 4, n3[0] if n3 else 0)
            sheet.write(nrows, 5, n4[0] if n4 else 0)
            sheet.write(nrows, 6, n5[0] if n5 else 0)

    excel.save(pathDir+excelFile)

def rawMaterialTest(data, excel, text, date):
    sheetType = TableType.rawMaterialTest
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页

    for i in range(len(text)):
        if re.match(r'.*原材料及试验检测', text[i]):
            # 逐标段分析
            pos = i+1
            while(True):
                if len(text[pos]) <= 20 and re.match(r'.*标', text[pos]):
                    sheet.write(nrows, 0, date) # 写入监理月报日期
                    sheet.write(nrows, 1, date) # 写入数据周期
                    # 标段名称
                    name = re.findall(r'(.+?标)', text[pos])
                    if name: sheet.write(nrows, 2, name[0])
                    # 没有实际完成工作，跳过该标段
                    pos += 1
                    if not re.match(r'.*取.*', text[pos]):
                        pos += 1
                        nrows += 1
                        continue
                    # 取样数量、验收批次
                    num = re.findall(r'取.*?(\d+)[组]', text[pos])
                    if num:
                        # print(num)
                        sheet.write(nrows, 3, num[0])
                        sheet.write(nrows, 5, '检测频次满足规范检测要求')
                    nrows += 1
                pos += 1
                if pos >= i+20: break
            break
    excel.save(pathDir+excelFile)

def qualityProblemDescription(data, excel, text, date):
    sheetType = TableType.qualityProblemDescription
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页

    sheet.write(nrows, 0, date) # 写入监理月报日期
    sheet.write(nrows, 1, date) # 写入数据周期

    for i in range(len(text)):
        if re.match(r'.*监理中心原材料抽检', text[i]):
            num1 = re.findall(r'.*共计(\d+)', text[i+1])
            if num1: sheet.write(nrows, 2, num1[0])
            num2 = re.findall(r'.*原材料抽检.*?(\d+)', text[i+2])
            if num2: sheet.write(nrows, 3, num2[0])
            num3 = re.findall(r'.*共计(\d+)', text[i+2])
            if num3 and num2: sheet.write(nrows, 4, int(num3[0])-int(num2[0]))
            sheet.write(nrows, 5, '是')
        elif re.match(r'.*施工期物探检测情况', text[i]):
            num1 = re.findall(r'自检.*?(\d+)次', text[i+1])
            num2 = re.findall(r'数量.*?(\d+)', text[i+1])
            if num1: sheet.write(nrows, 6, num1[0])
            if num2: sheet.write(nrows, 7, num2[0])
        elif re.match(r'.*质量生产周例会', text[i]):
            num1 = re.findall(r'质量生产周例会(\d+)', text[i])
            num2 = re.findall(r'整改通知书(\d+)', text[i])
            if num1: sheet.write(nrows, 14, num1[0])
            if num2: sheet.write(nrows, 15, num2[0])
    
    sheet.write(nrows, 11, '无')
    sheet.write(nrows, 12, '无')
    sheet.write(nrows, 13, '无')

    excel.save(pathDir+excelFile)

def standardProcess(data, excel, text, date):
    sheetType = TableType.standardProcess
    nrows, ncols = getSheetRowsCols(sheets[sheetType.value], data)
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页

    # 施工工艺、强制条文
    sheet.write(nrows, 0, date) # 写入监理月报日期
    sheet.write(nrows, 1, date) # 写入数据周期
    sheet.write(nrows, 2, '施工工艺')

    sheet.write(nrows+1, 0, date) # 写入监理月报日期
    sheet.write(nrows+1, 1, date) # 写入数据周期
    sheet.write(nrows+1, 2, '强制条文')

    for i in range(len(text)):
        if re.match(r'.*施工工艺示范手册执行情况', text[i]):
            num1 = re.findall(r'计划应用(\d+)项', text[i+1])
            if num1: sheet.write(nrows, 3, num1[0])
            num2 = re.findall(r'实际应用(\d+)项', text[i+1])
            if num2: sheet.write(nrows, 4, num2[0])
            num3 = re.findall(r'工序.*?(\d+)项', text[i+1])
            if num3: sheet.write(nrows, 5, num3[0])
            num4 = re.findall(r'良好.*?(\d+)项', text[i+1])
            if num4: sheet.write(nrows, 6, num4[0])
            num5 = re.findall(r'应用率.*?(\d+)', text[i+1])
            if num5:
                sheet.write(nrows, 7, num5[1]+'%')
                sheet.write(nrows, 8, num5[0]+'%')
        elif re.match(r'.*强制性条文执行计划落实', text[i]):
            num1 = re.findall(r'(\d+)条', text[i+1])
            if len(num1) >= 2:
                sheet.write(nrows+1, 3, num1[0])
                sheet.write(nrows+1, 4, num1[1])
            elif len(num1) >= 1:
                sheet.write(nrows+1, 3, num1[0])
            num2 = re.findall(r'执行率(\d+)', text[i+1])
            if num2: sheet.write(nrows+1, 8, num2[0]+'%')
    
    excel.save(pathDir+excelFile)

def getDate(text):
    for i in range(20):
        if re.match(r'\d{4}年\d{0,2}月', text[i]):
            print('report date:', text[i])
            date = datetime.strptime(text[i], '%Y年%m月').date()
            return date

if __name__ == '__main__':
    # saveAsDocx(pathDir) # 将doc形式的文件另存为docx
    fileName = getAllFile(pathDir)
    initExcel(pathDir, excelFile, True) # 初始化结果表格
    getAllTextData(fileName)
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
sheets = ['文件清单', '参见单位', '安全培训数据', '危险源管理情况统计表', '安全隐患排查治理一览表',
        '安全生产费用投入统计表', '违章隐患统计表', '各标段承包人人员投入情况', '标段投入设备情况',
        '本月主要完成工程量统计表', '本月工程形象进度面貌表', '本月进度计划完成情况', '进度描述', '本月锚杆监理检测结果一览表',
        '本月爆破振动检测成果表', '工序质量验评表', '原材料到货统计', '各标段检验批验收统计', '开挖质量控制表',
        '原材料试验监测情况表', '安全监测设备监测情况表', '安全检测仪器安装情况', '质量问题描述表', '合同台账',
        '本月各标段分包报审报备情况统计表', '分包计划一览表', '新增单价统计表', '合同变更会签单', '工程索赔表',
        '本月预付款支付与回扣统计表', '月进度支付统计表', '标准工艺'] # Excel中的sheet列表名称
sheetsHeads = [
    [['数据表', '数据情况', '记录数']],
    [['监理月报日期', '数据周期', '标段名称', '承包人', '备注', '标段简称']],
    [['监理月报日期', '数据周期', '监理安全周例会（次数）', '是否组织安全教育培训', '培训内容', '是否建立安全管理制度', '职业健康采取的措施是否得当', '职业健康采取的措施内容', '安全生产工作情况', '是否制定隐患排查治理执行手册和建立隐患排查治理机制']],
    [['监理月报日期', '数据周期', '危险源名称', '可能导致的危害', '危险源区域', '危险源类别', '风险级别', '责任单位', '控制措施', '控制状态', '关闭情况']],
    [['监理月报日期', '数据周期', '隐患简题', '评估等级', '专业分类', '归属单位/标段', '治理期限', '是否消除', '未消除的隐患当月整改进展情况']],
    [['监理月报日期', '数据周期', '标段', '费用（万元）', '累计（万元）', '备注']],
    [['监理月报日期', '数据周期', '施工单位', '本月检查次数', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容', '违章、隐患整改内容'],
        ['监理月报日期', '数据周期', '施工单位', '本月检查次数', '文明施工', '习惯性违章', '安全防护', '设备设施', '交通管理', '安全用电', '脚手架', '消防安全', '合计']],
    [['监理月报日期', '数据周期', '标段', '施工单位', '项目经理', '安全总监', '技术负责人', '项目部配备的管理人员（人）', '作业人员（人）', '其中特种作业人员（人）', '合计（人）', '备注']],
    [['监理月报日期', '数据周期', '标段', '设备/设施名称', '型号规格', '数量', '备注']],
    [['监理月报日期', '数据周期', '标段', '项目名称', '项目内容', '单位', '合同工程量', '月计划工程量', '本月实际完成', '本月完成占月计划（%）', '累计完成工程量', '累计完成占合同（%）']],
    [['监理月报日期', '数据周期', '标段', '工程部位', '主要设计特性描述', '形象进度描述']],
    [['监理月报日期', '数据周期', '标段', '标段别名', '计划', '实际', '完成率', '原因', '处理情况']],
    [['监理月报日期', '数据周期', '施工进度计划', '专项检查', '生产周进度分析例会', '周生产协调例会', '施工专题会']],
    [['监理月报日期', '数据周期', '标段', '工程部位', '受检数量', '检测数量', '检测结果', '检测结果', '检测结果', '检测结果', '总体评价', '标段'],
        ['监理月报日期', '数据周期', '标段', '工程部位', '受检数量', '检测数量', 'I级', 'II级', 'III级', 'IV级', '总体评价', '标段']],
    [['监理月报日期', '数据周期', '测点编号', '通道', '振动最大速度值V(cm/s)', '主频(Hz)', '量程(cm/s)', '灵敏度（V/m/s）', '允许振动速度(cm/s)', '是否合格', '布点类型', '爆心距R(m)', '监测点位置']],
    [['监理月报日期', '数据周期', '序号', '施工标段', '施工内容', '本月', '本月', '本月', '本月', '本年累计', '本年累计', '本年累计', '本年累计', '开工至今累计', '开工至今累计', '开工至今累计', '开工至今累计'],
        ['监理月报日期', '数据周期', '序号', '施工标段', '施工内容', '单元工程数量（个）', '合格', '优良', '优良率（%）', '单元工程数量（个）', '合格', '优良', '优良率（%）', '单元工程数量（个）', '合格', '优良', '优良率（%）']],
    [['监理月报日期', '数据周期', '标段', '材料名称', '单位', '重量', '生产厂家', '质量证明文件']],
    [['监理月报日期', '数据周期', '施工标段', '单位工程名称', '本月', '本月', '本月', '本年累计', '本年累计', '本年累计', '开工至今累计', '开工至今累计', '开工至今累计'],
        ['监理月报日期', '数据周期', '施工标段', '单位工程名称', '检验批数量（个）', '合格（个）', '合格率（%）', '检验批数量（个）', '合格（个）', '合格率（%）', '检验批数量（个）', '合格（个）', '合格率（%）']],
    [['监理月报日期', '数据周期', '标段', '序号', '工程部位', '设计开挖量（m³）', '测量断面数量（条）', '超欠挖情况（m3）', '超欠挖情况（m3）', '最大超欠挖尺寸（cm）', '最大超欠挖尺寸（cm）', '平均超欠挖尺寸（cm）', '平均超欠挖尺寸（cm）', '半孔率（%）', '平整度'],
        ['监理月报日期', '数据周期', '标段', '序号', '工程部位', '设计开挖量（m³）', '测量断面数量（条）', '超挖', '欠挖', '超挖', '欠挖', '超挖', '欠挖', '半孔率（%）', '平整度']],
    [['监理月报日期', '数据周期', '标段', '取样数量', '验收批次', '是否符合要求']],
    [['监理月报日期', '数据周期', '标段', '部位', '多点位移计', '锚杆应力计', '孔隙水压力计', '钢筋应力计', '水位', '测缝针', '压应力计', '土体位移计']],
    [['监理月报日期', '数据周期', '监测部位', '仪器及设备名称', '单位', '设计数量', '本月完成量', '累计完成量', '损坏量', '完成率', '完好率']],
    [['监理月报日期', '数据周期', '监理中心验收批次', '监理中心原材料抽检次数', '监理中心半成品抽检次数', '是否符合要求', '物探自检次数', '物探自检数量','物探监理检测次数', '物探监理检测数量', '物探受检总量', '存在的问题', '处理措施', '安全监测仪器质量评定情况', '质量生产周例会', '整改通知书', '检查']],
    [['监理月报日期', '数据周期', '合同名称', '合同编号', '合同价款（元）', '合同起止日期', '施工单位', '合同状态', '备注', '标段']],
    [['监理月报日期', '数据周期', '标段', '劳务分包计划', '专业分包计划', '劳务分包申请', '专业分包申请', '劳务分包合同及安全管理协议', '专业分包合同及安全管理协议', '备注']],
    [['监理月报日期', '数据周期', '标段', '分包工程项目名称', '分包项目内容（部位）', '分包形式', '工程数量', '拟分包额（万元）', '施工总承包合同额（万元）', '占总合同比例', '备注']],
    [['监理月报日期', '数据周期', '合同名称', '合同编号', '工程变更项目名称', '新增单价项目名称', '单位', '新增单价编号', '审批单价', '备注', '标段']],
    [['监理月报日期', '数据周期', '合同项目名称', '合同编号', '施工单位', '合同变更项目名称', '项目名称', '单位', '工程量', '单价（元）', '合计（元）', '备注', '标段']],
    [['监理月报日期', '数据周期', '序号', '合同项目名称', '合同编号', '施工单位', '索赔意向名称', '监理审核情况', '业主审批情况', '备注', '标段']],
    [['监理月报日期', '数据周期', '合同项目名称', '合同编号', '施工单位', '合同总额(元）', '工程预付款总金额(元）', '已支付金额', '本月扣回金额（元）', '累计扣回金额（元）', '未扣回金额（元）', '备注', '标段']],
    [['监理月报日期', '数据周期', '合同项目', '合同编号', '施工单位', '合同总额（元）', '施工申报（元）', '监理审核（元）', '业主审批（元）', '累计完成', '累计完成比例', '备注']],
    [['监理月报日期', '数据周期', '名称', '计划', '实际', '工序数量', '工序应用效果良好', '工序应用率', '应用率']]
]

# 原表格中数据项之前的无用行
sheetsHeadRows = [0, 1, 0, 1, 1, 0, 2, 1, 1, 1, 1, 0, 0, 2, 1, 2, 0, 2, 2, 0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0]
# 原表格中数据项之前的无用列
sheetsHeadCols = [0, 1, 0, 1, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 1, 1, 0]
# 总表相比原表在数据项之前增加的列数
newLines       = [1, 1, 1, 1, 1, 1, 2, 3, 2, 1, 1, 0, 0, 2, 2, 2, 0, 1, 3, 0, 0, 2, 0, 1, 1, 1, 1, 1, 2, 1, 1, 0]

# 表格的类型，不同类型对应最终Excel中不同的sheet
class TableType(enum.Enum):
    default = -1
    filelist = 0
    contractor = 1 # 承包人相关表，对应表格“参建单位”
    secureTrainingData = 2 # 对应表格'安全培训数据'
    hazardManagement = 3 # 对应表格'危险源管理情况统计表'
    safetyInvestigation = 4 # 对应表格'安全隐患排查治理一览表'
    safetyProdutionCost = 5 # 对应表格'安全生产费用投入统计表'
    violation = 6 # 对应表格'违章隐患统计表'
    personnelInvestment = 7 # 对应表格'各标段承包人人员投入情况'
    equipmentInvestment = 8 # 对应表格'各标段设备投入情况'
    monthlyProjectCompleted = 9 # 对应表格'本月主要完成工程量统计表'
    monthlyProjectDescription = 10 # 对应表格'本月工程形象进度面貌'
    scheduleCompleted = 11 # 对应表格'本月进度计划完成情况'
    scheduleDescription = 12 # 对应表格'进度描述'
    anchorRodTest = 13 # 对应表格'本月锚杆监理检测结果一览表'
    blastingMonitoring = 14 # 对应表格'本月爆破振动监测成果表'
    projectQualityEvaluation = 15 # 对应表格'工序质量验评表'
    rawMaterialArrival = 16 # 对应表格'原材料到货统计'
    inspectionLotAcceptance = 17 # 对应表格'各标段检验批验收统计'
    miningQuality = 18 # 对应表格'开挖质量控制表'
    rawMaterialTest = 19 # 对应表格'原材料试验监测情况表'
    securityEquipmentMonitoring = 20 # 对应表格'安全监测设备监测情况表'
    securityEquipmentSetting = 21 # 对应表格'安全监测仪器安装情况表'
    qualityProblemDescription = 22 # 对应表格'质量问题描述表'
    contract = 23 # 对应表格'合同台账'
    subcontractSubmission = 24 # 对应表格'本月各标段分包报审报备情况统计表'
    subcontractPlan = 25 # 对应表格'分包计划一览表'
    newUnitPrice = 26 # 对应表格'新增单价统计表'
    contractModification = 27 # 对应表格'合同变更会签单'
    claimIndembity = 28 # 对应表格'工程索赔表'
    advancePaymentRebate = 29 # 对应表格'本月预付款与回扣统计表'
    monthlyProgressPayment = 30 # 对应表格'月进度款支付统计表'
    standardProcess = 31 # 对应表格'标准工艺'

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
    cells = head.cells[:] # 表头的每个单元格（合并单元格重复显示）
    head_content = []
    for cell in cells:
        head_content.append(cell.text)
    # print(head_content)
    numOfColumns = len(cells) # 表的列数
    determinTableType(table, head_content, numOfColumns, pathDir, excelFile, date)
    # print('--------------------------')

# 判定传入head对应的表的类型
def determinTableType(table, head, num, pathDir, excelFile, date):
    specialType = TableType.default
    table_type = TableType.default
    if num == 4:
        if(_match(head[0], "序号") and _match(head[1], "标段名称") and _match(head[2], "承包人") and _match(head[3], "备注")):
            # 表类型为“承包人”相关类型，Excel对应“参建单位”sheet
            table_type = TableType.contractor
    elif num == 5:
        if _match(head[0], "序号") and _match(head[1], "设备/设施名称"):
            # 表类型为“设备投入”相关类型，Excel对应“各标段投入设备情况”sheet
            table_type = TableType.equipmentInvestment
        elif _match(head[0], "序号") and _match(head[1], "标段名称") and _match(head[3], "主要设计特性描述"):
            # 表类型为“形象进度”相关类型，Excel对应“本月工程形象进度面貌”sheet
            table_type = TableType.monthlyProjectDescription
    elif num == 8:
        if _match(head[0], "序号") and _match(head[1], "隐患简题"):
            # 表类型为“安全隐患排查”相关类型，Excel对应“安全隐患排查治理一览表”sheet
            table_type = TableType.safetyInvestigation
        elif _match(head[0], "标段") and _match(head[1], "工程部位") and _match(head[2], "检测数量"):
            # 表类型为“锚杆检测”相关类型，Excel对应“本月锚杆监理检测结果一览表”sheet
            specialType = TableType.anchorRodTest
            table_type = TableType.anchorRodTest
        elif _match(head[0], "序号") and _match(head[1], "合同名称") and _match(head[2], "合同编号"):
            # 表类型为“合同台账”相关类型，Excel对应“合同台账表”sheet
            table_type = TableType.contract
        elif _match(head[0], "序号") and _match(head[1], "合同项目名称") and _match(head[4], "索赔意向名称"):
            # 表类型为“索赔”相关类型，Excel对应“工程索赔表”sheet
            table_type = TableType.claimIndembity
    elif num == 9:
        if _match(head[0], "施工单位") and _match(head[1], "项目经理"):
            # 表类型为“承包人员投入”相关类型，Excel对应“各标段承包人人员投入情况”sheet
            table_type = TableType.personnelInvestment
        elif _match(head[0], "标段") and _match(head[1], "工程部位") and _match(head[4], "受检根数"):
            # 表类型为“锚杆检测”相关类型，Excel对应“本月锚杆监理检测结果一览表”sheet
            table_type = TableType.anchorRodTest
        elif _match(head[0], "监测部位") and _match(head[1], "仪器及设备名称"):
            # 表类型为“设备安装”相关类型，Excel对应“安全监测仪器安装情况表”sheet
            table_type = TableType.securityEquipmentSetting
        elif _match(head[0], "序号") and _match(head[1], "标段") and _match(head[2], "劳务分包计划"):
            # 表类型为“分包报审报备”相关类型，Excel对应“本月各标段分包报审报备情况统计表”sheet
            table_type = TableType.subcontractSubmission
        elif _match(head[0], "序号") and _match(head[1], "合同名称") and _match(head[3], "工程变更项目名称"):
            # 表类型为“新增单价”相关类型，Excel对应“新增单价统计表”sheet
            table_type = TableType.newUnitPrice
    elif num == 10:
        if _match(head[0], "序号") and _match(head[1], "危险源名称"):
            # 表类型为“危险源”相关类型，Excel对应“危险源管理情况统计表”sheet
            table_type = TableType.hazardManagement
        elif _match(head[0], "序号") and _match(head[1], "项目名称") and _match(head[4], "合同工程量"):
            # 表类型为“月工程量”相关类型，Excel对应“本月主要完成工程量统计表”sheet
            table_type = TableType.monthlyProjectCompleted
        elif _match(head[0], "序号") and _match(head[1], "标段") and _match(head[2], "分包工程项目名称"):
            # 表类型为“分包计划”相关类型，Excel对应“分包计划一览表”sheet
            table_type = TableType.subcontractPlan
        elif _match(head[0], "测点编号") and _match(head[1], "通道"):
            # 表类型为“爆炸监测”相关类型，Excel对应“本月爆破振动监测成果表”sheet
            table_type = TableType.blastingMonitoring
    elif num == 11:
        if _match(head[0], "施工单位") and _match(head[1], "本月检查次数"):
            # 表类型为“违章”相关类型，Excel对应“违章隐患统计表”sheet
            table_type = TableType.violation
        elif _match(head[0], "测点编号") and _match(head[1], "通道"):
            # 表类型为“爆炸监测”相关类型，Excel对应“本月爆破振动监测成果表”sheet
            table_type = TableType.blastingMonitoring
        elif _match(head[0], "序号") and _match(head[1], "合同项目名称") and _match(head[4], "合同变更项目名称" ):
            # 表类型为“合同变更”相关类型，Excel对应“合同变更会签单”sheet
            table_type = TableType.contractModification
        elif _match(head[0], "序号") and _match(head[1], "合同项目名称") and _match(head[4], "合同总额(元)" ):
            # 表类型为“预付款与回扣”相关类型，Excel对应“本月预付款支付与回扣统计表”sheet
            table_type = TableType.advancePaymentRebate
        elif _match(head[0], "序号") and _match(head[1], "合同项目") and _match(head[8], "累计完成" ):
            # 表类型为“月进度款”相关类型，Excel对应“月进度款支付统计表”sheet
            table_type = TableType.monthlyProgressPayment
    elif num == 12:
        if _match(head[0], "序号") and _match(head[1], "施工标段") and _match(head[2], "单位工程名称"):
            # 表类型为“检验批验收”相关类型，Excel对应“各标段检验批验收统计表”sheet
            table_type = TableType.inspectionLotAcceptance
        elif _match(head[0], "序号") and _match(head[1], "工程部位") and _match(head[2], "设计开挖量(m³)"):
            # 表类型为“开挖质量”相关类型，Excel对应“开挖质量控制表”sheet
            table_type = TableType.miningQuality
    elif num == 20 or num == 15:
        if _match(head[0], "序号") and _match(head[1], "施工标段"):
            # 表类型为“工程质量验评”相关类型，Excel对应“工序质量验评表”sheet
            table_type = TableType.projectQualityEvaluation
    
    if table_type != TableType.default:
        data = xlrd.open_workbook(pathDir+excelFile)
        nrows, ncols = getSheetRowsCols(sheets[table_type.value], data)
        writeSheetContent(table, table_type, nrows, ncols, pathDir, excelFile, date, specialType)


# 获取将要写入的sheet的现有行数列数
def getSheetRowsCols(sheetName, data):
    sheet = data.sheet_by_name(sheetName)
    return sheet.nrows, sheet.ncols

# 写入sheet的具体内容
def writeSheetContent(table, sheetType, nrows, ncols, pathDir, excelFile, date, specialType=-1):
    data = xlrd.open_workbook(pathDir+excelFile, formatting_info=True)
    excel = copy(data) # 完成xlrd对象向xlwt的转换
    sheet = excel.get_sheet(sheetType.value) # 获得需要的操作页
    headerRows = getTableHeaderRows(table, sheetType.value) # 获取数据项前的无用行数
    headerCols = getTableHeaderCols(table, sheetType.value) # 获取数据项前的无用列数

    # 写入
    if specialType == TableType.default:
        for i in range(headerRows, len(table.rows)):
            sheet.write(nrows, 0, date)
            sheet.write(nrows, 1, date)

            row = table.rows[i]
            cells = row.cells[:]
            for k in range(headerCols, len(cells)):
                n = k+newLines[sheetType.value]
                sheet.write(nrows, n, cells[k].text)
            nrows += 1
    elif specialType == TableType.anchorRodTest:
        for i in range(headerRows, len(table.rows)):
            sheet.write(nrows, 0, date)
            sheet.write(nrows, 1, date)

            row = table.rows[i]
            cells = row.cells[:]
            for k in range(headerCols, len(cells)):
                n = k+newLines[sheetType.value]+(1 if k >= 2 else 0)
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
    if(s2 == s1.replace("\n", "").replace("\r", "").replace(" ", "").replace("（","(").replace("）",")")):
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
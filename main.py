from infoExtraction import *
from formExtraction import *

from sklearn.manifold import TSNE
from sklearn.cluster import KMeans
from sklearn import preprocessing
from sklearn import metrics
from sklearn import ensemble
from decimal import Decimal
import matplotlib.pyplot as plt

from sklearn.linear_model import BayesianRidge, LinearRegression, ElasticNet
from sklearn.svm import SVR
from sklearn.model_selection import cross_val_score    # 交叉验证
from sklearn.metrics import explained_variance_score, mean_absolute_error, mean_squared_error, r2_score  

import numpy as np
import pandas as pd
 
class secureModel:
    def __init__(self):
        self.tableName = [
            "安全培训数据",
            "危险源管理情况统计表",
            "安全隐患排查治理一览表",
            "违章隐患统计表"
        ]

        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-01', '2018-02', '2018-03', '2018-04', '2018-05', '2018-06', '2018-07',
                '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03',
                '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        # self.result = [[0, 1, 1, 4, 4, 0, 17, 4, 1, 3, 2, 1, 7, 0, 2], [0, 1, 1, 3, 3, 0, 8, 4, 1, 3, 2, 1, 7, 0, 2], [0, 1, 1, 4, 4, 0, 9, 25, 1, 0, 0, 22, 24, 0, 15], [0, 1, 1, 4, 4, 1, 17, 5, 16, 9, 3, 3, 5, 0, 4], [0, 1, 1, 4, 4, 0, 16, 11, 12, 9, 10, 0, 18, 0, 1], [0, 1, 1, 4, 4, 0, 13, 2, 0, 4, 11, 1, 10, 3, 0], [0, 1, 1, 4, 4, 0, 12, 4, 0, 3, 32, 3, 64, 0, 0], [0, 1, 1, 4, 4, 0, 22, 8, 0, 9, 10, 2, 8, 2, 4], [0, 1, 1, 3, 3, 0, 13, 17, 5, 15, 9, 4, 5, 1, 4], [0, 1, 1, 4, 4, 2, 13, 17, 5, 15, 9, 4, 5, 1, 4], [0, 1, 1, 4, 4, 0, 14, 10, 6, 12, 7, 2, 8, 0, 12], [0, 1, 1, 4, 4, 0, 6, 9, 4, 4, 4, 0, 4, 0, 1], [3, 1, 1, 5, 5, 0, 5, 6, 0, 2, 5, 0, 28, 0, 4], [0, 1, 1, 3, 3, 0, 1, 8, 1, 2, 1, 1, 1, 0, 1], [4, 1, 1, 5, 5, 0, 4, 14, 0, 6, 10, 0, 23, 0, 2], [5, 1, 1, 7, 7, 0, 6, 13, 2, 15, 9, 3, 13, 3, 9], [5, 1, 1, 26, 26, 0, 3, 4, 0, 5, 7, 11, 0, 4, 9], [4, 1, 1, 26, 26, 0, 3, 3, 1, 5, 0, 5, 1, 1, 3], [4, 1, 1, 28, 28, 1, 2, 2, 0, 3, 5, 0, 2, 0, 0], [0, 1, 1, 26, 26, 1, 5, 17, 8, 18, 17, 0, 9, 2, 10], [0, 1, 1, 27, 27, 0, 5, 15, 5, 6, 4, 1, 15, 0, 0], [0, 1, 1, 26, 26, 0, 4, 6, 1, 9, 2, 1, 5, 1, 0], [0, 1, 1, 24, 24, 0, 3, 7, 1, 6, 4, 1, 34, 7, 2], [0, 1, 1, 23, 23, 0, 3, 21, 0, 10, 0, 0, 28, 5, 2], [0, 1, 1, 19, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 1, 1, 19, 19, 0, 2, 2, 0, 0, 0, 0, 0, 0, 0], [0, 1, 1, 26, 26, 0, 4, 3, 0, 0, 0, 0, 0, 0, 12]]
        self.getVector() # 获取特征向量

        self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
    
    def getVector(self):
        for date in self.month:
            v = []
            self.getSecureTrainingData(date, v) # 获取安全培训数据特征向量
            self.getHazardManagement(date, v) # 获取危险源管理相关特征
            self.getSafetyInvestigation(date, v) # 获取安全隐患排查治理相关特征
            self.getViolation(date, v) # 获取违章隐患特征
            self.result.append(v)
        self.printResult()
        # print(self.result)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            print("\n")
            print("report date:", self.month[i])
            score = 100
            if self.result[i][0] <= 0:
                score -= 5
                print("-5：未召开监理安全周例会")
            elif self.result[i][0] <= 1:
                score -= 3
                print("-3：监理安全周例会只召开1次")
            elif self.result[i][0] <= 2:
                score -= 1
                print("-1：监理安全周例会只召开2次")
            if self.result[i][1] < 1:
                score -= 5
                print("-5：未组织安全培训教育")
            if self.result[i][2] < 1:
                score -= 5
                print("-5：职业健康采取措施不当")
            if self.result[i][3] > 3:
                score -= 3
                print("-3：危险源数目超过3次")
            elif self.result[i][3] > 5:
                score -= 5
                print("-5：危险源数目超过5次")
            elif self.result[i][3] > 7:
                score -= 10
                print("-10：危险源数目超过7次")
            if self.result[i][4] < self.result[i][3]:
                score -= 10
                print("-10：存在未受控的危险源")
            if self.result[i][5] == 2:
                score -= 10
                print("-10：存在未消除的隐患")
            if self.result[i][6] < 5:
                score -= 5
                print("-5：本月违章检查次数不足5次")
            elif self.result[i][6] < 10:
                score -= 3
                print("-3：本月违章检查次数不足10次")
            if self.result[i][-1] > 50:
                score -= 10
                print("-10：本月违章总次数超过50次")
            elif self.result[i][-1] > 40:
                score -= 7
                print("-7：本月违章总次数超过40次")
            elif self.result[i][-1] > 30:
                score -= 5
                print("-5：本月违章总次数超过30次")
            print(self.month[i], "安全总分：", score)

    def preProcess(self):
        self.result = preprocessing.normalize(self.result)

    def KMeansModel(self):
        kmeans_model = KMeans(n_clusters=3).fit_predict(self.result)
        labels = kmeans_model
        print(metrics.silhouette_score(self.result, labels, metric='euclidean'))
        self.showResult(labels)

    def showResult(self, labels):
        x_tsne = TSNE(n_components=2, n_iter=250).fit_transform(self.result)
        plt.scatter(x_tsne[:, 0], x_tsne[:, 1], s=50, c=labels)
        plt.show()

    '''
    从安全培训数据表中提取3个特征，包括监理安全周例会次数，是否组织安全培训教育（0或1），
    职业健康采取的措施是否得当（0或1）
    '''
    def getSecureTrainingData(self, date, v):
        sheetType = TableType.secureTrainingData
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 2) # 安全周例会次数
                v.append(int(value1) if value1 else 0)

                value2 = sheet.cell_value(i, 3) # 是否组织安全培训教育
                value3 = sheet.cell_value(i, 6) # 职业健康采取的措施是否得当
                v.append(1 if "是" in value2 else 0)
                v.append(1 if "是" in value3 else 0)
                return
        
        for i in range(3): v.append(0) # 未检测到当月相关数据
                    
    '''
    从危险源管理情况中提取2个特征，包括累计危险源数目与受控危险源数目
    '''
    def getHazardManagement(self, date, v):
        sheetType = TableType.hazardManagement
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        value1 = 0
        value2 = 0
        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 += 1
                if "有" in sheet.cell_value(i, 8): value2 += 1
        v.append(value1)
        v.append(value2)

    '''
    从安全隐患排查治理中提取1个特征
    0：无隐患 1：有隐患，已消除 2：有隐患，未消除
    '''
    def getSafetyInvestigation(self, date, v):
        sheetType = TableType.safetyInvestigation
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 2) # 隐患简题
                if len(value1) <= 1:
                    v.append(0)
                    return
                else:
                    value2 = sheet.cell_value(i, 7) # 是否消除
                    v.append(2 if "否" in value2 else 1) 
                    return

        v.append(0) # 未检测到当月相关数据

    '''
    从违章隐患统计中提取10个特征
    本月检查次数、8项内容违章数、违章总数
    '''
    def getViolation(self, date, v):
        sheetType = TableType.violation
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date and "合计" in sheet.cell_value(i, 2):
                for k in range(3, ncols):
                    value = sheet.cell_value(i, k)
                    v.append(int(value) if value.isdigit() else 0)
                return
        
        for i in range(10): v.append(0) # 未检测到当月相关数据

class qualityModel:
    def __init__(self):
        self.tableName = [
            "本月锚杆监理监测结果一览表",
            "工序质量验评表",
            "本月爆破振动监测成果表",
            "各标段检验批验收统计"
        ]
        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-01', '2018-02', '2018-03', '2018-04', '2018-05', '2018-06', '2018-07',
                '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03',
                '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']
                
        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量

        self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getAnchorRodTest(date, v) # 获取锚杆监理监测结果数据
            self.getBlastingMonitoring(date, v) # 获取爆破振动监测数据
            self.getProjectQualityEvaluation(date, v) #获取工序质量验评数据
            self.getInspectionLotAcceptance(date, v) # 获取各标段检验批验收数据

            self.result.append(v)

        # self.dataProcess()
        # self.printResult()
        # print(self.result)

    def dataProcess(self):
        _m = 0
        for i in self.result:
            _m = len(i) if _m < len(i) else _m
        for i in self.result:
            if len(i) < _m:
                for _ in range(_m-len(i)):
                    i.append(0)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], len(self.result[i]), self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            print("\n")
            print("report date:", self.month[i])
            score = 100
            if self.result[i][0] == 0:
                score -= 5
                print("-5：无锚杆监测相关结果")
            else:
                if self.result[i][3] > 0 or self.result[0][9] > 0:
                    score -= 3
                    print("-3：锚杆监测结果存在III级")
                if self.result[i][4] > 0 or self.result[0][10] > 0:
                    score -= 5
                    print("-5：锚杆监测结果存在IV级")
                if self.result[i][5] == 0 or self.result[0][11] == 0:
                    score -= 10
                    print("-10：锚杆监测结果存在不合格")
            
            if self.result[i][12] == 0:
                score -= 10
                print("-10：爆破振动监测不合格")
            
            if self.result[i][15] == 0:
                score -= 3
                print("-3：无工序质量验评相关结果")
            elif self.result[i][15] < 80:
                score -= 10
                print("-10：工序质量验评低于80%")
            elif self.result[i][15] < 90:
                score -= 7
                print("-7：工序质量验评低于90%")
            elif self.result[i][15] < 95:
                score -= 5
                print("-5：工序质量验评低于95%")
            elif self.result[i][15] < 97:
                score -= 2
                print("-2：工序质量验评低于97%")
            
            if self.result[i][18] == 0:
                score -= 3
                print("-3：无检验批验收相关结果")
            elif self.result[i][18] < 80:
                score -= 10
                print("-10：各标段检验批验收优良率低于80%")
            elif self.result[i][18] < 90:
                score -= 7
                print("-7：各标段检验批验收优良率低于90%")
            elif self.result[i][18] < 95:
                score -= 5
                print("-5：各标段检验批验收优良率低于95%")
            elif self.result[i][18] < 97:
                score -= 2
                print("-2：各标段检验批验收优良率低于97%")


            
            print(self.month[i], "质量总分：", score)

    def preProcess(self):
        self.result = preprocessing.normalize(self.result)

    def KMeansModel(self):
        kmeans_model = KMeans(n_clusters=3).fit_predict(self.result)
        labels = kmeans_model
        print(metrics.silhouette_score(self.result, labels, metric='euclidean'))
        self.showResult(labels)

    def showResult(self, labels):
        x_tsne = TSNE(n_components=2, n_iter=250).fit_transform(self.result)
        plt.scatter(x_tsne[:, 0], x_tsne[:, 1], s=50, c=labels)
        plt.show()

    '''
    从锚杆监理监测结果中提取6个特征
    '''
    def getAnchorRodTest(self, date, v):
        sheetType = TableType.anchorRodTest
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                if "小计" in sheet.cell_value(i, 2):
                    value1 = sheet.cell_value(i, 5) # 检测数量
                    v.append(int(value1) if value1 else 0)

                    value2 = sheet.cell_value(i, 6) # I级
                    v.append(int(value2) if value2 else 0)
                    
                    value3 = sheet.cell_value(i, 7) # II级
                    v.append(int(value3) if value3 else 0)

                    value4 = sheet.cell_value(i, 8) # III级
                    v.append(int(value4) if value4 else 0)

                    value5 = sheet.cell_value(i, 9) # IV级
                    v.append(int(value5) if value5 else 0)

                    v.append(1)

        for i in range(12-len(v)):
            v.append(0)

    '''
    从爆破振动监测中提取1个特征
    '''
    def getBlastingMonitoring(self, date, v):
        sheetType = TableType.blastingMonitoring
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        flag = True
        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 9) # 合格
                if "合格" not in value1:
                    flag = False
                    break
        v.append(1 if flag else 0)
    
    '''
    从工序质量验评中提取3个特征
    '''
    def getProjectQualityEvaluation(self, date, v):
        sheetType = TableType.projectQualityEvaluation
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date and "合计" in sheet.cell_value(i, 2):
                value1 = sheet.cell_value(i, 7) # 合格数量
                v.append(int(value1) if value1 and value1.isdigit() else 0)

                value2 = sheet.cell_value(i, 10) # 优良数量
                v.append(int(value2) if value2 and value2.isdigit() else 0)
                    
                value3 = sheet.cell_value(i, 11) # 优良率
                v.append(float(value3) if value3 and ("." in value3 or value3.isdigit()) else 0)
                return
        for i in range(3):
            v.append(0)

    '''
    从各标段检验批验收中提取3个特征
    '''
    def getInspectionLotAcceptance(self, date, v):
        sheetType = TableType.inspectionLotAcceptance
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date and "合计" in sheet.cell_value(i, 2):
                value1 = sheet.cell_value(i, 4) # 合格数量
                v.append(int(value1) if value1 and value1.isdigit() else 0)

                value2 = sheet.cell_value(i, 5) # 优良数量
                v.append(int(value2) if value2 and value2.isdigit() else 0)
                    
                value3 = sheet.cell_value(i, 6) # 优良率
                v.append(float(value3) if value3 and ("." in value3 or value3.isdigit()) else 0)
                return
        for i in range(3):
            v.append(0)

class progressModel:
    def __init__(self):
        self.tableName = [
            "本月主要完成工程量统计表", "本月进度计划完成情况", "进度描述"]
        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-01', '2018-02', '2018-03', '2018-04', '2018-05', '2018-06', '2018-07',
                '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03',
                '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']
        self.procedure = [] # 存储对应月的进度情况
        self.procedureAvg = []

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量
        self.dataProcess()

        self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
        self.regressionModel()
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getMonthlyProjectCompleted(date, v) # 获取主要完成工程量
            self.getScheduleCompleted(date, v) # 获取本月进度计划完成情况
            self.getScheduleDescription(date, v) # 获取进度描述

            self.result.append(v)

        self.printResult()
    
    def dataProcess(self):
        _m = 0
        for i in self.procedure:
            _m = len(i) if _m < len(i) else _m

        for i in self.procedure:
            if len(i) < _m:
                for _ in range(_m-len(i)):
                    i.append(0)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            self.procedureAvg.append(self.result[i][1])
            print("\n")
            print("report date:", self.month[i])
            score = 100
            if self.result[i][0] == 0:
                score -= 15
                print("-15：本月进度情况丢失或没有进度")
            elif self.result[i][0] <= 50:
                score -= 10
                print("-10：本月进度完成占计划比小于50%")
            elif self.result[i][0] <= 75:
                score -= 5
                print("-5：本月进度完成占计划比小于75%")

            if self.result[i][2] <= 0.5:
                score -= 5
                print("-5：本月进度计划完成情况不足50%")
            elif self.result[i][2] <= 0.75:
                score -= 3
                print("-3: 本月进度计划完成情况不足75%")

            if self.result[i][3] == 0:
                score -= 3
                print("-3：没有施工进度计划")
            if self.result[i][4] == 0:
                score -= 3
                print("-3：没有专项检查")
            if self.result[i][5] == 0:
                score -= 3
                print("-3：没有生产周进度分析例会")
            if self.result[i][6] == 0:
                score -= 3
                print("-3：没有施工专题会")
            
            print(self.month[i], "进度总分：", score)

    def regressionModel(self):
        print("-----------------------------")
        # 自变量
        X = self.procedure
        # 因变量
        y = self.procedureAvg

        # 设置交叉验证次数
        n_folds = 5

        # 建立贝叶斯岭回归模型
        br_model = BayesianRidge()

        # 普通线性回归
        lr_model = LinearRegression()

        # 弹性网络回归模型
        etc_model = ElasticNet()

        # 支持向量机回归
        svr_model = SVR()

        # 梯度增强回归模型对象
        gbr_model = ensemble.GradientBoostingRegressor()

        # 不同模型的名称列表
        model_names = ['BayesianRidge', 'LinearRegression', 'ElasticNet', 'SVR', 'GBR']
        # 不同回归模型
        model_dic = [br_model, lr_model, etc_model, svr_model, gbr_model]
        # 交叉验证结果
        cv_score_list = []
        # 各个回归模型预测的y值列表
        pre_y_list = []

        # 读出每个回归模型对象
        for model in model_dic:
            # 将每个回归模型导入交叉检验
            scores = cross_val_score(model, X, y, cv=n_folds)
            # 将交叉检验结果存入结果列表
            cv_score_list.append(scores)
            # 将回归训练中得到的预测y存入列表
            pre_y_list.append(model.fit(X, y).predict(X))

        ### 模型效果指标评估 ###
        # 获取样本量，特征数
        # n_sample, n_feature = X.shape
        # 回归评估指标对象列表
        model_metrics_name = [explained_variance_score, mean_absolute_error, mean_squared_error, r2_score]
        # 回归评估指标列表
        model_metrics_list = []
        # 循环每个模型的预测结果
        for pre_y in pre_y_list:
            # 临时结果列表
            tmp_list = []
            # 循环每个指标对象
            for mdl in model_metrics_name:
                # 计算每个回归指标结果
                tmp_score = mdl(y, pre_y)
                # 将结果存入临时列表
                tmp_list.append(tmp_score)
            # 将结果存入回归评估列表
            model_metrics_list.append(tmp_list)
        df_score = pd.DataFrame(cv_score_list, index=model_names)
        df_met = pd.DataFrame(model_metrics_list, index=model_names, columns=['ev', 'mae', 'mse', 'r2'])
        print(df_score)
        print(df_met)


    def KMeansModel(self):
        kmeans_model = KMeans(n_clusters=3).fit_predict(self.result)
        labels = kmeans_model
        print(metrics.silhouette_score(self.result, labels, metric='euclidean'))
        self.showResult(labels)

    def showResult(self, labels):
        x_tsne = TSNE(n_components=2, n_iter=250).fit_transform(self.result)
        plt.scatter(x_tsne[:, 0], x_tsne[:, 1], s=50, c=labels)
        plt.show()
    
    def isFloatNum(self, st):
        s=st.split('.')
        if len(s)>2:
            return False
        else:
            for si in s:
                if not si.isdigit():
                    return False
            return True

    '''
    从本月主要完成工程量中提取本月完成情况
    '''
    def getMonthlyProjectCompleted(self, date, v):
        sheetType = TableType.monthlyProjectCompleted
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        num = 0
        num1 = 0
        num2 = 0
        temp = []
        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                
                value1 = sheet.cell_value(i, 8) # 本月完成占计划比
                value2 = sheet.cell_value(i, 10) # 累计完成占比

                if self.isFloatNum(value1):
                    num += 1

                    temp.append(float(value1) if value1 and self.isFloatNum(value1) else 0)

                    num1 += float(value1.replace('%', '')) if value1 and self.isFloatNum(value1.replace('%', '')) else 0
                    num2 += float(value2.replace('%', '')) if value2 and self.isFloatNum(value2.replace('%', '')) else 0

        # print(date, num, num1, num2)

        v.append(num1/num if num is not 0 else 0)
        v.append(num2/num if num is not 0 else 0)
        self.procedure.append(temp)

    '''
    从本月进度计划完成情况中提取本月完成情况
    '''
    def getScheduleCompleted(self, date, v):
        sheetType = TableType.scheduleCompleted
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        num1 = 0
        num2 = 0
        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 4)
                value2 = sheet.cell_value(i, 5)
                num1 += int(value1) if value1 else 0
                num2 += int(value2) if value2 else 0
        
        v.append(num2/num1 if num1 is not 0 else 0)

    '''
    从进度描述中提取本月完成情况
    '''
    def getScheduleDescription(self, date, v):
        sheetType = TableType.scheduleDescription
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 2)
                v.append(int(value1) if value1 else 0)

                value2 = sheet.cell_value(i, 3)
                v.append(int(value2) if value2 else 0)

                value3 = sheet.cell_value(i, 4)
                v.append(int(value3) if value3 else 0)

                value4 = sheet.cell_value(i, 5)
                v.append(int(value4) if value4 else 0)

                value5 = sheet.cell_value(i, 6)
                v.append(int(value5) if value5 else 0)

                return
        for i in range(5):
            v.append(0)

class economicModel:
    def __init__(self):
        self.tableName = [
            "合同变更会签单", "本月预付款支付与回扣统计表", "月进度支付统计表"]
        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-01', '2018-02', '2018-03', '2018-04', '2018-05', '2018-06', '2018-07',
                '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03',
                '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']
        self.economics = [] # 存储对应月的技经情况

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量

        # self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
        self.regressionModel()
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getMonthlyProgressPayment(date, v) # 获取合同管理及投资控制数据特征向量

            self.result.append(v)

        self.dataProcess(self.result)
        # self.printResult()
        # print(self.result)
    
    def dataProcess(self, tab):
        _m = 0
        for i in tab:
            _m = len(i) if _m < len(i) else _m

        for i in tab:
            if len(i) < _m:
                for _ in range(_m-len(i)):
                    i.append(0)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            _s = 0
            _n = int(len(self.result[i])/3)
            for k in range(_n):
                _s += self.result[i][k*3-1]
            if _n is not 0:
                self.economics.append(_s/_n)
            else:
                self.economics.append(0)
        
        for i in range(len(self.month)):
            # print("\n")
            # print("report date:", self.month[i])
            if float(self.economics[i]) != 0:
                avg = float(self.economics[i])
            else:
                avg = float(self.economics[i+1])
            self.procedureAvg.append(avg)
            # print("总进度均值：", avg)
        
    def regressionModel(self):
        # 自变量
        X = self.result
        
        # 因变量
        y = self.economics

        # 设置交叉验证次数
        n_folds = 5

        # 建立贝叶斯岭回归模型
        br_model = BayesianRidge()

        # 普通线性回归
        lr_model = LinearRegression()

        # 弹性网络回归模型
        etc_model = ElasticNet()

        # 支持向量机回归
        svr_model = SVR()

        # 梯度增强回归模型对象
        gbr_model = ensemble.GradientBoostingRegressor()

        # 不同模型的名称列表
        model_names = ['BayesianRidge', 'LinearRegression', 'ElasticNet', 'SVR', 'GBR']
        # 不同回归模型
        model_dic = [br_model, lr_model, etc_model, svr_model, gbr_model]
        # 交叉验证结果
        cv_score_list = []
        # 各个回归模型预测的y值列表
        pre_y_list = []

        # 读出每个回归模型对象
        for model in model_dic:
            # 将每个回归模型导入交叉检验
            scores = cross_val_score(model, X, y, cv=n_folds)
            # 将交叉检验结果存入结果列表
            cv_score_list.append(scores)
            # 将回归训练中得到的预测y存入列表
            pre_y_list.append(model.fit(X, y).predict(X))

        ### 模型效果指标评估 ###
        # 获取样本量，特征数
        # n_sample, n_feature = X.shape
        # 回归评估指标对象列表
        model_metrics_name = [explained_variance_score, mean_absolute_error, mean_squared_error, r2_score]
        # 回归评估指标列表
        model_metrics_list = []
        # 循环每个模型的预测结果
        for pre_y in pre_y_list:
            # 临时结果列表
            tmp_list = []
            # 循环每个指标对象
            for mdl in model_metrics_name:
                # 计算每个回归指标结果
                tmp_score = mdl(y, pre_y)
                # 将结果存入临时列表
                tmp_list.append(tmp_score)
            # 将结果存入回归评估列表
            model_metrics_list.append(tmp_list)
        df_score = pd.DataFrame(cv_score_list, index=model_names)
        df_met = pd.DataFrame(model_metrics_list, index=model_names, columns=['ev', 'mae', 'mse', 'r2'])
        print(df_score)
        print(df_met)

    def preProcess(self):
        self.result = preprocessing.normalize(self.result)

    def KMeansModel(self):
        kmeans_model = KMeans(n_clusters=3).fit_predict(self.result)
        labels = kmeans_model
        print(metrics.silhouette_score(self.result, labels, metric='euclidean'))
        self.showResult(labels)

    def showResult(self, labels):
        x_tsne = TSNE(n_components=2, n_iter=250).fit_transform(self.result)
        plt.scatter(x_tsne[:, 0], x_tsne[:, 1], s=50, c=labels)
        plt.show()
    
    def isFloatNum(self, st):
        s=st.split('.')
        if len(s)>2:
            return False
        else:
            for si in s:
                if not si.isdigit():
                    return False
            return True

    '''
    从月进度支付统计表中提取每个合同德累计完成比例
    '''
    def getMonthlyProgressPayment(self, date, v):
        sheetType = TableType.monthlyProgressPayment
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        value4 = '0'

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date and "合计" not in sheet.cell_value(i, 2):
                value1 = sheet.cell_value(i, 5) # 合同总额
                v.append(int(float(value1.replace('%', ''))) if value1 and self.isFloatNum(value1.replace('%', '')) else 0)
                value2 = sheet.cell_value(i, 9) # 累计完成
                v.append(int(float(value2.replace('%', ''))) if value2 and self.isFloatNum(value2.replace('%', '')) else 0)
                value3 = sheet.cell_value(i, 10) # 累计完成比例
                v.append(float(value3.replace('%', '')) if value3 and self.isFloatNum(value3.replace('%', '')) else 0)
            elif sheet.cell_value(i, 0) == date and "合计" in sheet.cell_value(i, 2):
                value4 = sheet.cell_value(i, 10) # 合计完成比例
        self.economics.append(float(value4.replace('%', '')) if self.isFloatNum(value4.replace('%', '')) else 0)

def main():
    # m = secureModel()
    # m = qualityModel()
    # m = progressModel()
    m = economicModel()

if __name__ == "__main__":
    main()
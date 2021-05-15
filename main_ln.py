from infoExtraction_ln import *
from formExtraction_ln import *
from fake import *

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

        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-12', '2019-01', '2019-02', '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
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
            self.getSecureData(date, v) # 获取安全培训数据特征向量

            self.result.append(v)
        # self.printResult()
        # print(self.result)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            # print("\n")
            print("report date:", self.month[i])
            score = 100
            if self.result[i][0] <= 2:
                score -= 2
                print("-2：安全文函数目小于3份")
            if self.result[i][1] <= 4:
                score -= 5
                print("-5：监理工作数目小于4次")
            elif self.result[i][1] <= 10:
                score -= 2
                print("-2: 监理工作数目小于10次")
            if self.result[i][2] < 3:
                score -= 5
                print("-5：安全周例会次数小于3次")
            if self.result[i][3] > 10:
                score -= 10
                print("-10：违规整改次数超过10次")
            elif self.result[i][3] > 3:
                score -= 3
                print("-3：违规整改次数超过3次")
            if self.result[i][-1] != 1:
                score -= 10
                print("-10：存在未受控的危险源")
            
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
    从安全文明施工中提取4个特征，包括安全文函、监理工作、监理安全周例会次数、违规整改次数、危险源是否受控（失控为0，受控为1）
    '''
    def getSecureData(self, date, v):
        sheetType = TableType.secureConstruction
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 1) # 安全文函
                v.append(int(value1) if value1 else 0)

                value2 = sheet.cell_value(i, 2) # 监理工作
                v.append(int(value2) if value2 else 0)
                
                value3 = sheet.cell_value(i, 3) # 安全周例会次数
                v.append(int(value3) if value3 else 0)

                value4 = sheet.cell_value(i, 4) # 违规整改次数
                v.append(int(value4) if value4 else 0)

                value5 = sheet.cell_value(i, 5) # 危险源情况
                v.append(int(value5) if value5 else 0)

                return
        for i in range(5): v.append(0) # 未检测到当月相关数据

class qualityModel:
    def __init__(self):

        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-12', '2019-01', '2019-02', '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量

        self.ruleModel() # 规则模型
        self.preProcess() # 数据归一化预处理
        self.KMeansModel() # KMeans聚类
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getConstructionQuality(date, v) # 获取施工质量控制数据
            self.getQualityAcception(date, v) # 获取工程质量验收数据

            self.result.append(v)

        self.dataProcess()
        # self.printResult()
        # print(self.result)

    def dataProcess(self):
        print("enter")
        _m = 0
        for i in self.result:
            _m = len(i) if _m < len(i) else _m
        print(_m)
        for i in self.result:
            if len(i) < _m:
                for _ in range(_m-len(i)):
                    i.append(0)
    
    def printResult(self):
        for i in range(len(self.result)):
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            print("\n")
            print("report date:", self.month[i])
            score = 100
            if self.result[i][0] <= 2:
                score -= 3
                print("-3：施工管理体系文件数目小于3份")
            elif self.result[i][0] <= 6:
                score -= 2
                print("-2：施工管理体系文件数目小于6份")
            
            if self.result[i][1] <= 10:
                score -= 5
                print("-5：施工文件数目小于10份")
            
            if self.result[i][2] == 0:
                score -= 1
                print("-1：无设计图纸")
            
            if self.result[i][4] <= 3:
                score -= 5
                print("-5：工程技术联系单小于3份")
            elif self.result[i][4] <= 6:
                score -= 2
                print("-2：工程技术联系单小于6份")
            
            if self.result[i][-1] <= 90:
                score -= 10
                print("-10: 工程质量验收优良率低于90%")
            elif self.result[i][-1] <= 95:
                score -= 5
                print("-5: 工程质量验收优良率低于95%")
                
            
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
    从施工质量控制中提取5个特征，包括质量管理体系文件、施工文件、设计图纸、设计通知单、工程技术联系单及备忘录
    '''
    def getConstructionQuality(self, date, v):
        sheetType = TableType.constructionQuality
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 1) # 质量管理体系文件
                v.append(int(value1) if value1 else 0)

                value2 = sheet.cell_value(i, 2) # 施工文件
                v.append(int(value2) if value2 else 0)
                
                value3 = sheet.cell_value(i, 3) # 设计图纸
                v.append(int(value3) if value3 else 0)

                value4 = sheet.cell_value(i, 4) # 设计通知单
                v.append(int(value4) if value4 else 0)

                value5 = sheet.cell_value(i, 5) # 工程技术联系单及备忘录
                v.append(int(value5) if value5 else 0)

                return
        for i in range(5): v.append(0) # 未检测到当月相关数据

    '''
    从工程质量验收中提取4个特征，包括两个合格率和两个优秀率
    '''
    def getQualityAcception(self, date, v):
        start = len(v)
        sheetType = TableType.qualityAcception
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 6) # 合格率
                v.append(float(value1.replace('%', '')) if value1 else 0)

                value2 = sheet.cell_value(i, 7) # 优秀率
                v.append(float(value2.replace('%', '')) if value2 else 0)
                
                if len(v)-start >= 4: return

        if len(v)-start < 4:
            for i in range(4-len(v)+start): v.append(0) # 用0补足当月数据

class progressModel:
    def __init__(self):

        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-12', '2019-01', '2019-02', '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']
        self.process = [] # 存储对应月的进度均值
        self.procedureAvg = []

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量

        self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
        self.regressionModel()
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getProcessData(date, v) # 获取工程完成情况数据特征向量

            self.result.append(v)

        self.dataProcess()

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
            print(self.month[i], self.result[i])

    def ruleModel(self):
        for i in range(len(self.month)):
            _s = 0
            _n = int(len(self.result[i])/3)
            for k in range(_n):
                _s += self.result[i][k*3-1]
            if _n is not 0:
                self.process.append(_s/_n)
            else:
                self.process.append(0)
        
        # print(self.process)
        
        for i in range(len(self.month)):
            # print("\n")
            # print("report date:", self.month[i])
            # avg = float(self.process[i])-float(self.process[i-1]) if i > 0 else float(self.process[i])
            if float(self.process[i]) != 0:
                avg = float(self.process[i])
            elif i > 0 and i < len(self.month)-1: 
                avg = (float(self.process[i-1]) + float(self.process[i+1])) / 2
            else:
                avg = float(self.process[i+1])
            self.procedureAvg.append(avg)
            # print("总进度均值：", avg)

    def regressionModel(self):
        # 自变量
        X = self.result
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
    从工程完成情况中提取3*n个特征，包括设计量，累计完成，累计完成比例
    '''
    def getProcessData(self, date, v):
        sheetType = TableType.constructionComplete
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 4) # 设计量
                if value1 and self.isFloatNum(value1):
                    v.append(int(float(value1)) if value1 and self.isFloatNum(value1) else 0)

                    value2 = sheet.cell_value(i, 5) # 累计完成
                    v.append(int(float(value2)) if value2 and self.isFloatNum(value2) else 0)
                    
                    value3 = sheet.cell_value(i, 6) # 累计完成比例
                    v.append(int(float(value3.replace('%', ''))) if value3 and self.isFloatNum(value3.replace('%', '')) else 0)
                else:
                    value1 = sheet.cell_value(i, 5) # 设计量
                    v.append(int(float(value1)) if value1 and self.isFloatNum(value1) else 0)

                    value2 = sheet.cell_value(i, 6) # 累计完成
                    v.append(int(float(value2)) if value2 and self.isFloatNum(value2) else 0)
                    
                    value3 = sheet.cell_value(i, 7) # 累计完成比例
                    v.append(int(float(value3.replace('%', ''))) if value3 and self.isFloatNum(value3.replace('%', '')) else 0)

class economicModel:
    def __init__(self):

        # month存储月报日期，对应的特征存在result相同index位置
        self.result = []
        self.month = ['2018-12', '2019-01', '2019-02', '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']
        self.economics = [] # 存储对应月的技经均值
        self.procedureAvg = []

        self.excel = xlrd.open_workbook(pathDir+excelFile, formatting_info=True) # 打开对应表格

        self.getVector() # 获取特征向量

        self.ruleModel() # 规则模型
        # self.preProcess() # 数据归一化预处理
        # self.KMeansModel() # KMeans聚类
        self.regressionModel()
    
    def getVector(self):
        self.result = []
        for date in self.month:
            v = []
            self.getContractData(date, v) # 获取合同管理及投资控制数据特征向量

            self.result.append(v)

        self.dataProcess()
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
    从合同管理中提取3*n个特征，包括本期结算，累计结算，结算率
    '''
    def getContractData(self, date, v):
        sheetType = TableType.contractManagement
        nrows, ncols = getSheetRowsCols(sheets[sheetType.value], self.excel)
        sheet = self.excel.sheet_by_index(sheetType.value)

        for i in range(1, nrows):
            if sheet.cell_value(i, 0) == date:
                value1 = sheet.cell_value(i, 2) # 本期结算
                v.append(int(float(value1)) if value1 and self.isFloatNum(value1) else 0)

                value2 = sheet.cell_value(i, 3) # 累计结算
                v.append(int(float(value2)) if value2 and self.isFloatNum(value2) else 0)
                
                value3 = sheet.cell_value(i, 4) # 累计结算比例
                v.append(int(float(value3.replace('%', ''))) if value3 and self.isFloatNum(value3.replace('%', '')) else 0)


def main():
    # m_secure = secureModel()
    # m_quality = qualityModel()
    # m_process = progressModel()
    m_economic = economicModel()

if __name__ == "__main__":
    main()
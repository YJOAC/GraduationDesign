import pickle

import jieba
import matplotlib.pyplot as plt
import xlrd
from gensim.models import Doc2Vec
from gensim.models.doc2vec import TaggedDocument
from sklearn.cluster import KMeans
# from sklearn.cluster import DBSCAN
from sklearn.decomposition import PCA
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score

from collections import Counter

dates = ['2018-01', '2018-02', '2018-03', '2018-04', '2018-05', '2018-06', '2018-07',
            '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03',
            '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
            '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']

def get_data(sheetnames):
    workbook = xlrd.open_workbook_xls('result.xls')
    data = {}
    meta = []
    for sheet_name in sheetnames:
        sheet = workbook.sheet_by_name(sheet_name)
        for col in range(2, sheet.ncols):
            item = sheet.cell_value(0, col).replace('\n', '\\n')
            meta.append(f'{sheet_name}::{item}')
        for row in range(1, sheet.nrows):
            date = sheet.cell_value(row, 0)
            if date not in dates: continue
            if date not in data:
                data[date] = {}
            for col in range(2, sheet.ncols):
                item = sheet.cell_value(0, col).replace('\n', '\\n')
                k = f'{sheet_name}::{item}'
                if k not in data[date]:
                    data[date][k] = []
                data[date][k].append(sheet.cell_value(row, col).replace('\n', '\\n'))
    with open('data.pkl', 'wb') as out:
        pickle.dump({
            'meta': meta,
            'data': data
        }, out)


def analyse_data():
    with open('data.pkl', 'rb') as fin:
        pkl = pickle.load(fin)
    meta, data = pkl['meta'], pkl['data']
    documents = []
    for date in data:
        words = []
        print(date, end='\t')
        for m in meta:
            words.append('，'.join(data[date].get(m, [])))
        words = '。'.join(words)
        words = jieba.lcut(words)
        documents.append(TaggedDocument(words, [date]))
    model = Doc2Vec(documents)
    date_vec = {}
    for date in data:
        date_vec[date] = model.docvecs[date]
    with open('vec.pkl', 'wb') as out:
        pickle.dump(date_vec, out)


scores_secure = {
    '2018-12': '95',
    '2018-02': '95',
    '2018-03': '90',
    '2018-04': '93',
    '2018-05': '93',
    '2018-06': '93',
    '2019-03': '92',
    '2019-04': '87',
    '2018-01': '98',
    '2018-10': '83',
    '2018-11': '93',
    '2018-07': '93',
    '2018-08': '93',
    '2018-09': '93',
    '2019-01': '92',
    '2019-02': '95',
    '2019-05': '87',
    '2019-08': '85',
    '2019-09': '85',
    '2019-10': '90',
    '2019-11': '85',
    '2019-12': '85',
    '2019-06': '92',
    '2019-07': '92',
    '2020-01': '90',
    '2020-03': '90',
    '2020-04': '90',
    '2020-02': '75'
}

scores_quality = {
    '2018-01': '90',
    '2018-02': '82',
    '2018-03': '92',
    '2018-04': '90',
    '2018-05': '90',
    '2018-06': '90',
    '2018-07': '95',
    '2018-08': '87',
    '2018-09': '85',
    '2018-10': '88',
    '2018-11': '82',
    '2018-12': '90',
    '2019-01': '87',
    '2019-02': '84',
    '2019-03': '87',
    '2019-04': '88',
    '2019-05': '80',
    '2019-06': '85',
    '2019-07': '88',
    '2019-08': '90',
    '2019-09': '90',
    '2019-10': '90',
    '2019-11': '90',
    '2019-12': '90',
    '2020-01': '89',
    '2020-02': '89',
    '2020-03': '89',
    '2020-04': '84',
}


def analyse_result(scores):
    with open('vec.pkl', 'rb') as fin:
        date_vec = pickle.load(fin)
    labels, data = zip(*date_vec.items())
    pca = PCA(n_components=2)
    x_pca = pca.fit_transform(data)
    k = 5
    k_model = KMeans(n_clusters=k)
    k_model.fit(data)
    # k_model = DBSCAN(eps=0.24, min_samples=2)
    # k_model.fit(x_pca)
    # k = len(set(k_model.labels_))
    groups = [[] for _ in range(k)]
    color_list = ['grey', 'darkviolet', 'r', 'g', 'b', 'c', 'm', 'y', 'k', 'darkorange',
                  'lightgreen', 'gold', 'turquoise', 'plum', 'tan', 'khaki', 'pink', 'skyblue', 'lawngreen', 'salmon']
    colors, xs, ys = [], [], []
    for i, l in enumerate(k_model.labels_):
        groups[l].append(labels[i])
        colors.append(color_list[l % len(color_list)])
        xs.append(x_pca[i][0])
        ys.append(x_pca[i][1])
    for group in groups:
        print(sorted(group))
    plt.colormaps()
    plt.scatter(xs, ys, c=colors, s=10)
    for i, (x, y) in enumerate(zip(xs, ys)):
        plt.text(x - 0.17, y + 0.015, f'{labels[i]}({scores[labels[i]]})',
                 fontdict={'size': 8, 'color': color_list[k_model.labels_[i] % len(color_list)]})
    plt.xticks([])
    plt.yticks([])
    plt.axis('off')
    plt.show()

    # with open('data.pkl', 'rb') as fin:
    #     pkl = pickle.load(fin)
    # meta, data = pkl['meta'], pkl['data']
    # for group in groups:
    #     print('=' * 80)
    #     for date in group:
    #         words = []
    #         print(date, end='\t')
    #         for m in meta:
    #             words.append(' '.join(data[date].get(m, [])))
    #         words = ' '.join(words)
    #         counter = Counter(jieba.lcut(words))
    #         items = sorted(counter.items(), key=lambda x: x[1], reverse=True)
    #         print([x[0] for x in items])
    #     print('=' * 80)
    #     print()


def try_clf(scores, method='RF'):
    n_class = 3
    with open('vec.pkl', 'rb') as fin:
        date_vec = pickle.load(fin)
    if method == 'LR':
        clf = LogisticRegression(C=1, random_state=1, solver='liblinear')
    elif method == 'KNN':
        clf = KNeighborsClassifier(n_neighbors=2)
    elif method == 'SVM':
        clf = SVC(C=3.0)
    elif method == 'DT':
        clf = DecisionTreeClassifier()
    elif method == 'RF':
        clf = RandomForestClassifier(n_estimators=50, max_features="auto", max_depth=2,
                                     min_samples_split=3, bootstrap=True, random_state=0)
    else:
        raise AssertionError(f'Unknown method "{method}".')
    items = sorted(scores.items(), key=lambda x: int(x[1]))
    step = (len(items) - 1) // n_class + 1
    data = []
    for i, (date, _) in enumerate(items):
        data.append([date, i // step])
    items = sorted(data, key=lambda x: x[0])
    # items = sorted(scores.items(), key=lambda x: x[0])
    # max_score = max(map(int, scores.values()))
    # min_score = min(map(int, scores.values()))
    # step = (max_score - min_score) // (n_class - 1)
    n_train = 20
    train_items = items[:n_train]
    test_items = items[n_train:]
    # train_x, train_y = [], []
    # for date, score in train_items:
    #     train_x.append(date_vec[date])
    #     train_y.append((int(score) - min_score) // step)
    # test_x, test_y = [], []
    # for date, score in test_items:
    #     test_x.append(date_vec[date])
    #     test_y.append((int(score) - min_score) // step)
    train_x, train_y = zip(*train_items)
    train_x = [date_vec[x] for x in train_x]
    test_x, test_y = zip(*test_items)
    test_x = [date_vec[x] for x in test_x]
    # print(train_y)
    clf.fit(train_x, train_y)
    test_p = clf.predict(test_x)
    print(test_y)
    print(test_p.tolist())
    accuracy = accuracy_score(test_y, test_p)
    print(f'[{method}] Accuracy: {accuracy}')


def main():
    # get_data(["安全培训数据", "危险源管理情况统计表", "安全隐患排查治理一览表", "违章隐患统计表"])
    get_data(["本月锚杆监理检测结果一览表", "工序质量验评表", "本月爆破振动检测成果表", "各标段检验批验收统计"])
    # get_data(['工程完成情况'])
    # get_data(['合同管理及投资控制'])
    analyse_data()
    # analyse_result(scores_secure)
    analyse_result(scores_quality)
    # analyse_result(scores_process)
    # analyse_result(scores_economics)
    for method in ['LR', 'KNN', 'SVM', 'DT', 'RF']:
        # try_clf(scores_secure, method)
        try_clf(scores_quality, method)
        # try_clf(scores_process, method)
        # try_clf(scores_economics, method)
    # try_clf()


if __name__ == '__main__':
    main()

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

dates = ['2018-12', '2019-01', '2019-02', '2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11',
                '2019-12', '2020-01', '2020-02', '2020-03', '2020-04']

def get_data(sheetnames):
    workbook = xlrd.open_workbook_xls('result_ln.xls')
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
                data[date][k].append(str(sheet.cell_value(row, col)).replace('\n', '\\n'))
    with open('data_ln.pkl', 'wb') as out:
        pickle.dump({
            'meta': meta,
            'data': data
        }, out)


def analyse_data():
    with open('data_ln.pkl', 'rb') as fin:
        pkl = pickle.load(fin)
    meta, data = pkl['meta'], pkl['data']
    # print(data)
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
    with open('vec_ln.pkl', 'wb') as out:
        pickle.dump(date_vec, out)


scores_secure = {
    '2018-12': '95',
    '2019-04': '90',
    '2019-01': '92',
    '2019-02': '95',
    '2019-05': '92',
    '2019-08': '95',
    '2019-09': '98',
    '2019-10': '98',
    '2019-11': '100',
    '2019-12': '98',
    '2019-06': '90',
    '2019-07': '92',
    '2020-01': '98',
    '2020-03': '88',
    '2020-04': '98',
    '2020-02': '95'
}

scores_quality = {
    '2018-12': '98',
    '2019-04': '95',
    '2019-01': '95',
    '2019-02': '92',
    '2019-05': '85',
    '2019-08': '89',
    '2019-09': '94',
    '2019-10': '92',
    '2019-11': '94',
    '2019-12': '92',
    '2019-06': '87',
    '2019-07': '89',
    '2020-01': '89',
    '2020-03': '89',
    '2020-04': '94',
    '2020-02': '84'
}

scores_process = {
    '2018-12': '7',
    '2019-04': '16',
    '2019-01': '16',
    '2019-02': '32',
    '2019-05': '64',
    '2019-08': '18',
    '2019-09': '53',
    '2019-10': '55',
    '2019-11': '48',
    '2019-12': '63',
    '2019-06': '63',
    '2019-07': '59',
    '2020-01': '49',
    '2020-03': '50',
    '2020-04': '51',
    '2020-02': '49'
}

scores_economics = {
    '2018-12': '7',
    '2019-04': '6',
    '2019-01': '8',
    '2019-02': '32',
    '2019-05': '6',
    '2019-08': '16',
    '2019-09': '16',
    '2019-10': '17',
    '2019-11': '18',
    '2019-12': '21',
    '2019-06': '9',
    '2019-07': '14',
    '2020-01': '20',
    '2020-03': '20',
    '2020-04': '20',
    '2020-02': '20'
}


def analyse_result(scores):
    with open('vec_ln.pkl', 'rb') as fin:
        date_vec = pickle.load(fin)
    labels, data = zip(*date_vec.items())
    pca = PCA(n_components=2)
    x_pca = pca.fit_transform(data)
    k = 3
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
    # for group in groups:
    #     print(sorted(group))
    plt.colormaps()
    plt.scatter(xs, ys, c=colors, s=10)
    for i, (x, y) in enumerate(zip(xs, ys)):
        plt.text(x, y, f'{labels[i]}({scores[labels[i]]})',
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
    with open('vec_ln.pkl', 'rb') as fin:
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
    n_train = int(0.7 * len(items))
    # print(items)
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
    # get_data(['安全文明施工'])
    # get_data(['施工质量控制', '工程质量验收'])
    # get_data(['工程完成情况'])
    get_data(['合同管理及投资控制'])
    analyse_data()
    # analyse_result(scores_secure)
    # analyse_result(scores_quality)
    # analyse_result(scores_process)
    analyse_result(scores_economics)
    for method in ['LR', 'KNN', 'SVM', 'DT', 'RF']:
        # try_clf(scores_secure, method)
        # try_clf(scores_quality, method)
        # try_clf(scores_process, method)
        try_clf(scores_economics, method)
    # try_clf()


if __name__ == '__main__':
    main()

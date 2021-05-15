from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score
from collections import Counter
import random

test_data = [[17, 2, 4, 3, 1], [11, 2, 4, 4, 1], [6, 4, 3, 2, 1], [2, 4, 4, 6, 1], [8, 4, 4, 8, 1], [2, 4, 3, 4, 1],
             [7, 4, 4, 6, 1], [7, 9, 4, 5, 1], [6, 10, 4, 1, 1], [4, 6, 4, 1, 1], [5, 14, 4, 3, 1], [5, 9, 4, 1, 1],
             [21, 9, 4, 2, 1], [0, 3, 0, 0, 1], [5, 9, 3, 1, 1], [3, 8, 4, 6, 1]]


def get_score(xs):
    score = 100
    if xs[0] <= 2:
        score -= 2
    if xs[1] <= 4:
        score -= 5
    elif xs[1] <= 10:
        score -= 2
    if xs[2] < 3:
        score -= 5
    if xs[3] > 10:
        score -= 10
    elif xs[3] > 3:
        score -= 3
    if xs[-1] != 1:
        score -= 10
    return score


def try_clf(train_x, train_y, test_x, test_y, method='RF'):
    if method == 'LR':
        clf = LogisticRegression(C=1, random_state=1, solver='liblinear')
    elif method == 'KNN':
        clf = KNeighborsClassifier(n_neighbors=2)
    elif method == 'SVM':
        clf = SVC(C=3.0)
    elif method == 'DT':
        clf = DecisionTreeClassifier()
    elif method == 'RF':
        # clf = RandomForestClassifier(n_estimators=50, max_features="auto", max_depth=2,
        #                              min_samples_split=3, bootstrap=True, random_state=0)
        clf = RandomForestClassifier()
    else:
        raise AssertionError(f'Unknown method "{method}".')
    clf.fit(train_x, train_y)
    test_p = clf.predict(test_x)
    # print(test_y)
    # print(test_p.tolist())
    accuracy = accuracy_score(test_y, test_p)
    print(f'[{method}] Accuracy: {accuracy}')


def get_clf_func(n_class, scores):
    scores.sort()
    n_threshold = n_class - 1
    step = (len(scores) + 1) // n_class
    thresholds = []
    for i in range(n_threshold):
        thresholds.append(scores[(i + 1) * step - 1])

    def wrapper(x):
        for j in range(len(thresholds)):
            if x <= thresholds[j]:
                return j
        return len(thresholds)
    return wrapper


def main():
    scores = [get_score(data) for data in test_data]
    clf_func = get_clf_func(n_class=5, scores=scores)
    test_label = [clf_func(get_score(data)) for data in test_data]
    temp_data = zip(*test_data)
    ranges = [(min(data), max(data)) for data in temp_data]
    train_data = [[random.randint(ranges[i][0], ranges[i][1]) for i in range(len(ranges))] for _ in range(10000)]
    train_label = [clf_func(get_score(data)) for data in train_data]
    for method in ['LR', 'KNN', 'SVM', 'DT', 'RF']:
        try_clf(train_data, train_label, test_data, test_label, method)


if __name__ == '__main__':
    main()

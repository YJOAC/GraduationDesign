from prettytable import PrettyTable
from pyhanlp import *

class HpWord:
    def __init__(self, wid, text, rel, postag, head):
        self.id = wid
        self.text = text
        self.rel = rel
        self.postag = postag
        self.head = head

    def __repr__(self):
        return f'<id::{self.id}, text::{self.text}, rel::{self.rel}, postag::{self.postag}, head::{self.head}>'

class HpParser:
    def __init__(self, s):
        des = HanLP.parseDependency(s)
        self.words = []
        self.root = None
        # print(des)
        for i in des:
            self.words.append(HpWord(i.ID - 1, i.LEMMA, i.DEPREL, i.POSTAG, i.HEAD.ID - 1))
            if self.words[-1].rel == '核心关系':
                self.root = self.words[-1]

    def __repr__(self):
        return '\n'.join(map(str, self.words))

    def get_subject(self, verb_word):
        for word in self.words:
            if word.rel == '主谓关系' and word.head == verb_word.id:
                return word
        return None

    def get_object(self, verb_word):
        for word in self.words:
            if word.rel == '动宾关系' and word.head == verb_word.id:
                return word
        return None

    def get_roots(self):
        r = [self.root]
        rid_set = set()
        rid_set.add(self.root.id)
        for word in self.words:
            if word.rel == '并列关系' and word.head in rid_set:
                r.append(word)
                rid_set.add(word.id)
        return r

    def span_context(self, c_word):
        no_rel_set = {'动宾关系', '主谓关系', '并列关系'}

        def _span(_c):
            _min, _max = _c.id, _c.id
            for _w in self.words:
                if _w.head == _c.id and _w.rel not in no_rel_set:
                    if _w.rel == '标点符号':
                        break
                    __min, __max = _span(_w)
                    _min = min(_min, __min)
                    _max = max(_max, __max)
            return _min, _max

        min_id, max_id = _span(c_word)
        return ''.join(word.text for word in self.words[min_id:max_id + 1])

    def parse(self):
        roots = self.get_roots()
        tb = PrettyTable()
        tb.field_names = ['项目名称', '属性名称', '属性值']
        # print('主语:', self.get_subject(self.root))
        for i, root in enumerate(roots):
            # print('\t谓语:', root)
            # print('\t宾语:', self.get_object(root))
            # print()
            row = [self.span_context(self.get_subject(self.root))] if i == (len(roots) - 1) // 2 else ['']
            tb.add_row(row + [self.span_context(root), self.span_context(self.get_object(root))])
        return tb

def main():
    s = '1#排水廊道开挖进尺35.2m，累计完成进尺124.2m，占设计总长度25.48%。'
    parser = HpParser(s)
    result = parser.parse()
    print('文本:', s)
    print('分析结果如下:')
    print(result)

if __name__ == '__main__':
    main()

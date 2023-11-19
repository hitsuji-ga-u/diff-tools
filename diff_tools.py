import abc
from enum import Enum
import csv
from typing import Any
import openpyxl as px
from pprint  import pprint


class SignalExistError(Exception):
    def __init__(self, message='     Sinals already exists.   '):
        self.message = message
        super().__init__(self.message)


class Diff(Enum):
    SAME = 0
    ADD = 1
    DEL = 2
    CHANGE = 3


class Comparison(object, metaclass=abc.ABCMeta):
    def add_attribute(self, name):
        self.properties.append(name)


class Signals():
    def __init__(self, key, properties):
        pass

def cast2numeric(e_str):
    if any(c.isalpha() for c in e_str):
        return e_str

    # if e_str doesn't have any str, Suppose it's numeric.
    if '.' in e_str:
        return float(e_str)

    return int(e_str)


class Signals(object):

    def __init__(self, name, key, attributes, datas):
        self.identifer = str(name)
        self.key = key
        self.key_pos = 0
        self.attributes = [SignalAttribute(attr) for attr in attributes]
        self.signals = []

        if key not in attributes:
            raise ValueError(' cannot find key value')

        self.key_pos = attributes.index(key)

        Signal.init_sgnals()

        for data in datas:
            signal = Signal(key, data[self.key_pos])

            for idx, attr in enumerate(attributes):
                if attr == key:
                    continue
                signal.add_attribute(SignalAttribute(attr, data[idx]))
            self.signals.append(signal)

    def __getitem__(self, sgnl):
        idx = self.item_list_str().index(sgnl.name)
        return self.signals[idx]

    def __gt__(self, other):
        return len(self.signals) > len(other.signals)

    def __str__(self):
        ret_str = f'{self.identifer}\n'
        for sgnl in self.signals:
            ret_str += f'{str(sgnl)}\n'
        return ret_str

    def has(self, item):
        return item in self.signals

    def get_attr_like(self, attribute):
        print(f'        Signals > get_attr_like  {self.identifer}')
        print(f'        searching .... {attribute}')
        for attr in self.attributes:
            print(f'        found {attr}')
            if attr == attribute:
                print(f'        {attribute} == {attr}')
                return attr
        else:
            raise ValueError


    def item_list(self):
        Signal.init_sgnals()
        return {Signal(self.key, sg.value) for sg in self.signals}

    def item_list_str(self):
        return [item.name for item in self.signals]

    def attribute_set(self):
        return set(self.attributes)



class Attribute(object, metaclass=abc.ABCMeta):
    def __init__(self, name, value=''):
        self.name = name
        self.value = value


    def __eq__(self, other):
        pass

    def add_attribute(self, attribute):
        raise ValueError


class Signal(Attribute):

    _values = {}

    def __init__(self, key, value):
        super().__init__(key, value)
        self.attributes = []

    def __new__(cls, key, value):
        if value in cls._values:
            raise SignalExistError(f'Signal({key}, {value}) already exists.')

        cls._values[value] = super().__new__(cls)

        return cls._values[value]

    def __eq__(self, other):
        if len(self.attributes) != len(other.attributes):
            return False
        for my_attr, your_attr in zip(self.attributes, other.attributes):
            if my_attr != your_attr:
                return False
        return True

    def __gt__(self, other):
        return self.value > other.value

    def __str__(self):
        out = [f'{self.name}: {self.value}']
        out.extend([str(attr) for attr in self.attributes])
        return ',\t'.join(out)

    def __repr__(self):
        return f'Signal(\'{self.name}\', {self.value})'

    def __hash__(self):
        return hash(self.value)

    def add_attribute(self, attribute):
        self.attributes.append(attribute)

    @classmethod
    def init_sgnals(cls):
        cls._values = {}


class SignalAttribute(Attribute):

    attribute_groups = [
        {'name', 'Name'},
    ]

    def __eq__(self, other):
        return (self.value == other.value) and (self.same_name_with(other.name))

    def __str__(self):
        return f'{self.name}: {self.value}'

    def __repr__(self):
        return f'SignalAttribute(\'{self.name}\')'

    def __hash__(self):
        for group in self.attribute_groups:
            if self.name in group:
                return hash(frozenset(group))

        return hash(self.name)

    def same_name_with(self, name):
        print(f'            same_name_with      {self.name}, {name}.   {self.name == name}')
        if self.name == name:
            return True

        for group in self.attribute_groups:
            if self.name in group:
                if name in group:
                    return True

        return False



class SignalDiffTool(object):
    def __init__(self) -> None:
        self.comparisons = []
        self.diffs = []
        self.items = []

    def add_comparison(self, name, data, key):
        '''比較対象を生成する。'''

        table = data.split('\n')
        attributes = [e.strip() for e in table[0].split(',')]

        datas = []
        for sgnl in table[1:]:
            datas.append([cast2numeric(e.strip()) for e in sgnl.split(',')])

        self.comparisons.append(Signals(name, key, attributes, datas))
        return self 

    def compare(self):
        """itemの比較を行う。全てのitemに対して、新規、削除、変更を定義する。"""
        all_items_set = set()
        for comparison in self.comparisons:
            all_items_set |= comparison.item_list()
        all_items = sorted(list(all_items_set))

        all_attributes = []
        for comparison in self.comparisons:
            for attr in comparison.attributes:
                if not attr in set(all_attributes):
                    all_attributes.append(attr)
        print(all_attributes)

        key = self.comparisons[0].key
        datas = []
        for item in all_items:
            print(item)
            data = []
            for attr in all_attributes:
                if attr.name == key:
                    data.append(item.value)
                    continue

                diff = Diff.SAME
                for idx in range(len(self.comparisons) - 1):
                    before = self.comparisons[idx]
                    after = self.comparisons[idx+1]
                    if not before.has(item) and after.has(item):
                        diff = Diff.ADD
                    elif before.has(item) and not after.has(item):
                        diff = Diff.DEL
                    else:
                        if before[item] != after[item]:
                            diff = Diff.CHANGE
                data.append(diff)
            datas.append(data)

        attributes_str = [attr.name for attr in all_attributes]
        self.diffs = Signals('diff', self.comparisons[0].key, attributes_str, datas)

    def export_csv(self):
        print('SignalDiffTools > export_csv')
        out = ''
        k_pos = self.diffs.key_pos
        attrs = [self.diffs.attributes[k_pos]]
        print(self.diffs.attributes)

        for idx, attr in enumerate(self.diffs.attributes):
            if idx == k_pos:
                continue
            print('')
            print('')
            print(f'    search {attr}')

            for c in self.comparisons:
                print(f'    search at {c.identifer}')
                if not attr in c.attribute_set():
                    attrs.append(SignalAttribute('-'))
                else:
                    attrs.append(c.get_attr_like(attr))
        print(attrs)
        attrs = [self.diffs.attributes[k_pos]] + []
        self.diffs.attributes[:k_pos] + self.diffs.attributes[k_pos+1:]
        attrs_str = [attr.name for attr in self.diffs.attributes]
        print(', '.join([self.diffs.key] + attrs_str[:k_pos]  + attrs_str[k_pos+1:]))

        print()

        return out


def excel2csv(path):
    '''
    エクセルの1枚目のシートをcsv形式に変換

    Args: path
    - path (str): エクセルのパス

    Returns:
    - str: エクセルに入力されている値。csv形式。
    '''

    try:
        wb = px.load_workbook(path)
    except PermissionError:
        print(f' please  close the {path}') 
        return
    ws = wb.worksheets[0]

    csv = []
    for r in ws.iter_rows(min_row=ws.min_row, min_col=ws.min_column, values_only=True):
        csv.append(list(map(lambda x: str(x).strip() if not x is None else '', r)))

    no_empty_value_csv = [r for r in csv if any(r)]

    return '\n'.join(', '.join(row) for row in no_empty_value_csv)


class ExcelWriter(object):

    def __init__(self, path) -> None:
        self.path = path

    def write(self, txt):
        print(txt)


def main(*name_data_keys):


    tm = SignalDiffTool()

    for n, d, k in name_data_keys:
        tm.add_comparison(n, excel2csv(d), key=k)


    tm.compare()

    writer = ExcelWriter('diff.excl')
    writer.write(tm.export_csv())




# オブジェクトを作成させる


# オブジェクトの比較をする


# オブジェクトの比較結果を出力する

if __name__ == "__main__":


    data_path1 = 'diff-tools/sample/Signals_ver1.xlsx'
    data_path2 = 'diff-tools/sample/Signals_ver2.xlsx'


    main((1, data_path1, 'id'), (2, data_path2, 'id'))




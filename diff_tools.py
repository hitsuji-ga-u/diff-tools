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


class Attribute(object):

    attribute_group = [
        {'min', '最小値'}
    ]

    def __init__(self, name, value=''):
        self.name = name
        self.value = value

    def set(self, value):
        return Attribute(self.name, value)


    def __repr__(self):
        return f'Attribute(\'{self.name}\', \'{self.value}\')'

    def __eq__(self, other):
        if self.value != other.value:
            return False

        if self.name.lower() == other.name.lower():
            return True

        for group in self.attribute_group:
            if not self.name in group:
                continue

            if other.name in group:
                return True

        return False

class UndifinedAttribute(Attribute):

    def __init__(self, name):
        super().__init__('Undifined')

    def __bool__(self):
        return False

    def set(self, value):
        return Attribute(self.name, '-')


class Recode(object):

    def __init__(self, key, key_value, attributes):
        self.key = key
        self.key_value = key_value
        self.attributes = attributes

    def __repr__(self):
        return f'Recode(\'{self.key}\'=\'{self.key_value}\')'



class Table(object):

    def __init__(self, name, datas, key=0):
        self.name = name
        self.recode_name_list = []

        if isinstance(key, int):
            self.key_idx = key
            self.key = datas[0][key]
        elif isinstance(key, str):
            self.key_idx = datas[0].index(key)
            self.key = key

        key_attr = [Attribute(self.key)]
        other_attr = [Attribute(attr_name) for attr_name in datas[0][:self.key_idx] + datas[0][self.key_idx+1:]]
        self.attributes = key_attr + other_attr

        self.recodes = []
        for idx, rec_data in enumerate(datas[1:]):
            if len(rec_data) != len(self.attributes):
                raise ValueError(f' not match the length of {idx=} recode and attributes ({len(self.attributes)}) ')

            recode_name = rec_data[self.key_idx]
            if not recode_name:
                raise ValueError(f'\n\n\t the key of line {idx} recode is None. \n')
            if recode_name in self.recode_name_list:
                raise SignalExistError(f'{recode_name} Recode already exists.')
            self.recode_name_list.append(recode_name)
            self.recodes.append(Recode(self.key, recode_name,
                                       [attr.set(data) for attr, data in zip(self.attributes, rec_data)]))

    def __repr__(self):
        return f'Table(\'{self.name}\')'

    def get_attr(self, attr):
        if not attr in self.attributes:
            return UndifinedAttribute('Undifined')

        return self.attributes[self.attributes.index(attr)]



class ExcelTool(object):

    def __init__(self, path, sheet_name=None):
        self.path = path
        self.name = sheet_name

        try:
            wb = px.load_workbook(path)
        except PermissionError:
            raise PermissionError(f'\n\n\t please  close the {path}. \n') 
        if sheet_name is None:
            self.ws = wb.worksheets[0]
        else:
            self.ws = wb[sheet_name]

    def to_list(self, min_row=None, max_row=None, min_col=None, max_col=None):
        """セルの値をリストにして出力。値の先頭末尾の空白は削除。空白の場合は空文字。"""
        min_row = self.ws.min_row if min_row is None else min_row
        min_col = self.ws.min_column if min_col is None else min_col
        max_row = self.ws.max_row if max_row is None else max_row
        max_col = self.ws.max_column if max_col is None else max_col

        out_list = []
        for r in self.ws.iter_rows(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col, values_only=True):
            out_list.append(list(map(lambda x: str(x).strip() if not x is None else '', r)))

        return out_list


class ExcelWriter(object):

    def __init__(self, path) -> None:
        self.path = path

    def write(self, txt):
        print(txt)





# オブジェクトを作成させる


# オブジェクトの比較をする


# オブジェクトの比較結果を出力する

if __name__ == "__main__":


    data_path1 = 'diff-tools/sample/Signals_ver1.xlsx'
    data_path2 = 'diff-tools/sample/Signals_ver2.xlsx'

    datas1 = ExcelTool(data_path1).to_list()
    datas2 = ExcelTool(data_path2).to_list(max_row=13)
    t1 = Table(1, datas1)
    t2 = Table(2, datas2)

    all_tables = [t1, t2]
    all_attributes = all_tables[0].attributes.copy()
    for t in all_tables[1:]:
        for attr in t.attributes:
            if not attr in all_attributes:
                all_attributes.append(attr)

    all_recodes = all_tables[0].recode_name_list.copy()
    for t in all_tables[1:]:
        for recode_name in t.recode_name_list:
            if not recode_name in all_recodes:
                all_recodes.append(recode_name)

    comparison_each_attr = [(t, t.get_attr(attr)) for attr in all_attributes[1:] for t in all_tables]
    comparison_attrs_names = [f'{table.name}: {attr.name}' for table, attr in comparison_each_attr]
    print(comparison_each_attr)
    print(comparison_attrs_names)


 



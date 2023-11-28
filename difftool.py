from enum import Enum
from pathlib import Path


class SignalExistError(Exception):
    def __init__(self, message='     Sinals already exists.   '):
        self.message = message
        super().__init__(self.message)


class Judgement(Enum):
    NOJUDGEMENT = 0
    SAME = 1
    ADD = 2
    DEL = 3
    CHANGE = 4



class Attribute(object):

    attribute_group = [
        {'min', '最小値'},
        {'Name', 'Signal Name'},
        {'ID', 'CAN ID'},
        {'Cycle Time [ms]', 'Cycl time[ms]'},
        {'Signal Description', 'Value table'},
        {'Size', 'Size[Bit]'},
        {'TX', 'Transmitter'}
    ]

    def __init__(self, name, value=''):
        self.name = name
        self.value = value

    def __repr__(self):
        return f'Attribute(\'{self.name}\', \'{self.value}\')'

    def __eq__(self, other):
        return self.value == other.value and self.same_name(other)

    def set(self, value):
        return Attribute(self.name, value)

    def same_name(self, attr):
        if self.name.lower() == attr.name.lower():
            return True

        for group in self.attribute_group:
            if not self.name in group:
                continue
            if attr.name in group:
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

    def __init__(self, key_idx, attributes):
        self.key = attributes[key_idx]
        self.attributes = attributes


    def __repr__(self):
        return f'Recode({self.attributes.index(self.key)}, {self.attributes})'

    def get_value(self, attr):
        for my_attr in self.attributes:
            if my_attr.same_name(attr):
                return my_attr.value
        else:
            return ''


class Table(object):

    def __init__(self, name, datas, key=0):
        self.name = name
        self.recodes = []
        self.attributes = []
        self.keys_list = []


        if isinstance(key, int):
            key_idx = key
        elif isinstance(key, str):
            key_idx = datas[0].index(key)

        self.attributes = [Attribute(attr_name) for attr_name in datas[0]]
        self.key = self.attributes[key_idx]

        for idx, rec_data in enumerate(datas[1:]):
            if len(rec_data) != len(self.attributes):
                raise ValueError(f' not match the number of {idx=} recode attributes ({len(rec_data)}) and the number of table attributes ({len(self.attributes)}) ')

            recode_name = rec_data[key_idx]
            if not recode_name:
                raise ValueError(f'\n\n\t the key of line {idx} recode is None. \n')
            if self.has_recode(recode_name):
                raise SignalExistError(f'{recode_name} Recode already exists.')
            attributes = [attr.set(data) for attr, data in zip(self.attributes, rec_data)]
            self.recodes.append(Recode(key_idx, attributes))
            self.keys_list.append(attributes[key_idx])


    def __repr__(self):
        return f'Table(\'{self.name}\')'

    def get_attr(self, attr):
        if not attr in self.attributes:
            return UndifinedAttribute('Undifined')

        return self.attributes[self.attributes.index(attr)]

    def get_field(self, key, attr):
        attr_idx = self.attributes.index(attr)
        recode = self.recodes[self.keys_list.index(key)]
        return recode.attributes[attr_idx]

    def has_attr(self, attr):
        return attr in self.attributes

    def has_recode(self, recode):
        if isinstance(recode, str):
            return recode in [rec.get_value(rec.key) for rec in self.recodes]
        elif isinstance(recode, Attribute):
            return recode in self.keys_list
        else:
            # todo: 主キー名、主キー以外を渡されたときの対処
            raise ValueError()



class DiffTool(object):
    def __init__(self):
        self.tables = []
        self.all_keys = []
        self.key = None
        self.comparison_attributes = []

    def add_table(self, name, data_list, key=0):
        t = Table(name, data_list, key)
        self.tables.append(t)
        self.key = t.key
        for attr in t.attributes:
            if attr == self.key:
                continue
            if not attr in self.comparison_attributes:
                self.comparison_attributes.append(attr)

        for key in t.keys_list:
            if not key in self.all_keys:
                self.all_keys.append(key)

    def comapre(self):
        # Tableクラスに入れたいので、リストを作成する
        diff_datas = [[attr.name for attr in [self.key] + self.comparison_attributes]]
        for key in self.all_keys:
            diff_data = [key.value]
            for attr in self.comparison_attributes:
                result = Judgement.NOJUDGEMENT
                for i in range(len(self.tables) - 1):
                    t1 = self.tables[i]
                    t2 = self.tables[i+1]
                    # いずれかのテーブルで、判定したい属性が未定義ならその属性の判定は判定無しとする。
                    if not t1.has_attr(attr) or not t2.has_attr(attr):
                        result = Judgement.NOJUDGEMENT
                        continue

                    # まだ判定がされていない場合
                    if result == Judgement.NOJUDGEMENT:
                        if t1.has_recode(key) and not t2.has_recode(key):
                            result = Judgement.DEL
                        elif not t1.has_recode(key) and t2.has_recode(key):
                            result = Judgement.ADD
                        elif t1.get_field(key, attr) == t2.get_field(key, attr):
                            result = Judgement.SAME
                        else:
                            result = Judgement.CHANGE
                    # すでに判定されている（テーブルが3以上ある）場合
                    else:
                        if result == Judgement.SAME and t1.get_field(key, attr) == t2.get_field(key, attr):
                            result = Judgement.SAME
                        else:
                            result = Judgement.CHANGE
                diff_data.append(result.name)
            diff_datas.append(diff_data)

        self.tables.append(Table('diff', diff_datas))

    def out(self):
        all_datas = []
        all_attributes = [self.key.name]
        all_attributes.extend([f'{t.name}:\n{t.get_attr(attr).name}' for attr in self.comparison_attributes for t in self.tables])
        all_datas.append(all_attributes)

        # 出力用に全て文字列に
        for key in self.all_keys:
            data = [key.value]
            for comp_attr in self.comparison_attributes:
                for t in self.tables:
                    if not t.has_attr(comp_attr):
                        data.append('')
                        continue
                    if not t.has_recode(key):
                        data.append('-')
                        continue

                    data.append(t.get_field(key, comp_attr).value)
            all_datas.append(data)
        return all_datas

    def get_border_diff(self):
        # thin_line_col = [0]
        # thin_line_col.extend([c for c in range(1, len(all_tables)*(len(comparison_attributes)+1), len(all_tables))])
        return [0] + [c for c in range(1, len(self.tables)*(len(self.comparison_attributes)+1), len(self.tables))]

    def get_border_attr(self):
        attr_border = []
        for c in range(len(self.comparison_attributes)):
            for j in range(len(self.tables) - 1):
                attr_border.append(2 + c*len(self.tables) + j)
        return attr_border

    def get_border_keys(self):
        return [r for r in range(2, len(self.all_keys)+1)]

    def get_border_attr_and_keys(self):
        return  [0, 1, len(self.all_keys)+1]

    def get_attr_num(self):
        return len(self.comparison_attributes) * len(self.tables) + 1

    def get_keys_num(self):
        return len(self.all_keys)

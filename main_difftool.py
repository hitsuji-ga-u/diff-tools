from mylibs.excel import ExcelHelper
from difftool import DiffTool, Judgement

if __name__ == "__main__":

    data_path1 = 'sample/Signals_ver1.xlsx'
    data_path2 = 'sample/Signals_ver2.xlsx'
    out_path = 'sample/Signals_diff.xlsx'

    datas1 = ExcelHelper(data_path1).to_list(min_col=2, min_row=2, max_row=12, max_col=8)
    datas2 = ExcelHelper(data_path2).to_list(min_col=2, min_row=2, max_row=13, max_col=7)

    diff_tool = DiffTool()
    diff_tool.add_table(1, datas1)
    diff_tool.add_table(2, datas2)

    diff_tool.comapre()

    diff_data = diff_tool.out()

    out = ExcelHelper(out_path)

    out.write(diff_data)

    out.fill(value=Judgement.SAME.name, color='BDD7EE')
    out.fill(value=Judgement.ADD.name, color='C6E0B4')
    out.fill(value=Judgement.DEL.name, color='DBDBDB')
    out.fill(value=Judgement.CHANGE.name, color='FFE699')
    out.fill(value=Judgement.NOJUDGEMENT.name, color='C6ACD9')
    out.fill(row=[0], color='F8CBAD')

    thin_line_col = diff_tool.get_border_diff()
    dot_line_col = diff_tool.get_border_attr()
    thin_line_row = diff_tool.get_border_attr_and_keys()
    dot_line_row = diff_tool.get_border_keys()

    out.line(cols=thin_line_col, type='thin')
    out.line(cols=dot_line_col, type='dotted')
    out.line(rows=thin_line_row, type='thin')
    out.line(rows=dot_line_row, type='dotted')

    out.clear_more_than(diff_tool.get_attr_num(), diff_tool.get_keys_num() + 1)
    out.font('Meiryo UI')

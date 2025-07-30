import shutil
import os

from numpy.f2py.auxfuncs import throw_error

from workbook import Workbook
from myutil import *
from officeLoader import OfficeLoader
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.format.calc.direct.cell.borders import Side
from ooodev.formatters.formatter_table import FormatterTable, FormatTableItem


def foo():
    wb_src = Workbook(read_only=True, filepath=r"F:\客户风险\数据\2024年11月重点客户风险排查情况表.xlsx", visible=False)
    sheet1_props = wb_src.doc.sheets[0].get_custom_properties()


def gen_xls(src_file: str, tgt_wb: Workbook, sheet_n, data_cell_name: str, date_str: str = None,
            date_cell_name: str = None,
            order: [] = None,
            sum_cells_list: [] = None, merge_list: [] = None, merge_idx_name: str = None,
            idx_row: int = None) -> None:
    print("open workbook")
    src_wb = Workbook(read_only=True, filepath=src_file, visible=False)
    data = array2df(src_wb.get_used_value(0))
    if order is not None:
        data = reorder_dataframe_columns(data, order)

    merge_index = None
    if merge_idx_name is not None:
        labels, uniques = pd.factorize(data[merge_idx_name])  # 因子化第一列
        merge_index = labels + 1  # 標籤從1開始
        if idx_row is not None:
            data.insert(idx_row, 'merge_index', merge_index)

    src_wb.close()
    print("copy data")
    tgt_wb.set_pandas_range(data, sheet_n, data_cell_name)
    if date_str is not None or date_cell_name is not None:
        print("set date")
        tgt_wb.doc.sheets[sheet_n][date_cell_name].value = date_str

    if merge_list is not None:
        for cell in merge_list:
            print(f"merge {cell}")
            tgt_wb.merge_same_cells(sheet_n, cell, merge_index)

    if sum_cells_list is not None:
        for cell in sum_cells_list:
            print(f"sum {sum_cells_list}")
            tgt_wb.sum_col(sheet_n, cell)


def key_customers(template_path: str, data_path: str, result_path: str, date_str: str) -> None:
    template_file = os.path.join(template_path, "重点客户风险排查情况表-模板.xlsx")
    src_list = list(map(lambda x: os.path.join(data_path, x),
                        ['借新还旧汇总.xlsx', '借新还旧明细.xlsx', '逾期贷款.xlsx']))

    orders = [
        ['贷款发放行名称', '贷款笔数', '发放金额', '贷款余额'],
        ['贷款发放行名称', '贷款客户名称', '借新还旧次数', '发放日期', '到期日期',
         '发放金额', '贷款余额', '欠本天数', '欠息天数', '五级分类', '贷款发放类型'],
        ['贷款发放行名称', '贷款客户名称', '发放日期', '到期日期', '发放金额',
         '贷款余额', '欠本天数', '欠息天数', '五级分类', '贷款发放类型']
    ]
    result_file = os.path.join(result_path, f"{date_str}重点客户风险排查情况表.xlsx")
    if not os.path.exists(result_file):
        shutil.copy2(template_file, result_file)
    wb_tgt = Workbook(read_only=False, filepath=result_file, visible=True)
    data = [
        [src_list[0], wb_tgt, 0, 'A5', date_str, 'A2', orders[0], ['B4', 'C4', 'D4'], None],
        [src_list[1], wb_tgt, 1, 'A4', date_str, 'A2', orders[1], None, ['A4'], '贷款发放行名称'],
        [src_list[2], wb_tgt, 2, 'A5', date_str, 'A2', orders[2], ['E4', 'F4'], ['A5'], '贷款发放行名称']
    ]

    for d in data:
        gen_xls(*d)

    wb_tgt.save()
    wb_tgt.close()


def covering_up_asset_quality(template_path: str, data_path: str, result_path: str, date_str: str) -> None:
    template_file = os.path.join(template_path, "疑似掩盖资产质量贷款台账-模板.xlsx")
    src_path = os.path.join(data_path, '借新还旧.xlsx')
    orders = [
        ['贷款发放行名称', '贷款客户名称', '贷款发放类型', '发放日期', '到期日期',
         '发放金额', '贷款余额', '五级分类', '欠本天数', '欠息天数'],
        ['上报时间', '贷款发放行名称', '贷款客户名称', '贷款发放类型', '发放日期', '到期日期',
         '发放金额', '贷款余额', '五级分类', '欠本天数', '欠息天数'],
    ]

    result_file = os.path.join(result_path, f"{date_str}疑似掩盖资产质量贷款台账.xlsx")
    if not os.path.exists(result_file):
        shutil.copy2(template_file, result_file)
    wb_tgt = Workbook(read_only=False, filepath=result_file, visible=True)
    data = [
        [src_path, wb_tgt, 0, 'A5', date_str, 'A2', orders[0], ['F4', 'G4'], ['A5', 'B5']],
        [src_path, wb_tgt, 1, 'A4', date_str, 'A2', orders[1], None, None],
    ]
    for d in data:
        gen_xls(*d)

    wb_tgt.save()
    wb_tgt.close()


def bank_data_tables(template_path: str, data_path: str, result_path: str, date_str: str,
                     visible: bool = True) -> None:
    template_file = os.path.join(template_path, "昭通市银行业对公客户贷款相关台账-模板.xlsx")
    src_list = list(map(lambda x: os.path.join(data_path, x),
                        ['多头授信.xlsx', '五级分类.xlsx', '前20大客户.xlsx', '前20大关注.xlsx', '前20大不良.xlsx',
                         '煤炭企业.xlsx']))

    orders = [
        ['CUSTOMERNAME', 'BANKNAME', 'COUNT_CREDIT', 'STARTDATE', 'DUEDATE', 'LOANBALANCE', 'CUSTOMER_LOANBALANCE',
         'FIVECLASSIFY'],
        ['客户名称', '贷款发放行', '发放日期', '到期日期', '发放金额', '贷款余额',
         '贷款余额小计', '五级分类', ],
        ['贷款客户名称', '发放机构', '发放金额', '贷款余额明细', '客户贷款余额', '贷款余额', '五级分类'],
        ['贷款客户名称', '贷款发放机构', '发放日期', '到期日期', '发放金额', '贷款余额明细', '客户贷款余额',
         '贷款余额'],
        ['贷款客户名称', '发放机构', '发放金额', '贷款余额明细', '客户贷款余额', '全部对公客户不良贷款余额',
         '五级分类'],
        ['贷款客户名称', '机构名称', '发放日期', '到期日期', '发放金额', '客户贷款余额', '贷款余额', '五级分类']
    ]
    result_file = os.path.join(result_path, f"{date_str}昭通市银行业对公客户贷款相关台账.xlsx")
    if not os.path.exists(result_file):
        shutil.copy2(template_file, result_file)
    wb_tgt = Workbook(read_only=False, filepath=result_file, visible=visible)
    # src_file tgt_wb sheet_n, data_cell_name date_str date_cell_name orde sum_cells_list merge_list merge_idx_name    idx_row:
    data = [
        [src_list[0], wb_tgt, 0, 'A5', date_str, 'A2', orders[0], ['G4'], ['A5', 'B5', 'C5', 'D5', 'H5'], orders[0][0],
         0],
        [src_list[1], wb_tgt, 1, 'A5', date_str, 'A2', orders[1], ['F4', 'G4'], ['A5', 'B5', 'C5', 'F5'], orders[1][0],
         0],
        [src_list[2], wb_tgt, 2, 'A6', date_str, 'A2', orders[2], ['E5'], ['A6', 'B6', 'C6', 'F6', 'G6'], orders[2][0],
         0],
        [src_list[3], wb_tgt, 3, 'A6', date_str, 'A2', orders[3], ['F5','G5'], ['A6', 'B6', 'C6', 'H6', 'I6'], orders[3][0],
         0],
        [src_list[4], wb_tgt, 4, 'A6', date_str, 'A2', orders[4], ['E5'], ['A6', 'B6', 'C6', 'F6', 'G6'],
         orders[4][0], 0],
        [src_list[5], wb_tgt, 5, 'A6', date_str, 'A2', orders[5], ['F5', 'H5'], ['A6', 'B6', 'C6', 'G6'],
         orders[5][0], 0],

    ]
    for d in data:
        gen_xls(*d)

    data2 = [
        [wb_tgt, 2, 'G6', 'G6'],
        [wb_tgt, 3, 'I6', 'I6'],
        [wb_tgt, 4, 'G6', 'G6'],
    ]
    for d in data2:
        print('foo')
        foo(*d)

    wb_tgt.save()
    wb_tgt.close()


def foo(workbook, sheet_n, cell1, cell2):
    sheet = workbook.doc.sheets[sheet_n]

    used_rng = workbook.doc.sheets[sheet_n].find_used_range_obj()
    col = get_cell_col_name(cell1)
    cell1_list = convert_cell_name_to_list(cell1)
    end_cell = f"{col}{used_rng.end_row_index + 1}"

    range_name = f"{cell1}:{end_cell}"
    range_list = convert_range_name_to_list(range_name)
    all_balance = workbook.doc.sheets[sheet_n][cell1].value
    for i in range(cell1_list[1], used_rng.end_row_index + 1):
        cell = sheet.get_cell(col=cell1_list[0], row=i)
        value = sheet.get_cell(col=cell1_list[0] - 1, row=i).value / all_balance
        cell.set_val(f"{round(value * 100, 2)}%")
    # rng = workbook.doc.sheets[sheet_n].get_range(range_name=range_name)
    # fl = FormatterTable(format=(".2f", ">9"), idxs=(range_list[1], range_list[3]))
    # print(range_list)
    # fl.col_formats.append(
    #     FormatTableItem(
    #         format=(".2%", ">9"),
    #         idxs_inc=(range_list[0],),
    #         row_idxs_exc=(range_list[1], range_list[3]),
    #     )
    # )


def tech_companies(template_path: str, data_path: str, result_path: str, date_str: str,
                   visible: bool = True):
    template_file = os.path.join(template_path, "昭通市科技型企业和高新企业贷款相关台账-模板.xlsx")
    src_list = list(map(lambda x: os.path.join(data_path, x), ['科技型企业.xlsx', '高新企业.xlsx']))

    orders = [
        ['贷款客户名称', '机构名称', '发放日期', '到期日期', '发放金额', '贷款余额', '五级分类'],
        ['贷款客户名称', '机构名称', '发放日期', '到期日期', '发放金额', '贷款余额', '五级分类'],
    ]
    result_file = os.path.join(result_path, f"{date_str}昭通市科技型企业和高新企业贷款相关台账.xlsx")
    if not os.path.exists(result_file):
        shutil.copy2(template_file, result_file)
    wb_tgt = Workbook(read_only=False, filepath=result_file, visible=visible)
    data = [
        [src_list[0], wb_tgt, 0, 'A5', date_str, 'A2', orders[0], ['E4', 'F4'], ['A5', 'B5'], orders[0][0], None],
        [src_list[1], wb_tgt, 1, 'A5', date_str, 'A2', orders[1], ['E4', 'F4'], ['A5', 'B5'], orders[1][0], None], ]

    for d in data:
        gen_xls(*d)

    wb_tgt.save()
    wb_tgt.close()


def test111():
    src_wb = Workbook(read_only=False, filepath=r'F:\客户风险\数据\test_merge.xlsx', visible=True)
    data = array2df(src_wb.get_used_value(0))
    labels, uniques = pd.factorize(data['col1'])  # 因子化第一列
    i = labels + 1  # 標籤從1開始

    print("索引i =", i.tolist())
    src_wb.merge_cells_by_index(0, "B2", i)
    src_wb.merge_same_cells(0, "C2", i)
    src_wb.merge_same_cells(0, "D2")
    src_wb.close()


def main():
    data_path = r'F:\客户风险\数据'
    template_path = r'F:\客户风险\teml'
    result_path = r'F:\客户风险\202503'
    date_str = "2025年3月"
    src_files = [
        '借新还旧汇总.xlsx', '借新还旧明细.xlsx', '逾期贷款.xlsx', '借新还旧.xlsx', '多头授信.xlsx', '五级分类.xlsx',
        '前20大客户.xlsx', '前20大关注.xlsx', '前20大不良.xlsx', '煤炭企业.xlsx', '科技型企业.xlsx', '高新企业.xlsx'
    ]
    check_result = check_files_exist(data_path, src_files)
    for filename, exists in check_result['details'].items():
        status = "存在" if exists else "不存在"
        if not exists:
            print(f"{filename}: {status}")

    if not check_result['all_exist']:
        raise FileNotFoundError(f"部分文件不存在")

    key_customers(template_path, data_path, result_path, date_str)

    covering_up_asset_quality(template_path, data_path, result_path, date_str)
    bank_data_tables(template_path, data_path, result_path, date_str)
    tech_companies(template_path, data_path, result_path, date_str)
    office_loader = OfficeLoader()
    office_loader.close()


if __name__ == '__main__':
    raise SystemExit(main())
# test111()

from __future__ import annotations
import pandas as pd
from typing import Tuple
from typing import Union
import os
import re


def is_number_regex(s):
    pattern = r'^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$'
    return re.fullmatch(pattern, s.strip()) is not None


def number_to_rounded_str(n: Union[int, float], digits: int = 2) -> str:
    if isinstance(n, int):
        return str(n)
    elif isinstance(n, float):
        # 检查小数部分是否为 0（如 5.0, 10.0）
        if n.is_integer():
            return str(int(n))
        else:
            # 四舍五入并保留指定位数
            rounded = round(n, digits)
            # 如果 digits=0，直接返回整数形式
            if digits == 0:
                return str(int(rounded))
            else:
                return f"{rounded}"
    else:
        print()
        raise ValueError(f"Input n = {n} must be int or float")


def auto_convert_objects(df):
    # 遍历所有object类型的列
    for col in df.select_dtypes(include='object').columns:
        # 原始列的缺失值数量
        original_nulls = df[col].isnull().sum()
        # 尝试转换为数值
        # print(f'col:{col}')
        # print(df[col])
        # if df[col].shape[1] !=1:
        #     raise Exception(f"the sheet has {df[col].shape[1]} cols named {col}")
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        new_nulls = numeric_series.isnull().sum()

        # 判断是否可以转换为数值类型
        if new_nulls == original_nulls:
            df[col] = numeric_series.astype('float64')
        else:
            # 转换为字符串类型
            df[col] = df[col].astype('string')
    return df


def array2df(data_set: Tuple[Tuple, ...]) -> pd.DataFrame:
    # 识别列名行（假设列名行所有字段非空）
    columns = next(item for item in data_set if all(field != '' for field in item))

    # 提取数据行（排除列名行）
    rows = [list(item) for item in data_set if item != columns]

    # 创建DataFrame
    df = pd.DataFrame(rows, columns=columns)
    return auto_convert_objects(df)


#       type  label              value                      description
# 9              a6                0.0
# 10  /10000     a7         6413960.45
# 11     增减值     a8         -22.203955
# 12  /10000     a9         8945017.34
# 定义处理逻辑的函数
def process_value_to_str(row: [], decimal_places: int = 2) -> str:
    value = row['value']

    if value == '':  # value为空，无法处理
        return ''
    if not is_number_regex(value):
        return value

    if type(value) is str:
        value = float(value)

    if row['type'] == '':  # type为空，不处理
        return number_to_rounded_str(value, decimal_places)

    # 根据 type 进行运算
    if row['type'] == '增减值':
        return ("增加" if value >= 0 else "减少") + number_to_rounded_str(abs(value), decimal_places)

    operators = '+-*/'
    for op in operators:
        if row['type'].startswith(op):
            sign = op
            number = float(row['type'][1:])

            result = value
            # print(f"符号: {sign}, 数字: {number}")
            if sign == '/':
                result = value / number
            elif sign == '*':
                result = value * number
            elif sign == '+':
                result = value + number
            elif sign == '-':
                result = value - number
            # 返回整数或浮点数的字符串形式（根据实际需求调整）
            return number_to_rounded_str(result, decimal_places)

    else:
        return number_to_rounded_str(value, decimal_places)  # 未知type，不处理


def get_cell_col_name(cell_name: str) -> str:
    match = re.match(r"^([A-Za-z]+)(\d+)$", cell_name)
    if not match:
        raise ValueError("Invalid input format")
    letters = match.group(1).upper()
    return letters


def convert_cell_name_to_list(cell_name: str) -> list:
    # 使用正则表达式分割字母和数字部分
    match = re.match(r"^([A-Za-z]+)(\d+)$", cell_name)
    if not match:
        raise ValueError("Invalid input format")

    letters = match.group(1).upper()
    numbers = match.group(2)

    # 转换字母部分为数值（类似Excel列编号）
    column = 0
    for c in letters:
        column = column * 26 + (ord(c) - ord('A'))

    # 转换数字部分为整数
    row = int(numbers) - 1

    return [column, row]


def convert_range_name_to_list(range_name) -> list:
    # 分割范围字符串为起始和结束单元格
    parts = range_name.split(':')
    if len(parts) != 2:
        raise ValueError("Invalid range format. Expected format like 'A1:C3'")

    start = convert_cell_name_to_list(parts[0])
    end = convert_cell_name_to_list(parts[1])

    # 组合结果：[起始列, 起始行, 结束列, 结束行]
    return start + end


def convert_list_to_range_name(lst):
    """
    将 [列号, 行号] 转换为 Excel 的 A1 格式字符串（如 [1,1] -> "A1"）

    Args:
        lst: 包含两个整数的列表，格式为 [列号, 行号]

    Returns:
        str: 对应的 Excel 单元格地址

    Raises:
        ValueError: 输入格式错误或数值不合法
    """
    # 验证输入格式
    if len(lst) != 2:
        raise ValueError("输入必须为包含两个元素的列表，例如 [1,1]")
    column, row = lst[0], lst[1]

    # 验证列和行是否为合法正整数
    if not isinstance(column, int) or not isinstance(row, int):
        raise ValueError("列和行必须为整数")
    if column < 1 or row < 1:
        raise ValueError("列和行必须大于等于 1")

    # 转换列号为字母部分（如 1 -> A，27 -> AA）
    letters = ""
    n = column
    while n > 0:
        n -= 1  # 调整为从0开始计算
        remainder = n % 26
        letters = chr(ord('A') + remainder) + letters
        n = n // 26

    return f"{letters}{row}"


def reorder_dataframe_columns(df, new_order):
    # 检查new_order中的列是否都存在于DataFrame中
    missing_columns = [col for col in new_order if col not in df.columns]
    if missing_columns:
        raise ValueError(f"以下列不存在于DataFrame中: {missing_columns}")
    # 按新顺序筛选列
    return df[new_order]


def check_files_exist(directory: str, file_list: []) -> []:
    """
    检查指定目录下是否存在给定的所有文件

    参数:
        directory (str): 要检查的目录路径
        file_list (list): 要检查的文件名列表

    返回:
        dict: 包含每个文件是否存在的结果字典，
              以及一个总结字段表示是否所有文件都存在
    """
    result = {
        'all_exist': True,
        'details': {}
    }

    # 检查目录是否存在
    if not os.path.isdir(directory):
        raise ValueError(f"目录不存在: {directory}")

    # 检查每个文件是否存在
    for filename in file_list:
        file_path = os.path.join(directory, filename)
        exists = os.path.isfile(file_path)
        result['details'][filename] = exists

        if not exists:
            result['all_exist'] = False

    return result

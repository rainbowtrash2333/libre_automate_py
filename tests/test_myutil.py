import pytest
import pandas as pd
import os
import tempfile
from typing import Tuple, Union
from src.libre_automate_py.myutil import (
    is_number_regex,
    number_to_rounded_str,
    auto_convert_objects,
    array2df,
    process_value_to_str,
    get_cell_col_name,
    convert_cell_name_to_list,
    convert_range_name_to_list,
    convert_list_to_range_name,
    reorder_dataframe_columns,
    check_files_exist
)


class TestMyUtilFunctions:
    """测试myutil模块中的工具函数"""
    
    def test_is_number_regex(self):
        """测试数字正则表达式验证函数"""
        # 正整数
        assert is_number_regex("123") is True
        assert is_number_regex("0") is True
        
        # 负整数
        assert is_number_regex("-123") is True
        assert is_number_regex("-0") is True
        
        # 正浮点数
        assert is_number_regex("123.45") is True
        assert is_number_regex("0.123") is True
        assert is_number_regex(".123") is True
        
        # 负浮点数
        assert is_number_regex("-123.45") is True
        assert is_number_regex("-.123") is True
        
        # 科学计数法
        assert is_number_regex("1.23e10") is True
        assert is_number_regex("1.23E-10") is True
        assert is_number_regex("-1.23e+5") is True
        
        # 非数字
        assert is_number_regex("abc") is False
        assert is_number_regex("12.34.56") is False
        assert is_number_regex("") is False
        assert is_number_regex("  ") is False
        
        # 带空格的数字
        assert is_number_regex("  123  ") is True
        assert is_number_regex("  -123.45  ") is True
    
    def test_number_to_rounded_str(self):
        """测试数字转换为字符串函数"""
        # 整数
        assert number_to_rounded_str(123) == "123"
        assert number_to_rounded_str(-456) == "-456"
        assert number_to_rounded_str(0) == "0"
        
        # 整型浮点数
        assert number_to_rounded_str(5.0) == "5"
        assert number_to_rounded_str(-10.0) == "-10"
        
        # 普通浮点数
        assert number_to_rounded_str(3.14159, 2) == "3.14"
        assert number_to_rounded_str(3.14159, 4) == "3.1416"
        assert number_to_rounded_str(3.14159, 0) == "3"
        
        # 负浮点数
        assert number_to_rounded_str(-2.71828, 3) == "-2.718"
        
        # 错误输入
        with pytest.raises(ValueError):
            number_to_rounded_str("not_a_number")
    
    def test_auto_convert_objects(self):
        """测试DataFrame对象类型自动转换函数"""
        # 创建测试DataFrame
        df = pd.DataFrame({
            'numbers': ['1', '2', '3.5', '4'],
            'mixed': ['1', 'text', '3', ''],
            'text': ['apple', 'banana', 'cherry', 'date'],
            'empty': ['', '', '', '']
        })
        
        result = auto_convert_objects(df)
        
        # 验证纯数字列被转换为float64
        assert result['numbers'].dtype == 'float64'
        
        # 验证混合列被转换为string
        assert result['mixed'].dtype == 'string'
        assert result['text'].dtype == 'string'
    
    def test_array2df(self):
        """测试数组转DataFrame函数"""
        # 测试数据
        data = (
            ('Name', 'Age', 'City'),
            ('Alice', '25', 'New York'),
            ('Bob', '30', 'Los Angeles'),
            ('', '', ''),  # 空行
            ('Charlie', '35', 'Chicago')
        )
        
        result = array2df(data)
        
        # 验证DataFrame结构
        assert list(result.columns) == ['Name', 'Age', 'City']
        assert len(result) == 3  # 应该排除空行和列名行
        assert result.iloc[0]['Name'] == 'Alice'
        assert result.iloc[1]['Name'] == 'Bob'
        assert result.iloc[2]['Name'] == 'Charlie'
    
    def test_process_value_to_str(self):
        """测试值处理转字符串函数"""
        # 空值
        row = {'type': '', 'value': ''}
        assert process_value_to_str(row) == ''
        
        # 非数字值
        row = {'type': '', 'value': 'text'}
        assert process_value_to_str(row) == 'text'
        
        # 普通数字，无类型
        row = {'type': '', 'value': '123.45'}
        assert process_value_to_str(row) == '123.45'
        
        # 增减值类型
        row = {'type': '增减值', 'value': '10.5'}
        assert process_value_to_str(row) == '增加10.5'
        
        row = {'type': '增减值', 'value': '-5.2'}
        assert process_value_to_str(row) == '减少5.2'
        
        # 除法运算
        row = {'type': '/10000', 'value': '50000'}
        assert process_value_to_str(row) == '5'
        
        # 乘法运算
        row = {'type': '*100', 'value': '0.25'}
        assert process_value_to_str(row) == '25'
        
        # 加法运算
        row = {'type': '+10', 'value': '20'}
        assert process_value_to_str(row) == '30'
        
        # 减法运算
        row = {'type': '-5', 'value': '15'}
        assert process_value_to_str(row) == '10'
    
    def test_get_cell_col_name(self):
        """测试获取单元格列名函数"""
        assert get_cell_col_name("A1") == "A"
        assert get_cell_col_name("B10") == "B"
        assert get_cell_col_name("Z100") == "Z"
        assert get_cell_col_name("AA1") == "AA"
        assert get_cell_col_name("AB123") == "AB"
        
        # 小写输入
        assert get_cell_col_name("a1") == "A"
        assert get_cell_col_name("bc23") == "BC"
        
        # 错误输入
        with pytest.raises(ValueError):
            get_cell_col_name("1A")
        with pytest.raises(ValueError):
            get_cell_col_name("A")
        with pytest.raises(ValueError):
            get_cell_col_name("123")
    
    def test_convert_cell_name_to_list(self):
        """测试单元格名转列表函数"""
        # 基本测试
        assert convert_cell_name_to_list("A1") == [0, 0]
        assert convert_cell_name_to_list("B2") == [1, 1]
        assert convert_cell_name_to_list("Z26") == [25, 25]
        
        # 多字母列
        assert convert_cell_name_to_list("AA1") == [26, 0]
        assert convert_cell_name_to_list("AB2") == [27, 1]
        
        # 小写输入
        assert convert_cell_name_to_list("a1") == [0, 0]
        assert convert_cell_name_to_list("bc10") == [54, 9]
        
        # 错误输入
        with pytest.raises(ValueError):
            convert_cell_name_to_list("1A")
    
    def test_convert_range_name_to_list(self):
        """测试范围名转列表函数"""
        # 基本范围
        assert convert_range_name_to_list("A1:B2") == [0, 0, 1, 1]
        assert convert_range_name_to_list("C3:E5") == [2, 2, 4, 4]
        
        # 大范围
        assert convert_range_name_to_list("A1:Z100") == [0, 0, 25, 99]
        
        # 多字母列
        assert convert_range_name_to_list("AA1:AB2") == [26, 0, 27, 1]
        
        # 错误输入
        with pytest.raises(ValueError):
            convert_range_name_to_list("A1")
        with pytest.raises(ValueError):
            convert_range_name_to_list("A1:B2:C3")
    
    def test_convert_list_to_range_name(self):
        """测试列表转范围名函数"""
        # 基本测试
        assert convert_list_to_range_name([1, 1]) == "A1"
        assert convert_list_to_range_name([2, 10]) == "B10"
        assert convert_list_to_range_name([26, 1]) == "Z1"
        
        # 多字母列
        assert convert_list_to_range_name([27, 1]) == "AA1"
        assert convert_list_to_range_name([28, 2]) == "AB2"
        
        # 错误输入
        with pytest.raises(ValueError):
            convert_list_to_range_name([1])
        with pytest.raises(ValueError):
            convert_list_to_range_name([0, 1])
        with pytest.raises(ValueError):
            convert_list_to_range_name([1, 0])
        with pytest.raises(ValueError):
            convert_list_to_range_name(['a', 1])
    
    def test_reorder_dataframe_columns(self):
        """测试DataFrame列重排序函数"""
        # 创建测试DataFrame
        df = pd.DataFrame({
            'C': [1, 2, 3],
            'A': [4, 5, 6],
            'B': [7, 8, 9]
        })
        
        # 重新排序
        new_order = ['A', 'B', 'C']
        result = reorder_dataframe_columns(df, new_order)
        
        assert list(result.columns) == ['A', 'B', 'C']
        
        # 部分列排序
        partial_order = ['B', 'A']
        result_partial = reorder_dataframe_columns(df, partial_order)
        assert list(result_partial.columns) == ['B', 'A']
        
        # 错误输入 - 不存在的列
        with pytest.raises(ValueError):
            reorder_dataframe_columns(df, ['A', 'D'])
    
    def test_check_files_exist(self):
        """测试文件存在性检查函数"""
        # 创建临时目录和文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 创建测试文件
            test_file1 = os.path.join(temp_dir, "test1.txt")
            test_file2 = os.path.join(temp_dir, "test2.txt")
            
            with open(test_file1, 'w') as f:
                f.write("test content")
            
            # 测试文件存在性检查
            file_list = ["test1.txt", "test2.txt", "test3.txt"]
            result = check_files_exist(temp_dir, file_list)
            
            assert result['all_exist'] is False
            assert result['details']['test1.txt'] is True
            assert result['details']['test2.txt'] is False
            assert result['details']['test3.txt'] is False
            
            # 创建第二个文件
            with open(test_file2, 'w') as f:
                f.write("test content 2")
            
            # 重新测试（只检查存在的文件）
            existing_files = ["test1.txt", "test2.txt"]
            result2 = check_files_exist(temp_dir, existing_files)
            
            assert result2['all_exist'] is True
            assert result2['details']['test1.txt'] is True
            assert result2['details']['test2.txt'] is True
        
        # 测试不存在的目录
        with pytest.raises(ValueError):
            check_files_exist("/nonexistent/directory", ["file.txt"])


class MyUtilTestData:
    """myutil模块测试数据类"""
    
    @staticmethod
    def get_number_validation_cases():
        """获取数字验证测试用例"""
        return {
            'valid_integers': ['0', '123', '-456', '+789'],
            'valid_floats': ['0.0', '123.45', '-67.89', '+0.123', '.5', '-.5'],
            'valid_scientific': ['1e10', '1.23e-5', '-2.5E+3', '1.0e0'],
            'invalid_numbers': ['abc', '12.34.56', 'e10', '1.2.3e4', '++123', '--456'],
            'edge_cases': ['', '  ', '  123  ', '  -45.67  ']
        }
    
    @staticmethod
    def get_number_formatting_cases():
        """获取数字格式化测试用例"""
        return [
            {'input': 123, 'digits': 2, 'expected': '123'},
            {'input': 5.0, 'digits': 2, 'expected': '5'},
            {'input': 3.14159, 'digits': 2, 'expected': '3.14'},
            {'input': 3.14159, 'digits': 4, 'expected': '3.1416'},
            {'input': 3.14159, 'digits': 0, 'expected': '3'},
            {'input': -2.71828, 'digits': 3, 'expected': '-2.718'},
            {'input': 0.0, 'digits': 2, 'expected': '0'},
            {'input': -0.0, 'digits': 2, 'expected': '0'}
        ]
    
    @staticmethod
    def get_dataframe_conversion_cases():
        """获取DataFrame转换测试用例"""
        return [
            {
                'name': 'pure_numbers',
                'data': {
                    'col1': ['1', '2', '3'],
                    'col2': ['1.5', '2.5', '3.5']
                },
                'expected_types': {'col1': 'float64', 'col2': 'float64'}
            },
            {
                'name': 'mixed_data',
                'data': {
                    'col1': ['1', 'text', '3'],
                    'col2': ['apple', 'banana', 'cherry']
                },
                'expected_types': {'col1': 'string', 'col2': 'string'}
            },
            {
                'name': 'with_nulls',
                'data': {
                    'col1': ['1', '', '3'],
                    'col2': ['2.5', None, '4.5']
                },
                'expected_types': {'col1': 'string', 'col2': 'float64'}
            }
        ]
    
    @staticmethod
    def get_array_to_dataframe_cases():
        """获取数组转DataFrame测试用例"""
        return [
            {
                'name': 'simple_table',
                'data': (
                    ('Name', 'Age', 'City'),
                    ('Alice', '25', 'New York'),
                    ('Bob', '30', 'Los Angeles')
                ),
                'expected_shape': (2, 3),
                'expected_columns': ['Name', 'Age', 'City']
            },
            {
                'name': 'with_empty_rows',
                'data': (
                    ('ID', 'Value'),
                    ('1', '100'),
                    ('', ''),
                    ('2', '200'),
                    ('3', '300')
                ),
                'expected_shape': (3, 2),  # 空行被过滤
                'expected_columns': ['ID', 'Value']
            },
            {
                'name': 'numeric_data',
                'data': (
                    ('Product', 'Price', 'Quantity'),
                    ('Apple', '1.2', '100'),
                    ('Banana', '0.8', '150'),
                    ('Orange', '1.5', '80')
                ),
                'expected_shape': (3, 3),
                'expected_columns': ['Product', 'Price', 'Quantity']
            }
        ]
    
    @staticmethod
    def get_value_processing_cases():
        """获取值处理测试用例"""
        return [
            # 空值和非数字
            {'type': '', 'value': '', 'expected': ''},
            {'type': '', 'value': 'text', 'expected': 'text'},
            
            # 普通数字
            {'type': '', 'value': '123.45', 'expected': '123.45'},
            {'type': '', 'value': 123.45, 'expected': '123.45'},
            
            # 增减值
            {'type': '增减值', 'value': '10.5', 'expected': '增加10.5'},
            {'type': '增减值', 'value': '-5.2', 'expected': '减少5.2'},
            {'type': '增减值', 'value': 0, 'expected': '增加0'},
            
            # 数学运算
            {'type': '/10000', 'value': '50000', 'expected': '5'},
            {'type': '*100', 'value': '0.25', 'expected': '25'},
            {'type': '+10', 'value': '20', 'expected': '30'},
            {'type': '-5', 'value': '15', 'expected': '10'},
            
            # 特殊情况
            {'type': '/0', 'value': '100', 'expected': 'inf'},  # 除零
            {'type': 'unknown', 'value': '123', 'expected': '123'}  # 未知类型
        ]
    
    @staticmethod
    def get_cell_reference_cases():
        """获取单元格引用测试用例"""
        return [
            # 基本单元格
            {'cell': 'A1', 'list': [0, 0], 'col': 'A'},
            {'cell': 'B10', 'list': [1, 9], 'col': 'B'},
            {'cell': 'Z100', 'list': [25, 99], 'col': 'Z'},
            
            # 多字母列
            {'cell': 'AA1', 'list': [26, 0], 'col': 'AA'},
            {'cell': 'AB123', 'list': [27, 122], 'col': 'AB'},
            {'cell': 'AZ999', 'list': [51, 998], 'col': 'AZ'},
            
            # 小写输入
            {'cell': 'a1', 'list': [0, 0], 'col': 'A'},
            {'cell': 'bc23', 'list': [54, 22], 'col': 'BC'}
        ]
    
    @staticmethod
    def get_range_reference_cases():
        """获取范围引用测试用例"""
        return [
            # 基本范围
            {'range': 'A1:B2', 'list': [0, 0, 1, 1]},
            {'range': 'C3:E5', 'list': [2, 2, 4, 4]},
            {'range': 'A1:Z100', 'list': [0, 0, 25, 99]},
            
            # 多字母列范围
            {'range': 'AA1:AB2', 'list': [26, 0, 27, 1]},
            {'range': 'AZ1:BA10', 'list': [51, 0, 52, 9]},
            
            # 单行/单列范围
            {'range': 'A1:A10', 'list': [0, 0, 0, 9]},
            {'range': 'A1:J1', 'list': [0, 0, 9, 0]}
        ]
    
    @staticmethod
    def get_file_check_scenarios():
        """获取文件检查测试场景"""
        return [
            {
                'scenario': 'all_exist',
                'files': ['file1.txt', 'file2.txt', 'file3.txt'],
                'existing': ['file1.txt', 'file2.txt', 'file3.txt'],
                'expected_all_exist': True
            },
            {
                'scenario': 'some_missing',
                'files': ['file1.txt', 'file2.txt', 'missing.txt'],
                'existing': ['file1.txt', 'file2.txt'],
                'expected_all_exist': False
            },
            {
                'scenario': 'all_missing',
                'files': ['missing1.txt', 'missing2.txt'],
                'existing': [],
                'expected_all_exist': False
            },
            {
                'scenario': 'empty_list',
                'files': [],
                'existing': [],
                'expected_all_exist': True
            }
        ]
    
    @staticmethod
    def get_column_reorder_scenarios():
        """获取列重排序测试场景"""
        return [
            {
                'original_columns': ['C', 'A', 'B', 'D'],
                'new_order': ['A', 'B', 'C', 'D'],
                'expected_success': True
            },
            {
                'original_columns': ['X', 'Y', 'Z'],
                'new_order': ['Z', 'X'],
                'expected_success': True
            },
            {
                'original_columns': ['A', 'B', 'C'],
                'new_order': ['A', 'B', 'C', 'D'],  # D不存在
                'expected_success': False,
                'expected_error': ValueError
            },
            {
                'original_columns': ['Col1', 'Col2'],
                'new_order': [],
                'expected_success': True
            }
        ]
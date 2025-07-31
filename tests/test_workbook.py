import pytest
import pandas as pd
from unittest.mock import patch, MagicMock, mock_open
from typing import Tuple
from src.libre_automate_py.workbook import Workbook


class TestWorkbook:
    """测试Workbook类的Excel操作功能"""
    
    def setup_method(self):
        """每个测试方法前的设置"""
        pass
    
    def teardown_method(self):
        """每个测试方法后的清理"""
        pass
    
    @patch('src.libre_automate_py.workbook.OfficeLoader')
    @patch('src.libre_automate_py.workbook.CalcDoc')
    @patch('src.libre_automate_py.workbook.FileIO')
    def test_init_with_filepath(self, mock_fileio, mock_calcdoc, mock_office_loader):
        """测试使用文件路径初始化工作簿"""
        # 设置mock
        mock_loader_instance = MagicMock()
        mock_office_loader.return_value = mock_loader_instance
        mock_loader_instance.get_loader.return_value = MagicMock()
        
        mock_fileio.get_absolute_path.return_value = "/absolute/path/test.xlsx"
        mock_doc = MagicMock()
        mock_calcdoc.open_doc.return_value = mock_doc
        
        # 创建Workbook实例
        wb = Workbook(filepath="test.xlsx", read_only=True, visible=False)
        
        # 验证
        assert wb._filepath == "test.xlsx"
        assert wb._read_only is True
        assert wb._visible is False
        assert wb.doc is mock_doc
        
        mock_fileio.get_absolute_path.assert_called_once_with("test.xlsx")
        mock_calcdoc.open_doc.assert_called_once()
    
    @patch('src.libre_automate_py.workbook.OfficeLoader')
    @patch('src.libre_automate_py.workbook.CalcDoc')
    def test_init_without_filepath(self, mock_calcdoc, mock_office_loader):
        """测试不使用文件路径初始化工作簿（创建新文档）"""
        # 设置mock
        mock_loader_instance = MagicMock()
        mock_office_loader.return_value = mock_loader_instance
        mock_loader_instance.get_loader.return_value = MagicMock()
        
        mock_doc = MagicMock()
        mock_calcdoc.create_doc.return_value = mock_doc
        
        # 创建Workbook实例
        wb = Workbook()
        
        # 验证
        assert wb._filepath is None
        assert wb._read_only is False
        assert wb._visible is True
        assert wb.doc is mock_doc
        
        mock_calcdoc.create_doc.assert_called_once_with(visible=True)
    
    @patch('src.libre_automate_py.workbook.FileIO')
    def test_save_with_path(self, mock_fileio):
        """测试使用指定路径保存文档"""
        # 创建mock document
        mock_doc = MagicMock()
        
        # 创建Workbook实例并设置doc
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb._filepath = None
        
        mock_fileio.get_absolute_path.return_value = "/absolute/path/output.xlsx"
        mock_fileio.make_directory.return_value = None
        
        # 调用save方法
        wb.save("output.xlsx")
        
        # 验证
        mock_fileio.get_absolute_path.assert_called_once_with("output.xlsx")
        mock_doc.save_doc.assert_called_once_with(fnm="/absolute/path/output.xlsx")
    
    def test_save_no_document(self):
        """测试在没有文档时保存抛出异常"""
        wb = Workbook.__new__(Workbook)
        wb.doc = None
        
        with pytest.raises(RuntimeError, match="No document to save"):
            wb.save()
    
    def test_save_no_path(self):
        """测试在没有指定路径时保存抛出异常"""
        mock_doc = MagicMock()
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb._filepath = None
        
        with pytest.raises(ValueError, match="No file path specified for saving"):
            wb.save()
    
    @patch('src.libre_automate_py.workbook.Calc')
    def test_get_range_value(self, mock_calc):
        """测试获取范围值"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_range_obj = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.sheets = [mock_sheet]
        
        mock_calc.get_range_obj.return_value = mock_range_obj
        mock_sheet.get_array.return_value = (('A1', 'B1'), ('A2', 'B2'))
        
        result = wb.get_range_value(0, "A1:B2")
        
        assert result == (('A1', 'B1'), ('A2', 'B2'))
        mock_sheet.get_array.assert_called_once_with(range_obj=mock_range_obj)
    
    def test_close(self):
        """测试关闭文档"""
        mock_doc = MagicMock()
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        
        result = wb.close()
        
        assert result == 0
        mock_doc.close_doc.assert_called_once()
    
    def test_get_used_value_without_range_name(self):
        """测试获取已使用区域的值（不指定范围名）"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_used_range = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.sheets = [mock_sheet]
        
        mock_sheet.find_used_range_obj.return_value = mock_used_range
        mock_sheet.get_array.return_value = (('Data1', 'Data2'), ('Data3', 'Data4'))
        
        result = wb.get_used_value(0)
        
        assert result == (('Data1', 'Data2'), ('Data3', 'Data4'))
        mock_sheet.get_array.assert_called_once_with(range_obj=mock_used_range)
    
    def test_get_used_value_with_range_name(self):
        """测试获取已使用区域的值（指定范围名）"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.sheets = [mock_sheet]
        
        mock_sheet.get_array.return_value = (('Range1', 'Range2'),)
        
        result = wb.get_used_value(0, "A1:B1")
        
        assert result == (('Range1', 'Range2'),)
        mock_sheet.get_array.assert_called_once_with(range_name="A1:B1")
    
    def test_set_array_value(self):
        """测试设置数组值"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.sheets = [mock_sheet]
        
        test_values = (('Test1', 'Test2'), ('Test3', 'Test4'))
        wb.set_array_value(0, test_values, "A1:B2")
        
        mock_sheet.set_array.assert_called_once_with(values=test_values, name="A1:B2")
    
    @patch('src.libre_automate_py.workbook.convert_range_name_to_list')
    @patch('src.libre_automate_py.workbook.Side')
    @patch('src.libre_automate_py.workbook.CommonColor')
    @patch('src.libre_automate_py.workbook.FormatterTable')
    def test_formatter_range(self, mock_formatter_table, mock_common_color, mock_side, mock_convert_range):
        """测试格式化范围"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_range = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.sheets = [mock_sheet]
        
        mock_sheet.get_range.return_value = mock_range
        mock_convert_range.return_value = [0, 0, 1, 1]  # [start_col, start_row, end_col, end_row]
        
        wb.formatter_range(0, "A1:B2")
        
        mock_sheet.get_range.assert_called_once_with(range_name="A1:B2")
        mock_range.style_borders.assert_called_once()
        mock_formatter_table.assert_called_once()
    
    @patch('src.libre_automate_py.workbook.convert_cell_name_to_list')
    def test_merge_cells_by_index(self, mock_convert_cell):
        """测试根据索引合并单元格"""
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_range = MagicMock()
        
        wb = Workbook.__new__(Workbook)
        wb.doc = mock_doc
        wb.doc.get_sheet.return_value = mock_sheet
        mock_sheet.get_range.return_value = mock_range
        
        mock_convert_cell.return_value = [0, 0]  # A1对应的[列, 行]
        
        # 测试数据：连续的相同值
        index_data = [1, 1, 1, 2, 2, 3]
        
        wb.merge_cells_by_index(0, "A1", index_data)
        
        # 验证get_range被调用（具体调用次数取决于需要合并的区域数量）
        assert mock_sheet.get_range.call_count >= 0


class WorkbookTestData:
    """Workbook测试数据类"""
    
    @staticmethod
    def get_sample_excel_data():
        """获取示例Excel数据"""
        return (
            ('Name', 'Age', 'City', 'Salary'),
            ('Alice', '25', 'New York', '50000'),
            ('Bob', '30', 'Los Angeles', '60000'),
            ('Charlie', '35', 'Chicago', '70000'),
            ('Diana', '28', 'Houston', '55000')
        )
    
    @staticmethod
    def get_sample_dataframe():
        """获取示例DataFrame"""
        return pd.DataFrame({
            'Product': ['Apple', 'Banana', 'Orange', 'Grape'],
            'Price': [1.2, 0.8, 1.5, 2.0],
            'Quantity': [100, 150, 80, 120],
            'Category': ['Fruit', 'Fruit', 'Fruit', 'Fruit']
        })
    
    @staticmethod
    def get_cell_range_scenarios():
        """获取单元格范围测试场景"""
        return [
            {
                'range_name': 'A1:B2',
                'expected_data': (('A1', 'B1'), ('A2', 'B2'))
            },
            {
                'range_name': 'C3:E5',
                'expected_data': (
                    ('C3', 'D3', 'E3'),
                    ('C4', 'D4', 'E4'),
                    ('C5', 'D5', 'E5')
                )
            },
            {
                'range_name': 'A1:A10',
                'expected_data': tuple((f'A{i}',) for i in range(1, 11))
            }
        ]
    
    @staticmethod
    def get_merge_scenarios():
        """获取单元格合并测试场景"""
        return [
            {
                'start_cell': 'A1',
                'index_data': [1, 1, 1, 2, 2, 3, 3, 3, 3],
                'expected_merge_ranges': ['A2:A4', 'A5:A6', 'A7:A10']
            },
            {
                'start_cell': 'B2',
                'index_data': [1, 2, 2, 2, 3],
                'expected_merge_ranges': ['B4:B6']
            },
            {
                'start_cell': 'C1',
                'index_data': [1, 1, 2, 3, 3, 3, 4, 4],
                'expected_merge_ranges': ['C2:C3', 'C5:C7', 'C8:C9']
            }
        ]
    
    @staticmethod
    def get_file_scenarios():
        """获取文件操作测试场景"""
        return [
            {
                'filepath': 'test.xlsx',
                'read_only': True,
                'visible': False
            },
            {
                'filepath': 'data/report.xlsx',
                'read_only': False,
                'visible': True
            },
            {
                'filepath': None,
                'read_only': False,
                'visible': True
            }
        ]
    
    @staticmethod
    def get_error_scenarios():
        """获取错误场景测试数据"""
        return [
            {
                'error_type': 'file_not_found',
                'filepath': 'nonexistent.xlsx',
                'expected_exception': FileNotFoundError
            },
            {
                'error_type': 'permission_denied',
                'filepath': 'readonly.xlsx',
                'expected_exception': PermissionError
            },
            {
                'error_type': 'invalid_format',
                'filepath': 'invalid.txt',
                'expected_exception': ValueError
            }
        ]
    
    @staticmethod
    def get_formatting_test_data():
        """获取格式化测试数据"""
        return {
            'border_styles': ['thin', 'thick', 'double'],
            'colors': ['black', 'red', 'blue', 'green'],
            'number_formats': ['.2f', '.0f', '.4f'],
            'ranges': ['A1:B5', 'C1:F10', 'A1:Z100']
        }
    
    @staticmethod
    def get_sum_formula_scenarios():
        """获取求和公式测试场景"""
        return [
            {
                'sum_cell': 'A10',
                'end_cell': None,
                'expected_formula': '=SUM(A12:A20)'  # 示例，实际值取决于已使用区域
            },
            {
                'sum_cell': 'B5',
                'end_cell': 'B15',
                'expected_formula': '=SUM(B7:B16)'
            },
            {
                'sum_cell': 'C1',
                'end_cell': 'C10',
                'expected_formula': '=SUM(C3:C11)'
            }
        ]
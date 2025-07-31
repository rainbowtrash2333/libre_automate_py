import pytest
from unittest.mock import patch, MagicMock
from typing import Sequence
from src.libre_automate_py.word import Word


class TestWord:
    """测试Word类的文档操作功能"""
    
    def setup_method(self):
        """每个测试方法前的设置"""
        pass
    
    def teardown_method(self):
        """每个测试方法后的清理"""
        pass
    
    @patch('src.libre_automate_py.word.OfficeLoader')
    @patch('src.libre_automate_py.word.WriteDoc')
    @patch('src.libre_automate_py.word.FileIO')
    def test_init_with_filepath(self, mock_fileio, mock_writedoc, mock_office_loader):
        """测试使用文件路径初始化Word文档"""
        # 设置mock
        mock_loader_instance = MagicMock()
        mock_office_loader.return_value = mock_loader_instance
        mock_loader_instance.get_loader.return_value = MagicMock()
        
        mock_fileio.get_absolute_path.return_value = "/absolute/path/test.docx"
        mock_doc = MagicMock()
        mock_writedoc.open_doc.return_value = mock_doc
        
        # 创建Word实例
        word = Word(filepath="test.docx", read_only=True, visible=False)
        
        # 验证
        assert word._filepath == "test.docx"
        assert word._read_only is True
        assert word._visible is False
        assert word.doc is mock_doc
        
        mock_fileio.get_absolute_path.assert_called_once_with("test.docx")
        mock_writedoc.open_doc.assert_called_once()
    
    @patch('src.libre_automate_py.word.OfficeLoader')
    @patch('src.libre_automate_py.word.WriteDoc')
    def test_init_without_filepath(self, mock_writedoc, mock_office_loader):
        """测试不使用文件路径初始化Word文档（创建新文档）"""
        # 设置mock
        mock_loader_instance = MagicMock()
        mock_office_loader.return_value = mock_loader_instance
        mock_loader_instance.get_loader.return_value = MagicMock()
        
        mock_doc = MagicMock()
        mock_writedoc.create_doc.return_value = mock_doc
        
        # 创建Word实例
        word = Word()
        
        # 验证
        assert word._filepath is None
        assert word._read_only is False
        assert word._visible is True
        assert word.doc is mock_doc
        
        mock_writedoc.create_doc.assert_called_once_with(visible=True)
    
    @patch('src.libre_automate_py.word.FileIO')
    def test_save_with_path(self, mock_fileio):
        """测试使用指定路径保存文档"""
        # 创建mock document
        mock_doc = MagicMock()
        
        # 创建Word实例并设置doc
        word = Word.__new__(Word)
        word.doc = mock_doc
        word._filepath = None
        
        mock_fileio.get_absolute_path.return_value = "/absolute/path/output.docx"
        mock_fileio.make_directory.return_value = None
        
        # 调用save方法
        word.save("output.docx")
        
        # 验证
        mock_fileio.get_absolute_path.assert_called_once_with("output.docx")
        mock_doc.save_doc.assert_called_once_with(fnm="/absolute/path/output.docx")
    
    def test_save_no_document(self):
        """测试在没有文档时保存抛出异常"""
        word = Word.__new__(Word)
        word.doc = None
        
        with pytest.raises(RuntimeError, match="No document to save"):
            word.save()
    
    def test_save_no_path(self):
        """测试在没有指定路径时保存抛出异常"""
        mock_doc = MagicMock()
        word = Word.__new__(Word)
        word.doc = mock_doc
        word._filepath = None
        
        with pytest.raises(ValueError, match="No file path specified for saving"):
            word.save()
    
    def test_close(self):
        """测试关闭文档"""
        mock_doc = MagicMock()
        word = Word.__new__(Word)
        word.doc = mock_doc
        
        word.close()
        
        mock_doc.close_doc.assert_called_once()
    
    @patch('src.libre_automate_py.word.Write')
    def test_get_content_text(self, mock_write):
        """测试获取文档内容文本"""
        mock_doc = MagicMock()
        mock_text_doc = MagicMock()
        mock_cursor = MagicMock()
        
        word = Word.__new__(Word)
        word.doc = mock_doc
        
        mock_write.get_text_doc.return_value = mock_text_doc
        mock_write.get_cursor.return_value = mock_cursor
        mock_write.get_all_text.return_value = "这是文档内容测试文本"
        
        result = word.get_content_text()
        
        assert result == "这是文档内容测试文本"
        mock_write.get_text_doc.assert_called_once_with(doc=mock_doc)
        mock_write.get_cursor.assert_called_once_with(mock_text_doc)
        mock_write.get_all_text.assert_called_once_with(mock_cursor)
    
    @patch('src.libre_automate_py.word.Lo')
    def test_italicize_all(self, mock_lo):
        """测试将指定短语设为斜体"""
        mock_doc = MagicMock()
        mock_cursor = MagicMock()
        mock_page_cursor = MagicMock()
        mock_searchable = MagicMock()
        mock_search_desc = MagicMock()
        mock_matches = MagicMock()
        mock_match_tr = MagicMock()
        
        word = Word.__new__(Word)
        word.doc = mock_doc
        
        # 设置mock返回值
        word.doc.get_cursor.return_value = mock_cursor
        word.doc.get_view_cursor.return_value = mock_page_cursor
        word.doc.qi.return_value = mock_searchable
        mock_searchable.createSearchDescriptor.return_value = mock_search_desc
        mock_searchable.findAll.return_value = mock_matches
        mock_matches.getCount.return_value = 3
        mock_matches.getByIndex.return_value = mock_match_tr
        mock_lo.qi.return_value = mock_match_tr
        mock_match_tr.getString.return_value = "测试"
        mock_page_cursor.get_page.return_value = 1
        mock_cursor.get_string.return_value = "这是一个测试文档"
        
        result = word.italicize_all("测试")
        
        assert result == 3
        mock_search_desc.setSearchString.assert_called_once_with("测试")
        mock_searchable.findAll.assert_called_once_with(mock_search_desc)
    
    def test_replace_words(self):
        """测试批量替换单词"""
        mock_doc = MagicMock()
        word = Word.__new__(Word)
        word.doc = mock_doc
        
        # Mock replace_word方法
        word.replace_word = MagicMock(side_effect=[2, 1, 3])
        
        old_words = ["old1", "old2", "old3"]
        new_words = ["new1", "new2", "new3"]
        
        result = word.replace_words(old_words, new_words)
        
        assert result == 6  # 2 + 1 + 3
        assert word.replace_word.call_count == 3
    
    @patch('src.libre_automate_py.word.Lo')
    def test_replace_word(self, mock_lo):
        """测试替换单个单词"""
        mock_doc = MagicMock()
        mock_replaceable = MagicMock()
        mock_replace_desc = MagicMock()
        
        word = Word.__new__(Word)
        word.doc = mock_doc
        
        word.doc.qi.return_value = mock_replaceable
        mock_replaceable.createSearchDescriptor.return_value = mock_replace_desc
        mock_lo.qi.return_value = mock_replace_desc
        mock_replaceable.replaceAll.return_value = 5
        
        result = word.replace_word("旧词", "新词")
        
        assert result == 5
        mock_replace_desc.setSearchString.assert_called_once_with("旧词")
        mock_replace_desc.setReplaceString.assert_called_once_with("新词")
        mock_replaceable.replaceAll.assert_called_once_with(mock_replace_desc)


class WordTestData:
    """Word测试数据类"""
    
    @staticmethod
    def get_sample_document_content():
        """获取示例文档内容"""
        return """
        这是一个测试文档。
        
        第一段包含一些基本信息。这里有重复的词语：测试、文档、信息。
        
        第二段包含更多内容。这里也有一些重复的词语：内容、信息、数据。
        
        第三段是总结部分。这里总结了前面的内容和信息。
        
        结束。
        """
    
    @staticmethod
    def get_search_and_replace_scenarios():
        """获取搜索和替换测试场景"""
        return [
            {
                'search_phrase': '测试',
                'expected_matches': 3,
                'replacement': '实验',
                'expected_replacements': 3
            },
            {
                'search_phrase': '信息',
                'expected_matches': 4,
                'replacement': '数据',
                'expected_replacements': 4
            },
            {
                'search_phrase': '内容',
                'expected_matches': 3,
                'replacement': '材料',
                'expected_replacements': 3
            },
            {
                'search_phrase': '不存在的词',
                'expected_matches': 0,
                'replacement': '替换词',
                'expected_replacements': 0
            }
        ]
    
    @staticmethod
    def get_batch_replace_scenarios():
        """获取批量替换测试场景"""
        return [
            {
                'old_words': ['老板', '经理', '员工'],
                'new_words': ['CEO', '总监', '团队成员'],
                'expected_total_replacements': 15
            },
            {
                'old_words': ['公司', '部门', '项目'],
                'new_words': ['企业', '团队', '任务'],
                'expected_total_replacements': 8
            },
            {
                'old_words': ['开发', '测试', '部署'],
                'new_words': ['编程', '验证', '发布'],
                'expected_total_replacements': 12
            }
        ]
    
    @staticmethod
    def get_document_formats():
        """获取文档格式测试数据"""
        return [
            {
                'format': 'docx',
                'filepath': 'test.docx',
                'mime_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            },
            {
                'format': 'doc',
                'filepath': 'legacy.doc',
                'mime_type': 'application/msword'
            },
            {
                'format': 'odt',
                'filepath': 'libreoffice.odt',
                'mime_type': 'application/vnd.oasis.opendocument.text'
            },
            {
                'format': 'rtf',
                'filepath': 'richtext.rtf',
                'mime_type': 'application/rtf'
            }
        ]
    
    @staticmethod
    def get_text_formatting_scenarios():
        """获取文本格式化测试场景"""
        return [
            {
                'text': '重要文本',
                'format_type': 'bold',
                'expected_result': True
            },
            {
                'text': '斜体文本',
                'format_type': 'italic',
                'expected_result': True
            },
            {
                'text': '下划线文本',
                'format_type': 'underline',
                'expected_result': True
            },
            {
                'text': '删除线文本',
                'format_type': 'strikethrough',
                'expected_result': True
            }
        ]
    
    @staticmethod
    def get_file_operation_scenarios():
        """获取文件操作测试场景"""
        return [
            {
                'operation': 'open_existing',
                'filepath': 'existing_document.docx',
                'read_only': True,
                'visible': False,
                'expected_success': True
            },
            {
                'operation': 'create_new',
                'filepath': None,
                'read_only': False,
                'visible': True,
                'expected_success': True
            },
            {
                'operation': 'save_as',
                'filepath': 'new_document.docx',
                'read_only': False,
                'visible': True,
                'expected_success': True
            }
        ]
    
    @staticmethod
    def get_error_scenarios():
        """获取错误场景测试数据"""
        return [
            {
                'error_type': 'file_not_found',
                'filepath': 'nonexistent.docx',
                'expected_exception': FileNotFoundError
            },
            {
                'error_type': 'permission_denied',
                'filepath': 'readonly.docx',
                'expected_exception': PermissionError
            },
            {
                'error_type': 'invalid_format',
                'filepath': 'invalid.txt',
                'expected_exception': ValueError
            },
            {
                'error_type': 'corrupted_file',
                'filepath': 'corrupted.docx',
                'expected_exception': RuntimeError
            }
        ]
    
    @staticmethod
    def get_complex_document_content():
        """获取复杂文档内容测试数据"""
        return {
            'title': '年度报告',
            'sections': [
                {
                    'title': '执行摘要',
                    'content': '本年度公司业绩良好，销售额增长15%，利润增长12%。'
                },
                {
                    'title': '财务状况',
                    'content': '资产总额1000万元，负债300万元，净资产700万元。现金流充足。'
                },
                {
                    'title': '市场分析',
                    'content': '市场竞争激烈，但公司产品具有竞争优势。客户满意度达到95%。'
                },
                {
                    'title': '未来展望',
                    'content': '预计明年销售额将增长20%，计划开拓新市场，推出新产品。'
                }
            ],
            'footer': '报告日期：2024年12月31日'
        }
    
    @staticmethod
    def get_search_patterns():
        """获取搜索模式测试数据"""
        return [
            {
                'pattern': r'\d+%',
                'description': '百分比数字',
                'expected_matches': ['15%', '12%', '95%', '20%']
            },
            {
                'pattern': r'\d+万元',
                'description': '金额（万元）',
                'expected_matches': ['1000万元', '300万元', '700万元']
            },
            {
                'pattern': r'公司|企业',
                'description': '公司相关词汇',
                'expected_matches': ['公司', '公司', '公司']
            },
            {
                'pattern': r'\d{4}年\d{1,2}月\d{1,2}日',
                'description': '日期格式',
                'expected_matches': ['2024年12月31日']
            }
        ]
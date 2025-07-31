import pytest
import unittest.mock as mock
from unittest.mock import patch, MagicMock
import threading
import time
from src.libre_automate_py.officeLoader import OfficeLoader


class TestOfficeLoader:
    """测试OfficeLoader类的单例模式和线程安全性"""
    
    def setup_method(self):
        """每个测试方法前重置单例"""
        OfficeLoader._instance = None
        OfficeLoader._loader = None
    
    def teardown_method(self):
        """每个测试方法后清理资源"""
        if OfficeLoader._instance:
            OfficeLoader.close()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_singleton_creation(self, mock_lo):
        """测试单例模式的创建"""
        mock_loader = MagicMock()
        mock_lo.load_office.return_value = mock_loader
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        # 创建第一个实例
        instance1 = OfficeLoader()
        
        # 创建第二个实例
        instance2 = OfficeLoader()
        
        # 验证是同一个实例
        assert instance1 is instance2
        
        # 验证Lo.load_office只被调用一次
        mock_lo.load_office.assert_called_once()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_get_loader_success(self, mock_lo):
        """测试成功获取loader"""
        mock_loader = MagicMock()
        mock_lo.load_office.return_value = mock_loader
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        # 创建实例
        OfficeLoader()
        
        # 获取loader
        loader = OfficeLoader.get_loader()
        
        assert loader is mock_loader
    
    def test_get_loader_not_initialized(self):
        """测试在未初始化时获取loader抛出异常"""
        with pytest.raises(RuntimeError, match="OfficeLoader instance not initialized"):
            OfficeLoader.get_loader()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_close_office(self, mock_lo):
        """测试关闭Office连接"""
        mock_loader = MagicMock()
        mock_lo.load_office.return_value = mock_loader
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        # 创建实例
        instance = OfficeLoader()
        
        # 关闭连接
        OfficeLoader.close()
        
        # 验证状态被重置
        assert OfficeLoader._instance is None
        assert OfficeLoader._loader is None
        
        # 验证Lo.close_office被调用
        mock_lo.close_office.assert_called_once()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_close_not_initialized(self, mock_lo):
        """测试在未初始化时关闭不会报错"""
        OfficeLoader.close()
        
        # 验证Lo.close_office没有被调用
        mock_lo.close_office.assert_not_called()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_context_manager(self, mock_lo):
        """测试上下文管理器"""
        mock_loader = MagicMock()
        mock_lo.load_office.return_value = mock_loader
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        # 使用上下文管理器
        with OfficeLoader.context() as loader:
            assert loader is mock_loader
        
        # 验证退出时关闭了连接
        mock_lo.close_office.assert_called_once()
        assert OfficeLoader._instance is None
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_thread_safety(self, mock_lo):
        """测试多线程环境下的线程安全性"""
        mock_loader = MagicMock()
        mock_lo.load_office.return_value = mock_loader
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        instances = []
        errors = []
        
        def create_instance():
            try:
                instance = OfficeLoader()
                instances.append(instance)
            except Exception as e:
                errors.append(e)
        
        # 创建多个线程同时创建实例
        threads = []
        for _ in range(10):
            thread = threading.Thread(target=create_instance)
            threads.append(thread)
        
        # 启动所有线程
        for thread in threads:
            thread.start()
        
        # 等待所有线程完成
        for thread in threads:
            thread.join()
        
        # 验证没有错误
        assert len(errors) == 0
        
        # 验证所有实例都是同一个
        assert len(instances) == 10
        for instance in instances:
            assert instance is instances[0]
        
        # 验证Lo.load_office只被调用一次
        mock_lo.load_office.assert_called_once()
    
    @patch('src.libre_automate_py.officeLoader.Lo')
    def test_reinitialize_after_close(self, mock_lo):
        """测试关闭后可以重新初始化"""
        mock_loader1 = MagicMock()
        mock_loader2 = MagicMock()
        mock_lo.load_office.side_effect = [mock_loader1, mock_loader2]
        mock_lo.ConnectSocket.return_value = MagicMock()
        
        # 第一次创建
        instance1 = OfficeLoader()
        loader1 = OfficeLoader.get_loader()
        
        # 关闭
        OfficeLoader.close()
        
        # 第二次创建
        instance2 = OfficeLoader()
        loader2 = OfficeLoader.get_loader()
        
        # 验证是不同的loader
        assert loader1 is mock_loader1
        assert loader2 is mock_loader2
        
        # 验证Lo.load_office被调用两次
        assert mock_lo.load_office.call_count == 2


# 测试数据
class OfficeLoaderTestData:
    """OfficeLoader测试数据类"""
    
    @staticmethod
    def get_mock_loader_config():
        """获取模拟loader配置"""
        return {
            'connection_type': 'socket',
            'timeout': 30,
            'visible': True
        }
    
    @staticmethod
    def get_thread_test_scenarios():
        """获取线程测试场景"""
        return [
            {'thread_count': 5, 'expected_instances': 5},
            {'thread_count': 10, 'expected_instances': 10},
            {'thread_count': 20, 'expected_instances': 20},
        ]
    
    @staticmethod
    def get_error_scenarios():
        """获取错误场景测试数据"""
        return [
            {
                'exception': ConnectionError("Connection failed"),
                'expected_error': ConnectionError
            },
            {
                'exception': RuntimeError("Office not available"),
                'expected_error': RuntimeError
            },
            {
                'exception': TimeoutError("Connection timeout"),
                'expected_error': TimeoutError
            }
        ]
import threading
from contextlib import contextmanager
from typing import Optional
from ooodev.loader import Lo

class OfficeLoader:
    _instance: Optional["OfficeLoader"] = None
    _lock = threading.Lock()
    _loader = None

    def __new__(cls):
        # 双重检查锁确保线程安全
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    # 初始化 Office 连接
                    cls._loader = Lo.load_office(Lo.ConnectSocket())
        return cls._instance

    @classmethod
    def get_loader(cls):
        if cls._instance is None:
            raise RuntimeError("OfficeLoader instance not initialized")
        return cls._loader

    @classmethod
    def close(cls):
        if cls._instance is not None:
            Lo.close_office()
            cls._instance = None  # 允许重新初始化
            cls._loader = None

    @classmethod
    @contextmanager
    def context(cls):
        """上下文管理，自动处理资源"""
        try:
            yield cls.get_loader()
        finally:
            cls.close()
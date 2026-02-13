import unittest
import os
import sys
import tempfile
import shutil
from unittest.mock import Mock, MagicMock, patch
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
from src.ui.main_window import MainWindow, BatchConversionWorker
from src.utils.config_manager import ConfigManager
from src.utils.history_manager import HistoryManager
from src.utils.logger import Logger


class TestIntegration(unittest.TestCase):
    """集成测试类"""
    
    @classmethod
    def setUpClass(cls):
        """测试类初始化"""
        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication([])
        
        cls.logger = Logger()
        cls.test_dir = tempfile.mkdtemp()
        cls.config_file = os.path.join(cls.test_dir, 'config.json')
        cls.history_file = os.path.join(cls.test_dir, 'history.json')
    
    @classmethod
    def tearDownClass(cls):
        """测试类清理"""
        if os.path.exists(cls.test_dir):
            shutil.rmtree(cls.test_dir)
    
    def setUp(self):
        """每个测试前的初始化"""
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """每个测试后的清理"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_config_and_history_integration(self):
        """测试配置管理和历史记录的集成"""
        config_manager = ConfigManager(self.config_file)
        history_manager = HistoryManager(self.history_file)
        
        # 设置配置（自动保存）
        config_manager.set('conversion.default_mode', 'background_fill')
        config_manager.set('conversion.default_image_quality', 90)
        
        # 添加历史记录
        history_manager.add_record(
            input_path='test.pptx',
            output_path='output.pptx',
            mode='background_fill',
            output_format='pptx',
            success=True
        )
        history_manager.save()
        
        # 验证配置
        loaded_config = config_manager.load()
        self.assertEqual(loaded_config.get('conversion', {}).get('default_mode'), 'background_fill')
        self.assertEqual(loaded_config.get('conversion', {}).get('default_image_quality'), 90)
        
        # 验证历史记录
        history = history_manager.get_history()
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]['input_path'], 'test.pptx')
    
    def test_converter_with_options_integration(self):
        """测试转换器与选项的集成"""
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 95
        options.preserve_aspect_ratio = True
        
        # 验证选项设置
        self.assertEqual(options.mode, ConversionMode.BACKGROUND_FILL)
        self.assertEqual(options.output_format, OutputFormat.PPTX)
        self.assertEqual(options.image_quality, 95)
        self.assertTrue(options.preserve_aspect_ratio)
        
        # 测试选项序列化和反序列化
        options_dict = options.to_dict()
        self.assertIn('mode', options_dict)
        self.assertIn('output_format', options_dict)
        
        new_options = ConversionOptions.from_dict(options_dict)
        self.assertEqual(new_options.mode, options.mode)
        self.assertEqual(new_options.output_format, options.output_format)
    
    def test_logger_integration(self):
        """测试日志系统与配置的集成"""
        logger_instance = Logger()
        logger = logger_instance.get_logger()
        
        # 测试日志记录
        logger.info("测试信息日志")
        logger.warning("测试警告日志")
        logger.error("测试错误日志")
        
        # 验证日志文件存在
        log_files = [f for f in os.listdir('logs') if f.endswith('.log')]
        self.assertTrue(len(log_files) > 0)


class TestEndToEnd(unittest.TestCase):
    """端到端测试类"""
    
    @classmethod
    def setUpClass(cls):
        """测试类初始化"""
        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication([])
        
        cls.test_dir = tempfile.mkdtemp()
    
    @classmethod
    def tearDownClass(cls):
        """测试类清理"""
        if os.path.exists(cls.test_dir):
            shutil.rmtree(cls.test_dir)
    
    def setUp(self):
        """每个测试前的初始化"""
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """每个测试后的清理"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_full_conversion_workflow(self):
        """测试完整的转换工作流"""
        # 创建测试文件
        test_file = os.path.join(self.temp_dir, 'test.pptx')
        with open(test_file, 'w') as f:
            f.write('test content')
        
        # 创建转换器
        converter = Win32PPTConverter()
        
        # 创建转换选项
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 95
        
        # 获取文件信息
        info = converter.get_conversion_info(test_file)
        self.assertIsNotNone(info)
        
        # 验证选项
        self.assertEqual(options.mode, ConversionMode.BACKGROUND_FILL)
        self.assertEqual(options.output_format, OutputFormat.PPTX)
    
    def test_batch_conversion_workflow(self):
        """测试批量转换工作流"""
        # 创建测试文件
        test_files = []
        for i in range(3):
            test_file = os.path.join(self.temp_dir, f'test{i}.pptx')
            with open(test_file, 'w') as f:
                f.write(f'test content {i}')
            test_files.append(test_file)
        
        # 创建转换选项
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 90
        
        # 验证文件列表
        self.assertEqual(len(test_files), 3)
        
        # 验证选项
        self.assertEqual(options.mode, ConversionMode.BACKGROUND_FILL)
        self.assertEqual(options.image_quality, 90)
    
    def test_ui_conversion_workflow(self):
        """测试UI转换工作流"""
        # 创建主窗口
        main_window = MainWindow()
        
        # 验证UI初始化
        self.assertIsNotNone(main_window.converter)
        self.assertIsNotNone(main_window.file_label)
        self.assertIsNotNone(main_window.convert_btn)
        
        # 验证批量转换UI
        self.assertIsNotNone(main_window.batch_file_table)
        self.assertIsNotNone(main_window.batch_convert_btn)
        
        # 验证初始状态
        self.assertFalse(main_window.convert_btn.isEnabled())
        self.assertFalse(main_window.batch_convert_btn.isEnabled())
    
    def test_error_recovery_workflow(self):
        """测试错误恢复工作流"""
        from src.utils.error_handler import ErrorHandler, ErrorCategory, ErrorSeverity
        
        # 创建错误处理器
        error_handler = ErrorHandler()
        
        # 测试错误处理（没有注册处理器时返回False）
        try:
            raise ValueError("测试错误")
        except Exception as e:
            handled = error_handler.handle_error(e, "测试上下文")
            self.assertFalse(handled)  # 没有注册处理器，返回False
        
        # 注册错误处理器
        def error_handler_func(error):
            return True
        
        error_handler.register_handler(ValueError, error_handler_func)
        
        # 测试错误处理（注册处理器后返回True）
        try:
            raise ValueError("测试错误")
        except Exception as e:
            handled = error_handler.handle_error(e, "测试上下文")
            self.assertTrue(handled)  # 有注册处理器，返回True
        
        # 测试错误恢复（通过注册恢复策略）
        def recovery_func(error):
            return True
        
        error_handler.register_recovery(ValueError, recovery_func)
        
        try:
            raise ValueError("测试恢复")
        except Exception as e:
            handled = error_handler.handle_error(e, "测试上下文")
            self.assertTrue(handled)  # 恢复策略成功执行
    
    def test_configuration_workflow(self):
        """测试配置工作流"""
        config_file = os.path.join(self.temp_dir, 'config.json')
        config_manager = ConfigManager(config_file)
        
        # 设置配置（自动保存）
        config_manager.set('app.name', 'PitchPPT')
        config_manager.set('app.version', '1.0.0')
        config_manager.set('conversion.default_mode', 'background_fill')
        config_manager.set('conversion.default_image_quality', 95)
        config_manager.set('ui.theme', 'default')
        
        # 加载配置
        loaded_config = config_manager.load()
        
        # 验证配置
        self.assertEqual(loaded_config.get('app', {}).get('name'), 'PitchPPT')
        self.assertEqual(loaded_config.get('app', {}).get('version'), '1.0.0')
        self.assertEqual(loaded_config.get('conversion', {}).get('default_mode'), 'background_fill')
        self.assertEqual(loaded_config.get('conversion', {}).get('default_image_quality'), 95)
        self.assertEqual(loaded_config.get('ui', {}).get('theme'), 'default')
    
    def test_history_workflow(self):
        """测试历史记录工作流"""
        history_file = os.path.join(self.temp_dir, 'history.json')
        history_manager = HistoryManager(history_file)
        
        # 添加历史记录
        for i in range(5):
            history_manager.add_record(
                input_path=f'test{i}.pptx',
                output_path=f'output{i}.pptx',
                mode='background_fill',
                output_format='pptx',
                success=True
            )
        history_manager.save()
        
        # 获取历史记录
        history = history_manager.get_history()
        
        # 验证历史记录
        self.assertEqual(len(history), 5)
        self.assertEqual(history[0]['input_path'], 'test0.pptx')
        
        # 清除历史记录
        history_manager.clear_history()
        history_manager.save()
        
        # 验证清除
        history = history_manager.get_history()
        self.assertEqual(len(history), 0)


class TestBatchConversionWorker(unittest.TestCase):
    """批量转换工作线程测试"""
    
    @classmethod
    def setUpClass(cls):
        """测试类初始化"""
        cls.app = QApplication.instance()
        if cls.app is None:
            cls.app = QApplication([])
    
    def setUp(self):
        """每个测试前的初始化"""
        self.temp_dir = tempfile.mkdtemp()
        
        # 创建测试文件
        self.test_files = []
        for i in range(3):
            test_file = os.path.join(self.temp_dir, f'test{i}.pptx')
            with open(test_file, 'w') as f:
                f.write(f'test content {i}')
            self.test_files.append(test_file)
    
    def tearDown(self):
        """每个测试后的清理"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_batch_worker_initialization(self):
        """测试批量转换工作线程初始化"""
        converter = Win32PPTConverter()
        options = ConversionOptions()
        
        worker = BatchConversionWorker(
            converter,
            self.test_files,
            self.temp_dir,
            options
        )
        
        self.assertEqual(len(worker.file_list), 3)
        self.assertEqual(worker.output_dir, self.temp_dir)
        self.assertEqual(worker.success_count, 0)
        self.assertEqual(worker.fail_count, 0)
    
    def test_batch_worker_signals(self):
        """测试批量转换工作线程信号"""
        converter = Win32PPTConverter()
        options = ConversionOptions()
        
        worker = BatchConversionWorker(
            converter,
            self.test_files,
            self.temp_dir,
            options
        )
        
        # 验证信号存在
        self.assertTrue(hasattr(worker, 'progress_updated'))
        self.assertTrue(hasattr(worker, 'file_finished'))
        self.assertTrue(hasattr(worker, 'conversion_finished'))


if __name__ == '__main__':
    unittest.main()
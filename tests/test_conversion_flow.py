"""
转换流程测试脚本
在 tests 目录下进行规范测试
"""
import sys
import os
import unittest
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt

from src.ui.main_window import MainWindow
from src.core.converter import ConversionOptions, OutputFormat, ConversionMode


class TestConversionFlow(unittest.TestCase):
    """测试完整的转换流程"""
    
    @classmethod
    def setUpClass(cls):
        """测试类初始化"""
        cls.app = QApplication(sys.argv)
        cls.test_pptx = Path(__file__).parent / "TCTSlide.pptx"
        cls.output_dir = Path(__file__).parent / "test_output"
        cls.output_dir.mkdir(exist_ok=True)
        
    def setUp(self):
        """每个测试用例初始化"""
        self.window = MainWindow()
        
    def tearDown(self):
        """每个测试用例清理"""
        self.window.close()
        self.window.deleteLater()
        
    def test_01_basic_ui_load(self):
        """测试UI基本加载"""
        self.assertIsNotNone(self.window)
        self.assertIn("PitchPPT", self.window.windowTitle())
        print("✓ UI加载测试通过")
        
    def test_02_step1_file_selection(self):
        """测试第一步文件选择"""
        if not self.test_pptx.exists():
            self.skipTest(f"测试文件不存在: {self.test_pptx}")
            
        # 模拟添加文件
        self.window.add_file(str(self.test_pptx))
        self.assertEqual(self.window.current_input_file, str(self.test_pptx))
        self.assertTrue(self.window.step1_next_btn.isEnabled())
        print("✓ 文件选择测试通过")
        
    def test_03_step2_settings(self):
        """测试第二步设置页面"""
        if not self.test_pptx.exists():
            self.skipTest(f"测试文件不存在: {self.test_pptx}")
            
        # 添加文件并进入第二步
        self.window.add_file(str(self.test_pptx))
        self.window.go_to_step2()
        
        # 验证第二步显示
        self.assertEqual(self.window.content_stack.currentIndex(), 1)
        
        # 测试设置选项
        self.assertIsNotNone(self.window.output_format_combo)
        self.assertIsNotNone(self.window.mode_combo)
        self.assertIsNotNone(self.window.quality_slider)
        print("✓ 设置页面测试通过")
        
    def test_04_conversion_options_creation(self):
        """测试转换选项创建"""
        options = ConversionOptions()
        options.output_format = OutputFormat.PPTX
        options.mode = ConversionMode.BACKGROUND_FILL
        options.image_quality = 95
        
        self.assertEqual(options.output_format, OutputFormat.PPTX)
        self.assertEqual(options.mode, ConversionMode.BACKGROUND_FILL)
        self.assertEqual(options.image_quality, 95)
        print("✓ 转换选项创建测试通过")
        
    def test_05_slide_range_parsing(self):
        """测试幻灯片范围解析"""
        # 测试各种范围格式
        test_cases = [
            ("1-5", (1, 5)),
            ("1-5,8", (1, 8)),
            ("1-5,8,10-12", (1, 12)),
            ("5", (5, 5)),
        ]
        
        for range_str, expected in test_cases:
            result = self.window._parse_slide_range(range_str)
            self.assertEqual(result, expected, f"解析 {range_str} 失败")
            
        print("✓ 幻灯片范围解析测试通过")
        
    def test_06_advanced_options(self):
        """测试高级选项设置"""
        if not self.test_pptx.exists():
            self.skipTest(f"测试文件不存在: {self.test_pptx}")
            
        # 添加文件并进入第二步
        self.window.add_file(str(self.test_pptx))
        self.window.go_to_step2()
        
        # 验证高级选项组件存在
        self.assertIsNotNone(self.window.image_format_combo)
        self.assertIsNotNone(self.window.dpi_combo)
        self.assertIsNotNone(self.window.compression_combo)
        
        # 测试选项值
        self.window.quality_slider.setValue(90)
        self.assertEqual(self.window.quality_slider.value(), 90)
        
        print("✓ 高级选项测试通过")


def run_tests():
    """运行测试"""
    print("=" * 60)
    print("PitchPPT 转换流程测试")
    print("=" * 60)
    
    # 创建测试套件
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestConversionFlow)
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # 输出结果
    print("\n" + "=" * 60)
    if result.wasSuccessful():
        print("✅ 所有测试通过！")
    else:
        print(f"❌ 测试失败: {len(result.failures)} 个失败, {len(result.errors)} 个错误")
    print("=" * 60)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)

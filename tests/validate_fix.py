import sys
import os

# 添加项目根目录到sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

def test_imports():
    """测试所有必要的导入"""
    print("测试导入...")
    try:
        from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
        print("✅ 核心模块导入成功")
    except Exception as e:
        print(f"❌ 核心模块导入失败: {e}")
        return False
    
    try:
        from src.ui.fixed_main_window import MainWindow
        print("✅ UI模块导入成功")
    except Exception as e:
        print(f"❌ UI模块导入失败: {e}")
        return False
    
    try:
        from src.utils.logger import Logger
        print("✅ 日志模块导入成功")
    except Exception as e:
        print(f"❌ 日志模块导入失败: {e}")
        return False
    
    try:
        from src.utils.config_manager import ConfigManager
        print("✅ 配置管理模块导入成功")
    except Exception as e:
        print(f"❌ 配置管理模块导入失败: {e}")
        return False
    
    try:
        from src.utils.history_manager import HistoryManager
        print("✅ 历史记录模块导入成功")
    except Exception as e:
        print(f"❌ 历史记录模块导入失败: {e}")
        return False
    
    return True

def test_converter():
    """测试转换器功能"""
    print("\n测试转换器...")
    try:
        from src.core import Win32PPTConverter
        converter = Win32PPTConverter()
        
        # 测试初始化
        success = converter._initialize_powerpoint()
        if success:
            print("✅ 转换器初始化成功")
        else:
            print("❌ 转换器初始化失败")
            return False
        
        return True
    except Exception as e:
        print(f"❌ 转换器测试失败: {e}")
        return False

def test_ui_creation():
    """测试UI创建"""
    print("\n测试UI创建...")
    try:
        from PyQt5.QtWidgets import QApplication
        from src.ui.fixed_main_window import MainWindow
        import sys
        
        # 创建一个临时的应用程序实例进行测试
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)
        
        window = MainWindow()
        print("✅ UI创建成功")
        print(f"  - 窗口标题: {window.windowTitle()}")
        print(f"  - 转换器存在: {window.converter is not None}")
        
        return True
    except Exception as e:
        print(f"❌ UI创建失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_options():
    """测试转换选项"""
    print("\n测试转换选项...")
    try:
        from src.core import ConversionOptions, ConversionMode, OutputFormat
        
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 95
        
        is_valid, msg = options.validate()
        if is_valid:
            print("✅ 转换选项验证成功")
            return True
        else:
            print(f"❌ 转换选项验证失败: {msg}")
            return False
    except Exception as e:
        print(f"❌ 转换选项测试失败: {e}")
        return False

def main():
    print("=" * 60)
    print("PitchPPT 功能测试")
    print("=" * 60)
    
    tests = [
        ("导入测试", test_imports),
        ("转换器测试", test_converter),
        ("UI创建测试", test_ui_creation),
        ("选项测试", test_options),
    ]
    
    results = []
    for name, test_func in tests:
        result = test_func()
        results.append((name, result))
    
    print("\n" + "=" * 60)
    print("测试结果汇总:")
    print("=" * 60)
    
    passed = 0
    for name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{name:.<30} {status}")
        if result:
            passed += 1
    
    print("=" * 60)
    print(f"总计: {len(results)}, 通过: {passed}, 失败: {len(results)-passed}")
    print(f"成功率: {passed/len(results)*100:.1f}%")
    print("=" * 60)
    
    return passed == len(results)

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
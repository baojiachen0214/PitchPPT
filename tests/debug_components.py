import os
import sys
import tempfile
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

# 添加项目根目录到sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

# 设置应用属性（必须在创建QApplication之前）
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

app = QApplication(sys.argv)

def test_basic_imports():
    """测试基本导入"""
    print("测试1: 基本导入...")
    try:
        from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
        print("✅ 核心模块导入成功")
        
        from src.ui.main_window import MainWindow
        print("✅ UI模块导入成功")
        
        from src.utils.logger import Logger
        print("✅ 工具模块导入成功")
        
        return True
    except Exception as e:
        print(f"❌ 导入失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_converter_creation():
    """测试转换器创建"""
    print("\n测试2: 转换器创建...")
    try:
        from src.core import Win32PPTConverter
        converter = Win32PPTConverter()
        print("✅ 转换器创建成功")
        return True
    except Exception as e:
        print(f"❌ 转换器创建失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_converter_initialization():
    """测试转换器初始化"""
    print("\n测试3: 转换器初始化...")
    try:
        from src.core import Win32PPTConverter
        converter = Win32PPTConverter()
        
        # 尝试初始化PowerPoint
        success = converter._initialize_powerpoint()
        print(f"✅ 转换器初始化: {'成功' if success else '失败'}")
        return success
    except Exception as e:
        print(f"❌ 转换器初始化失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_ui_creation():
    """测试UI创建"""
    print("\n测试4: UI创建...")
    try:
        from src.ui.main_window import MainWindow
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

def test_conversion_options():
    """测试转换选项"""
    print("\n测试5: 转换选项...")
    try:
        from src.core import ConversionOptions, ConversionMode, OutputFormat
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 95
        
        # 测试验证
        is_valid, error_msg = options.validate()
        print(f"✅ 转换选项验证: {'成功' if is_valid else '失败'}")
        if not is_valid:
            print(f"  - 错误: {error_msg}")
        return is_valid
    except Exception as e:
        print(f"❌ 转换选项测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_batch_worker():
    """测试批量转换工作线程"""
    print("\n测试6: 批量转换工作线程...")
    try:
        from src.ui.main_window import BatchConversionWorker
        from src.core import Win32PPTConverter, ConversionOptions
        
        converter = Win32PPTConverter()
        options = ConversionOptions()
        
        # 创建临时文件用于测试
        with tempfile.TemporaryDirectory() as temp_dir:
            # 创建一个假的PPT文件用于测试
            test_file = os.path.join(temp_dir, "test.pptx")
            with open(test_file, "wb") as f:
                f.write(b"PK\x03\x04")  # ZIP文件头，模拟PPTX文件
                
            worker = BatchConversionWorker(
                converter,
                [test_file],
                temp_dir,
                options
            )
            print("✅ 批量转换工作线程创建成功")
            print(f"  - 文件数量: {len(worker.file_list)}")
            print(f"  - 输出目录: {worker.output_dir}")
            return True
    except Exception as e:
        print(f"❌ 批量转换工作线程创建失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_error_handling():
    """测试错误处理"""
    print("\n测试7: 错误处理...")
    try:
        from src.utils.error_handler import ErrorHandler
        error_handler = ErrorHandler()
        
        # 注册错误处理器
        def test_handler(error):
            return True
        
        error_handler.register_handler(Exception, test_handler)
        
        # 测试错误处理
        try:
            raise ValueError("测试错误")
        except Exception as e:
            handled = error_handler.handle_error(e, "测试上下文")
            print(f"✅ 错误处理: {'成功' if handled else '失败'}")
            return handled
            
    except Exception as e:
        print(f"❌ 错误处理测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    print("=" * 60)
    print("PitchPPT 组件完整性测试")
    print("=" * 60)
    
    results = [
        ("基本导入", test_basic_imports()),
        ("转换器创建", test_converter_creation()),
        ("转换器初始化", test_converter_initialization()),
        ("UI创建", test_ui_creation()),
        ("转换选项", test_conversion_options()),
        ("批量转换工作线程", test_batch_worker()),
        ("错误处理", test_error_handling())
    ]
    
    # 输出测试结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{test_name:.<30} {status}")
    
    print("=" * 60)
    print(f"总计: {total} 个测试")
    print(f"通过: {passed} 个")
    print(f"失败: {total - passed} 个")
    print(f"成功率: {passed/total*100:.1f}%")
    print("=" * 60)
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
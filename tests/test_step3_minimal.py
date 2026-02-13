"""
最小化测试 - 验证 go_to_step3 闪退问题
"""
import sys
import os
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer

from src.ui.main_window import MainWindow


def test_go_to_step3():
    """测试进入第三步是否闪退"""
    print("=" * 60)
    print("最小化测试: 验证 go_to_step3 闪退问题")
    print("=" * 60)
    
    # 创建应用
    app = QApplication(sys.argv)
    
    # 创建主窗口
    print("1. 创建主窗口...")
    window = MainWindow()
    print("   ✓ 主窗口创建成功")
    
    # 模拟添加文件
    test_pptx = Path(__file__).parent / "TCTSlide.pptx"
    if not test_pptx.exists():
        print(f"   ⚠ 测试文件不存在: {test_pptx}")
        print("   跳过文件测试，仅测试UI切换")
    else:
        print(f"2. 添加测试文件: {test_pptx.name}")
        window.add_file(str(test_pptx))
        print(f"   ✓ 文件添加成功: {window.current_input_file}")
    
    # 进入第二步
    print("3. 进入第二步（设置页面）...")
    window.go_to_step2()
    print(f"   ✓ 第二步切换成功，当前索引: {window.content_stack.currentIndex()}")
    
    # 进入第三步（关键测试）
    print("4. 进入第三步（转换页面）...")
    try:
        window.go_to_step3()
        print(f"   ✓ 第三步切换成功，当前索引: {window.content_stack.currentIndex()}")
        print("   ✅ 测试通过！没有闪退")
        success = True
    except Exception as e:
        print(f"   ❌ 测试失败！发生异常: {e}")
        import traceback
        traceback.print_exc()
        success = False
    
    # 等待一下看是否有延迟崩溃
    print("5. 等待2秒，检查是否有延迟崩溃...")
    timer = QTimer()
    timer.singleShot(2000, app.quit)
    app.exec_()
    print("   ✓ 2秒等待完成，没有延迟崩溃")
    
    print("=" * 60)
    if success:
        print("✅ 所有测试通过！")
    else:
        print("❌ 测试失败！")
    print("=" * 60)
    
    return success


if __name__ == "__main__":
    success = test_go_to_step3()
    sys.exit(0 if success else 1)

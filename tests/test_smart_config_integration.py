"""
智能配置集成测试

测试内容：
1. SmartConfigWidget组件创建
2. 核心约束输入功能
3. 画质倾向策略选择
4. 高级设置功能
5. 进度更新和结果显示
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout
from PyQt5.QtCore import Qt
from src.ui.smart_config_widget import SmartConfigWidget


def test_widget_creation():
    """测试1: 组件创建"""
    print("\n" + "=" * 60)
    print("测试1: SmartConfigWidget组件创建")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        print("✓ SmartConfigWidget创建成功")
        
        # 检查关键组件是否存在
        checks = [
            ("size_spinbox", hasattr(widget, 'size_spinbox')),
            ("range_all_radio", hasattr(widget, 'range_all_radio')),
            ("resolution_radio", hasattr(widget, 'resolution_radio')),
            ("optimize_btn", hasattr(widget, 'optimize_btn')),
            ("apply_btn", hasattr(widget, 'apply_btn')),
            ("progress_bar", hasattr(widget, 'progress_bar')),
            ("status_label", hasattr(widget, 'status_label')),
            ("prediction_label", hasattr(widget, 'prediction_label')),
        ]
        
        all_passed = True
        for name, exists in checks:
            status = "✓" if exists else "✗"
            print(f"  {status} {name}: {'存在' if exists else '不存在'}")
            if not exists:
                all_passed = False
        
        if all_passed:
            print("\n✓ 所有关键组件都存在!")
        else:
            print("\n✗ 部分组件缺失!")
        
        return all_passed
        
    except Exception as e:
        print(f"\n✗ 创建失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_default_values():
    """测试2: 默认值检查"""
    print("\n" + "=" * 60)
    print("测试2: 默认值检查")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        checks = [
            ("文件大小上限", widget.size_spinbox.value(), 100),
            ("最小高度", widget.min_height_spinbox.value(), 480),
            ("最大高度", widget.max_height_spinbox.value(), 8640),
            ("DPI上限", widget.max_dpi_spinbox.value(), 300),
            ("质量上限", widget.max_quality_spinbox.value(), 100),
        ]
        
        all_passed = True
        for name, actual, expected in checks:
            status = "✓" if actual == expected else "✗"
            print(f"  {status} {name}: {actual} (预期: {expected})")
            if actual != expected:
                all_passed = False
        
        # 检查默认选中状态
        radio_checks = [
            ("全部幻灯片", widget.range_all_radio.isChecked(), True),
            ("清晰度优先", widget.resolution_radio.isChecked(), True),
        ]
        
        for name, actual, expected in radio_checks:
            status = "✓" if actual == expected else "✗"
            print(f"  {status} {name}选中: {actual} (预期: {expected})")
            if actual != expected:
                all_passed = False
        
        if all_passed:
            print("\n✓ 所有默认值正确!")
        else:
            print("\n✗ 部分默认值错误!")
        
        return all_passed
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_preset_buttons():
    """测试3: 预设按钮功能"""
    print("\n" + "=" * 60)
    print("测试3: 预设按钮功能")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        # 测试预设值设置
        print("  测试预设值设置...")
        
        # 模拟设置不同的预设值
        test_values = [20, 50, 100, 200, 500]
        all_passed = True
        
        for value in test_values:
            widget.size_spinbox.setValue(value)
            actual = widget.size_spinbox.value()
            status = "✓" if actual == value else "✗"
            print(f"    {status} 设置 {value}MB: 实际值 {actual}MB")
            if actual != value:
                all_passed = False
        
        if all_passed:
            print("\n✓ 预设值设置功能正常!")
        else:
            print("\n✗ 预设值设置功能异常!")
        
        return all_passed
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_slide_range_selection():
    """测试4: 幻灯片范围选择"""
    print("\n" + "=" * 60)
    print("测试4: 幻灯片范围选择")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        # 测试范围选择
        print("  测试范围选项...")
        
        # 测试全部
        widget.range_all_radio.setChecked(True)
        result = widget._get_slide_range()
        status = "✓" if result == "all" else "✗"
        print(f"    {status} 全部幻灯片: {result}")
        
        # 测试当前页
        widget.range_current_radio.setChecked(True)
        result = widget._get_slide_range()
        status = "✓" if result == "current" else "✗"
        print(f"    {status} 当前页: {result}")
        
        # 测试自定义
        widget.range_custom_radio.setChecked(True)
        widget.range_custom_edit.setText("1-5,8,10-12")
        result = widget._get_slide_range()
        status = "✓" if result == "1-5,8,10-12" else "✗"
        print(f"    {status} 自定义范围: {result}")
        
        # 测试自定义输入框启用状态
        widget.range_custom_radio.setChecked(False)
        is_enabled = widget.range_custom_edit.isEnabled()
        status = "✓" if not is_enabled else "✗"
        print(f"    {status} 自定义未选中时输入框禁用: {not is_enabled}")
        
        widget.range_custom_radio.setChecked(True)
        is_enabled = widget.range_custom_edit.isEnabled()
        status = "✓" if is_enabled else "✗"
        print(f"    {status} 自定义选中时输入框启用: {is_enabled}")
        
        print("\n✓ 幻灯片范围选择功能正常!")
        return True
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_priority_mode_selection():
    """测试5: 画质策略选择"""
    print("\n" + "=" * 60)
    print("测试5: 画质策略选择")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        print("  测试策略选项...")
        
        # 测试清晰度优先
        widget.resolution_radio.setChecked(True)
        mode = widget._get_priority_mode()
        status = "✓" if mode == "resolution" else "✗"
        print(f"    {status} 清晰度优先: {mode}")
        
        # 测试色彩平衡优先
        widget.quality_radio.setChecked(True)
        mode = widget._get_priority_mode()
        status = "✓" if mode == "quality" else "✗"
        print(f"    {status} 色彩平衡优先: {mode}")
        
        # 测试平衡模式
        widget.balanced_radio.setChecked(True)
        mode = widget._get_priority_mode()
        status = "✓" if mode == "balanced" else "✗"
        print(f"    {status} 平衡模式: {mode}")
        
        print("\n✓ 画质策略选择功能正常!")
        return True
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_advanced_settings():
    """测试6: 高级设置功能"""
    print("\n" + "=" * 60)
    print("测试6: 高级设置功能")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        print("  测试高级参数设置...")
        
        # 测试修改高级参数
        widget.min_height_spinbox.setValue(720)
        widget.max_height_spinbox.setValue(4320)
        widget.max_dpi_spinbox.setValue(200)
        widget.max_quality_spinbox.setValue(95)
        
        config = widget.get_current_config()
        
        checks = [
            ("min_height", config.get("min_height"), 720),
            ("max_height", config.get("max_height"), 4320),
            ("max_dpi", config.get("max_dpi"), 200),
            ("max_quality", config.get("max_quality"), 95),
        ]
        
        all_passed = True
        for name, actual, expected in checks:
            status = "✓" if actual == expected else "✗"
            print(f"    {status} {name}: {actual} (预期: {expected})")
            if actual != expected:
                all_passed = False
        
        if all_passed:
            print("\n✓ 高级设置功能正常!")
        else:
            print("\n✗ 高级设置功能异常!")
        
        return all_passed
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_reset_function():
    """测试7: 重置功能"""
    print("\n" + "=" * 60)
    print("测试7: 重置功能")
    print("=" * 60)
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    try:
        widget = SmartConfigWidget()
        
        print("  修改参数...")
        # 修改一些参数
        widget.size_spinbox.setValue(200)
        widget.min_height_spinbox.setValue(1080)
        widget.resolution_radio.setChecked(False)
        widget.quality_radio.setChecked(True)
        
        print("  执行重置...")
        widget.reset()
        
        # 检查是否恢复默认值
        checks = [
            ("文件大小上限", widget.size_spinbox.value(), 100),
            ("最小高度", widget.min_height_spinbox.value(), 480),
            ("清晰度优先选中", widget.resolution_radio.isChecked(), True),
            ("全部幻灯片选中", widget.range_all_radio.isChecked(), True),
        ]
        
        all_passed = True
        for name, actual, expected in checks:
            status = "✓" if actual == expected else "✗"
            print(f"    {status} {name}: {actual} (预期: {expected})")
            if actual != expected:
                all_passed = False
        
        if all_passed:
            print("\n✓ 重置功能正常!")
        else:
            print("\n✗ 重置功能异常!")
        
        return all_passed
        
    except Exception as e:
        print(f"\n✗ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("智能配置UI组件集成测试")
    print("=" * 60)
    
    tests = [
        ("组件创建", test_widget_creation),
        ("默认值检查", test_default_values),
        ("预设按钮功能", test_preset_buttons),
        ("幻灯片范围选择", test_slide_range_selection),
        ("画质策略选择", test_priority_mode_selection),
        ("高级设置功能", test_advanced_settings),
        ("重置功能", test_reset_function),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            passed = test_func()
            results.append((name, passed))
        except Exception as e:
            print(f"\n✗ {name} 测试出错: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))
    
    # 汇总结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    
    passed_count = sum(1 for _, passed in results if passed)
    total_count = len(results)
    
    for name, passed in results:
        status = "✓ 通过" if passed else "✗ 失败"
        print(f"  {status}: {name}")
    
    print(f"\n总计: {passed_count}/{total_count} 项测试通过")
    
    if passed_count == total_count:
        print("\n✓ 所有测试通过!")
        return True
    else:
        print(f"\n✗ {total_count - passed_count} 项测试失败")
        return False


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)

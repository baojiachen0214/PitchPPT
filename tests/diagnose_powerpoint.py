"""
PowerPoint COM接口诊断脚本
"""
import sys
import os
import win32com.client
import pythoncom

def check_powerpoint_availability():
    """检查PowerPoint COM接口可用性"""
    print("=" * 60)
    print("PowerPoint COM接口诊断")
    print("=" * 60)
    
    # 检查win32com模块
    try:
        import win32com.client
        print("✅ win32com模块导入成功")
    except ImportError as e:
        print(f"❌ win32com模块导入失败: {e}")
        return False
    
    # 检查pythoncom模块
    try:
        import pythoncom
        print("✅ pythoncom模块导入成功")
    except ImportError as e:
        print(f"❌ pythoncom模块导入失败: {e}")
        return False
    
    # 尝试不同的PowerPoint版本
    powerpoint_versions = [
        "PowerPoint.Application",
        "PowerPoint.Application.16",  # Office 2016+
        "PowerPoint.Application.15",  # Office 2013
        "PowerPoint.Application.14",  # Office 2010
    ]
    
    powerpoint_available = False
    powerpoint_version = ""
    
    for version in powerpoint_versions:
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch(version)
            powerpoint_version = powerpoint.Version
            print(f"✅ PowerPoint {version} 可用 - 版本: {powerpoint_version}")
            powerpoint_available = True
            
            # 测试基本功能
            try:
                presentations = powerpoint.Presentations.Count
                print(f"  - 打开的演示文稿数量: {presentations}")
            except Exception as e:
                print(f"  - 获取演示文稿信息失败: {e}")
            
            # 清理
            try:
                powerpoint.Quit()
            except:
                pass
            
            pythoncom.CoUninitialize()
            break
            
        except Exception as e:
            print(f"❌ PowerPoint {version} 不可用: {e}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    # 如果标准版本都不可用，尝试EnsureDispatch
    if not powerpoint_available:
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
            powerpoint_version = powerpoint.Version
            print(f"✅ PowerPoint (EnsureDispatch) 可用 - 版本: {powerpoint_version}")
            powerpoint_available = True
            
            try:
                powerpoint.Quit()
            except:
                pass
            
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"❌ PowerPoint EnsureDispatch 也失败: {e}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    print("\n" + "=" * 60)
    
    if powerpoint_available:
        print("✅ PowerPoint COM接口诊断通过")
        print(f"检测到的PowerPoint版本: {powerpoint_version}")
        return True
    else:
        print("❌ PowerPoint COM接口诊断失败")
        print("可能的原因:")
        print("1. Microsoft Office未安装")
        print("2. Office版本与Python版本不匹配(32位/64位)")
        print("3. Office安装损坏")
        print("4. 权限问题")
        return False

def test_ppt_export():
    """测试PPT导出功能"""
    print("\n" + "=" * 60)
    print("PPT导出功能测试")
    print("=" * 60)
    
    # 检查是否有测试PPT文件
    test_files = []
    for root, dirs, files in os.walk("."):
        for file in files:
            if file.endswith((".ppt", ".pptx")):
                test_files.append(os.path.join(root, file))
    
    if not test_files:
        print("❌ 未找到测试PPT文件")
        print("请将PPT文件放在项目目录中")
        return False
    
    print(f"找到 {len(test_files)} 个PPT文件:")
    for test_file in test_files:
        print(f"  - {test_file}")
    
    # 使用第一个文件进行测试
    test_file = test_files[0]
    print(f"\n使用文件进行测试: {test_file}")
    
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = False
        powerpoint.DisplayAlerts = False
        
        # 打开PPT文件
        presentation = powerpoint.Presentations.Open(os.path.abspath(test_file))
        slide_count = presentation.Slides.Count
        print(f"✅ 成功打开PPT，共 {slide_count} 张幻灯片")
        
        # 测试导出第一张幻灯片
        temp_dir = "test_export"
        os.makedirs(temp_dir, exist_ok=True)
        
        test_slide = presentation.Slides(1)
        export_path = os.path.join(temp_dir, "test_export.jpg")
        
        try:
            test_slide.Export(export_path, "JPG")
            if os.path.exists(export_path) and os.path.getsize(export_path) > 0:
                file_size = os.path.getsize(export_path) / 1024
                print(f"✅ 幻灯片导出成功: {export_path} ({file_size:.1f} KB)")
            else:
                print("❌ 幻灯片导出失败: 文件为空或不存在")
        except Exception as e:
            print(f"❌ 幻灯片导出失败: {e}")
        
        # 清理
        presentation.Close()
        powerpoint.Quit()
        pythoncom.CoUninitialize()
        
        print("✅ PPT导出功能测试完成")
        return True
        
    except Exception as e:
        print(f"❌ PPT导出功能测试失败: {e}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return False

if __name__ == "__main__":
    # 检查PowerPoint可用性
    powerpoint_ok = check_powerpoint_availability()
    
    if powerpoint_ok:
        # 测试PPT导出功能
        export_ok = test_ppt_export()
        
        print("\n" + "=" * 60)
        if powerpoint_ok and export_ok:
            print("✅ 所有诊断测试通过")
            print("PowerPoint转换功能应该可以正常工作")
        else:
            print("❌ 诊断测试失败")
            print("请检查PowerPoint安装和配置")
    else:
        print("\n❌ PowerPoint不可用，无法进行导出测试")
    
    print("=" * 60)
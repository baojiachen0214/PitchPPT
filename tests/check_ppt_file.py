"""
检查PPT文件路径和权限问题
"""
import os
import win32com.client
import pythoncom
from pathlib import Path

def check_ppt_file():
    """检查PPT文件的可访问性"""
    ppt_path = Path("tests/TCTSlide.pptx")
    
    print("=" * 60)
    print("PPT文件检查")
    print("=" * 60)
    
    # 检查文件是否存在
    if not ppt_path.exists():
        print(f"❌ 文件不存在: {ppt_path}")
        return False
    
    print(f"✅ 文件存在: {ppt_path}")
    
    # 检查文件大小
    file_size = ppt_path.stat().st_size
    print(f"📊 文件大小: {file_size} 字节 ({file_size/1024/1024:.2f} MB)")
    
    # 检查文件权限
    if os.access(ppt_path, os.R_OK):
        print("✅ 文件可读")
    else:
        print("❌ 文件不可读")
        return False
    
    # 检查绝对路径
    abs_path = ppt_path.absolute()
    print(f"📁 绝对路径: {abs_path}")
    
    # 检查路径长度（Windows路径限制）
    if len(str(abs_path)) > 260:
        print("⚠️ 路径长度超过Windows限制，可能有问题")
    else:
        print("✅ 路径长度正常")
    
    # 尝试直接使用PowerPoint打开
    print("\n" + "=" * 60)
    print("PowerPoint COM接口测试")
    print("=" * 60)
    
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        
        # 尝试不同的打开方式
        print("尝试打开方式1: 相对路径")
        try:
            presentation = powerpoint.Presentations.Open(str(ppt_path))
            print("✅ 相对路径打开成功")
            presentation.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception as e:
            print(f"❌ 相对路径打开失败: {e}")
        
        print("\n尝试打开方式2: 绝对路径")
        try:
            presentation = powerpoint.Presentations.Open(str(abs_path))
            print("✅ 绝对路径打开成功")
            presentation.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception as e:
            print(f"❌ 绝对路径打开失败: {e}")
        
        print("\n尝试打开方式3: 使用OpenEx")
        try:
            # 尝试使用不同的打开方法
            presentation = powerpoint.Presentations.Open(str(abs_path), True, True, False)
            print("✅ OpenEx打开成功")
            presentation.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception as e:
            print(f"❌ OpenEx打开失败: {e}")
        
        powerpoint.Quit()
        pythoncom.CoUninitialize()
        
    except Exception as e:
        print(f"❌ PowerPoint初始化失败: {e}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    return False

def check_file_format():
    """检查文件格式"""
    ppt_path = Path("tests/TCTSlide.pptx")
    
    print("\n" + "=" * 60)
    print("文件格式检查")
    print("=" * 60)
    
    # 检查文件头
    try:
        with open(ppt_path, 'rb') as f:
            header = f.read(8)
            print(f"文件头: {header.hex()}")
            
            # PPTX文件头应该是PK开头（ZIP格式）
            if header.startswith(b'PK'):
                print("✅ 文件格式: PPTX (ZIP格式)")
            elif header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
                print("✅ 文件格式: PPT (OLE格式)")
            else:
                print("❌ 未知文件格式")
                return False
    except Exception as e:
        print(f"❌ 读取文件头失败: {e}")
        return False
    
    return True

if __name__ == "__main__":
    # 检查文件格式
    format_ok = check_file_format()
    
    if format_ok:
        # 检查文件可访问性
        access_ok = check_ppt_file()
        
        print("\n" + "=" * 60)
        if format_ok and access_ok:
            print("✅ 文件检查通过")
            print("文件格式和可访问性都正常")
        else:
            print("❌ 文件检查失败")
            print("请检查文件路径、权限和格式")
    else:
        print("\n❌ 文件格式检查失败")
    
    print("=" * 60)
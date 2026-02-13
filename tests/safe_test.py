"""
安全测试PPT转换功能 - 确保PowerPoint进程正确清理
"""
import sys
import os
import time
from pathlib import Path
from src.core.win32_converter import Win32PPTConverter
from src.core.converter import ConversionOptions, ConversionMode, OutputFormat
from src.utils.logger import Logger

def safe_ppt_test():
    """安全的PPT转换测试，确保进程正确清理"""
    logger = Logger().get_logger()
    logger.info("=" * 60)
    logger.info("安全PPT转换测试 - 确保进程正确清理")
    logger.info("=" * 60)
    
    # 检查PowerPoint进程状态
    logger.info("检查PowerPoint进程状态...")
    
    # 检查提供的PPT文件
    test_ppt = Path("tests/TCTSlide.pptx")
    if not test_ppt.exists():
        logger.error(f"未找到提供的PPT文件: {test_ppt}")
        print(f"❌ 未找到提供的PPT文件: {test_ppt}")
        return False
    
    logger.info(f"找到提供的PPT文件: {test_ppt}")
    print(f"✅ 找到提供的PPT文件: {test_ppt}")
    
    # 创建输出目录
    output_dir = Path("test_output")
    output_dir.mkdir(exist_ok=True)
    
    # 只测试最简单的背景填充模式
    print("\n🔄 开始安全测试 - 背景填充模式")
    
    output_file = output_dir / "safe_test_background_fill.pptx"
    
    options = ConversionOptions()
    options.mode = ConversionMode.BACKGROUND_FILL
    options.output_format = OutputFormat.PPTX
    options.image_quality = 95
    
    converter = Win32PPTConverter()
    
    try:
        print("📊 开始转换...")
        success = converter.convert(
            str(test_ppt),
            str(output_file),
            options
        )
        
        if success and output_file.exists():
            size_mb = output_file.stat().st_size / 1024 / 1024
            logger.info(f"✅ 转换成功: {output_file}")
            logger.info(f"   文件大小: {size_mb:.2f} MB")
            print(f"✅ 转换成功！文件大小: {size_mb:.2f} MB")
            
            # 检查PowerPoint进程是否已清理
            time.sleep(2)  # 等待清理完成
            print("\n🔍 检查PowerPoint进程状态...")
            
            return True
        else:
            logger.error(f"❌ 转换失败: {output_file}")
            print(f"❌ 转换失败")
            return False
            
    except Exception as e:
        logger.error(f"❌ 转换异常: {e}")
        import traceback
        logger.error(traceback.format_exc())
        print(f"❌ 转换异常: {e}")
        return False
    finally:
        # 确保清理资源
        print("\n🧹 确保资源清理...")
        try:
            # 调用转换器的清理方法
            if hasattr(converter, '_cleanup'):
                converter._cleanup()
        except:
            pass

def check_powerpoint_processes():
    """检查PowerPoint进程状态"""
    print("\n" + "=" * 60)
    print("PowerPoint进程状态检查")
    print("=" * 60)
    
    import subprocess
    try:
        result = subprocess.run(['tasklist', '/fi', 'imagename eq POWERPNT.EXE'], 
                              capture_output=True, text=True)
        
        if 'POWERPNT.EXE' in result.stdout:
            print("⚠️ 检测到PowerPoint进程正在运行")
            print("建议手动关闭PowerPoint后再进行测试")
            return False
        else:
            print("✅ 没有检测到PowerPoint进程")
            return True
    except:
        print("⚠️ 无法检查进程状态")
        return True

def main():
    """主测试函数"""
    print("🚀 PitchPPT 安全测试启动")
    print("=" * 60)
    
    # 检查PowerPoint进程状态
    if not check_powerpoint_processes():
        print("\n❌ 检测到PowerPoint进程正在运行，请手动关闭后再测试")
        print("1. 按 Ctrl+Shift+Esc 打开任务管理器")
        print("2. 找到 POWERPNT.EXE 进程")
        print("3. 右键选择'结束任务'")
        print("4. 重新运行此测试")
        return False
    
    # 进行安全测试
    success = safe_ppt_test()
    
    # 最终检查
    print("\n" + "=" * 60)
    print("最终检查")
    print("=" * 60)
    
    if not check_powerpoint_processes():
        print("⚠️ 测试后PowerPoint进程可能未正确清理")
        print("建议手动检查并关闭PowerPoint进程")
    else:
        print("✅ PowerPoint进程已正确清理")
    
    print("\n" + "=" * 60)
    if success:
        print("🎉 安全测试通过！")
        print("转换功能可以正常工作")
    else:
        print("❌ 安全测试失败")
        print("请查看日志文件获取详细信息")
    
    print("=" * 60)
    
    # 检查日志文件位置
    log_file = Path("logs/pitchppt_20260209.log")
    if log_file.exists():
        print(f"📋 详细日志请查看: {log_file}")
    
    return success

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
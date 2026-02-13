"""
兼容性更好的PPT转换器 - 借鉴成功案例
"""
import os
import tempfile
import shutil
from pathlib import Path
import win32com.client
import pythoncom
from PIL import Image
import threading
import time
from typing import Dict, Any, List

class CompatiblePPTConverter:
    """
    兼容性更好的PPT转换器
    借鉴成功案例的解决方案
    """
    
    def __init__(self):
        self.powerpoint = None
        self._lock = threading.Lock()
        
    def _initialize_powerpoint(self) -> bool:
        """初始化PowerPoint，借鉴成功案例的谨慎处理"""
        try:
            if self.powerpoint is None:
                pythoncom.CoInitialize()
                
                # 尝试不同的PowerPoint版本
                versions = [
                    "PowerPoint.Application",
                    "PowerPoint.Application.16",
                    "PowerPoint.Application.15", 
                    "PowerPoint.Application.14"
                ]
                
                for version in versions:
                    try:
                        self.powerpoint = win32com.client.Dispatch(version)
                        
                        # 借鉴成功案例：谨慎处理属性
                        try:
                            self.powerpoint.DisplayAlerts = False
                            print(f"✅ 已禁用PowerPoint警告对话框 ({version})")
                        except:
                            print(f"⚠️ 设置DisplayAlerts失败 ({version})")
                        
                        # 借鉴成功案例：检查并设置Visible属性
                        try:
                            current_visible = self.powerpoint.Visible
                            print(f"📊 PowerPoint当前可见状态: {current_visible}")
                            
                            if not current_visible:
                                self.powerpoint.Visible = True
                                print("✅ PowerPoint窗口已设置为可见")
                        except:
                            print("⚠️ 设置Visible属性失败")
                        
                        print(f"✅ PowerPoint COM接口初始化成功: {version}")
                        return True
                        
                    except Exception as e:
                        print(f"❌ 尝试初始化 {version} 失败: {e}")
                        if self.powerpoint:
                            try:
                                self.powerpoint.Quit()
                            except:
                                pass
                        self.powerpoint = None
                        pythoncom.CoUninitialize()
                
                return False
            return True
            
        except Exception as e:
            print(f"❌ PowerPoint初始化失败: {e}")
            return False
    
    def _cleanup(self):
        """清理资源"""
        with self._lock:
            if self.powerpoint:
                try:
                    # 关闭所有打开的演示文稿
                    while self.powerpoint.Presentations.Count > 0:
                        self.powerpoint.Presentations(1).Close()
                    
                    self.powerpoint.Quit()
                    print("✅ PowerPoint实例已关闭")
                except Exception as e:
                    print(f"❌ 关闭PowerPoint时出错: {e}")
                finally:
                    self.powerpoint = None
            
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _validate_image_file(self, image_path: str) -> bool:
        """验证图片文件的有效性"""
        try:
            if not os.path.exists(image_path):
                return False
                
            file_size = os.path.getsize(image_path)
            if file_size < 500:  # 小于500字节
                return False
                
            # 尝试使用PIL打开
            try:
                with Image.open(image_path) as img:
                    width, height = img.size
                    if width < 5 or height < 5:
                        return False
            except:
                # 如果PIL无法打开，但文件大小合理，也认为是有效的
                if file_size > 5000:
                    return True
                return False
                
            return True
        except:
            return os.path.exists(image_path) and os.path.getsize(image_path) > 500
    
    def _export_slide_to_image(self, slide, output_path: str, quality: int = 95) -> bool:
        """导出幻灯片为图片，借鉴成功案例的重试机制"""
        temp_png = output_path + "_temp.png"
        
        # 借鉴成功案例：添加重试机制
        for attempt in range(2):
            try:
                print(f"📷 导出幻灯片为图片，尝试 {attempt + 1}")
                
                # 导出为PNG
                slide.Export(temp_png, "PNG")

                # 验证临时PNG图片
                if self._validate_image_file(temp_png):
                    # 转为JPG
                    with Image.open(temp_png) as img:
                        rgb_img = img.convert("RGB")
                        rgb_img.save(output_path, "JPEG", quality=quality, optimize=True)
                    
                    # 验证JPG
                    if self._validate_image_file(output_path):
                        print(f"✅ 幻灯片JPG转换成功: {output_path}")
                        return True
                    else:
                        print(f"⚠️ 生成的JPG文件无效: {output_path}")
                        if attempt == 0:
                            continue
                else:
                    print(f"⚠️ PNG临时文件导出失败: {temp_png}")
                    
                    # 借鉴成功案例：尝试重新导出一次
                    if attempt == 0:
                        print("🔄 尝试重新导出...")
                        time.sleep(0.5)
                        continue
                    
            except Exception as e:
                print(f"⚠️ 导出幻灯片为图片失败 (尝试 {attempt + 1}): {e}")
                if attempt == 0:
                    continue
        
        # 如果所有重试都失败，尝试直接导出为JPG
        print("🔄 PNG导出失败，尝试直接导出JPG")
        try:
            slide.Export(output_path, "JPG")
            if self._validate_image_file(output_path):
                print(f"✅ 直接导出JPG成功: {output_path}")
                return True
            else:
                print(f"❌ 直接导出的JPG文件无效: {output_path}")
        except Exception as e:
            print(f"❌ 直接导出JPG也失败: {e}")
            
        finally:
            # 清理临时文件
            if os.path.exists(temp_png):
                try:
                    os.remove(temp_png)
                except:
                    pass
        
        return False
    
    def convert_background_fill(self, input_ppt: str, output_ppt: str) -> bool:
        """
        背景填充模式转换
        完全借鉴成功案例的实现
        """
        temp_dir = None
        powerpoint = None
        
        try:
            print("=" * 60)
            print("开始背景填充模式转换")
            print("=" * 60)
            
            # 1. 初始化PowerPoint
            if not self._initialize_powerpoint():
                return False
            
            # 2. 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="ppt_to_image_")
            print(f"📁 创建临时目录: {temp_dir}")
            
            # 3. 打开原PPT
            print(f"📂 打开PPT文件: {input_ppt}")
            
            # 借鉴成功案例：使用绝对路径
            abs_input_path = os.path.abspath(input_ppt)
            print(f"📁 绝对路径: {abs_input_path}")
            
            # 尝试不同的打开方法
            presentation = None
            try:
                presentation = self.powerpoint.Presentations.Open(abs_input_path)
                print("✅ PPT文件打开成功")
            except Exception as e:
                print(f"❌ PPT文件打开失败: {e}")
                
                # 尝试备用方案：使用OpenEx
                try:
                    presentation = self.powerpoint.Presentations.Open(abs_input_path, True, True, False)
                    print("✅ 使用OpenEx打开成功")
                except Exception as e2:
                    print(f"❌ OpenEx也失败: {e2}")
                    return False
            
            slide_count = presentation.Slides.Count
            print(f"📊 成功打开PPT，共 {slide_count} 张幻灯片")
            
            # 获取幻灯片尺寸信息
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            print(f"📐 幻灯片尺寸: {slide_width:.1f} x {slide_height:.1f} 点")
            
            # 4. 导出为图片
            print("🖼️ 开始导出幻灯片为图片...")
            image_files = []
            
            for i in range(1, slide_count + 1):
                jpg_path = os.path.join(temp_dir, f"slide_{i:03d}.jpg")
                print(f"📷 导出幻灯片 {i}/{slide_count}: slide_{i:03d}.jpg")

                try:
                    if self._export_slide_to_image(presentation.Slides(i), jpg_path, 95):
                        image_files.append(jpg_path)
                        print(f"✅ 幻灯片 {i} JPG 转换成功")
                    else:
                        print(f"❌ 幻灯片 {i} JPG 转换失败")
                        
                except Exception as e:
                    print(f"❌ 导出幻灯片 {i} 失败: {e}")
                    continue

            if not image_files:
                print("❌ 错误：没有成功导出任何图片")
                return False
                
            print(f"✅ 成功导出 {len(image_files)} 张JPG图片")
            presentation.Close()
            
            # 5. 重新打开PPT作为模板，设置背景
            print("🎨 重新打开PPT，设置JPG图片为背景...")
            
            # 使用原文件作为模板
            template_presentation = self.powerpoint.Presentations.Open(abs_input_path)
            
            # 确保模板幻灯片数量与图片数量匹配
            template_slide_count = template_presentation.Slides.Count
            image_count = len(image_files)
            print(f"📊 模板幻灯片数量: {template_slide_count}, JPG图片数量: {image_count}")
            
            if template_slide_count != image_count:
                print(f"⚠️ 警告：幻灯片数量({template_slide_count})与图片数量({image_count})不匹配")
                # 如果模板幻灯片少于图片数量，添加幻灯片
                while template_presentation.Slides.Count < image_count:
                    # 复制最后一张幻灯片
                    last_slide = template_presentation.Slides(template_presentation.Slides.Count)
                    new_slide = last_slide.Duplicate()
                    print(f"➕ 添加了新幻灯片，当前总数: {template_presentation.Slides.Count}")
            
            # 6. 处理每张幻灯片
            processed_count = 0
            for i, image_file in enumerate(image_files, 1):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        print(f"🎨 处理幻灯片 {i}/{len(image_files)} (JPG)...")
                        
                        # 借鉴成功案例：关键修复 - 设置FollowMasterBackground为False
                        try:
                            slide.FollowMasterBackground = False
                            print(f"✅ 幻灯片 {i} 已禁用跟随母版背景")
                        except Exception as e:
                            print(f"⚠️ 设置FollowMasterBackground失败: {e}")
                        
                        # 借鉴成功案例：设置为空白版式（避免占位符文本）
                        try:
                            slide.Layout = 12  # ppLayoutBlank = 12
                            print(f"✅ 幻灯片 {i} 已设置为空白版式")
                        except Exception as e:
                            print(f"⚠️ 设置空白版式失败: {e}")
                        
                        # 借鉴成功案例：彻底清空幻灯片内容（包括占位符）
                        try:
                            shape_count = slide.Shapes.Count
                            deleted_count = 0
                            # 从后往前删除所有形状，包括占位符
                            for j in range(shape_count, 0, -1):
                                try:
                                    shape = slide.Shapes(j)
                                    # 删除所有形状，包括占位符
                                    shape.Delete()
                                    deleted_count += 1
                                except:
                                    pass
                            print(f"🗑️ 清空了 {deleted_count} 个元素（包括占位符）")
                        except Exception as e:
                            print(f"⚠️ 清空幻灯片内容时出错: {e}")
                        
                        # 借鉴成功案例：设置背景图片（多方案备用）
                        background_set = False
                        abs_image_path = os.path.abspath(image_file)

                        # 方法1：使用UserPicture设置JPG背景
                        try:
                            slide.Background.Fill.UserPicture(abs_image_path)
                            # 等待一下让设置生效
                            time.sleep(0.1)
                            
                            # 验证背景是否设置成功
                            try:
                                fill_type = slide.Background.Fill.Type
                                if fill_type == 6:  # msoFillPicture = 6
                                    background_set = True
                                    print(f"✅ 方法1成功：幻灯片 {i} JPG背景设置完成")
                                else:
                                    print(f"⚠️ 方法1设置JPG后验证失败")
                            except:
                                print(f"⚠️ 方法1验证背景失败")
                                
                        except Exception as e:
                            print(f"⚠️ 方法1失败：{e}")
                        
                        # 如果UserPicture失败，使用备用方案
                        if not background_set:
                            try:
                                # 获取幻灯片尺寸
                                slide_width = template_presentation.PageSetup.SlideWidth
                                slide_height = template_presentation.PageSetup.SlideHeight
                                
                                # 添加图片铺满整个幻灯片
                                picture = slide.Shapes.AddPicture(abs_image_path, False, True, 0, 0, slide_width, slide_height)
                                # 将图片移到最底层（作为背景）
                                try:
                                    picture.ZOrder(0)  # 发送到底层
                                except:
                                    pass
                                background_set = True
                                print(f"✅ 备用方案成功：幻灯片 {i} JPG图片作为背景添加完成")
                            except Exception as e:
                                print(f"❌ 备用方案失败：{e}")
                        
                        if background_set:
                            processed_count += 1
                        else:
                            print(f"❌ 幻灯片 {i} 背景设置完全失败")
                        
                    except Exception as e:
                        print(f"❌ 设置幻灯片 {i} 背景失败: {e}")
            
            # 7. 保存最终文件
            print(f"💾 保存最终文件: {output_ppt}")
            
            try:
                abs_output_path = os.path.abspath(output_ppt)
                template_presentation.SaveAs(abs_output_path)
                
                if os.path.exists(abs_output_path):
                    file_size = os.path.getsize(abs_output_path) / 1024 / 1024
                    print(f"✅ 文件保存成功: {abs_output_path} ({file_size:.2f} MB)")
                    print(f"✅ 成功处理 {processed_count}/{len(image_files)} 张幻灯片")
                    success = True
                else:
                    print("❌ 文件保存失败")
                    success = False
                    
            except Exception as e:
                print(f"❌ 保存文件失败: {e}")
                success = False
            
            template_presentation.Close()
            return success
            
        except Exception as e:
            print(f"❌ 转换过程中发生异常: {e}")
            import traceback
            print(traceback.format_exc())
            return False
            
        finally:
            # 清理资源
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                    print(f"🗑️ 清理临时目录: {temp_dir}")
                except Exception as e:
                    print(f"⚠️ 清理临时目录失败: {e}")
            
            self._cleanup()

def test_compatible_converter():
    """测试兼容性转换器"""
    print("🚀 测试兼容性PPT转换器")
    
    # 测试文件
    input_ppt = "tests/TCTSlide.pptx"
    output_ppt = "test_output/compatible_test.pptx"
    
    # 确保输出目录存在
    os.makedirs("test_output", exist_ok=True)
    
    converter = CompatiblePPTConverter()
    
    success = converter.convert_background_fill(input_ppt, output_ppt)
    
    print("\n" + "=" * 60)
    if success:
        print("🎉 兼容性转换器测试成功！")
        print("转换功能现在应该可以正常工作")
    else:
        print("❌ 兼容性转换器测试失败")
        print("可能需要进一步调试或使用其他方案")
    print("=" * 60)
    
    return success

if __name__ == "__main__":
    test_compatible_converter()
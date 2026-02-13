from src.core.converter import PPTConverter, ConversionOptions, ConversionMode, OutputFormat, CompressionLevel, ImageFormat
from src.core.progress_tracker import ProgressTracker, ConversionStage
from src.utils.logger import Logger
import win32com.client
import pythoncom
import os
import tempfile
import shutil
from PIL import Image
import threading
import time
import atexit
import base64
from typing import Dict, Any, List, Callable

# 全局实例列表，用于程序退出时清理
_converter_instances = []

def _cleanup_all_converters():
    """程序退出时清理所有转换器实例"""
    for converter in _converter_instances:
        try:
            converter._cleanup()
        except:
            pass

# 注册退出清理函数
atexit.register(_cleanup_all_converters)

class Win32PPTConverter(PPTConverter):
    """
    基于Win32 COM接口的PPT转换器实现
    使用PowerPoint应用程序对象进行操作
    """
    
    def __init__(self, progress_callback: Callable[[float, str], None] = None):
        self.logger = Logger().get_logger()
        self.powerpoint = None
        self._lock = threading.Lock()
        self._progress_tracker = ProgressTracker(progress_callback)
        self._progress_callback = progress_callback
        self._temp_dirs = []  # 跟踪所有创建的临时目录

        # 注册到全局列表
        _converter_instances.append(self)

    def _create_temp_dir(self, prefix: str = "pitchppt_") -> str:
        """
        创建安全的临时目录，带路径验证和错误处理
        
        Args:
            prefix: 临时目录前缀
            
        Returns:
            str: 临时目录路径
            
        Raises:
            OSError: 如果创建临时目录失败
        """
        max_attempts = 3
        last_error = None
        
        for attempt in range(max_attempts):
            try:
                # 尝试创建临时目录
                temp_dir = tempfile.mkdtemp(prefix=prefix)
                
                # 验证目录是否可写
                test_file = os.path.join(temp_dir, "write_test.tmp")
                try:
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                except Exception as e:
                    # 如果不可写，清理并重试
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass
                    raise OSError(f"临时目录不可写: {e}")
                
                # 转换为绝对路径
                temp_dir = os.path.abspath(temp_dir)
                
                # 记录到跟踪列表
                self._temp_dirs.append(temp_dir)
                
                self.logger.info(f"✓ 成功创建临时目录: {temp_dir}")
                return temp_dir
                
            except Exception as e:
                last_error = e
                self.logger.warning(f"创建临时目录失败 (尝试 {attempt + 1}/{max_attempts}): {e}")
                if attempt < max_attempts - 1:
                    time.sleep(0.2)  # 短暂等待后重试
        
        raise OSError(f"无法创建临时目录: {last_error}")

    def _cleanup_temp_dirs(self):
        """清理所有临时目录"""
        for temp_dir in self._temp_dirs[:]:  # 使用切片避免修改列表时的问题
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    self.logger.debug(f"已清理临时目录: {temp_dir}")
            except Exception as e:
                self.logger.warning(f"清理临时目录失败 {temp_dir}: {e}")
            finally:
                if temp_dir in self._temp_dirs:
                    self._temp_dirs.remove(temp_dir)

    @property
    def progress_tracker(self):
        """获取进度跟踪器"""
        return self._progress_tracker
    
    def __del__(self):
        """析构时确保清理"""
        try:
            self._cleanup()
            if self in _converter_instances:
                _converter_instances.remove(self)
        except:
            pass
    
    def _initialize_powerpoint(self) -> bool:
        """
        初始化PowerPoint COM对象
        借鉴成功案例：谨慎处理Visible属性，避免兼容性问题
        
        Returns:
            bool: 初始化是否成功
        """
        try:
            # 检查现有的PowerPoint实例是否仍然有效
            if self.powerpoint is not None:
                try:
                    # 尝试访问一个属性来验证实例是否仍然有效
                    _ = self.powerpoint.Visible
                    self.logger.info("PowerPoint实例仍然有效，复用现有实例")
                    return True
                except Exception as e:
                    self.logger.warning(f"现有PowerPoint实例已失效: {e}")
                    self.powerpoint = None
            
            # 初始化COM库
            pythoncom.CoInitialize()
            
            # 尝试不同的PowerPoint版本
            powerpoint_versions = [
                "PowerPoint.Application",
                "PowerPoint.Application.16",  # Office 2016+
                "PowerPoint.Application.15",  # Office 2013
                "PowerPoint.Application.14",  # Office 2010
            ]
            
            for version in powerpoint_versions:
                try:
                    self.powerpoint = win32com.client.Dispatch(version)
                    
                    # 借鉴成功案例：谨慎处理属性设置
                    try:
                        self.powerpoint.DisplayAlerts = False  # 禁用警告对话框
                        self.logger.info("已禁用PowerPoint警告对话框")
                    except Exception as alert_error:
                        self.logger.warning(f"设置DisplayAlerts失败: {alert_error}")
                    
                    # 借鉴成功案例：谨慎处理Visible属性
                    try:
                        # 先尝试获取当前状态
                        current_visible = self.powerpoint.Visible
                        self.logger.info(f"PowerPoint当前可见状态: {current_visible}")
                        
                        # 如果当前不可见，尝试设置为可见（避免兼容性问题）
                        if not current_visible:
                            self.powerpoint.Visible = True
                            self.logger.info("PowerPoint窗口已设置为可见")
                    except Exception as visible_error:
                        self.logger.warning(f"设置Visible属性失败，使用默认设置: {visible_error}")
                    
                    self.logger.info(f"PowerPoint COM接口初始化成功: {version}")
                    return True
                except Exception as e:
                    self.logger.warning(f"尝试初始化 {version} 失败: {e}")
                    if self.powerpoint:
                        try:
                            self.powerpoint.Quit()
                        except:
                            pass
                    self.powerpoint = None
            
            # 如果所有版本都失败，尝试使用CreateObject
            try:
                self.powerpoint = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
                
                # 同样谨慎处理属性
                try:
                    self.powerpoint.DisplayAlerts = False
                except:
                    pass
                
                self.logger.info("PowerPoint COM接口初始化成功 (使用EnsureDispatch)")
                return True
            except Exception as e:
                self.logger.error(f"所有PowerPoint初始化尝试都失败: {e}")
                return False
        except Exception as e:
            self.logger.error(f"初始化PowerPoint失败: {e}")
            return False
    
    def _cleanup(self, force_kill=True):
        """
        清理资源，关闭PowerPoint实例
        使用更强制的方法确保进程被关闭
        
        Args:
            force_kill: 是否强制终止进程，默认为True。在单文件转换完成后应该为False，避免影响其他转换
        """
        import time
        
        # 先清理所有临时目录
        self._cleanup_temp_dirs()
        
        with self._lock:
            if self.powerpoint:
                try:
                    # 关闭所有打开的演示文稿
                    try:
                        while self.powerpoint.Presentations.Count > 0:
                            try:
                                self.powerpoint.Presentations(1).Close()
                            except Exception as e:
                                self.logger.warning(f"关闭演示文稿时出错: {e}")
                                break
                    except Exception as e:
                        self.logger.warning(f"获取演示文稿列表时出错: {e}")
                    
                    # 尝试正常退出
                    try:
                        self.powerpoint.Quit()
                        self.logger.info("PowerPoint实例已关闭")
                    except Exception as e:
                        self.logger.warning(f"PowerPoint Quit失败: {e}")
                    
                except Exception as e:
                    self.logger.error(f"关闭PowerPoint时出错: {e}")
                finally:
                    self.powerpoint = None
            
            # 等待一段时间让文件系统完成写入
            time.sleep(0.5)
            
            # 强制终止残留的PowerPoint进程（仅在需要时）
            if force_kill:
                self._kill_powerpoint_processes()
    
    def _kill_powerpoint_processes(self):
        """
        强制终止残留的PowerPoint进程
        """
        try:
            import subprocess
            import time
            
            # 先尝试正常关闭
            result = subprocess.run(
                ['taskkill', '/F', '/IM', 'POWERPNT.EXE'],
                capture_output=True,
                text=True,
                timeout=5
            )
            
            if result.returncode == 0:
                self.logger.info("已强制终止PowerPoint进程")
                # 等待进程完全退出
                time.sleep(1)
            elif "没有找到进程" in result.stderr or "No tasks" in result.stderr:
                self.logger.debug("没有残留的PowerPoint进程")
            else:
                self.logger.warning(f"终止PowerPoint进程时出错: {result.stderr}")
                
        except subprocess.TimeoutExpired:
            self.logger.error("终止PowerPoint进程超时")
        except Exception as e:
            self.logger.error(f"终止PowerPoint进程失败: {e}")
    
    def _update_progress(self, value: float, task: str = None):
        """
        更新进度
        
        Args:
            value: 进度值 (0.0-1.0)
            task: 当前任务描述
        """
        with self._lock:
            self._progress = max(0.0, min(1.0, value))
            if task:
                self._current_task = task
            
        self.logger.debug(f"进度更新: {self._progress:.1%} - {task or '未知任务'}")
    
    def _validate_image_file(self, image_path: str) -> bool:
        """
        验证图片文件的有效性
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            bool: 文件是否有效
        """
        try:
            if not os.path.exists(image_path):
                return False
                
            file_size = os.path.getsize(image_path)
            if file_size < 500:  # 放宽到500字节
                return False
                
            # 尝试使用PIL打开，如果失败则使用文件大小判断
            try:
                with Image.open(image_path) as img:
                    width, height = img.size
                    if width < 5 or height < 5:  # 放宽到5像素
                        return False
            except:
                # 如果PIL无法打开，但文件大小合理，也认为是有效的
                if file_size > 5000:  # 大于5KB
                    return True
                return False
                
            return True
        except Exception as e:
            self.logger.warning(f"验证图片文件失败 {image_path}: {e}")
            # 放宽验证，如果文件存在且有一定大小，就认为是有效的
            return os.path.exists(image_path) and os.path.getsize(image_path) > 500
    
    def _export_slide_to_image(self, slide, output_path: str, image_config=None, presentation=None) -> bool:
        """
        导出幻灯片为图片
        支持多种格式和清晰度定制
        
        Args:
            slide: 幻灯片对象
            output_path: 输出路径
            image_config: ImageExportConfig 配置对象
            presentation: 演示文稿对象（用于设置高分辨率）
            
        Returns:
            bool: 是否成功
        """
        from src.core.converter import ImageExportConfig, ImageFormat
        
        if image_config is None:
            image_config = ImageExportConfig()
        
        # 获取导出格式
        export_format = image_config.format.value.upper()
        temp_png = output_path + "_temp.png"
        
        # 获取DPI设置
        dpi = image_config.get_effective_dpi()

        # 借鉴成功案例：添加重试机制
        for attempt in range(2):  # 最多重试2次
            try:
                self.logger.debug(f"导出幻灯片为图片，格式: {export_format}, DPI: {dpi}, 尝试 {attempt + 1}")

                # 获取原始尺寸（以点为单位，1点 = 1/72英寸）
                if presentation:
                    try:
                        orig_width = presentation.PageSetup.SlideSize.Width
                        orig_height = presentation.PageSetup.SlideSize.Height
                        orig_ratio = orig_width / orig_height
                        self.logger.info(f"原始幻灯片尺寸: {orig_width:.0f}x{orig_height:.0f} 点 (宽高比: {orig_ratio:.4f})")
                    except:
                        orig_width = 720  # 默认10英寸
                        orig_height = 540
                        orig_ratio = 4/3
                else:
                    orig_width = 720
                    orig_height = 540
                    orig_ratio = 4/3

                # 计算目标导出尺寸（以像素为单位）
                export_width = None
                export_height = None
                needs_post_scale = False
                needs_ratio_correction = False

                # PowerPoint Export方法的最大像素限制：100,000,000像素（1亿像素）
                MAX_PIXELS = 100_000_000

                if image_config.use_custom_resolution and image_config.custom_height > 0:
                    # 使用自定义高度
                    export_height = image_config.custom_height
                    export_width = int(export_height * orig_ratio)
                    total_pixels = export_width * export_height

                    # 检查是否超过PowerPoint的像素限制
                    if total_pixels > MAX_PIXELS:
                        self.logger.warning(f"目标尺寸 {export_width}x{export_height} = {total_pixels:,} 像素，超过PowerPoint最大限制 {MAX_PIXELS:,} 像素")
                        # 计算最大允许尺寸（保持宽高比）
                        scale_factor = (MAX_PIXELS / total_pixels) ** 0.5
                        export_width = int(export_width * scale_factor)
                        export_height = int(export_height * scale_factor)
                        total_pixels = export_width * export_height
                        self.logger.warning(f"将使用最大允许尺寸: {export_width}x{export_height} = {total_pixels:,} 像素")
                        needs_post_scale = True

                    self.logger.info(f"目标导出尺寸: {export_width}x{export_height}px ({total_pixels:,} 像素)")
                elif dpi > 96:
                    # 使用DPI计算导出尺寸
                    # 原始像素尺寸（96 DPI基准）
                    base_width = int(orig_width * 96 / 72)  # 点转像素
                    base_height = int(orig_height * 96 / 72)
                    scale_factor = dpi / 96.0
                    export_width = int(base_width * scale_factor)
                    export_height = int(base_height * scale_factor)
                    total_pixels = export_width * export_height

                    # 检查是否超过PowerPoint的像素限制
                    if total_pixels > MAX_PIXELS:
                        self.logger.warning(f"DPI缩放尺寸 {export_width}x{export_height} = {total_pixels:,} 像素，超过PowerPoint最大限制 {MAX_PIXELS:,} 像素")
                        # 计算最大允许尺寸（保持宽高比）
                        scale_factor = (MAX_PIXELS / total_pixels) ** 0.5
                        export_width = int(export_width * scale_factor)
                        export_height = int(export_height * scale_factor)
                        total_pixels = export_width * export_height
                        self.logger.warning(f"将使用最大允许尺寸: {export_width}x{export_height}px ({total_pixels:,} 像素)")
                        needs_post_scale = True

                    self.logger.debug(f"DPI缩放导出: {base_width}x{base_height} -> {export_width}x{export_height} (DPI: {dpi})")

                # 使用Export方法导出图片
                # Export(FileName, FilterName, ScaleWidth, ScaleHeight)
                # ScaleWidth和ScaleHeight的单位是"点"（points），不是像素！
                # 1点 = 1/72英寸
                # 关键修复：不使用ScaleWidth和ScaleHeight参数，让PowerPoint使用默认尺寸
                # 然后用PIL调整到目标尺寸
                try:
                    # 先使用默认尺寸导出
                    slide.Export(temp_png, "PNG")
                    self.logger.debug(f"使用默认尺寸导出PNG")
                    needs_post_scale = True
                except Exception as e:
                    self.logger.warning(f"默认尺寸导出失败: {e}")
                    if attempt == 0:
                        import time
                        time.sleep(0.5)
                        continue
                    return False

                # 验证临时PNG图片
                if not self._validate_image_file(temp_png):
                    self.logger.warning(f"PNG临时文件导出失败或文件无效: {temp_png}")
                    if attempt == 0:
                        import time
                        time.sleep(0.5)
                        continue
                    return False
                
                # 根据目标格式进行转换
                with Image.open(temp_png) as img:
                    # 获取PowerPoint导出的原始图片尺寸
                    exported_width, exported_height = img.size
                    exported_ratio = exported_width / exported_height
                    
                    self.logger.info(f"PowerPoint导出尺寸: {exported_width}x{exported_height}px (宽高比: {exported_ratio:.4f})")
                    self.logger.info(f"原始PPT尺寸: {orig_width:.0f}x{orig_height:.0f}点 (宽高比: {orig_ratio:.4f})")
                    
                    # 计算目标尺寸 - 关键：使用原始PPT的宽高比
                    target_width = None
                    target_height = None
                    
                    # 应用自定义分辨率（如果需要）
                    if image_config.use_custom_resolution and image_config.custom_height > 0:
                        # 使用自定义高度，但保持原始PPT的宽高比
                        target_height = image_config.custom_height
                        target_width = int(target_height * orig_ratio)
                        self.logger.info(f"使用PIL缩放到目标分辨率: {exported_width}x{exported_height} -> {target_width}x{target_height}")
                    elif dpi > 96:
                        # 使用DPI计算目标尺寸，保持原始PPT的宽高比
                        base_width = int(orig_width * 96 / 72)
                        base_height = int(orig_height * 96 / 72)
                        scale_factor = dpi / 96.0
                        target_width = int(base_width * scale_factor)
                        target_height = int(base_height * scale_factor)
                        self.logger.info(f"使用PIL缩放到DPI目标: {exported_width}x{exported_height} -> {target_width}x{target_height}")
                    else:
                        # 使用原始PPT的宽高比重新计算尺寸
                        # 保持PowerPoint导出的高度，调整宽度以匹配原始PPT的宽高比
                        target_height = exported_height
                        target_width = int(target_height * orig_ratio)
                        self.logger.info(f"修正宽高比: {exported_width}x{exported_height} -> {target_width}x{target_height}")
                    
                    # 检查是否需要缩放
                    if target_width and target_height:
                        # 使用LANCZOS重采样进行高质量缩放
                        img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                        self.logger.info(f"✓ 图片已缩放至: {target_width}x{target_height}px")
                    
                    # 最终验证：检查长宽比是否正确
                    final_width, final_height = img.size
                    final_ratio = final_width / final_height
                    
                    # 检查是否需要修正（允许0.1%的误差）
                    if abs(final_ratio - orig_ratio) / orig_ratio > 0.001:
                        self.logger.warning(f"检测到长宽比不匹配: {final_ratio:.4f} vs 目标 {orig_ratio:.4f}")
                        
                        # 修正长宽比（保持高度不变，调整宽度）
                        corrected_height = final_height
                        corrected_width = int(corrected_height * orig_ratio)
                        
                        self.logger.info(f"最终修正长宽比: {final_width}x{final_height} -> {corrected_width}x{corrected_height}")
                        img = img.resize((corrected_width, corrected_height), Image.Resampling.LANCZOS)
                    
                    # 根据格式保存
                    if image_config.format == ImageFormat.PNG:
                        # PNG格式
                        save_kwargs = {'optimize': image_config.optimize}
                        if image_config.transparent_background:
                            img = img.convert("RGBA")
                        else:
                            img = img.convert("RGB")
                        img.save(output_path, "PNG", **save_kwargs)
                        
                    elif image_config.format in (ImageFormat.JPG, ImageFormat.JPEG):
                        # JPG格式
                        img = img.convert("RGB")
                        save_kwargs = {
                            'quality': image_config.quality,
                            'optimize': image_config.optimize,
                            'progressive': image_config.progressive
                        }
                        img.save(output_path, "JPEG", **save_kwargs)
                        
                    elif image_config.format == ImageFormat.TIFF:
                        # TIFF格式
                        save_kwargs = {'compression': 'tiff_lzw' if image_config.compression_level != CompressionLevel.NONE else None}
                        img.save(output_path, "TIFF", **save_kwargs)
                        
                    elif image_config.format == ImageFormat.BMP:
                        # BMP格式
                        img = img.convert("RGB")
                        img.save(output_path, "BMP")
                        
                    elif image_config.format == ImageFormat.GIF:
                        # GIF格式
                        if image_config.transparent_background:
                            img = img.convert("RGBA")
                        else:
                            img = img.convert("RGB")
                        img.save(output_path, "GIF", optimize=image_config.optimize)
                        
                    elif image_config.format == ImageFormat.WEBP:
                        # WebP格式
                        save_kwargs = {
                            'quality': image_config.quality,
                            'method': 6 if image_config.optimize else 0
                        }
                        if image_config.transparent_background:
                            img = img.convert("RGBA")
                            save_kwargs['lossless'] = False
                        else:
                            img = img.convert("RGB")
                        img.save(output_path, "WEBP", **save_kwargs)
                        
                    else:
                        # 默认JPG
                        img = img.convert("RGB")
                        img.save(output_path, "JPEG", quality=image_config.quality, optimize=True)
                
                # 验证输出文件
                if self._validate_image_file(output_path):
                    self.logger.info(f"✓ 幻灯片导出成功: {output_path} ({image_config.format.value.upper()})")
                    return True
                else:
                    self.logger.warning(f"生成的文件无效: {output_path}")
                    if attempt == 0:
                        continue
                        
            except Exception as e:
                self.logger.warning(f"导出幻灯片为图片失败 (尝试 {attempt + 1}): {e}")
                if attempt == 0:
                    continue
        
        # 清理临时文件
        if os.path.exists(temp_png):
            try:
                os.remove(temp_png)
            except:
                pass
        
        return False
    
    def convert(self, input_path: str, output_path: str, options: ConversionOptions = None) -> bool:
        """
        执行PPT转换操作
        使用精细化进度跟踪系统
        """
        if options is None:
            options = ConversionOptions()
            
        self.logger.info(f"=" * 60)
        self.logger.info(f"📄 开始转换任务")
        self.logger.info(f"   输入文件: {input_path}")
        self.logger.info(f"   输出文件: {output_path}")
        self.logger.info(f"   转换模式: {options.mode.value}")
        self.logger.info(f"   输出格式: {options.output_format.value}")
        self.logger.info(f"=" * 60)
        
        presentation = None
        success = False
        
        try:
            # 阶段1: 初始化PowerPoint (0-5%)
            self._progress_tracker.start_stage(ConversionStage.INITIALIZING)
            self.logger.info("[1/6] 正在初始化 PowerPoint...")
            if not self._initialize_powerpoint():
                self.logger.error("❌ PowerPoint 初始化失败")
                self._progress_tracker.complete(False)
                return False
            self.logger.info("✓ PowerPoint 初始化完成")
            self._progress_tracker.finish_stage("PowerPoint初始化完成")
            
            # 阶段2: 创建临时目录 (5-10%)
            self._progress_tracker.start_stage(ConversionStage.OPENING_FILE)
            try:
                temp_dir = self._create_temp_dir(prefix="pitchppt_")
                self.logger.info(f"[2/6] 创建临时目录: {temp_dir}")
                self._progress_tracker.finish_stage("临时目录创建完成")
            except OSError as e:
                self.logger.error(f"❌ 无法创建临时目录: {e}")
                self._progress_tracker.complete(False)
                return False
            
            # 阶段3: 打开PPT文件并分析 (10-15%)
            self._progress_tracker.start_stage(ConversionStage.ANALYZING)
            self.logger.info("[3/6] 正在打开并分析 PPT 文件...")
            abs_input_path = os.path.abspath(input_path)
            presentation = self.powerpoint.Presentations.Open(abs_input_path)
            slide_count = presentation.Slides.Count
            
            # 获取幻灯片尺寸信息
            try:
                slide_width = presentation.PageSetup.SlideSize.Width
                slide_height = presentation.PageSetup.SlideSize.Height
                self.logger.info(f"✓ 成功打开 PPT")
                self.logger.info(f"   幻灯片总数: {slide_count} 张")
                self.logger.info(f"   幻灯片尺寸: {slide_width:.0f} x {slide_height:.0f} (点)")
            except Exception as e:
                self.logger.warning(f"   无法获取幻灯片尺寸: {e}")
                self.logger.info(f"✓ 成功打开 PPT")
                self.logger.info(f"   幻灯片总数: {slide_count} 张")
            self._progress_tracker.finish_stage(f"分析完成，共{slide_count}张幻灯片")
            
            if slide_count == 0:
                self.logger.error("PPT中没有幻灯片")
                self._progress_tracker.complete(False)
                return False
                
            # 阶段4: 根据转换模式执行不同操作 (15-85%)
            self.logger.info(f"[4/6] 开始执行转换 ({options.mode.value})...")
            if options.mode == ConversionMode.BACKGROUND_FILL:
                success = self._convert_background_fill(presentation, output_path, options, slide_count)
            elif options.mode == ConversionMode.FOREGROUND_IMAGE:
                success = self._convert_foreground_image(presentation, output_path, options, slide_count)
            elif options.mode == ConversionMode.SLIDE_TO_IMAGE:
                success = self._convert_slide_to_image(presentation, output_path, options, slide_count)
            elif options.output_format == OutputFormat.PDF:
                # PDF模式：使用基于图片序列的新方法
                success = self._convert_to_pdf(presentation, output_path, options, slide_count)
            else:
                self.logger.error(f"❌ 不支持的转换模式: {options.mode.value}")
                self._progress_tracker.complete(False)
                return False
                
            # 阶段5: 保存文件 (85-95%)
            # 对于幻灯片转图片序列模式和PDF模式，不需要额外保存文件
            if success and options.mode not in (ConversionMode.SLIDE_TO_IMAGE, None) and options.output_format != OutputFormat.PDF:
                self._progress_tracker.start_stage(ConversionStage.SAVING)
                self.logger.info("[5/6] 正在保存文件...")
                self._progress_tracker.finish_stage("文件保存完成")
                self.logger.info("✓ 文件保存完成")
            elif success and options.mode == ConversionMode.SLIDE_TO_IMAGE:
                # 对于幻灯片转图片序列模式，直接更新进度
                self._progress_tracker.start_stage(ConversionStage.SAVING)
                self._progress_tracker.finish_stage("图片序列导出完成")
                self.logger.info("✓ 图片序列导出完成")
            elif success and options.output_format == OutputFormat.PDF:
                # 对于PDF模式，直接更新进度
                self._progress_tracker.start_stage(ConversionStage.SAVING)
                self._progress_tracker.finish_stage("PDF导出完成")
                self.logger.info("✓ PDF导出完成")
            
        except InterruptedError as e:
            self.logger.info(f"转换被用户终止: {e}")
            success = False
            
        except Exception as e:
            self.logger.error(f"转换过程中发生异常: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            success = False
            
        finally:
            # 阶段6: 清理资源 (95-100%)
            try:
                self._progress_tracker.start_stage(ConversionStage.CLEANING)
                self.logger.info("[6/6] 正在清理资源...")
            except:
                pass
            
            # 关闭演示文稿
            if presentation:
                try:
                    presentation.Close()
                    self.logger.info("   已关闭演示文稿")
                except Exception as e:
                    self.logger.warning(f"   关闭演示文稿失败: {e}")
            
            # 关闭PowerPoint（会自动清理所有临时目录）
            self._cleanup()
            
            try:
                self._progress_tracker.finish_stage("资源清理完成")
                self.logger.info("✓ 资源清理完成")
            except:
                pass
            
        self.logger.info(f"{'=' * 60}")
        self.logger.info(f"✅ 转换{'成功' if success else '失败'}")
        self.logger.info(f"{'=' * 60}")
        self._progress_tracker.complete(success)
        
        return success
    
    def _convert_background_fill(self, presentation, output_path: str, options: ConversionOptions, slide_count: int) -> bool:
        """
        背景填充模式转换
        使用精细化进度跟踪
        """
        try:
            temp_dir = self._create_temp_dir(prefix="pitchppt_bg_")
        except OSError as e:
            self.logger.error(f"无法创建临时目录: {e}")
            return False
        
        image_files = []
        
        self.logger.info(f"   临时目录: {temp_dir}")
        self.logger.info(f"   图片导出配置: 格式={options.image_export.format.value}, DPI={options.image_export.get_effective_dpi()}")
        
        try:
            # 阶段4.1: 导出每张幻灯片为图片 (15-50%)
            self._progress_tracker.start_stage(ConversionStage.EXPORTING_IMAGES, slide_count)
            self.logger.info(f"   开始导出 {slide_count} 张幻灯片为图片...")
            
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                if slide.SlideShowTransition.Hidden and not options.include_hidden_slides:
                    self.logger.info(f"   幻灯片 {i}/{slide_count}: 跳过隐藏幻灯片")
                    self._progress_tracker.step("跳过隐藏幻灯片")
                    continue
                    
                image_path = os.path.join(temp_dir, f"slide_{i:03d}.jpg")
                self.logger.info(f"   幻灯片 {i}/{slide_count}: 正在导出...")
                
                if self._export_slide_to_image(slide, image_path, options.image_export, presentation):
                    image_files.append(image_path)
                    self.logger.info(f"   幻灯片 {i}/{slide_count}: ✓ 导出成功 -> {image_path}")
                    self._progress_tracker.update_stage(i, f"导出幻灯片 {i}/{slide_count}")
                else:
                    self.logger.warning(f"   幻灯片 {i}/{slide_count}: ✗ 导出失败")
                    self._progress_tracker.step(f"导出幻灯片 {i} 失败")
                    
            if not image_files:
                self.logger.error("❌ 没有成功导出任何图片")
                return False
                
            self.logger.info(f"✓ 成功导出 {len(image_files)}/{slide_count} 张图片")
            self._progress_tracker.finish_stage(f"成功导出 {len(image_files)} 张图片")
            
            # 阶段4.2: 处理幻灯片背景 (50-85%)
            self._progress_tracker.start_stage(ConversionStage.PROCESSING_SLIDES, len(image_files))
            template_path = os.path.abspath(output_path)
            
            # 如果输出格式是PPTX，先保存为临时PPTX
            if options.output_format == OutputFormat.PPTX:
                temp_pptx = os.path.join(temp_dir, "template.pptx")
                presentation.SaveAs(temp_pptx)
                template_path = temp_pptx
            
            template_presentation = self.powerpoint.Presentations.Open(template_path)
            
            # 如果不包含隐藏幻灯片，需要删除模板中的隐藏幻灯片
            if not options.include_hidden_slides:
                slides_to_delete = []
                for i in range(1, template_presentation.Slides.Count + 1):
                    slide = template_presentation.Slides(i)
                    if slide.SlideShowTransition.Hidden:
                        slides_to_delete.append(i)
                
                # 从后往前删除，避免索引变化
                for slide_idx in reversed(slides_to_delete):
                    try:
                        template_presentation.Slides(slide_idx).Delete()
                        self.logger.info(f"删除隐藏幻灯片 {slide_idx}")
                    except Exception as e:
                        self.logger.warning(f"删除隐藏幻灯片 {slide_idx} 失败: {e}")
            
            # 确保幻灯片数量匹配
            while template_presentation.Slides.Count < len(image_files):
                last_slide = template_presentation.Slides(template_presentation.Slides.Count)
                last_slide.Duplicate()
                
            # 设置每张幻灯片的背景
            for i, image_file in enumerate(image_files, 1):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        self.logger.info(f"处理幻灯片 {i}/{len(image_files)} (JPG)...")
                        
                        # 借鉴成功案例：关键修复 - 设置FollowMasterBackground为False
                        try:
                            slide.FollowMasterBackground = False
                            self.logger.info(f"✓ 幻灯片 {i} 已禁用跟随母版背景")
                        except Exception as e:
                            self.logger.warning(f"设置FollowMasterBackground失败: {e}")
                        
                        # 借鉴成功案例：设置为空白版式（避免占位符文本）
                        try:
                            slide.Layout = 12  # ppLayoutBlank = 12
                            self.logger.info(f"✓ 幻灯片 {i} 已设置为空白版式")
                        except Exception as e:
                            self.logger.warning(f"设置空白版式失败: {e}")
                        
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
                            self.logger.info(f"清空了 {deleted_count} 个元素（包括占位符）")
                        except Exception as e:
                            self.logger.warning(f"清空幻灯片内容时出错: {e}")
                        
                        # 借鉴成功案例：设置背景图片（多方案备用）
                        background_set = False
                        abs_image_path = os.path.abspath(image_file)

                        # 方法1：使用UserPicture设置JPG背景
                        try:
                            slide.Background.Fill.UserPicture(abs_image_path)
                            # 等待一下让设置生效
                            import time
                            time.sleep(0.1)
                            
                            # 验证背景是否设置成功
                            try:
                                fill_type = slide.Background.Fill.Type
                                if fill_type == 6:  # msoFillPicture = 6
                                    background_set = True
                                    self.logger.info(f"✓ 方法1成功：幻灯片 {i} JPG背景设置完成")
                                else:
                                    self.logger.warning(f"方法1设置JPG后验证失败")
                            except:
                                self.logger.warning(f"方法1验证背景失败")
                                
                        except Exception as e:
                            self.logger.warning(f"方法1失败：{e}")
                        
                        # 如果UserPicture失败，使用备用方案
                        if not background_set:
                            try:
                                # 获取幻灯片尺寸
                                slide_width = template_presentation.PageSetup.SlideWidth
                                slide_height = template_presentation.PageSetup.SlideHeight
                                
                                # 获取图片实际尺寸，保持宽高比
                                with Image.open(abs_image_path) as img:
                                    img_width, img_height = img.size
                                    img_ratio = img_width / img_height
                                    slide_ratio = slide_width / slide_height
                                
                                # 计算保持宽高比的尺寸（使用"cover"模式：图片覆盖整个幻灯片，可能裁剪）
                                if img_ratio > slide_ratio:
                                    # 图片更宽，以高度为基准
                                    scaled_height = slide_height
                                    scaled_width = int(scaled_height * img_ratio)
                                else:
                                    # 图片更高，以宽度为基准
                                    scaled_width = slide_width
                                    scaled_height = int(scaled_width / img_ratio)
                                
                                # 计算居中位置
                                left = (slide_width - scaled_width) // 2
                                top = (slide_height - scaled_height) // 2
                                
                                # 添加图片，保持宽高比
                                picture = slide.Shapes.AddPicture(abs_image_path, False, True, left, top, scaled_width, scaled_height)
                                # 将图片移到最底层（作为背景）
                                try:
                                    picture.ZOrder(0)  # 发送到底层
                                except:
                                    pass
                                background_set = True
                                self.logger.info(f"✓ 备用方案成功：幻灯片 {i} JPG图片作为背景添加完成 (保持宽高比 {img_ratio:.3f})")
                            except Exception as e:
                                self.logger.error(f"备用方案失败：{e}")
                        
                        if not background_set:
                            self.logger.error(f"幻灯片 {i} 背景设置完全失败")
                        else:
                            self.logger.debug(f"已设置幻灯片 {i} 背景")
                        
                    except Exception as e:
                        self.logger.error(f"设置幻灯片 {i} 背景失败: {e}")
                
                # 更新进度
                self._progress_tracker.update_stage(i, f"处理幻灯片 {i}/{len(image_files)}")
            
            self._progress_tracker.finish_stage(f"完成 {len(image_files)} 张幻灯片处理")
            
            # 阶段5: 保存最终文件 (85-95%)
            self._progress_tracker.start_stage(ConversionStage.SAVING)
            final_success = False
            
            try:
                save_path = os.path.abspath(output_path)
                self.logger.info(f"[保存] 准备保存到: {save_path}")
                self.logger.info(f"[保存] 输出格式: {options.output_format.value}")
                
                if options.output_format == OutputFormat.PPTX:
                    self.logger.info(f"[保存] 调用 SaveAs...")
                    template_presentation.SaveAs(save_path)
                    self.logger.info(f"[保存] SaveAs 调用完成")
                elif options.output_format == OutputFormat.PDF:
                    # PowerPoint的PDF导出需要使用ExportAsFixedFormat方法
                    # ppFixedFormatTypePDF = 2
                    # ppFixedFormatIntentPrint = 2 (打印质量)
                    # ppFixedFormatIntentScreen = 1 (屏幕质量)
                    self.logger.info(f"[保存] 开始导出PDF到: {save_path}")
                    
                    try:
                        # 使用完整的参数导出PDF
                        template_presentation.ExportAsFixedFormat(
                            save_path,                    # 输出路径
                            2,                            # ppFixedFormatTypePDF
                            2,                            # ppFixedFormatIntentPrint (打印质量)
                            False,                         # 不包含文档属性
                            False,                         # 不保留IRM
                            False,                         # 不包含文档结构标签
                            False,                         # 不对缺失字体使用位图
                            False                          # 不使用PDF/A标准
                        )
                        self.logger.info(f"[保存] PDF导出完成: {save_path}")
                    except Exception as pdf_error:
                        # 如果完整参数失败，尝试使用最小参数
                        self.logger.warning(f"[保存] 完整参数PDF导出失败，尝试简化参数: {pdf_error}")
                        try:
                            template_presentation.ExportAsFixedFormat(
                                save_path,
                                2  # ppFixedFormatTypePDF
                            )
                            self.logger.info(f"[保存] PDF导出完成（简化参数）: {save_path}")
                        except Exception as e:
                            raise Exception(f"PDF导出失败（完整和简化参数都失败）: {e}")
                else:
                    self.logger.error(f"[保存] 不支持的输出格式: {options.output_format.value}")
                    
                # 验证文件是否真的被创建
                import time
                time.sleep(0.5)  # 等待文件系统完成写入
                
                if os.path.exists(save_path):
                    file_size = os.path.getsize(save_path)
                    self.logger.info(f"[保存] ✓ 文件已创建: {save_path} ({file_size} bytes)")
                    final_success = True
                else:
                    self.logger.error(f"[保存] ✗ 文件未创建: {save_path}")
                    final_success = False
                
            except Exception as e:
                self.logger.error(f"[保存] 保存文件失败: {e}")
                import traceback
                self.logger.error(f"[保存] 详细错误信息: {traceback.format_exc()}")
                
            template_presentation.Close()
            return final_success
            
        except Exception as e:
            self.logger.error(f"背景填充模式转换失败: {e}")
            return False
    
    def _convert_foreground_image(self, presentation, output_path: str, options: ConversionOptions, slide_count: int) -> bool:
        """
        前景图片模式转换
        """
        try:
            temp_dir = self._create_temp_dir(prefix="pitchppt_fg_")
        except OSError as e:
            self.logger.error(f"无法创建临时目录: {e}")
            return False
        
        image_files = []
        
        try:
            # 1. 导出每张幻灯片为图片
            self._update_progress(0.2, "导出幻灯片为前景图片")
            
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                if slide.SlideShowTransition.Hidden and not options.include_hidden_slides:
                    continue
                    
                image_path = os.path.join(temp_dir, f"slide_{i:03d}.jpg")
                
                if self._export_slide_to_image(slide, image_path, options.image_export, presentation):
                    image_files.append(image_path)
                    progress = 0.2 + 0.3 * (i / slide_count)
                    self._update_progress(progress, f"导出前景图片 {i}/{slide_count}")
                else:
                    self.logger.warning(f"导出幻灯片 {i} 作为前景图片失败")
                    
            if not image_files:
                self.logger.error("没有成功导出任何前景图片")
                return False
                
            self.logger.info(f"成功导出 {len(image_files)} 张前景图片")
            
            # 2. 创建新的PPTX，将图片作为前景对象
            self._update_progress(0.6, "创建前景图片幻灯片")
            
            # 创建新的演示文稿
            new_presentation = self.powerpoint.Presentations.Add()
            
            # 设置幻灯片大小与原始一致
            new_presentation.PageSetup.SlideWidth = presentation.PageSetup.SlideWidth
            new_presentation.PageSetup.SlideHeight = presentation.PageSetup.SlideHeight
            
            for i, image_file in enumerate(image_files, 1):
                # 添加新幻灯片
                slide = new_presentation.Slides.Add(i, 1)  # ppLayoutBlank = 1
                
                # 设置为空白版式
                slide.Layout = 12
                
                # 清空所有形状
                for j in range(slide.Shapes.Count, 0, -1):
                    try:
                        slide.Shapes(j).Delete()
                    except:
                        pass
                
                # 添加图片作为前景对象
                slide_width = new_presentation.PageSetup.SlideWidth
                slide_height = new_presentation.PageSetup.SlideHeight
                abs_image_path = os.path.abspath(image_file)
                
                try:
                    # 获取图片实际尺寸，保持宽高比
                    with Image.open(abs_image_path) as img:
                        img_width, img_height = img.size
                        img_ratio = img_width / img_height
                        slide_ratio = slide_width / slide_height
                    
                    # 计算保持宽高比的尺寸（使用"cover"模式：图片覆盖整个幻灯片，可能裁剪）
                    if img_ratio > slide_ratio:
                        # 图片更宽，以高度为基准
                        scaled_height = slide_height
                        scaled_width = int(scaled_height * img_ratio)
                    else:
                        # 图片更高，以宽度为基准
                        scaled_width = slide_width
                        scaled_height = int(scaled_width / img_ratio)
                    
                    # 计算居中位置
                    left = (slide_width - scaled_width) // 2
                    top = (slide_height - scaled_height) // 2
                    
                    # 添加图片并调整大小，保持宽高比
                    picture = slide.Shapes.AddPicture(
                        FileName=abs_image_path,
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=scaled_width,
                        Height=scaled_height
                    )
                    
                    # 将图片移到最顶层
                    try:
                        picture.ZOrder(1)  # 置于顶层
                    except:
                        pass
                    
                    self.logger.debug(f"已添加幻灯片 {i} 前景图片 (保持宽高比 {img_ratio:.3f})")
                except Exception as e:
                    self.logger.error(f"添加幻灯片 {i} 前景图片失败: {e}")
                    
                progress = 0.6 + 0.3 * (i / len(image_files))
                self._update_progress(progress, f"处理前景图片 {i}/{len(image_files)}")
            
            # 3. 保存最终文件
            self._update_progress(0.95, "保存前景图片PPT")
            final_success = False
            
            try:
                save_path = os.path.abspath(output_path)
                self.logger.info(f"[保存] 准备保存到: {save_path}")
                
                if options.output_format == OutputFormat.PPTX:
                    self.logger.info(f"[保存] 调用 SaveAs...")
                    new_presentation.SaveAs(save_path)
                    self.logger.info(f"[保存] SaveAs 调用完成")
                elif options.output_format == OutputFormat.PDF:
                    # PowerPoint的PDF导出需要使用ExportAsFixedFormat方法
                    # ppFixedFormatTypePDF = 2
                    # ppFixedFormatIntentPrint = 2 (打印质量)
                    # ppFixedFormatIntentScreen = 1 (屏幕质量)
                    self.logger.info(f"[保存] 开始导出前景图片PDF到: {save_path}")
                    
                    try:
                        # 使用完整的参数导出PDF
                        new_presentation.ExportAsFixedFormat(
                            save_path,                    # 输出路径
                            2,                            # ppFixedFormatTypePDF
                            2,                            # ppFixedFormatIntentPrint (打印质量)
                            False,                         # 不包含文档属性
                            False,                         # 不保留IRM
                            False,                         # 不包含文档结构标签
                            False,                         # 不对缺失字体使用位图
                            False                          # 不使用PDF/A标准
                        )
                        self.logger.info(f"[保存] 前景图片PDF导出完成: {save_path}")
                    except Exception as pdf_error:
                        # 如果完整参数失败，尝试使用最小参数
                        self.logger.warning(f"[保存] 完整参数PDF导出失败，尝试简化参数: {pdf_error}")
                        try:
                            new_presentation.ExportAsFixedFormat(
                                save_path,
                                2  # ppFixedFormatTypePDF
                            )
                            self.logger.info(f"[保存] 前景图片PDF导出完成（简化参数）: {save_path}")
                        except Exception as e:
                            raise Exception(f"前景图片PDF导出失败（完整和简化参数都失败）: {e}")
                else:
                    self.logger.error(f"[保存] 不支持的输出格式: {options.output_format.value}")
                    
                # 验证文件是否真的被创建
                import time
                time.sleep(0.5)  # 等待文件系统完成写入
                
                if os.path.exists(save_path):
                    file_size = os.path.getsize(save_path)
                    self.logger.info(f"[保存] ✓ 文件已创建: {save_path} ({file_size} bytes)")
                    final_success = True
                else:
                    self.logger.error(f"[保存] ✗ 文件未创建: {save_path}")
                    final_success = False
                
            except Exception as e:
                self.logger.error(f"[保存] 保存前景图片PPT失败: {e}")
                import traceback
                self.logger.error(f"[保存] 详细错误信息: {traceback.format_exc()}")
                
            new_presentation.Close()
            return final_success
            
        except Exception as e:
            self.logger.error(f"前景图片模式转换失败: {e}")
            return False

    def _convert_slide_to_image(self, presentation, output_path: str, options: ConversionOptions, slide_count: int) -> bool:
        """
        幻灯片转图片序列模式
        使用新的进度跟踪系统
        借鉴背景填充模式：先导出到临时目录，再移动到目标目录
        """
        temp_dir = None
        image_files = []
        
        try:
            # 创建临时目录（借鉴背景填充模式）
            temp_dir = self._create_temp_dir(prefix="pitchppt_imgseq_")
            
            # 阶段4: 导出幻灯片为图片序列 (15-85%)
            self._progress_tracker.start_stage(ConversionStage.EXPORTING_IMAGES, slide_count)
            
            # 导出每张幻灯片为图片到临时目录
            exported_index = 0
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                if slide.SlideShowTransition.Hidden and not options.include_hidden_slides:
                    self._progress_tracker.step("跳过隐藏幻灯片")
                    continue
                
                exported_index += 1
                    
                # 生成文件名（使用连续索引）
                img_format = options.image_export.format.value
                image_path = os.path.join(
                    temp_dir, 
                    f"slide_{exported_index:03d}.{img_format}"
                )
                
                # 使用新的导出方法
                if self._export_slide_to_image(slide, image_path, options.image_export, presentation):
                    image_files.append(image_path)
                    self._progress_tracker.update_stage(i, f"导出幻灯片 {i}/{slide_count}")
                    self.logger.debug(f"已导出幻灯片 {i} 为 {image_path}")
                else:
                    self.logger.error(f"导出幻灯片 {i} 失败")
                    self._progress_tracker.step(f"导出幻灯片 {i} 失败")
                    
            self._progress_tracker.finish_stage(f"成功导出 {len(image_files)} 张幻灯片图片")
            
            # 阶段5: 移动文件到目标目录 (85-95%)
            self._progress_tracker.start_stage(ConversionStage.SAVING)
            
            # 验证并确保输出目录存在且可写
            output_path = os.path.abspath(output_path)
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path, exist_ok=True)
                    self.logger.info(f"创建输出目录: {output_path}")
                except Exception as e:
                    self.logger.error(f"无法创建输出目录 {output_path}: {e}")
                    return False
            
            # 验证目录是否可写
            try:
                test_file = os.path.join(output_path, "write_test.tmp")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                self.logger.info(f"输出目录可写: {output_path}")
            except Exception as e:
                self.logger.error(f"输出目录不可写 {output_path}: {e}")
                return False
            
            # 移动文件到目标目录
            moved_count = 0
            for i, image_file in enumerate(image_files, 1):
                try:
                    filename = os.path.basename(image_file)
                    dest_file = os.path.join(output_path, filename)
                    
                    # 如果目标文件已存在，先删除
                    if os.path.exists(dest_file):
                        os.remove(dest_file)
                    
                    # 移动文件
                    shutil.move(image_file, dest_file)
                    moved_count += 1
                    self._progress_tracker.update_stage(i, f"移动文件 {i}/{len(image_files)}")
                    self.logger.debug(f"已移动文件: {filename}")
                except Exception as e:
                    self.logger.error(f"移动文件失败 {image_file}: {e}")
            
            self._progress_tracker.finish_stage(f"成功移动 {moved_count} 个文件到目标目录")
            
            # 验证输出目录中的文件
            import time
            time.sleep(0.5)  # 等待文件系统完成写入
            
            if os.path.exists(output_path) and os.path.isdir(output_path):
                file_count = len([f for f in os.listdir(output_path) if os.path.isfile(os.path.join(output_path, f))])
                self.logger.info(f"[保存] ✓ 图片序列目录已创建: {output_path} (共 {file_count} 个文件)")
                return True
            else:
                self.logger.error(f"[保存] ✗ 图片序列目录未创建: {output_path}")
                return False
            
        except Exception as e:
            self.logger.error(f"幻灯片转图片序列失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
    
    def _convert_to_pdf(self, presentation, output_path: str, options: ConversionOptions, slide_count: int) -> bool:
        """
        PDF导出模式
        基于图片序列模式：先导出所有幻灯片为图片，再合成PDF
        使用新的进度跟踪系统
        """
        temp_dir = None
        image_files = []
        
        try:
            # 创建临时目录
            temp_dir = self._create_temp_dir(prefix="pitchppt_pdf_")
            
            # 阶段1: 导出幻灯片为图片序列 (10-70%)
            self._progress_tracker.start_stage(ConversionStage.EXPORTING_IMAGES, slide_count)
            
            # 导出每张幻灯片为图片到临时目录
            exported_index = 0
            for i in range(1, slide_count + 1):
                slide = presentation.Slides(i)
                if slide.SlideShowTransition.Hidden and not options.include_hidden_slides:
                    self._progress_tracker.step("跳过隐藏幻灯片")
                    continue
                
                exported_index += 1
                    
                # 生成文件名（使用连续索引）
                img_format = options.image_export.format.value
                image_path = os.path.join(
                    temp_dir, 
                    f"slide_{exported_index:03d}.{img_format}"
                )
                
                # 使用导出方法
                if self._export_slide_to_image(slide, image_path, options.image_export, presentation):
                    image_files.append(image_path)
                    self._progress_tracker.update_stage(i, f"导出幻灯片 {i}/{slide_count}")
                    self.logger.debug(f"已导出幻灯片 {i} 为 {image_path}")
                else:
                    self.logger.error(f"导出幻灯片 {i} 失败")
                    self._progress_tracker.step(f"导出幻灯片 {i} 失败")
                    
            self._progress_tracker.finish_stage(f"成功导出 {len(image_files)} 张幻灯片图片")
            
            # 阶段2: 合成PDF (70-90%)
            self._progress_tracker.start_stage(ConversionStage.SAVING)
            self.logger.info(f"开始合成PDF，共 {len(image_files)} 张图片")
            
            try:
                from reportlab.lib.pagesizes import letter
                from reportlab.platypus import SimpleDocTemplate, Image, PageBreak
                from reportlab.lib.utils import ImageReader
                
                # 获取第一张图片的尺寸作为PDF页面尺寸
                if image_files:
                    with Image.open(image_files[0]) as first_img:
                        page_width, page_height = first_img.size
                        # 转换为points（1 point = 1/72 inch）
                        page_width_pts = page_width * 72 / 96  # 假设96 DPI
                        page_height_pts = page_height * 72 / 96
                else:
                    # 默认使用A4尺寸
                    page_width_pts = 595.28
                    page_height_pts = 841.89
                
                # 创建PDF文档
                doc = SimpleDocTemplate(
                    output_path,
                    pagesize=(page_width_pts, page_height_pts),
                    rightMargin=0,
                    leftMargin=0,
                    topMargin=0,
                    bottomMargin=0
                )
                
                # 创建PDF内容
                story = []
                for i, image_file in enumerate(image_files, 1):
                    try:
                        # 读取图片
                        img_reader = ImageReader(image_file)
                        img_width, img_height = img_reader.getSize()
                        
                        # 创建Image对象，保持原始尺寸
                        pdf_img = Image(img_reader, width=page_width_pts, height=page_height_pts)
                        story.append(pdf_img)
                        
                        # 添加分页符（除了最后一张）
                        if i < len(image_files):
                            story.append(PageBreak())
                        
                        self._progress_tracker.update_stage(i, f"合成PDF {i}/{len(image_files)}")
                        self.logger.debug(f"已添加图片 {i} 到PDF")
                    except Exception as e:
                        self.logger.error(f"添加图片 {i} 到PDF失败: {e}")
                
                # 构建PDF文档
                doc.build(story)
                self.logger.info(f"✓ PDF合成完成: {output_path}")
                
            except ImportError:
                self.logger.error("未安装reportlab库，无法合成PDF。请运行: pip install reportlab")
                return False
            except Exception as e:
                self.logger.error(f"合成PDF失败: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
                return False
                
            self._progress_tracker.finish_stage(f"成功合成PDF")
            
            # 验证文件是否真的被创建
            import time
            time.sleep(0.5)  # 等待文件系统完成写入
            
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                self.logger.info(f"[保存] ✓ PDF文件已创建: {output_path} ({file_size} bytes)")
                return True
            else:
                self.logger.error(f"[保存] ✗ PDF文件未创建: {output_path}")
                return False
            
        except Exception as e:
            self.logger.error(f"PDF导出失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

    def get_progress(self) -> float:
        """
        获取当前转换进度
        """
        progress, _ = self._progress_tracker.get_current_progress()
        return progress
    
    def get_conversion_info(self, input_path: str) -> Dict[str, Any]:
        """
        获取PPT文件的详细信息
        """
        result = {
            'success': False,
            'error': None,
            'file_info': {}
        }
        
        powerpoint = None
        presentation = None
        
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # 不设置Visible属性，避免某些PowerPoint版本的兼容性问题
            # powerpoint.Visible = False
            
            abs_path = os.path.abspath(input_path)
            presentation = powerpoint.Presentations.Open(abs_path)
            
            # 收集文件信息
            result['file_info'] = {
                'slide_count': presentation.Slides.Count,
                'title': getattr(presentation.BuiltInDocumentProperties, 'Title', 'Unknown'),
                'author': getattr(presentation.BuiltInDocumentProperties, 'Author', 'Unknown'),
                'created_time': getattr(presentation.BuiltInDocumentProperties, 'CreationDate', 'Unknown'),
                'modified_time': getattr(presentation.BuiltInDocumentProperties, 'LastModificationDate', 'Unknown'),
                'company': getattr(presentation.BuiltInDocumentProperties, 'Company', 'Unknown'),
                'application_name': presentation.Application.Name,
                'application_version': presentation.Application.Version,
                'page_setup': {
                    'slide_width': presentation.PageSetup.SlideWidth,
                    'slide_height': presentation.PageSetup.SlideHeight,
                    'orientation': presentation.PageSetup.Orientation
                }
            }
            
            result['success'] = True
            
        except Exception as e:
            result['error'] = str(e)
            self.logger.error(f"获取PPT信息失败: {e}")
            
        finally:
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
            if powerpoint:
                try:
                    powerpoint.Quit()
                except:
                    pass
            
        return result
    
    def batch_convert(self, input_files: List[str], output_dir: str, options: ConversionOptions = None) -> Dict[str, bool]:
        """
        批量转换多个PPT文件
        """
        if options is None:
            options = ConversionOptions()
            
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        results = {}
        
        for input_file in input_files:
            try:
                filename = os.path.basename(input_file)
                name_part = os.path.splitext(filename)[0]
                
                # 确保options参数正确传递
                if options and hasattr(options, 'output_format'):
                    output_file = os.path.join(output_dir, f"{name_part}_converted.{options.output_format.value}")
                else:
                    # 如果options无效，使用默认格式
                    options = ConversionOptions()
                    output_file = os.path.join(output_dir, f"{name_part}_converted.pptx")
                
                success = self.convert(input_file, output_file, options)
                results[input_file] = success
                
            except Exception as e:
                self.logger.error(f"批量转换文件失败 {input_file}: {e}")
                results[input_file] = False
                
        return results
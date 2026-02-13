"""
批处理转换工作线程

职责:
1. 批量处理多个PPT文件
2. 支持普通模式和智能优化模式
3. 提供详细的进度反馈
4. 支持暂停/终止功能
"""

import os
import logging
import threading
from typing import List, Callable, Optional
from PyQt5.QtCore import QThread, pyqtSignal

from src.core.win32_converter import Win32PPTConverter
from src.core.smart_optimizer_v4 import SmartOptimizerV4
from src.core.smart_optimizer_v5 import SmartOptimizerV5
from src.core.converter import ConversionOptions


class BatchConversionWorker(QThread):
    """批处理转换工作线程"""

    file_started = pyqtSignal(str)
    file_progress = pyqtSignal(str, float, str)
    file_finished = pyqtSignal(str, bool, str)
    batch_progress = pyqtSignal(float)
    batch_finished = pyqtSignal(dict)

    def __init__(self, 
                 files: List[str],
                 options: ConversionOptions,
                 output_dir: str,
                 is_smart_mode: bool = False,
                 target_size_mb: float = 10.0,
                 algorithm: str = "v4",
                 logger: Optional[logging.Logger] = None):
        super().__init__()
        
        self.files = files
        self.options = options
        self.output_dir = output_dir
        self.is_smart_mode = is_smart_mode
        self.target_size_mb = target_size_mb
        self.algorithm = algorithm
        self.logger = logger or logging.getLogger(__name__)
        
        self._paused = False
        self._stopped = False
        self._pause_condition = threading.Condition()
        
        # 共享的PowerPoint实例（用于批处理复用）
        self._shared_converter = None
        self._shared_optimizer = None

    def pause(self):
        self._paused = True
        self.logger.info("批处理已暂停")

    def resume(self):
        with self._pause_condition:
            self._paused = False
            self._pause_condition.notify_all()
        self.logger.info("批处理继续")

    def stop(self):
        self._stopped = True
        self.resume()
        self.logger.info("批处理已终止")

    def _check_pause(self):
        with self._pause_condition:
            while self._paused and not self._stopped:
                self._pause_condition.wait(0.1)

    def run(self):
        try:
            results = []
            total_files = len(self.files)
            
            self.logger.info(f"开始批处理，共 {total_files} 个文件")
            
            # 初始化共享实例
            if self.is_smart_mode:
                # 智能模式：先创建共享的converter，然后传给optimizer
                self._shared_converter = Win32PPTConverter()
                self._shared_converter._initialize_powerpoint()
                
                # 根据算法选择创建对应的optimizer
                if self.algorithm == "v5":
                    self._shared_optimizer = SmartOptimizerV5(
                        converter=self._shared_converter
                    )
                    self.logger.info("使用均衡画质贪心算法 V5")
                else:
                    self._shared_optimizer = SmartOptimizerV4(
                        logger=self.logger,
                        converter=self._shared_converter
                    )
                    self.logger.info("使用均分配额贪心算法 V4")
            else:
                self._shared_converter = Win32PPTConverter()
                self._shared_converter._initialize_powerpoint()
            
            self.logger.info("PowerPoint实例已预初始化，开始处理文件...")
            
            for i, input_file in enumerate(self.files):
                if self._stopped:
                    self.logger.info("批处理被用户终止")
                    break
                
                self._check_pause()
                
                self.file_started.emit(input_file)
                self.logger.info(f"[{i+1}/{total_files}] 开始处理: {os.path.basename(input_file)}")
                
                input_name = os.path.splitext(os.path.basename(input_file))[0]
                output_ext = self.options.output_format.value if self.options.output_format else "pptx"
                output_path = os.path.join(self.output_dir, f"{input_name}_converted.{output_ext}")
                
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(self.output_dir, f"{input_name}_converted_{counter}.{output_ext}")
                    counter += 1
                
                try:
                    if self.is_smart_mode:
                        success = self._process_smart_mode(input_file, output_path)
                    else:
                        success = self._process_normal_mode(input_file, output_path)
                    
                    self.file_finished.emit(input_file, success, output_path if success else "")
                    results.append({
                        'file': input_file,
                        'success': success,
                        'output': output_path if success else ""
                    })
                    
                    if success:
                        self.logger.info(f"[{i+1}/{total_files}] ✓ 成功: {os.path.basename(input_file)}")
                    else:
                        self.logger.warning(f"[{i+1}/{total_files}] ✗ 失败: {os.path.basename(input_file)}")
                    
                except Exception as e:
                    self.logger.error(f"[{i+1}/{total_files}] 异常: {os.path.basename(input_file)} - {e}")
                    self.file_finished.emit(input_file, False, "")
                    results.append({
                        'file': input_file,
                        'success': False,
                        'output': "",
                        'error': str(e)
                    })
                
                progress = (i + 1) / total_files
                self.batch_progress.emit(progress)
            
            success_count = sum(1 for r in results if r['success'])
            failed_count = sum(1 for r in results if not r['success'])
            
            self.logger.info(f"批处理完成: 成功 {success_count}/{total_files}, 失败 {failed_count}/{total_files}")
            
            self.batch_finished.emit({
                'total': total_files,
                'success': success_count,
                'failed': failed_count,
                'results': results
            })
            
        except Exception as e:
            self.logger.error(f"批处理线程异常: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.batch_finished.emit({
                'total': len(self.files),
                'success': 0,
                'failed': len(self.files),
                'results': [],
                'error': str(e)
            })
        finally:
            # 批处理完成后统一清理
            self._cleanup_shared_instances()

    def _cleanup_shared_instances(self):
        """清理共享实例"""
        if self._shared_converter:
            try:
                self._shared_converter._cleanup(force_kill=False)
                self.logger.info("共享Converter已清理")
            except Exception as e:
                self.logger.warning(f"清理共享Converter时出错: {e}")
            finally:
                self._shared_converter = None
        
        if self._shared_optimizer:
            try:
                self._shared_optimizer._cleanup()
                self.logger.info("共享Optimizer已清理")
            except Exception as e:
                self.logger.warning(f"清理共享Optimizer时出错: {e}")
            finally:
                self._shared_optimizer = None

    def _process_normal_mode(self, input_file: str, output_path: str) -> bool:
        """使用共享Converter处理单个文件"""
        def progress_callback(progress: float, message: str):
            self.file_progress.emit(input_file, progress, message)
        
        # 使用共享的converter，但设置临时的进度回调
        original_callback = self._shared_converter._progress_callback
        original_tracker_callback = self._shared_converter._progress_tracker.callback
        try:
            self._shared_converter._progress_callback = progress_callback
            self._shared_converter._progress_tracker.callback = progress_callback
            success = self._shared_converter.convert(input_file, output_path, self.options)
            return success
        finally:
            # 恢复原始回调
            self._shared_converter._progress_callback = original_callback
            self._shared_converter._progress_tracker.callback = original_tracker_callback

    def _process_smart_mode(self, input_file: str, output_path: str) -> bool:
        """使用共享Optimizer处理单个文件"""
        def progress_callback(message: str, progress: int):
            self.file_progress.emit(input_file, progress / 100.0, message)
        
        # 保存原始回调并设置新的回调
        original_callback = self._shared_converter._progress_callback
        original_tracker_callback = self._shared_converter._progress_tracker.callback
        
        def converter_progress_callback(progress: float, message: str):
            self.file_progress.emit(input_file, progress, message)
        
        try:
            # 设置converter的进度回调
            self._shared_converter._progress_callback = converter_progress_callback
            self._shared_converter._progress_tracker.callback = converter_progress_callback
            
            # Step 1: 计算优化参数
            self.logger.info(f"[智能模式] 开始计算优化参数: {input_file}")
            result = self._shared_optimizer.optimize(
                input_file, 
                self.target_size_mb,
                progress_callback
            )
            
            if not result.success:
                self.logger.error(f"[智能模式] 优化参数计算失败: {result.message}")
                return False
            
            self.logger.info(f"[智能模式] 优化参数计算完成，开始生成最终文件")
            
            # Step 2: 使用优化参数生成最终PPT文件
            if result.page_results:
                page_heights = [r.optimal_height for r in result.page_results]
                aspect_ratio = result.aspect_ratio
                
                success = self._export_with_page_heights(
                    input_file,
                    output_path,
                    page_heights,
                    aspect_ratio,
                    progress_callback
                )
                
                if success:
                    self.logger.info(f"[智能模式] 最终文件已生成: {output_path}")
                else:
                    self.logger.error(f"[智能模式] 最终文件生成失败")
                
                return success
            else:
                self.logger.error(f"[智能模式] 没有有效的优化结果")
                return False
                
        except Exception as e:
            self.logger.error(f"智能优化失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
        finally:
            # 恢复原始回调
            self._shared_converter._progress_callback = original_callback
            self._shared_converter._progress_tracker.callback = original_tracker_callback
    
    def _export_with_page_heights(self, input_file: str, output_path: str,
                                   page_heights: list, aspect_ratio: float,
                                   progress_callback=None) -> bool:
        """使用每页的最优高度导出完整PPT"""
        import tempfile
        import shutil
        import time
        
        try:
            self.logger.info(f"[导出] 使用优化高度导出PPT: {output_path}")
            
            # 使用共享的converter
            converter = self._shared_converter
            
            # 确保PowerPoint实例可用
            if not converter.powerpoint:
                self.logger.error("[导出] PowerPoint实例不可用")
                return False
            
            # 等待一小段时间，确保之前的操作已完成
            time.sleep(0.3)
            
            # 打开原始PPT
            abs_path = os.path.abspath(input_file)
            self.logger.info(f"[导出] 打开文件: {abs_path}")
            pres = converter.powerpoint.Presentations.Open(abs_path)
            
            try:
                slide_count = pres.Slides.Count
                self.logger.info(f"[导出] 幻灯片数量: {slide_count}")
                
                # 创建临时目录
                temp_dir = converter._create_temp_dir(prefix="smart_export_")
                self.logger.info(f"[导出] 临时目录: {temp_dir}")
                
                # 导出每一页为PNG
                image_files = []
                for page_num in range(1, slide_count + 1):
                    # 检查是否被停止
                    if self._stopped:
                        return False
                    
                    if progress_callback:
                        progress = 80 + int((page_num / slide_count) * 10)
                        progress_callback(f"正在导出第{page_num}/{slide_count}页...", progress)
                    
                    # 获取该页的最优高度
                    optimal_height = page_heights[page_num - 1]
                    optimal_width = int(optimal_height * aspect_ratio)
                    
                    # 导出为PNG
                    png_path = os.path.join(temp_dir, f"slide_{page_num:03d}.png")
                    
                    slide = pres.Slides(page_num)
                    slide.Export(png_path, "PNG", optimal_width, optimal_height)
                    
                    image_files.append(png_path)
                    self.logger.info(f"  第{page_num}页: H={optimal_height}px")
                
                # 保存原始PPT为临时文件作为模板
                template_path = os.path.join(temp_dir, "template.pptx")
                self.logger.info(f"[导出] 保存模板: {template_path}")
                try:
                    pres.SaveAs(template_path)
                except Exception as e:
                    self.logger.error(f"保存模板PPT失败: {e}")
                    return False
                
                # 关闭原始PPT
                try:
                    pres.Close()
                    self.logger.info(f"[导出] 原始PPT已关闭")
                except Exception as e:
                    self.logger.warning(f"关闭原始PPT失败: {e}")
                
                # 等待文件系统完成写入
                time.sleep(0.3)
                
                # 打开模板PPT
                self.logger.info(f"[导出] 打开模板: {template_path}")
                new_pres = converter.powerpoint.Presentations.Open(template_path)
                
                try:
                    # 设置每张幻灯片的背景为导出的图片
                    for i, image_file in enumerate(image_files, 1):
                        # 检查是否被停止
                        if self._stopped:
                            return False
                        
                        if i <= new_pres.Slides.Count:
                            if progress_callback:
                                progress = 90 + int((i / len(image_files)) * 8)
                                progress_callback(f"正在设置第{i}/{len(image_files)}页背景...", progress)
                            
                            slide = new_pres.Slides(i)
                            
                            try:
                                # 删除所有形状
                                for shape_idx in range(slide.Shapes.Count, 0, -1):
                                    try:
                                        slide.Shapes(shape_idx).Delete()
                                    except:
                                        pass
                                
                                # 设置背景填充为图片
                                slide.FollowMasterBackground = False
                                abs_image_path = os.path.abspath(image_file)
                                slide.Background.Fill.UserPicture(abs_image_path)
                                
                            except Exception as e:
                                self.logger.warning(f"设置第{i}页背景失败: {e}")
                    
                    # 保存新PPT
                    if progress_callback:
                        progress_callback("正在保存文件...", 98)
                    
                    save_path = os.path.abspath(output_path)
                    self.logger.info(f"[导出] 保存最终文件: {save_path}")
                    new_pres.SaveAs(save_path)
                    self.logger.info(f"[导出] PPT已保存: {save_path}")
                    
                    # 验证文件是否创建
                    time.sleep(0.5)
                    
                    if os.path.exists(save_path):
                        file_size = os.path.getsize(save_path)
                        self.logger.info(f"[导出] ✓ 文件已创建: {save_path} ({file_size} bytes)")
                        return True
                    else:
                        self.logger.error(f"[导出] ✗ 文件未创建: {save_path}")
                        return False
                    
                finally:
                    try:
                        new_pres.Close()
                        self.logger.info(f"[导出] 新PPT已关闭")
                    except Exception as e:
                        self.logger.warning(f"关闭新PPT失败: {e}")
                
            except Exception as e:
                self.logger.error(f"导出PPT失败: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
                # 尝试关闭presentation
                try:
                    pres.Close()
                except:
                    pass
                return False
                
        except Exception as e:
            self.logger.error(f"导出PPT失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

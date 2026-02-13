"""
智能处理工作线程 V4 - 每页独立调优

职责:
1. 使用SmartOptimizerV4/V5/V6计算每页的最优参数
2. 使用每页的最优高度导出完整PPT
3. 验证最终文件大小
4. 如果不符合预期，进行微调
"""

import os
import logging
import threading
from typing import Optional, Callable, List
from PyQt5.QtCore import QThread, pyqtSignal

from src.core.smart_optimizer_v4 import SmartOptimizerV4, OptimizationResult, PageOptimizationResult
from src.core.smart_optimizer_v5 import SmartOptimizerV5
from src.core.smart_optimizer_v6 import SmartOptimizerV6
from src.core.converter import ConversionOptions, ConversionMode, OutputFormat, ImageFormat


class SmartOptimizationWorkerV4(QThread):
    """智能处理工作线程 - 支持V4、V5和V6算法"""
    
    # 信号定义
    progress_updated = pyqtSignal(str, int)  # 消息, 进度(0-100)
    result_ready = pyqtSignal(dict)  # 结果字典
    error_occurred = pyqtSignal(str)  # 错误消息
    
    def __init__(self, 
                 pptx_path: str,
                 output_path: str,
                 target_size_mb: float,
                 dpi: int = 96,
                 algorithm: str = "v4",
                 include_hidden_slides: bool = True,
                 logger: Optional[logging.Logger] = None):
        super().__init__()
        
        self.pptx_path = pptx_path
        self.output_path = output_path
        self.target_size_mb = target_size_mb
        self.dpi = dpi
        self.algorithm = algorithm
        self.include_hidden_slides = include_hidden_slides
        self.logger = logger or logging.getLogger(__name__)
        
        # 控制标志
        self._paused = False
        self._stopped = False
        self._pause_condition = threading.Condition()
    
    def pause(self):
        """暂停处理"""
        self._paused = True
        self.logger.info("智能处理已暂停")
    
    def resume(self):
        """继续处理"""
        with self._pause_condition:
            self._paused = False
            self._pause_condition.notify_all()
        self.logger.info("智能处理继续")
    
    def stop(self):
        """终止处理"""
        self._stopped = True
        self.resume()  # 如果处于暂停状态，先唤醒线程
        self.logger.info("智能处理已终止")
    
    def _check_pause(self):
        """检查是否需要暂停"""
        with self._pause_condition:
            while self._paused and not self._stopped:
                self._pause_condition.wait(0.1)
    
    def cancel(self):
        """取消处理 - 使用正确的控制标志"""
        self._stopped = True
        self.logger.info("用户取消智能处理")
    
    def _progress_callback(self, message: str, progress: int):
        """进度回调"""
        self.logger.info(f"[DEBUG] _progress_callback 收到: progress={progress}, message={message}")
        # 检查是否暂停
        self._check_pause()
        # 检查是否终止
        if self._stopped:
            self.logger.info("[DEBUG] 检测到停止标志，抛出中断")
            raise InterruptedError("智能处理被用户终止")
        # 更新进度
        try:
            self.logger.info(f"[DEBUG] 发送信号 progress_updated: {progress}% - {message}")
            self.progress_updated.emit(message, progress)
            self.logger.info(f"[DEBUG] 信号已成功发送")
        except Exception as e:
            self.logger.error(f"[ERROR] 发送进度信号失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
    
    def _export_with_page_heights(self, page_heights: List[int], aspect_ratio: float,
                                  progress_callback=None) -> bool:
        """
        使用每页的最优高度导出完整PPT
        
        完全使用 PowerPoint COM 接口，不依赖 python-pptx
        
        Args:
            page_heights: 每页的高度列表（索引从0开始，对应原始PPT的页码）
            aspect_ratio: 宽高比（从优化结果中获取）
            progress_callback: 进度回调
            
        Returns:
            是否成功
        """
        try:
            from src.core.win32_converter import Win32PPTConverter
            import tempfile
            import shutil
            
            self.logger.info("\n使用每页的最优高度导出PPT...")
            
            def converter_progress(progress: float, message: str):
                if progress_callback:
                    progress_percent = int(80 + progress * 20)
                    progress_callback(f"正在导出: {message}", progress_percent)
            
            converter = Win32PPTConverter(converter_progress)
            
            if not converter._initialize_powerpoint():
                raise Exception("无法初始化PowerPoint")
            
            try:
                abs_path = os.path.abspath(self.pptx_path)
                pres = converter.powerpoint.Presentations.Open(abs_path)
                
                try:
                    slide_count = pres.Slides.Count
                    
                    temp_dir = converter._create_temp_dir(prefix="smart_export_")
                    
                    image_files = []
                    exported_page_indices = []
                    height_index = 0  # 用于跟踪page_heights的索引
                    
                    for page_num in range(1, slide_count + 1):
                        if self._stopped:
                            return False
                        
                        slide = pres.Slides(page_num)
                        
                        if not self.include_hidden_slides and slide.SlideShowTransition.Hidden:
                            self.logger.info(f"  跳过隐藏幻灯片 {page_num}")
                            continue
                        
                        if progress_callback:
                            progress = 80 + int((len(image_files) / slide_count) * 10)
                            progress_callback(f"正在导出第{page_num}/{slide_count}页...", progress)
                        
                        # 使用height_index来获取对应的高度
                        if height_index >= len(page_heights):
                            self.logger.error(f"高度索引越界: height_index={height_index}, page_heights长度={len(page_heights)}")
                            return False
                        
                        optimal_height = page_heights[height_index]
                        height_index += 1
                        optimal_width = int(optimal_height * aspect_ratio)
                        
                        png_path = os.path.join(temp_dir, f"slide_{page_num:03d}.png")
                        
                        slide.Export(png_path, "PNG", optimal_width, optimal_height)
                        
                        image_files.append(png_path)
                        exported_page_indices.append(page_num)
                        self.logger.info(f"  第{page_num}页: H={optimal_height}px")
                    
                    if not image_files:
                        self.logger.error("没有成功导出任何幻灯片")
                        return False
                    
                    if progress_callback:
                        progress_callback("正在创建新PPT...", 90)
                    
                    template_path = os.path.join(temp_dir, "template.pptx")
                    try:
                        pres.SaveAs(template_path)
                    except Exception as e:
                        self.logger.error(f"保存模板PPT失败: {e}")
                        return False
                    
                    try:
                        pres.Close()
                    except Exception as e:
                        self.logger.warning(f"关闭原始PPT失败（可能已自动关闭）: {e}")
                    
                    new_pres = converter.powerpoint.Presentations.Open(template_path)
                    
                    try:
                        if not self.include_hidden_slides:
                            slides_to_delete = []
                            for i in range(1, new_pres.Slides.Count + 1):
                                slide = new_pres.Slides(i)
                                if slide.SlideShowTransition.Hidden:
                                    slides_to_delete.append(i)
                            
                            for slide_idx in reversed(slides_to_delete):
                                try:
                                    new_pres.Slides(slide_idx).Delete()
                                    self.logger.info(f"删除隐藏幻灯片 {slide_idx}")
                                except Exception as e:
                                    self.logger.warning(f"删除隐藏幻灯片 {slide_idx} 失败: {e}")
                        
                        while new_pres.Slides.Count < len(image_files):
                            last_slide = new_pres.Slides(new_pres.Slides.Count)
                            last_slide.Duplicate()
                        
                        for i, image_file in enumerate(image_files, 1):
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
                                    
                                    # 设置背景填充为图片（使用绝对路径）
                                    slide.FollowMasterBackground = False
                                    abs_image_path = os.path.abspath(image_file)
                                    slide.Background.Fill.UserPicture(abs_image_path)
                                    
                                except Exception as e:
                                    self.logger.warning(f"设置第{i}页背景失败: {e}")
                        
                        # 保存新PPT
                        if progress_callback:
                            progress_callback("正在保存文件...", 98)
                        
                        new_pres.SaveAs(os.path.abspath(self.output_path))
                        self.logger.info(f"PPT已保存: {self.output_path}")
                        
                        return True
                        
                    finally:
                        try:
                            new_pres.Close()
                        except Exception as e:
                            self.logger.warning(f"关闭新PPT失败: {e}")
                    
                finally:
                    try:
                        pres.Close()
                    except:
                        pass
                    
            finally:
                converter._cleanup()
                
        except Exception as e:
            self.logger.error(f"导出PPT失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
    
    def _verify_and_fine_tune(self, opt_result: OptimizationResult,
                               progress_callback=None) -> dict:
        """
        验证最终文件大小，如果不符合预期则进行分尺度微调
        
        微调策略：
        1. 分多个尺度进行微调：50px, 20px, 10px, 5px, 1px
        2. 50px, 20px, 10px, 5px 各进行一次微调
        3. 1px 允许进行三次微调
        4. 所有尺度的微调整体称为"最终微调"
        5. 必须确保最终结果小于用户设定值
        6. 如果超过上限，退回到上一个满足条件的配置
        
        Args:
            opt_result: 优化结果
            progress_callback: 进度回调
            
        Returns:
            结果字典
        """
        target_size = self.target_size_mb
        
        # 获取实际文件大小
        actual_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
        size_error = abs(actual_size_mb - target_size) / target_size
        
        self.logger.info(f"\n[验证] 实际大小: {actual_size_mb:.2f}MB, 目标: {target_size}MB, 误差: {size_error:.1%}")
        
        result_dict = {
            'success': True,
            'page_heights': [r.optimal_height for r in opt_result.page_results],
            'target_size_mb': target_size,
            'base_volume_a_mb': opt_result.base_volume_a_mb,
            'target_per_page_mb': opt_result.target_per_page_mb,
            'total_pages': opt_result.total_pages,
            'estimated_size_mb': opt_result.estimated_final_size_mb,
            'actual_size_mb': actual_size_mb,
            'output_path': self.output_path,
            'message': f"处理完成! 实际大小: {actual_size_mb:.2f}MB (目标: {target_size}MB)"
        }
        
        # 如果误差已经在2%以内且不超过上限，无需微调
        if size_error <= 0.02 and actual_size_mb <= target_size:
            self.logger.info(f"误差{size_error:.1%}已在2%以内且未超过上限，无需微调")
            return result_dict
        
        # 检查是否超过上限
        if actual_size_mb > target_size:
            self.logger.warning(f"文件大小超过上限 ({actual_size_mb:.2f}MB > {target_size}MB)，需要降低")
        else:
            self.logger.info(f"文件大小低于上限 ({actual_size_mb:.2f}MB < {target_size}MB)，可以提高")
        
        if progress_callback:
            progress_callback(f"开始最终微调...", 90)
        
        self.logger.info(f"\n[最终微调] 开始分尺度微调...")
        
        # 定义微调尺度：50px, 20px, 10px, 5px 各一次，1px 三次
        scales = [
            (50, 1),   # 50px, 1次
            (20, 1),   # 20px, 1次
            (10, 1),   # 10px, 1次
            (5, 1),    # 5px, 1次
            (1, 3)     # 1px, 3次
        ]
        current_heights = result_dict['page_heights'].copy()
        step_counter = 0
        total_steps = sum(count for _, count in scales)
        
        # 记录满足条件的配置（用于回退）
        valid_configs = []
        if actual_size_mb <= target_size:
            valid_configs.append({
                'heights': current_heights.copy(),
                'size_mb': actual_size_mb
            })
        
        for step_size, repeat_count in scales:
            for repeat_idx in range(repeat_count):
                if self._stopped:
                    break
                
                step_counter += 1
                
                # 获取当前文件大小
                actual_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
                
                # 如果误差已在2%以内且未超过上限，停止微调
                size_error = abs(actual_size_mb - target_size) / target_size
                if size_error <= 0.02 and actual_size_mb <= target_size:
                    self.logger.info(f"误差{size_error:.1%}已在2%以内且未超过上限，微调完成")
                    break
                
                # 确定微调方向
                if actual_size_mb > target_size:
                    direction = -1  # 需要降低
                else:
                    direction = 1   # 可以提高
                
                if repeat_count == 1:
                    self.logger.info(f"\n[微调步骤{step_counter}/{total_steps}] 步长={step_size}px, 方向={'降低' if direction==-1 else '提高'}")
                else:
                    self.logger.info(f"\n[微调步骤{step_counter}/{total_steps}] 步长={step_size}px (第{repeat_idx+1}/{repeat_count}次), 方向={'降低' if direction==-1 else '提高'}")
                
                # 调整所有页面的高度
                adjusted_heights = []
                for height in current_heights:
                    new_height = height + direction * step_size
                    new_height = max(480, min(4000, new_height))
                    adjusted_heights.append(new_height)
                
                # 检查是否有实际变化
                if adjusted_heights == current_heights:
                    self.logger.info(f"  已到达边界，无法继续调整")
                    continue
                
                self.logger.info(f"  调整前高度: {current_heights[:3]}...")
                self.logger.info(f"  调整后高度: {adjusted_heights[:3]}...")
                
                # 更新进度 (90-99范围)
                if progress_callback:
                    progress_percent = 90 + int((step_counter / total_steps) * 9)
                    progress_callback(f"最终微调中(步长{step_size}px)...", progress_percent)
                
                # 重新导出
                success = self._export_with_page_heights(adjusted_heights, opt_result.aspect_ratio, None)
                
                if not success:
                    self.logger.warning(f"  微调导出失败，跳过此步骤")
                    continue
                
                # 获取新的文件大小
                new_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
                new_error = abs(new_size_mb - target_size) / target_size
                
                self.logger.info(f"  微调后大小: {new_size_mb:.2f}MB, 误差: {new_error:.1%}")
                
                # 更新当前配置
                current_heights = adjusted_heights
                result_dict['actual_size_mb'] = new_size_mb
                result_dict['page_heights'] = adjusted_heights
                
                # 记录满足条件的配置
                if new_size_mb <= target_size:
                    valid_configs.append({
                        'heights': adjusted_heights.copy(),
                        'size_mb': new_size_mb
                    })
                    self.logger.info(f"  ✓ 满足条件（≤{target_size}MB），记录配置")
                else:
                    self.logger.info(f"  ✗ 超过上限，继续尝试")
            
            # 检查是否已经满足条件
            actual_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
            size_error = abs(actual_size_mb - target_size) / target_size
            if size_error <= 0.02 and actual_size_mb <= target_size:
                break
        
        # 最终检查：如果仍然超过上限，回退到最近的满足条件的配置
        final_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
        
        if final_size_mb > target_size and valid_configs:
            self.logger.warning(f"\n[回退] 最终结果超过上限，回退到最近的满足条件的配置")
            
            # 使用最后一个满足条件的配置
            last_valid = valid_configs[-1]
            current_heights = last_valid['heights']
            
            # 重新导出
            success = self._export_with_page_heights(current_heights, opt_result.aspect_ratio, None)
            if success:
                final_size_mb = os.path.getsize(self.output_path) / (1024 * 1024)
                self.logger.info(f"  回退后大小: {final_size_mb:.2f}MB")
                result_dict['actual_size_mb'] = final_size_mb
                result_dict['page_heights'] = current_heights
        
        # 更新最终结果
        final_error = abs(final_size_mb - target_size) / target_size
        
        if final_size_mb <= target_size:
            if final_error <= 0.02:
                result_dict['message'] = f"处理完成! 实际大小: {final_size_mb:.2f}MB (目标: {target_size}MB, 误差{final_error:.1%})"
            else:
                result_dict['message'] = f"处理完成! 实际大小: {final_size_mb:.2f}MB (目标: {target_size}MB, 误差{final_error:.1%})"
        else:
            result_dict['message'] = f"处理完成(已尽力)! 实际大小: {final_size_mb:.2f}MB (超过上限{final_error:.1%})"
        
        self.logger.info(f"\n[最终结果] {result_dict['message']}")
        
        return result_dict
    
    def run(self):
        """执行智能处理和转换"""
        optimizer = None
        try:
            self.logger.info("=" * 60)
            self.logger.info(f"智能处理工作线程启动 - 算法: {self.algorithm.upper()}")
            self.logger.info("=" * 60)
            
            # ========== Phase 1: 优化计算 ==========
            self._progress_callback("正在计算最优参数...", 5)
            
            # 根据算法选择创建对应的optimizer
            if self.algorithm == "v6":
                optimizer = SmartOptimizerV6(include_hidden_slides=self.include_hidden_slides)
                self.logger.info("使用复杂度自适应算法 V6")
            elif self.algorithm == "v5":
                optimizer = SmartOptimizerV5(include_hidden_slides=self.include_hidden_slides)
                self.logger.info("使用迭代优化算法 V5")
            else:
                optimizer = SmartOptimizerV4(self.logger, include_hidden_slides=self.include_hidden_slides)
                self.logger.info("使用平均配额算法 V4")
            
            optimizer.set_dpi(self.dpi)
            # 设置停止状态回调
            optimizer.set_stopped_callback(lambda: self._stopped)
            opt_result = optimizer.optimize(
                self.pptx_path,
                self.target_size_mb,
                self._progress_callback
            )
            
            if self._stopped:
                self.logger.info("智能处理被用户终止")
                return
            
            if not opt_result.success:
                self.error_occurred.emit(opt_result.message)
                return
            
            self.logger.info(f"优化计算完成")
            self._progress_callback(f"优化完成，正在导出...", 80)
            
            # ========== Phase 2: 使用每页最优高度导出 ==========
            page_heights = [r.optimal_height for r in opt_result.page_results]
            
            success = self._export_with_page_heights(page_heights, opt_result.aspect_ratio, self._progress_callback)
            
            if self._stopped:
                self.logger.info("智能处理被用户终止")
                return
            
            if not success:
                self.error_occurred.emit("导出失败")
                return
            
            # ========== Phase 3: 验证和微调 ==========
            self._progress_callback("正在验证文件大小...", 90)
            
            result_dict = self._verify_and_fine_tune(opt_result, self._progress_callback)
            
            # ========== Phase 4: 返回结果 ==========
            self.progress_updated.emit("完成", 100)
            self.result_ready.emit(result_dict)
            
        except InterruptedError:
            self.logger.info("智能处理被用户终止")
            self.error_occurred.emit("用户终止")
        except Exception as e:
            self.logger.error(f"智能处理失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.error_occurred.emit(str(e))
        finally:
            # 确保清理optimizer资源
            if optimizer is not None:
                try:
                    optimizer._cleanup()
                    self.logger.info("Optimizer资源已清理")
                except Exception as e:
                    self.logger.warning(f"清理Optimizer资源时出错: {e}")

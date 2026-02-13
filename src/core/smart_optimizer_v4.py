"""
平均配额算法 V4 - 每页独立调优

核心概念:
- A: PPT中非图片内容的体积（通过删除所有图片内容后保存计算）
- C: 用户设置的目标上限
- N: PPT页数
- 目标每页图片体积 = (C - A) / N

优化策略:
1. 计算基准体积A（删除所有图片内容后保存）
2. 极端边界条件判断：
   - 第一页H=480px仍超过目标 → 提示用户设置太低，所有页面用最低配置
   - 第一页H=4000px仍未达到目标 → 提示用户设置太高，所有页面用最高配置
3. 每页独立调优，只调整高度H，DPI由用户设置
4. 第一页从默认值开始二分搜索
5. 后续页面参考前一页的最优H作为起点
6. 使用"正负突变"判断停止：从低于目标变为超过目标
7. 最终使用每页的最优H导出完整PPT

特点：
- 简单稳定，每页配额相同
- 适用于页面内容复杂度相近的PPT
"""

import os
import io
import tempfile
import logging
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Tuple, Optional, List, Dict
from dataclasses import dataclass

from src.core.win32_converter import Win32PPTConverter


@dataclass
class PageOptimizationResult:
    """单页优化结果"""
    page_num: int
    optimal_height: int
    actual_size_bytes: int
    target_size_bytes: int
    iterations: int
    hit_boundary: bool = False  # 是否命中极端边界


@dataclass
class OptimizationResult:
    """优化结果"""
    success: bool = False
    page_results: List[PageOptimizationResult] = None
    target_size_mb: float = 0.0
    base_volume_a_mb: float = 0.0
    target_per_page_mb: float = 0.0
    total_pages: int = 0
    estimated_final_size_mb: float = 0.0
    actual_final_size_mb: float = 0.0
    message: str = ""
    boundary_warning: str = ""  # 极端边界警告
    slide_width: float = 0.0  # 幻灯片宽度（点）
    slide_height: float = 0.0  # 幻灯片高度（点）
    aspect_ratio: float = 0.0  # 宽高比
    
    def __post_init__(self):
        if self.page_results is None:
            self.page_results = []


class SmartOptimizerV4:
    """
    智能处理器 V4 - 每页独立调优
    
    职责:
    1. 计算基准体积A（删除所有图片内容后保存）
    2. 极端边界条件判断
    3. 每页独立调优，只调整高度H
    4. 使用"正负突变"判断停止
    5. 返回每页的最优参数
    """
    
    def __init__(self, logger: Optional[logging.Logger] = None, converter: Optional[Win32PPTConverter] = None,
                 include_hidden_slides: bool = True):
        self.logger = logger or logging.getLogger(__name__)
        self.converter = converter  # 外部传入的converter（批处理时复用）
        self._own_converter = converter is None  # 是否自己创建的converter
        self.presentation = None
        self.pptx_path = None
        self.slide_count = 0
        self.slide_width = 0
        self.slide_height = 0
        self.default_dpi = 96  # 默认DPI，用户可修改
        self._stopped_callback = None  # 停止状态回调函数
        self.include_hidden_slides = include_hidden_slides  # 是否包含隐藏幻灯片
        self.hidden_slides = []  # 隐藏幻灯片索引列表
        self.visible_slide_count = 0  # 可见幻灯片数量
        
    def set_dpi(self, dpi: int):
        """设置DPI"""
        self.default_dpi = dpi
        self.logger.info(f"设置DPI: {dpi}")
        
    def set_stopped_callback(self, callback):
        """设置停止状态回调函数"""
        self._stopped_callback = callback
        
    def _is_stopped(self):
        """检查是否被停止"""
        if self._stopped_callback:
            return self._stopped_callback()
        return False
        
    def _initialize(self, pptx_path: str) -> bool:
        """初始化并打开PPT"""
        self.pptx_path = pptx_path
        
        # 如果没有外部传入的converter，则创建新的
        if self.converter is None:
            self.converter = Win32PPTConverter()
            self._own_converter = True
        
        if not self.converter._initialize_powerpoint():
            self.logger.error("无法初始化PowerPoint")
            return False
        
        try:
            abs_path = os.path.abspath(pptx_path)
            self.presentation = self.converter.powerpoint.Presentations.Open(abs_path)
            self.slide_count = self.presentation.Slides.Count
            
            # 获取幻灯片尺寸
            try:
                self.slide_width = self.presentation.PageSetup.SlideSize.Width
                self.slide_height = self.presentation.PageSetup.SlideSize.Height
            except:
                self.slide_width = 720
                self.slide_height = 540
            
            # 识别隐藏幻灯片
            self.hidden_slides = []
            for i in range(1, self.slide_count + 1):
                try:
                    slide = self.presentation.Slides(i)
                    if slide.SlideShowTransition.Hidden:
                        self.hidden_slides.append(i)
                except:
                    pass
            
            # 计算可见幻灯片数量
            if self.include_hidden_slides:
                self.visible_slide_count = self.slide_count
            else:
                self.visible_slide_count = self.slide_count - len(self.hidden_slides)
            
            self.logger.info(f"PPT: {self.slide_count}页, 可见: {self.visible_slide_count}页, 隐藏: {len(self.hidden_slides)}页, 尺寸: {self.slide_width:.0f}x{self.slide_height:.0f}")
            return True
            
        except Exception as e:
            self.logger.error(f"打开PPT失败: {e}")
            return False
    
    def _cleanup(self):
        """清理资源"""
        if self.presentation:
            try:
                self.presentation.Close()
            except:
                pass
            self.presentation = None
        
        # 只有是自己创建的converter才清理
        if self._own_converter and self.converter:
            self.converter._cleanup(force_kill=False)
            self.converter = None
    
    def _calculate_base_volume_a(self, progress_callback=None) -> float:
        """
        计算基准体积A
        
        方法:
        1. 拷贝PPT文件
        2. 删除PPT中的所有背景、页面上的所有原始内容（图片、形状、文本框等）
        3. 删除media文件夹中的所有图片
        4. 保留母版页（slideMasters、notesMasters）
        5. 仅保留注释和其他基本结构
        6. 保存后获取文件大小
        """
        try:
            self.logger.info("计算基准体积A...")
            
            if progress_callback:
                progress_callback("正在拷贝PPT文件...", 0)
            
            # 拷贝PPT文件
            with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
                temp_pptx_path = tmp.name
                shutil.copy2(self.pptx_path, temp_pptx_path)
            
            try:
                # 使用ZIP方法直接操作，删除media文件夹中的所有图片
                import zipfile
                import tempfile as tf
                
                # 创建临时目录
                temp_dir = tf.mkdtemp(prefix="ppt_base_volume_")
                
                try:
                    if progress_callback:
                        progress_callback("正在解压PPT文件...", 10)
                    
                    # 解压PPT
                    with zipfile.ZipFile(temp_pptx_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                    
                    # 首先找出被slideLayouts引用的媒体文件（母版页背景等）
                    media_to_keep = set()
                    slide_layouts_rels_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts', '_rels')
                    if os.path.exists(slide_layouts_rels_dir):
                        for rels_file in os.listdir(slide_layouts_rels_dir):
                            if rels_file.endswith('.xml.rels'):
                                file_path = os.path.join(slide_layouts_rels_dir, rels_file)
                                try:
                                    tree = ET.parse(file_path)
                                    root = tree.getroot()
                                    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                        target = rel.get('Target', '')
                                        if '../media/' in target:
                                            media_filename = os.path.basename(target)
                                            media_to_keep.add(media_filename)
                                except:
                                    pass
                    
                    # 删除media文件夹中不被引用的图片
                    media_dir = os.path.join(temp_dir, 'ppt', 'media')
                    if os.path.exists(media_dir):
                        media_files = os.listdir(media_dir)
                        deleted_count = 0
                        kept_count = 0
                        for media_file in media_files:
                            file_path = os.path.join(media_dir, media_file)
                            if os.path.isfile(file_path):
                                if media_file not in media_to_keep:
                                    os.remove(file_path)
                                    deleted_count += 1
                                else:
                                    kept_count += 1
                        self.logger.info(f"  删除了 {deleted_count} 个媒体文件，保留了 {kept_count} 个")
                    
                    if progress_callback:
                        progress_callback("正在处理嵌入对象...", 30)
                    
                    # 删除embeddings文件夹（嵌入对象）
                    embeddings_dir = os.path.join(temp_dir, 'ppt', 'embeddings')
                    if os.path.exists(embeddings_dir):
                        embedding_files = os.listdir(embeddings_dir)
                        for emb_file in embedding_files:
                            file_path = os.path.join(embeddings_dir, emb_file)
                            if os.path.isfile(file_path):
                                os.remove(file_path)
                        self.logger.info(f"  删除了 {len(embedding_files)} 个嵌入对象")
                    
                    if progress_callback:
                        progress_callback("正在清理幻灯片内容...", 50)
                    
                    # 保留slides文件夹但清空每个XML文件的内容
                    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
                    if os.path.exists(slides_dir):
                        slide_files = os.listdir(slides_dir)
                        cleaned_count = 0
                        for slide_file in slide_files:
                            if slide_file.endswith('.xml') and not slide_file.startswith('~'):
                                file_path = os.path.join(slides_dir, slide_file)
                                # 清空XML文件内容，只保留基本结构
                                with open(file_path, 'w', encoding='utf-8') as f:
                                    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                                    f.write('<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ')
                                    f.write('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
                                    f.write('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n')
                                    f.write('  <p:cSld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>\n')
                                    f.write('</p:sld>\n')
                                cleaned_count += 1
                        self.logger.info(f"  清空了 {cleaned_count} 个幻灯片XML文件")
                    
                    if progress_callback:
                        progress_callback("正在清理关系文件...", 70)
                    
                    # 清理slides/_rels文件夹中的关系文件，只保留notesSlide和slideLayout引用
                    slides_rels_dir = os.path.join(temp_dir, 'ppt', 'slides', '_rels')
                    if os.path.exists(slides_rels_dir):
                        rels_files = os.listdir(slides_rels_dir)
                        cleaned_rels_count = 0
                        for rels_file in rels_files:
                            if rels_file.endswith('.xml.rels'):
                                file_path = os.path.join(slides_rels_dir, rels_file)
                                try:
                                    tree = ET.parse(file_path)
                                    root = tree.getroot()
                                    
                                    # 只保留notesSlide和slideLayout引用，删除其他
                                    relationships_to_keep = []
                                    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                        rel_type = rel.get('Type', '')
                                        if 'notesSlide' in rel_type or 'slideLayout' in rel_type:
                                            relationships_to_keep.append(rel)
                                    
                                    # 清空并重新添加
                                    root.clear()
                                    for rel in relationships_to_keep:
                                        root.append(rel)
                                    
                                    # 保存
                                    tree.write(file_path, encoding='utf-8', xml_declaration=True)
                                    cleaned_rels_count += 1
                                except Exception as e:
                                    self.logger.warning(f"  处理 {rels_file} 时出错: {e}")
                        self.logger.info(f"  清理了 {cleaned_rels_count} 个关系文件")
                    
                    if progress_callback:
                        progress_callback("正在重新打包PPT文件...", 90)
                    
                    # 重新打包为PPTX（保留母版页、注释页等）
                    new_temp_pptx = os.path.join(temp_dir, "cleaned.pptx")
                    with zipfile.ZipFile(new_temp_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, temp_dir)
                                zip_out.write(file_path, arcname)
                    
                    # 替换原文件
                    shutil.move(new_temp_pptx, temp_pptx_path)
                    
                    # 获取文件大小
                    base_volume_a_mb = os.path.getsize(temp_pptx_path) / (1024 * 1024)
                    base_volume_a_kb = os.path.getsize(temp_pptx_path) / 1024
                    self.logger.info(f"  基准体积A: {base_volume_a_mb:.2f}MB ({base_volume_a_kb:.2f}KB)")
                    
                    if progress_callback:
                        progress_callback("基准体积计算完成", 100)
                    
                    return base_volume_a_mb
                    
                finally:
                    # 清理临时目录
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass
                    
            finally:
                # 清理临时文件
                try:
                    os.unlink(temp_pptx_path)
                except:
                    pass
            
        except Exception as e:
            self.logger.error(f"计算基准体积A失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return 0.0
    
    def _export_page_to_png(self, page_num: int, height: int) -> int:
        """
        导出单页为PNG格式，返回文件大小（字节）
        """
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                png_path = os.path.join(temp_dir, f"slide_{page_num:03d}.png")
                
                # 获取幻灯片
                slide = self.presentation.Slides(page_num)
                
                # 计算目标尺寸（保持宽高比）
                aspect_ratio = self.slide_width / self.slide_height
                target_width = int(height * aspect_ratio)
                
                # 导出为PNG
                slide.Export(png_path, "PNG", target_width, height)
                
                return os.path.getsize(png_path)
                    
        except Exception as e:
            self.logger.error(f"导出第{page_num}页失败: {e}")
            return 0
    
    def _check_boundary_conditions(self, target_size_bytes: int) -> Tuple[bool, str, int]:
        """
        检查极端边界条件
        
        Args:
            target_size_bytes: 目标每页大小（字节）
            
        Returns:
            (hit_boundary, warning_message, fixed_height)
            - hit_boundary: 是否命中极端边界
            - warning_message: 警告消息
            - fixed_height: 固定的高度（如果命中边界）
        """
        self.logger.info("\n检查极端边界条件...")
        
        # 找到第一个可见幻灯片
        first_visible_page = 1
        if not self.include_hidden_slides and self.hidden_slides:
            for i in range(1, self.slide_count + 1):
                if i not in self.hidden_slides:
                    first_visible_page = i
                    break
        
        # 测试最低配置（480px）
        min_height = 480
        min_size = self._export_page_to_png(first_visible_page, min_height)
        
        self.logger.info(f"  最低配置(H={min_height}px, 第{first_visible_page}页): {min_size/(1024*1024):.2f}MB")
        
        if min_size > target_size_bytes:
            # 最低配置仍然超过目标，用户设置太低
            warning = f"警告: 最低配置(H={min_height}px)的单页大小({min_size/(1024*1024):.2f}MB)仍超过目标({target_size_bytes/(1024*1024):.2f}MB)，目标设置过低。将使用最低配置处理所有页面。"
            self.logger.warning(warning)
            return True, warning, min_height
        
        # 测试最高配置（4000px）
        max_height = 4000
        max_size = self._export_page_to_png(first_visible_page, max_height)
        
        self.logger.info(f"  最高配置(H={max_height}px, 第{first_visible_page}页): {max_size/(1024*1024):.2f}MB")
        
        if max_size < target_size_bytes:
            # 最高配置仍未达到目标，用户设置太高
            warning = f"提示: 最高配置(H={max_height}px)的单页大小({max_size/(1024*1024):.2f}MB)仍未达到目标({target_size_bytes/(1024*1024):.2f}MB)，目标设置较高。将使用最高配置处理所有页面。"
            self.logger.warning(warning)
            return True, warning, max_height
        
        # 未命中极端边界
        self.logger.info("  未命中极端边界，正常优化")
        return False, "", 0
    
    def _optimize_single_page(self, page_num: int, target_size_bytes: int, 
                           start_height: int = 1080, progress_callback=None) -> PageOptimizationResult:
        """
        优化单页，找到最优高度
        
        使用"正负突变"判断停止：从低于目标变为超过目标
        
        Args:
            page_num: 页码
            target_size_bytes: 目标大小（字节）
            start_height: 起始高度
            progress_callback: 进度回调
            
        Returns:
            PageOptimizationResult
        """
        # 检查是否被停止
        if self._is_stopped():
            raise InterruptedError("优化被用户终止")
            
        self.logger.info(f"\n优化第{page_num}页: 目标 {target_size_bytes/(1024*1024):.2f}MB, 起始高度 {start_height}px")
        
        # 搜索范围
        height_min, height_max = 480, 4000
        current_height = start_height
        best_height = start_height
        best_size = float('inf')
        best_error = float('inf')
        iterations = 0
        
        # 阶段1: 二分搜索快速接近目标
        self.logger.info("  阶段1: 二分搜索")
        if progress_callback:
            progress_callback("二分搜索中...", 25)
        
        for iteration in range(6):
            # 检查是否被停止
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
                
            height_mid = (height_min + height_max) // 2
            
            page_size = self._export_page_to_png(page_num, height_mid)
            if page_size == 0:
                continue
            
            iterations += 1
            error = abs(page_size - target_size_bytes)
            
            self.logger.info(f"    迭代{iteration+1}: H={height_mid}px, 大小={page_size/(1024*1024):.2f}MB")
            
            # 记录最优结果（必须不大于目标）
            if page_size <= target_size_bytes and error < best_error:
                best_height = height_mid
                best_size = page_size
                best_error = error
            
            # 调整搜索范围
            if page_size > target_size_bytes:
                height_max = height_mid
            else:
                height_min = height_mid
            
            if height_max - height_min < 100:
                break
        
        # 阶段2: 分阶段精细调整
        current_height = best_height
        stage_steps = [100, 50, 10, 1]
        stage_names = ["粗调(100px)", "中调(50px)", "细调(10px)", "精调(1px)"]
        stage_progress = [50, 65, 80, 95]
        
        for stage_idx, (step, stage_name, stage_prog) in enumerate(zip(stage_steps, stage_names, stage_progress)):
            # 检查是否被停止
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
                
            self.logger.info(f"  阶段{stage_idx+2}: {stage_name}")
            
            if progress_callback:
                progress_callback(f"{stage_name}...", stage_prog)
            
            # 确定方向
            current_size = self._export_page_to_png(page_num, current_height)
            iterations += 1
            
            if current_size > target_size_bytes:
                direction = -1  # 需要降低
            else:
                direction = 1   # 可以提高
            
            self.logger.info(f"    当前: H={current_height}px, 大小={current_size/(1024*1024):.2f}MB, 方向={'降低' if direction==-1 else '提高'}")
            
            # 迭代调整，使用"正负突变"判断停止
            prev_below_target = (current_size <= target_size_bytes)
            
            while True:
                # 检查是否被停止
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")
                    
                new_height = current_height + direction * step
                
                # 检查边界（极端边界条件）
                if new_height < 480 or new_height > 4000:
                    self.logger.info(f"    到达极端边界，停止调整")
                    break
                
                new_size = self._export_page_to_png(page_num, new_height)
                iterations += 1
                
                if new_size == 0:
                    continue
                
                # 正负突变判断：从低于目标变为超过目标
                current_below_target = (new_size <= target_size_bytes)
                
                if prev_below_target and not current_below_target:
                    # 发生正负突变：从低于目标变为超过目标
                    self.logger.info(f"    发生正负突变: H={new_height}px, 大小={new_size/(1024*1024):.2f}MB > 目标")
                    self.logger.info(f"    上一个迭代步骤为最优: H={current_height}px")
                    # 回退到上一个高度
                    current_height = current_height
                    break
                
                # 更新状态
                prev_below_target = current_below_target
                current_height = new_height
                
                self.logger.info(f"    调整: H={current_height}px, 大小={new_size/(1024*1024):.2f}MB")
            
            # 更新最优结果
            if current_size <= target_size_bytes:
                best_height = current_height
                best_size = current_size
        
        # 最终验证
        final_size = self._export_page_to_png(page_num, best_height)
        iterations += 1
        
        self.logger.info(f"第{page_num}页优化完成: H={best_height}px, 大小={final_size/(1024*1024):.2f}MB, 迭代{iterations}次")
        
        return PageOptimizationResult(
            page_num=page_num,
            optimal_height=best_height,
            actual_size_bytes=final_size,
            target_size_bytes=target_size_bytes,
            iterations=iterations
        )
    
    def optimize(self, pptx_path: str, target_size_mb: float,
                progress_callback=None) -> OptimizationResult:
        """
        执行优化计算
        
        Args:
            pptx_path: 输入PPT路径
            target_size_mb: 目标文件大小（MB）
            progress_callback: 进度回调函数
            
        Returns:
            OptimizationResult
        """
        result = OptimizationResult()
        result.target_size_mb = target_size_mb
        
        self.logger.info("=" * 60)
        self.logger.info("开始智能处理 V4 - 每页独立调优")
        self.logger.info(f"目标大小: {target_size_mb}MB, DPI: {self.default_dpi}")
        self.logger.info("=" * 60)
        
        # 初始化
        if not self._initialize(pptx_path):
            result.message = "无法初始化PowerPoint"
            return result
        
        # 保存尺寸信息到结果
        result.slide_width = self.slide_width
        result.slide_height = self.slide_height
        result.aspect_ratio = self.slide_width / self.slide_height
        
        try:
            result.total_pages = self.slide_count
            
            # Step 1: 计算基准体积A
            if progress_callback:
                progress_callback("正在计算基准体积A...", 5)
            
            base_volume_a_mb = self._calculate_base_volume_a(progress_callback)
            result.base_volume_a_mb = base_volume_a_mb
            
            if base_volume_a_mb >= target_size_mb:
                result.message = f"基准体积A({base_volume_a_mb:.2f}MB)已超过目标"
                return result
            
            # Step 2: 计算目标每页图片体积（使用可见幻灯片数量）
            available_for_images = target_size_mb - base_volume_a_mb
            target_per_page_mb = available_for_images / self.visible_slide_count
            target_per_page_bytes = int(target_per_page_mb * 1024 * 1024)
            
            result.target_per_page_mb = target_per_page_mb
            
            self.logger.info(f"\n[Step 2] 计算目标每页图片体积")
            self.logger.info(f"  可用于图片: {available_for_images:.2f}MB")
            self.logger.info(f"  可见幻灯片: {self.visible_slide_count}页")
            self.logger.info(f"  目标每页: {target_per_page_mb:.2f}MB ({target_per_page_bytes} bytes)")
            
            # Step 3: 检查极端边界条件
            if progress_callback:
                progress_callback("正在检查极端边界条件...", 10)
            
            hit_boundary, warning_message, fixed_height = self._check_boundary_conditions(target_per_page_bytes)
            
            if hit_boundary:
                result.boundary_warning = warning_message
                result.message = warning_message
                
                # 使用固定高度处理所有页面
                self.logger.info(f"\n[Step 3] 使用固定高度处理所有页面: {fixed_height}px")
                
                page_results = []
                visible_page_index = 0
                for page_num in range(1, self.slide_count + 1):
                    # 跳过隐藏幻灯片
                    if not self.include_hidden_slides and page_num in self.hidden_slides:
                        self.logger.info(f"  第{page_num}页: 跳过隐藏幻灯片")
                        continue
                    
                    visible_page_index += 1
                    
                    if progress_callback:
                        progress = 10 + int((visible_page_index / self.visible_slide_count) * 80)
                        progress_callback(f"正在处理第{visible_page_index}/{self.visible_slide_count}页...", progress)
                    
                    # 使用固定高度导出
                    page_size = self._export_page_to_png(page_num, fixed_height)
                    
                    page_result = PageOptimizationResult(
                        page_num=page_num,
                        optimal_height=fixed_height,
                        actual_size_bytes=page_size,
                        target_size_bytes=target_per_page_bytes,
                        iterations=1,
                        hit_boundary=True
                    )
                    page_results.append(page_result)
                    
                    self.logger.info(f"  第{page_num}页: H={fixed_height}px, 大小={page_size/(1024*1024):.2f}MB")
                
                result.page_results = page_results
                
                # 计算预估最终大小
                total_image_size = sum(r.actual_size_bytes for r in page_results)
                estimated_final_size_mb = base_volume_a_mb + (total_image_size / (1024 * 1024))
                result.estimated_final_size_mb = estimated_final_size_mb
                
                self.logger.info(f"\n[Step 4] 预估最终大小")
                self.logger.info(f"  图片总大小: {total_image_size/(1024*1024):.2f}MB")
                self.logger.info(f"  预估最终大小: {estimated_final_size_mb:.2f}MB")
                
                result.success = True
                return result
            
            # Step 3: 每页独立调优
            if progress_callback:
                progress_callback("正在优化每一页...", 15)
            
            self.logger.info(f"\n[Step 3] 每页独立调优")
            
            page_results = []
            start_height = 1080  # 第一页的起始高度
            visible_page_index = 0  # 可见页面计数器
            
            for page_num in range(1, self.slide_count + 1):
                # 检查是否被停止
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")
                
                # 跳过隐藏幻灯片
                if not self.include_hidden_slides and page_num in self.hidden_slides:
                    self.logger.info(f"  第{page_num}页: 跳过隐藏幻灯片")
                    continue
                
                visible_page_index += 1
                
                # 计算当前页的进度范围
                page_progress_start = 15 + int((visible_page_index - 1) / self.visible_slide_count * 75)
                page_progress_end = 15 + int(visible_page_index / self.visible_slide_count * 75)
                
                # 创建页面级进度回调 - 使用默认参数捕获当前page_num值
                def page_progress_callback(msg, progress, current_page=page_num):
                    if progress_callback:
                        # 将页面内的进度映射到整体进度
                        overall_progress = page_progress_start + int((progress / 100) * (page_progress_end - page_progress_start))
                        progress_callback(f"第{current_page}页: {msg}", overall_progress)
                
                # 优化当前页
                page_result = self._optimize_single_page(
                    page_num, 
                    target_per_page_bytes,
                    start_height,
                    page_progress_callback
                )
                page_results.append(page_result)
                
                # 下一页的起始高度参考当前页的最优高度
                start_height = page_result.optimal_height
                
                self.logger.info(f"  第{page_num}页: H={page_result.optimal_height}px, "
                              f"大小={page_result.actual_size_bytes/(1024*1024):.2f}MB, "
                              f"迭代{page_result.iterations}次")
            result.page_results = page_results
            
            # Step 4: 计算预估最终大小
            total_image_size = sum(r.actual_size_bytes for r in page_results)
            estimated_final_size_mb = base_volume_a_mb + (total_image_size / (1024 * 1024))
            
            result.estimated_final_size_mb = estimated_final_size_mb
            
            self.logger.info(f"\n[Step 4] 预估最终大小")
            self.logger.info(f"  图片总大小: {total_image_size/(1024*1024):.2f}MB")
            self.logger.info(f"  预估最终大小: {estimated_final_size_mb:.2f}MB")
            self.logger.info(f"  目标大小: {target_size_mb}MB")
            self.logger.info(f"  误差: {abs(estimated_final_size_mb - target_size_mb)/target_size_mb:.1%}")
            
            # 完成
            result.success = True
            result.message = (
                f"处理完成! 每页已独立调优, "
                f"预估大小: {estimated_final_size_mb:.2f}MB (目标: {target_size_mb}MB)"
            )
            
            self.logger.info("\n" + "=" * 60)
            self.logger.info("处理完成")
            self.logger.info("=" * 60)
            
            return result
            
        finally:
            self._cleanup()

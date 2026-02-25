"""
复杂度自适应算法 V6 - 根据每页内容复杂度分配体积配额

核心思想：
- 在固定参考高度下导出每页为PNG，记录每页的实际文件大小
- 实际文件大小直接反映内容复杂度：大文件=内容复杂，小文件=内容简单
- 根据实际大小比例分配配额：复杂页面获得更多体积，简单页面获得较少体积

算法流程：
1. 计算基准体积A
2. 在参考高度(1080px)下导出每页，记录实际大小
3. 根据实际大小比例计算每页配额
4. 用各自的配额独立调优每页

优势：
- 内容复杂的页面获得更多体积，画质更好
- 内容简单的页面使用较少体积，节省空间
- 避免简单页面浪费体积、复杂页面画质不足的问题
"""

import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass, field
try:
    from src.core.win32_converter import Win32PPTConverter
    from src.utils.logger import Logger
except ImportError:
    from core.win32_converter import Win32PPTConverter
    from utils.logger import Logger


@dataclass
class PageComplexityInfo:
    """页面复杂度信息"""
    page_num: int
    reference_height: int
    actual_size: int


@dataclass
class PageOptimizationResult:
    """单页优化结果"""
    page_num: int
    optimal_height: int
    actual_size_bytes: int
    target_size_bytes: int
    iterations: int = 0
    hit_boundary: bool = False


@dataclass
class OptimizationResult:
    """优化结果"""
    success: bool = False
    message: str = ""
    target_size_mb: float = 0.0
    base_volume_a_mb: float = 0.0
    target_per_page_mb: float = 0.0
    estimated_final_size_mb: float = 0.0
    total_pages: int = 0
    page_results: List[PageOptimizationResult] = field(default_factory=list)
    slide_width: int = 0
    slide_height: int = 0
    aspect_ratio: float = 0.0
    boundary_warning: str = ""
    complexity_info: List[PageComplexityInfo] = field(default_factory=list)


class SmartOptimizerV6:
    """
    复杂度自适应算法 V6
    
    根据每页内容复杂度分配体积配额
    """
    
    def __init__(self, converter: Win32PPTConverter = None, default_dpi: int = 150,
                 include_hidden_slides: bool = True):
        self.logger = Logger().get_logger()
        self.default_dpi = default_dpi
        
        if converter:
            self.converter = converter
            self._own_converter = False
        else:
            self.converter = Win32PPTConverter()
            self._own_converter = True
        
        self.powerpoint = None
        self.presentation = None
        self.pptx_path = None
        self.slide_width = 0
        self.slide_height = 0
        self.slide_count = 0
        self._stop_flag = False
        self._temp_dir = None
        self._stopped_callback = None
        self.include_hidden_slides = include_hidden_slides
        self.hidden_slides = []
        self.visible_slide_count = 0
    
    def set_dpi(self, dpi: int):
        """设置DPI"""
        self.default_dpi = dpi
        self.logger.info(f"设置DPI: {dpi}")
    
    def set_stopped_callback(self, callback):
        """设置停止状态回调函数"""
        self._stopped_callback = callback
    
    def stop(self):
        """停止优化"""
        self._stop_flag = True
    
    def _is_stopped(self) -> bool:
        """检查是否被停止"""
        if self._stopped_callback:
            return self._stopped_callback()
        return self._stop_flag
    
    def _initialize(self, pptx_path: str) -> bool:
        """初始化PowerPoint和演示文稿"""
        try:
            self.pptx_path = pptx_path
            
            if self.converter.powerpoint is None:
                self.converter._initialize_powerpoint()
            
            self.powerpoint = self.converter.powerpoint
            if self.powerpoint is None:
                self.logger.error("无法初始化PowerPoint")
                return False
            
            abs_path = os.path.abspath(pptx_path)
            self.presentation = self.powerpoint.Presentations.Open(abs_path)
            
            self.slide_width = self.presentation.PageSetup.SlideWidth
            self.slide_height = self.presentation.PageSetup.SlideHeight
            self.slide_count = self.presentation.Slides.Count
            
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
            
            self._temp_dir = self.converter._create_temp_dir(prefix="smart_optimizer_v6_")
            
            self.logger.info(f"初始化成功: {self.slide_count}页, 可见: {self.visible_slide_count}页, 隐藏: {len(self.hidden_slides)}页, "
                           f"尺寸={self.slide_width}x{self.slide_height}")
            return True
            
        except Exception as e:
            self.logger.error(f"初始化失败: {e}")
            return False
    
    def _cleanup(self):
        """清理资源"""
        if self.presentation:
            try:
                self.presentation.Close()
            except:
                pass
            self.presentation = None
        
        if self._own_converter and self.converter:
            self.converter._cleanup(force_kill=False)
            self.converter = None
        
        if self._temp_dir and os.path.exists(self._temp_dir):
            try:
                import shutil
                shutil.rmtree(self._temp_dir, ignore_errors=True)
            except:
                pass
    
    def _export_page_to_png(self, page_num: int, height: int) -> int:
        """
        导出单页为PNG并返回文件大小
        
        Args:
            page_num: 页码（1-based）
            height: 导出高度（像素）
            
        Returns:
            文件大小（字节），失败返回0
        """
        try:
            slide = self.presentation.Slides(page_num)
            
            width = int(height * (self.slide_width / self.slide_height))
            
            temp_png = os.path.join(self._temp_dir, f"page_{page_num}_h{height}.png")
            
            slide.Export(temp_png, "PNG", width, height)
            
            if os.path.exists(temp_png):
                size = os.path.getsize(temp_png)
                return size
            
            return 0
            
        except Exception as e:
            self.logger.error(f"导出第{page_num}页失败: {e}")
            return 0
    
    def _calculate_base_volume_a(self, progress_callback=None) -> float:
        """
        计算基准体积A（与V4/V5完全一致）
        """
        try:
            self.logger.info("计算基准体积A...")
            
            if progress_callback:
                progress_callback("正在拷贝PPT文件...", 0)
            
            with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
                temp_pptx_path = tmp.name
                shutil.copy2(self.pptx_path, temp_pptx_path)
            
            try:
                temp_dir = tempfile.mkdtemp(prefix="ppt_base_volume_v6_")
                
                try:
                    if progress_callback:
                        progress_callback("正在解压PPT文件...", 10)
                    
                    with zipfile.ZipFile(temp_pptx_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                    
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
                    
                    media_dir = os.path.join(temp_dir, 'ppt', 'media')
                    if os.path.exists(media_dir):
                        media_files = os.listdir(media_dir)
                        deleted_count = 0
                        for media_file in media_files:
                            file_path = os.path.join(media_dir, media_file)
                            if os.path.isfile(file_path) and media_file not in media_to_keep:
                                os.remove(file_path)
                                deleted_count += 1
                        self.logger.info(f"  删除了 {deleted_count} 个媒体文件")
                    
                    if progress_callback:
                        progress_callback("正在处理嵌入对象...", 30)
                    
                    embeddings_dir = os.path.join(temp_dir, 'ppt', 'embeddings')
                    if os.path.exists(embeddings_dir):
                        for emb_file in os.listdir(embeddings_dir):
                            os.remove(os.path.join(embeddings_dir, emb_file))
                    
                    if progress_callback:
                        progress_callback("正在清理幻灯片内容...", 50)
                    
                    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
                    if os.path.exists(slides_dir):
                        for slide_file in os.listdir(slides_dir):
                            if slide_file.endswith('.xml') and not slide_file.startswith('~'):
                                file_path = os.path.join(slides_dir, slide_file)
                                with open(file_path, 'w', encoding='utf-8') as f:
                                    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                                    f.write('<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ')
                                    f.write('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
                                    f.write('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n')
                                    f.write('  <p:cSld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>\n')
                                    f.write('</p:sld>\n')
                    
                    if progress_callback:
                        progress_callback("正在清理关系文件...", 70)
                    
                    slides_rels_dir = os.path.join(temp_dir, 'ppt', 'slides', '_rels')
                    if os.path.exists(slides_rels_dir):
                        for rels_file in os.listdir(slides_rels_dir):
                            if rels_file.endswith('.xml.rels'):
                                file_path = os.path.join(slides_rels_dir, rels_file)
                                try:
                                    tree = ET.parse(file_path)
                                    root = tree.getroot()
                                    relationships_to_keep = []
                                    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                        rel_type = rel.get('Type', '')
                                        if 'notesSlide' in rel_type or 'slideLayout' in rel_type:
                                            relationships_to_keep.append(rel)
                                    root.clear()
                                    for rel in relationships_to_keep:
                                        root.append(rel)
                                    tree.write(file_path, encoding='utf-8', xml_declaration=True)
                                except:
                                    pass
                    
                    if progress_callback:
                        progress_callback("正在重新打包PPT文件...", 90)
                    
                    new_temp_pptx = os.path.join(temp_dir, "cleaned.pptx")
                    with zipfile.ZipFile(new_temp_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                if file == "cleaned.pptx":
                                    continue
                                file_path = os.path.join(root, file)
                                arc_name = os.path.relpath(file_path, temp_dir)
                                zip_out.write(file_path, arc_name)
                    
                    if os.path.exists(new_temp_pptx):
                        size_bytes = os.path.getsize(new_temp_pptx)
                        size_mb = size_bytes / (1024 * 1024)
                        self.logger.info(f"基准体积A: {size_mb:.2f}MB ({size_bytes} bytes)")
                        return size_mb
                    
                    return 0.5 * self.slide_count
                    
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
            finally:
                if os.path.exists(temp_pptx_path):
                    os.remove(temp_pptx_path)
                    
        except Exception as e:
            self.logger.warning(f"计算基准体积A失败，使用估算值: {e}")
            return 0.5 * self.slide_count
    
    def _analyze_page_complexity(self, reference_height: int = 1080,
                                  progress_callback=None) -> List[PageComplexityInfo]:
        """
        分析每页的内容复杂度
        
        在固定参考高度下导出每页，记录实际文件大小
        文件大小直接反映内容复杂度
        """
        complexity_info = []
        
        self.logger.info(f"\n[分析页面复杂度] 参考高度: {reference_height}px")
        
        visible_page_index = 0
        for page_num in range(1, self.slide_count + 1):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
            
            # 跳过隐藏幻灯片
            if not self.include_hidden_slides and page_num in self.hidden_slides:
                self.logger.info(f"  第{page_num}页: 跳过隐藏幻灯片")
                continue
            
            visible_page_index += 1
            
            if progress_callback:
                progress = int((visible_page_index / self.visible_slide_count) * 100)
                progress_callback(f"分析第{visible_page_index}/{self.visible_slide_count}页复杂度...", progress)
            
            actual_size = self._export_page_to_png(page_num, reference_height)
            
            info = PageComplexityInfo(
                page_num=page_num,
                reference_height=reference_height,
                actual_size=actual_size
            )
            complexity_info.append(info)
            
            self.logger.info(f"  第{page_num}页: 大小={actual_size/1024:.1f}KB, "
                           f"复杂度={'高' if actual_size > 500*1024 else '中' if actual_size > 200*1024 else '低'}")
        
        return complexity_info
    
    def _calculate_quotas_by_complexity(self, available_bytes: int,
                                        complexity_info: List[PageComplexityInfo]) -> List[int]:
        """
        根据内容复杂度计算每页的体积配额
        
        原则：实际文件大小大的页面（内容复杂）分配更多体积
        """
        if not complexity_info:
            return [available_bytes // self.visible_slide_count] * self.visible_slide_count
        
        total_size = sum(info.actual_size for info in complexity_info)
        
        if total_size == 0:
            return [available_bytes // self.visible_slide_count] * self.visible_slide_count
        
        quotas = []
        for info in complexity_info:
            weight = info.actual_size / total_size
            quota = int(available_bytes * weight)
            quotas.append(max(quota, 10 * 1024))
        
        total_quota = sum(quotas)
        if total_quota != available_bytes and quotas:
            diff = available_bytes - total_quota
            quotas[-1] += diff
        
        self.logger.info(f"\n[复杂度配额分配] 总可用: {available_bytes/(1024*1024):.2f}MB")
        self.logger.info(f"  参考总大小: {total_size/(1024*1024):.2f}MB")
        for i, (quota, info) in enumerate(zip(quotas, complexity_info)):
            weight = info.actual_size / total_size
            self.logger.info(f"  第{info.page_num}页: 配额={quota/(1024*1024):.2f}MB, "
                           f"权重={weight:.3f}, 参考大小={info.actual_size/(1024*1024):.2f}MB")
        
        return quotas
    
    def _optimize_single_page(self, page_num: int, target_size_bytes: int,
                              start_height: int = 1080,
                              progress_callback=None) -> PageOptimizationResult:
        """
        优化单页，找到最优高度
        """
        if self._is_stopped():
            raise InterruptedError("优化被用户终止")
            
        self.logger.info(f"\n优化第{page_num}页: 目标 {target_size_bytes/(1024*1024):.2f}MB, 起始高度 {start_height}px")
        
        height_min, height_max = 480, 4000
        current_height = start_height
        best_height = start_height
        best_size = float('inf')
        best_error = float('inf')
        iterations = 0
        
        self.logger.info("  阶段1: 二分搜索")
        if progress_callback:
            progress_callback("二分搜索中...", 25)
        
        for iteration in range(6):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
                
            height_mid = (height_min + height_max) // 2
            
            page_size = self._export_page_to_png(page_num, height_mid)
            if page_size == 0:
                continue
            
            iterations += 1
            error = abs(page_size - target_size_bytes)
            
            self.logger.info(f"    迭代{iteration+1}: H={height_mid}px, 大小={page_size/(1024*1024):.2f}MB")
            
            if page_size <= target_size_bytes and error < best_error:
                best_height = height_mid
                best_size = page_size
                best_error = error
            
            if page_size > target_size_bytes:
                height_max = height_mid
            else:
                height_min = height_mid
            
            if height_max - height_min < 100:
                break
        
        current_height = best_height
        stage_steps = [100, 50, 10, 1]
        stage_names = ["粗调(100px)", "中调(50px)", "细调(10px)", "精调(1px)"]
        stage_progress = [50, 65, 80, 95]
        
        for stage_idx, (step, stage_name, stage_prog) in enumerate(zip(stage_steps, stage_names, stage_progress)):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
                
            self.logger.info(f"  阶段{stage_idx+2}: {stage_name}")
            
            if progress_callback:
                progress_callback(f"{stage_name}...", stage_prog)
            
            current_size = self._export_page_to_png(page_num, current_height)
            iterations += 1
            
            if current_size > target_size_bytes:
                direction = -1
            else:
                direction = 1
            
            self.logger.info(f"    当前: H={current_height}px, 大小={current_size/(1024*1024):.2f}MB, 方向={'降低' if direction==-1 else '提高'}")
            
            prev_below_target = (current_size <= target_size_bytes)
            
            while True:
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")
                    
                new_height = current_height + direction * step
                
                if new_height < 480 or new_height > 4000:
                    self.logger.info(f"    到达极端边界，停止调整")
                    break
                
                new_size = self._export_page_to_png(page_num, new_height)
                iterations += 1
                
                if new_size == 0:
                    continue
                
                current_below_target = (new_size <= target_size_bytes)
                
                if prev_below_target and not current_below_target:
                    self.logger.info(f"    发生正负突变: H={new_height}px, 大小={new_size/(1024*1024):.2f}MB > 目标")
                    self.logger.info(f"    上一个迭代步骤为最优: H={current_height}px")
                    break
                
                prev_below_target = current_below_target
                current_height = new_height
                
                self.logger.info(f"    调整: H={current_height}px, 大小={new_size/(1024*1024):.2f}MB")
            
            if current_size <= target_size_bytes:
                best_height = current_height
                best_size = current_size
        
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
        执行V6复杂度自适应优化
        """
        result = OptimizationResult()
        result.target_size_mb = target_size_mb
        
        self.logger.info("=" * 60)
        self.logger.info("开始智能处理 V6 - 复杂度自适应算法")
        self.logger.info(f"目标大小: {target_size_mb}MB, DPI: {self.default_dpi}")
        self.logger.info("=" * 60)
        
        if not self._initialize(pptx_path):
            result.message = "无法初始化PowerPoint"
            return result
        
        result.slide_width = self.slide_width
        result.slide_height = self.slide_height
        result.aspect_ratio = self.slide_width / self.slide_height
        
        try:
            result.total_pages = self.slide_count
            
            if progress_callback:
                progress_callback("正在计算基准体积A...", 5)
            
            base_volume_a_mb = self._calculate_base_volume_a(progress_callback)
            result.base_volume_a_mb = base_volume_a_mb
            
            if base_volume_a_mb >= target_size_mb:
                result.message = f"基准体积A({base_volume_a_mb:.2f}MB)已超过目标"
                return result
            
            available_for_images = target_size_mb - base_volume_a_mb
            available_bytes = int(available_for_images * 1024 * 1024)
            
            avg_target_per_page_mb = available_for_images / self.visible_slide_count
            result.target_per_page_mb = avg_target_per_page_mb
            
            self.logger.info(f"\n[体积分配]")
            self.logger.info(f"  可用于图片: {available_for_images:.2f}MB")
            self.logger.info(f"  可见幻灯片: {self.visible_slide_count}页")
            self.logger.info(f"  平均每页配额: {avg_target_per_page_mb:.2f}MB")
            
            if progress_callback:
                progress_callback("正在分析页面复杂度...", 15)
            
            complexity_info = self._analyze_page_complexity(1080, progress_callback)
            result.complexity_info = complexity_info
            
            page_quotas = self._calculate_quotas_by_complexity(available_bytes, complexity_info)
            
            if progress_callback:
                progress_callback("正在优化每一页...", 40)
            
            self.logger.info(f"\n[每页独立调优]")
            
            page_results = []
            start_height = 1080
            visible_page_index = 0
            quota_index = 0
            
            for page_num in range(1, self.slide_count + 1):
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")
                
                # 跳过隐藏幻灯片
                if not self.include_hidden_slides and page_num in self.hidden_slides:
                    self.logger.info(f"  第{page_num}页: 跳过隐藏幻灯片")
                    continue
                
                target_bytes = page_quotas[quota_index]
                quota_index += 1
                visible_page_index += 1
                
                page_progress_start = 40 + int((visible_page_index - 1) / self.visible_slide_count * 55)
                page_progress_end = 40 + int(visible_page_index / self.visible_slide_count * 55)
                
                def page_progress_callback(msg, progress, current_page=page_num):
                    if progress_callback:
                        overall_progress = page_progress_start + int((progress / 100) * (page_progress_end - page_progress_start))
                        progress_callback(f"第{current_page}页: {msg}", overall_progress)
                
                page_result = self._optimize_single_page(
                    page_num,
                    target_bytes,
                    start_height,
                    page_progress_callback
                )
                page_results.append(page_result)
                start_height = page_result.optimal_height
                
                self.logger.info(f"  第{page_num}页: H={page_result.optimal_height}px, "
                               f"大小={page_result.actual_size_bytes/(1024*1024):.2f}MB")
            
            result.page_results = page_results
            
            total_image_size = sum(r.actual_size_bytes for r in page_results)
            estimated_final_size_mb = base_volume_a_mb + (total_image_size / (1024 * 1024))
            
            result.estimated_final_size_mb = estimated_final_size_mb
            
            self.logger.info(f"\n[预估最终大小]")
            self.logger.info(f"  图片总大小: {total_image_size/(1024*1024):.2f}MB")
            self.logger.info(f"  预估最终大小: {estimated_final_size_mb:.2f}MB")
            self.logger.info(f"  目标大小: {target_size_mb}MB")
            self.logger.info(f"  误差: {abs(estimated_final_size_mb - target_size_mb)/target_size_mb:.1%}")
            
            result.success = True
            result.message = (
                f"V6处理完成! 复杂度自适应分配, "
                f"预估大小: {estimated_final_size_mb:.2f}MB (目标: {target_size_mb}MB)"
            )
            
            self.logger.info("\n" + "=" * 60)
            self.logger.info("V6处理完成")
            self.logger.info("=" * 60)
            
            return result
            
        except InterruptedError:
            self.logger.info("优化被用户终止")
            result.message = "用户终止"
            return result
        except Exception as e:
            self.logger.error(f"优化失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            result.message = str(e)
            return result
        finally:
            self._cleanup()

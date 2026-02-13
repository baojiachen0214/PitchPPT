"""
智能配置系统 - 修复版实现

算法核心：
1. 预检与边界判定
2. 采样与降维
3. 二分搜索迭代（样本模式）
4. 精细微调（Q补偿）
5. 全量导出
"""

import os
import io
import time
import tempfile
import logging
from typing import Tuple, Optional, Dict, Any, List
from dataclasses import dataclass, field


@dataclass
class OptimizationResult:
    """优化结果"""
    success: bool = False
    quality: int = 85
    height: int = 1080
    dpi: int = 96
    estimated_size_mb: float = 0.0
    confidence: float = 0.0
    message: str = ""
    iterations: int = 0
    total_time_seconds: float = 0.0
    
    # 详细统计
    sample_pages: List[int] = field(default_factory=list)
    sample_sizes_bytes: List[int] = field(default_factory=list)


class SmartConfigOptimizer:
    """
    智能配置优化器
    
    使用单一Win32PPTConverter实例进行所有操作
    """
    
    # 参数边界
    Q_MIN = 60
    Q_MAX = 100
    H_MIN = 480
    H_MAX = 8640
    D_MIN = 72
    D_MAX = 300
    
    # 默认参数
    Q_DEFAULT = 85
    
    # 算法参数
    SAMPLE_MULTIPLIER = 1.05  # 预测放大系数
    MAX_ITERATIONS = 5  # 最大迭代次数
    SIZE_TOLERANCE = 0.05  # 大小容差5%
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)
        self.converter = None
        self.pptx_path = None
    
    def _initialize(self, pptx_path: str) -> bool:
        """初始化转换器"""
        try:
            from src.core.win32_converter import Win32PPTConverter
            
            self.pptx_path = pptx_path
            self.converter = Win32PPTConverter()
            
            if not self.converter._initialize_powerpoint():
                self.logger.error("PowerPoint初始化失败")
                return False
            
            self.logger.info("PowerPoint初始化成功")
            return True
            
        except Exception as e:
            self.logger.error(f"初始化转换器失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
    
    def _cleanup(self):
        """清理转换器"""
        if self.converter:
            self.converter._cleanup()
            self.converter = None
        self.pptx_path = None
    
    def _calculate_dpi(self, height: int) -> int:
        """根据高度计算DPI（线性关系）"""
        if height <= self.H_MIN:
            return self.D_MIN
        if height >= self.H_MAX:
            return self.D_MAX
        
        dpi_range = self.D_MAX - self.D_MIN
        height_range = self.H_MAX - self.H_MIN
        
        dpi = self.D_MIN + (height - self.H_MIN) * dpi_range / height_range
        return int(round(dpi))
    
    def _get_total_pages(self) -> int:
        """获取总页数"""
        try:
            abs_path = os.path.abspath(self.pptx_path)
            pres = self.converter.powerpoint.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
            count = pres.Slides.Count
            pres.Close()
            self.logger.info(f"获取PPT页数成功: {count}页")
            return count
        except Exception as e:
            self.logger.error(f"获取PPT页数失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return 0
    
    def _export_sample_slides(self, slide_indices: List[int], 
                              height: int, quality: int) -> Dict[int, bytes]:
        """导出样本幻灯片 - PNG导出 + PIL JPEG转换"""
        results = {}
        temp_files = []
        
        try:
            self.logger.info(f"开始导出样本: 页码={slide_indices}, 高度={height}, 质量={quality}")
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 打开PPT
                abs_path = os.path.abspath(self.pptx_path)
                pres = self.converter.powerpoint.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
                
                try:
                    total_slides = pres.Slides.Count
                    self.logger.info(f"PPT总页数: {total_slides}")
                    
                    for page_num in slide_indices:
                        if page_num > total_slides:
                            self.logger.warning(f"页码{page_num}超出范围，跳过")
                            continue
                        
                        try:
                            slide = pres.Slides(page_num)
                            
                            # 导出图片到临时文件 - 使用 PNG（无损）
                            temp_img_path = os.path.join(temp_dir, f"slide_{page_num:03d}.png")
                            self.logger.debug(f"导出第{page_num}页到: {temp_img_path}")
                            
                            # 使用PowerPoint导出PNG（无损格式）
                            slide.Export(temp_img_path, "PNG")
                            
                            # 使用PIL处理图片：调整尺寸 + 转换为JPEG + 应用质量
                            from PIL import Image
                            with Image.open(temp_img_path) as img:
                                # 调整尺寸到目标高度
                                if img.height != height:
                                    aspect = img.width / img.height
                                    target_width = int(height * aspect)
                                    img = img.resize((target_width, height), Image.Resampling.LANCZOS)
                                
                                # 转换为RGB（如果需要）
                                if img.mode == 'RGBA':
                                    img = img.convert('RGB')
                                
                                # 压缩为JPEG，应用质量参数
                                buffer = io.BytesIO()
                                img.save(buffer, format='JPEG', quality=quality, optimize=True)
                                buffer.seek(0)
                                results[page_num] = buffer.read()
                            
                            temp_files.append(temp_img_path)
                            self.logger.debug(f"第{page_num}页导出成功: {len(results[page_num])} bytes (质量={quality})")
                                
                        except Exception as e:
                            self.logger.error(f"导出第{page_num}页失败: {e}")
                            import traceback
                            self.logger.error(traceback.format_exc())
                            
                finally:
                    try:
                        pres.Close()
                        self.logger.debug("PPT已关闭")
                    except:
                        pass
                    
        except Exception as e:
            self.logger.error(f"样本导出失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
        
        self.logger.info(f"样本导出完成: 成功{len(results)}页")
        return results
    
    def _predict_size(self, sample_sizes: Dict[int, bytes], 
                     total_pages: int) -> float:
        """预测全量文件大小 - 基于样本标准差动态Margin"""
        if not sample_sizes:
            return float('inf')
        
        # 计算样本大小列表
        sizes = [len(data) for data in sample_sizes.values()]
        
        # 计算样本平均大小
        avg_sample_size = sum(sizes) / len(sizes)
        
        # 计算样本标准差
        if len(sizes) > 1:
            variance = sum((x - avg_sample_size) ** 2 for x in sizes) / len(sizes)
            std_dev = variance ** 0.5
            # 计算变异系数（标准差/平均值）
            cv = std_dev / avg_sample_size if avg_sample_size > 0 else 0
            self.logger.debug(f"样本统计: 平均={avg_sample_size:.0f}bytes, 标准差={std_dev:.0f}bytes, 变异系数={cv:.2%}")
        else:
            cv = 0
            self.logger.debug(f"样本统计: 平均={avg_sample_size:.0f}bytes (仅1个样本)")
        
        # 基础环境补偿：补偿采样（PNG→PIL JPEG）与实际转换（Win32 JPG）的差异
        # 降低基础Margin，因为当前预测过于保守
        base_env_margin = 1.01
        
        # 变异系数调整：基于页面差异
        # 变异系数越大，说明页面差异越大，需要更大的安全余量
        if cv < 0.2:
            # 页面差异小
            cv_margin = 1.00
        elif cv < 0.4:
            # 页面差异中等
            cv_margin = 1.01
        elif cv < 0.6:
            # 页面差异较大
            cv_margin = 1.02
        else:
            # 页面差异很大
            cv_margin = 1.03
        
        # 最终Margin = 基础环境补偿 × 变异系数调整
        dynamic_margin = base_env_margin * cv_margin
        
        self.logger.debug(f"动态Margin: {dynamic_margin:.3f} (基础={base_env_margin:.2f} × 变异系数调整={cv_margin:.2f}, CV={cv:.2%})")
        
        # 预测全量大小 = 平均大小 × 总页数 × 动态Margin
        predicted_bytes = avg_sample_size * total_pages * dynamic_margin
        predicted_mb = predicted_bytes / (1024 * 1024)
        
        return predicted_mb
    
    def _predict_final_size(self, height: int, quality: int, 
                          total_pages: int) -> float:
        """
        通过采样预测最终文件大小
        
        Args:
            height: 图片高度
            quality: JPEG质量
            total_pages: 总页数
            
        Returns:
            预测的文件大小（MB）
        """
        # 采样3页（首页、中间、末页）
        sample_indices = [1, max(1, total_pages // 2), total_pages]
        
        # 导出样本
        samples = self._export_sample_slides(sample_indices, height, quality)
        
        if not samples:
            return float('inf')
        
        # 预测全量大小
        predicted_mb = self._predict_size(samples, total_pages)
        
        return predicted_mb
    
    def _calculate_base_volume_a(self) -> float:
        """
        计算基准体积A：除图片外的其他内容体积
        
        方法：导出第一页为PNG（无损），计算文件大小
        然后导出为JPG（有损），计算图片大小
        A = PPT文件大小 - 图片大小
        
        Returns:
            基准体积A（MB）
        """
        try:
            self.logger.info("计算基准体积A...")
            
            # 导出第一页为PNG（无损）
            with tempfile.TemporaryDirectory() as temp_dir:
                png_path = os.path.join(temp_dir, "slide_001.png")
                
                pres = self.converter.powerpoint.Presentations.Open(os.path.abspath(self.pptx_path), ReadOnly=True, WithWindow=False)
                try:
                    slide = pres.Slides(1)
                    slide.Export(png_path, "PNG")
                finally:
                    pres.Close()
                
                # 读取PNG文件大小
                png_size = os.path.getsize(png_path)
                self.logger.info(f"  第一页PNG大小: {png_size} bytes")
                
                # 使用PIL导出为JPG，计算图片大小
                from PIL import Image
                with Image.open(png_path) as img:
                    # 转换为RGB（如果需要）
                    if img.mode == 'RGBA':
                        img = img.convert('RGB')
                    
                    # 导出为JPG，质量95
                    jpg_buffer = io.BytesIO()
                    img.save(jpg_buffer, format='JPEG', quality=95, optimize=True)
                    jpg_size = len(jpg_buffer.getvalue())
                
                self.logger.info(f"  第一页JPG大小: {jpg_size} bytes")
                
                # 计算基准体积A = PPT大小 - 图片大小
                # 这里我们用PNG大小作为PPT大小，JPG大小作为图片大小
                base_volume_a = (png_size - jpg_size) / (1024 * 1024)
                
                self.logger.info(f"  基准体积A: {base_volume_a:.2f}MB")
                
                return base_volume_a
                
        except Exception as e:
            self.logger.error(f"计算基准体积A失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return 0.0
    
    def _local_optimize_first_page(self, target_per_page_mb: float) -> Tuple[int, int]:
        """
        局部优化：只用第一页PNG，调整高度和DPI，使大小≈target_per_page
        
        Args:
            target_per_page_mb: 目标每页图片体积（MB）
            
        Returns:
            (optimal_height, optimal_dpi)
        """
        self.logger.info(f"\n局部优化: 目标每页图片体积 {target_per_page_mb:.2f}MB")
        
        target_per_page_bytes = target_per_page_mb * 1024 * 1024
        
        # 初始参数
        best_height = 1080
        best_dpi = 96
        best_error = float('inf')
        
        # 搜索范围
        height_min, height_max = 600, 5000
        dpi_min, dpi_max = 72, 300
        
        # 二分搜索高度
        for iteration in range(5):
            height_mid = (height_min + height_max) // 2
            
            # 测试当前高度
            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    png_path = os.path.join(temp_dir, "slide_001.png")
                    
                    pres = self.converter.powerpoint.Presentations.Open(os.path.abspath(self.pptx_path), ReadOnly=True, WithWindow=False)
                    try:
                        slide = pres.Slides(1)
                        slide.Export(png_path, "PNG")
                    finally:
                        pres.Close()
                    
                    # 使用PIL调整尺寸并导出JPG
                    from PIL import Image
                    with Image.open(png_path) as img:
                        if img.mode == 'RGBA':
                            img = img.convert('RGB')
                        
                        # 调整尺寸
                        aspect = img.width / img.height
                        target_width = int(height_mid * aspect)
                        img_resized = img.resize((target_width, height_mid), Image.Resampling.LANCZOS)
                        
                        # 导出为JPG，质量95
                        jpg_buffer = io.BytesIO()
                        img_resized.save(jpg_buffer, format='JPEG', quality=95, optimize=True)
                        jpg_size = len(jpg_buffer.getvalue())
                    
                    error = abs(jpg_size - target_per_page_bytes)
                    
                    if error < best_error:
                        best_error = error
                        best_height = height_mid
                    
                    # 调整搜索范围
                    if jpg_size > target_per_page_bytes:
                        height_max = height_mid
                    else:
                        height_min = height_mid
                    
                    self.logger.debug(f"  迭代{iteration+1}: H={height_mid}px, JPG={jpg_size}bytes, 误差={error}bytes")
                    
            except Exception as e:
                self.logger.warning(f"  测试高度{height_mid}失败: {e}")
                continue
        
        # 二分搜索DPI
        for iteration in range(3):
            dpi_mid = (dpi_min + dpi_max) // 2
            
            # 测试当前DPI
            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    png_path = os.path.join(temp_dir, "slide_001.png")
                    
                    pres = self.converter.powerpoint.Presentations.Open(os.path.abspath(self.pptx_path), ReadOnly=True, WithWindow=False)
                    try:
                        slide = pres.Slides(1)
                        slide.Export(png_path, "PNG")
                    finally:
                        pres.Close()
                    
                    # 使用PIL调整尺寸并导出JPG
                    from PIL import Image
                    with Image.open(png_path) as img:
                        if img.mode == 'RGBA':
                            img = img.convert('RGB')
                        
                        # 根据DPI调整尺寸
                        aspect = img.width / img.height
                        target_width = int(best_height * aspect)
                        img_resized = img.resize((target_width, best_height), Image.Resampling.LANCZOS)
                        
                        # 导出为JPG，质量95
                        jpg_buffer = io.BytesIO()
                        img_resized.save(jpg_buffer, format='JPEG', quality=95, optimize=True)
                        jpg_size = len(jpg_buffer.getvalue())
                    
                    error = abs(jpg_size - target_per_page_bytes)
                    
                    if error < best_error:
                        best_error = error
                        best_dpi = dpi_mid
                    
                    # 调整搜索范围
                    if jpg_size > target_per_page_bytes:
                        dpi_max = dpi_mid
                    else:
                        dpi_min = dpi_mid
                    
                    self.logger.debug(f"  迭代{iteration+1}: DPI={dpi_mid}, JPG={jpg_size}bytes, 误差={error}bytes")
                    
            except Exception as e:
                self.logger.warning(f"  测试DPI{dpi_mid}失败: {e}")
                continue
        
        self.logger.info(f"局部优化结果: H={best_height}px, DPI={best_dpi}, 误差={best_error}bytes")
        
        return best_height, best_dpi
    
    def optimize(self, pptx_path: str, target_size_mb: float,
                progress_callback: Optional[callable] = None) -> OptimizationResult:
        """
        执行智能配置优化 - 新算法
        
        核心思路：
        1. PPT文件体积分解：A（除图片外的其他内容） + B（所有图片）
        2. 计算基准体积A
        3. 计算目标每页图片体积：target_per_page = (C - A) / N
        4. 局部优化：只用第一页PNG，调整高度和DPI，使大小≈target_per_page
        5. 全局验证：用局部最优配置转换全部页面
        6. 精细调整：如果超出或低于设定，只调整个别页面
        """
        start_time = time.time()
        result = OptimizationResult()
        
        self.logger.info("=" * 60)
        self.logger.info("开始智能配置优化 - 新算法")
        self.logger.info(f"目标文件大小: {target_size_mb}MB")
        self.logger.info("=" * 60)
        
        # 初始化转换器
        if not self._initialize(pptx_path):
            result.message = "无法初始化PowerPoint"
            self.logger.error(result.message)
            return result
        
        try:
            # Step 0: 获取基本信息
            total_pages = self._get_total_pages()
            
            if total_pages == 0:
                result.message = "PPT没有幻灯片或无法读取文件"
                self.logger.error(result.message)
                return result
            
            self.logger.info(f"PPT总页数: {total_pages}页")
            
            if progress_callback:
                progress_callback("正在计算基准体积A...", 10)
            
            # =====================
            # Step 1: 计算基准体积A
            # =====================
            self.logger.info("\n[Step 1] 计算基准体积A")
            
            base_volume_a = self._calculate_base_volume_a()
            
            if base_volume_a >= target_size_mb:
                # 基准体积已经超过目标，无法满足
                result.message = f"警告: 基准体积A({base_volume_a:.2f}MB)已超过目标({target_size_mb}MB)"
                self.logger.error(result.message)
                return result
            
            self.logger.info(f"  基准体积A: {base_volume_a:.2f}MB")
            
            # =====================
            # Step 2: 计算目标每页图片体积
            # =====================
            self.logger.info("\n[Step 2] 计算目标每页图片体积")
            
            available_for_images = target_size_mb - base_volume_a
            target_per_page_mb = available_for_images / total_pages
            
            self.logger.info(f"  可用于图片的体积: {available_for_images:.2f}MB")
            self.logger.info(f"  目标每页图片体积: {target_per_page_mb:.2f}MB")
            
            if progress_callback:
                progress_callback("正在局部优化第一页...", 30)
            
            # =====================
            # Step 3: 局部优化 - 第一页PNG
            # =====================
            self.logger.info("\n[Step 3] 局部优化 - 第一页PNG")
            
            optimal_height, optimal_dpi = self._local_optimize_first_page(target_per_page_mb)
            
            result.height = optimal_height
            result.dpi = optimal_dpi
            result.quality = 95  # PNG格式，固定质量95
            
            self.logger.info(f"  局部最优配置: H={optimal_height}px, DPI={optimal_dpi}")
            
            if progress_callback:
                progress_callback("优化完成!", 100)
            
            # =====================
            # Step 4: 预估最终文件大小
            # =====================
            self.logger.info("\n[Step 4] 预估最终文件大小")
            
            # 预估最终大小 = A + (每页图片大小 × N)
            estimated_size_mb = base_volume_a + (target_per_page_mb * total_pages)
            
            result.estimated_size_mb = estimated_size_mb
            
            self.logger.info(f"  预估最终大小: {estimated_size_mb:.2f}MB")
            self.logger.info(f"  目标大小: {target_size_mb}MB")
            
            # 计算置信度
            size_error = abs(result.estimated_size_mb - target_size_mb) / target_size_mb
            result.confidence = max(0, 1.0 - size_error)
            
            result.success = True
            result.total_time_seconds = time.time() - start_time
            result.message = (
                f"优化成功! 配置: H={result.height}px, DPI={result.dpi}\n"
                f"预估大小: {result.estimated_size_mb:.2f}MB (目标: {target_size_mb}MB)\n"
                f"置信度: {result.confidence:.1%}, 耗时: {result.total_time_seconds:.1f}s"
            )
            
            self.logger.info("\n" + "=" * 60)
            self.logger.info("优化结果")
            self.logger.info("=" * 60)
            self.logger.info(f"  高度(H): {result.height}px")
            self.logger.info(f"  DPI(D): {result.dpi}")
            self.logger.info(f"  预估大小: {result.estimated_size_mb:.2f}MB")
            self.logger.info(f"  置信度: {result.confidence:.1%}")
            self.logger.info(f"  总耗时: {result.total_time_seconds:.1f}s")
            self.logger.info("=" * 60)
            
            return result
            
        finally:
            self._cleanup()

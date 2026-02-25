"""
Smart Optimizer V8

Per-slide joint optimization:
- image format (PNG/JPEG)
- image height
- JPEG quality

Goal:
maximize perceived clarity under a strict total size budget.
"""

import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from src.core.smart_optimizer_v6 import (
    SmartOptimizerV6,
    OptimizationResult,
    PageOptimizationResult,
)
from src.core.converter import ImageExportConfig, ImageFormat


@dataclass
class SlideProbe:
    page_num: int
    png_1080: int
    jpg88_1080: int
    complexity: float
    jpg_gain_ratio: float


class SmartOptimizerV8(SmartOptimizerV6):
    """V8: per-slide format/height/quality optimization."""

    def __init__(self, converter=None, default_dpi: int = 150, include_hidden_slides: bool = True):
        super().__init__(converter=converter, default_dpi=default_dpi, include_hidden_slides=include_hidden_slides)
        self._size_cache: Dict[Tuple[int, int, str, int], int] = {}
        self._export_call_count: int = 0
        self._max_export_calls: int = 0

    def _build_image_config(self, image_format: str, quality: int, height: int) -> ImageExportConfig:
        cfg = ImageExportConfig()
        cfg.format = ImageFormat(image_format.lower())
        cfg.quality = int(max(1, min(100, quality)))
        cfg.use_custom_resolution = True
        cfg.custom_height = int(height)
        cfg.custom_width = 0
        cfg.maintain_aspect_ratio = True
        cfg.optimize = True
        cfg.progressive = True
        return cfg

    def _export_page_size(self, page_num: int, height: int, image_format: str, quality: int = 95) -> int:
        key = (int(page_num), int(height), image_format.lower(), int(quality))
        if key in self._size_cache:
            return self._size_cache[key]

        self._export_call_count += 1
        if self._max_export_calls > 0 and self._export_call_count > self._max_export_calls:
            raise RuntimeError("V8 export-call budget exceeded")

        slide = self.presentation.Slides(page_num)
        ext = "jpg" if image_format.lower() in ("jpg", "jpeg") else image_format.lower()
        out_path = os.path.join(self._temp_dir, f"v8_p{page_num}_h{height}_q{quality}.{ext}")
        cfg = self._build_image_config(image_format=image_format, quality=quality, height=height)

        ok = self.converter._export_slide_to_image(slide, out_path, cfg, self.presentation)
        size = os.path.getsize(out_path) if ok and os.path.exists(out_path) else 0
        self._size_cache[key] = size
        return size

    def _visible_pages(self) -> List[int]:
        pages: List[int] = []
        for i in range(1, self.slide_count + 1):
            if not self.include_hidden_slides and i in self.hidden_slides:
                continue
            pages.append(i)
        return pages

    def _probe_slides(self, progress_callback=None) -> List[SlideProbe]:
        probes: List[SlideProbe] = []
        pages = self._visible_pages()
        for idx, page_num in enumerate(pages, 1):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
            if progress_callback:
                progress_callback(f"分析第{idx}/{len(pages)}页内容特征...", int((idx / max(len(pages), 1)) * 100))

            png_1080 = max(self._export_page_size(page_num, 1080, "png", 95), 1)
            jpg_1080 = max(self._export_page_size(page_num, 1080, "jpg", 88), 1)
            # >1 means JPEG saves bytes.
            gain_ratio = png_1080 / jpg_1080
            complexity = max(png_1080, jpg_1080)
            probes.append(
                SlideProbe(
                    page_num=page_num,
                    png_1080=png_1080,
                    jpg88_1080=jpg_1080,
                    complexity=float(complexity),
                    jpg_gain_ratio=float(gain_ratio),
                )
            )
        return probes

    def _allocate_quotas(self, available_bytes: int, probes: List[SlideProbe]) -> Dict[int, int]:
        quotas: Dict[int, int] = {}
        if not probes:
            return quotas

        # reserve a minimal per-slide image budget
        min_per_slide = 35 * 1024
        floor_sum = min_per_slide * len(probes)
        if available_bytes <= floor_sum:
            for p in probes:
                quotas[p.page_num] = max(8 * 1024, available_bytes // len(probes))
            return quotas

        remain = available_bytes - floor_sum
        # weighting: complexity + a small bonus for JPEG-friendly pages
        weights = []
        for p in probes:
            w = (p.complexity ** 0.62) * (1.0 + max(0.0, p.jpg_gain_ratio - 1.0) * 0.12)
            weights.append(max(w, 1.0))
        wsum = sum(weights)

        for p, w in zip(probes, weights):
            quotas[p.page_num] = min_per_slide + int(remain * (w / wsum))

        # rounding fix
        diff = available_bytes - sum(quotas.values())
        if probes and diff != 0:
            quotas[probes[-1].page_num] += diff
        return quotas

    def _score_candidate(self, height: int, image_format: str, quality: int) -> float:
        # perceptual utility proxy:
        # - height contributes strongly
        # - JPEG quality contributes sub-linearly
        # - PNG has no quantization loss bonus
        if image_format == "png":
            return float(height) * 1.0
        qf = max(0.0, min(1.0, quality / 100.0))
        return float(height) * (0.72 + 0.28 * (qf ** 0.7))

    def _best_height_under_quota(
        self,
        page_num: int,
        quota_bytes: int,
        image_format: str,
        quality: int,
        ref_size_1080: Optional[int] = None,
        h_min: int = 480,
        h_max: int = 4000,
    ) -> Tuple[int, int]:
        if ref_size_1080 is None or ref_size_1080 <= 0:
            ref_size_1080 = self._export_page_size(page_num, 1080, image_format, quality)
        ref_size_1080 = max(ref_size_1080, 1)
        est_h = int(1080 * ((max(quota_bytes, 1) / ref_size_1080) ** 0.55))
        est_h = max(h_min, min(h_max, est_h))
        lo = max(h_min, est_h - 700)
        hi = min(h_max, est_h + 700)
        best_h, best_s = h_min, 0

        # if lower bound already exceeds budget, keep lower bound
        s_lo = self._export_page_size(page_num, lo, image_format, quality)
        if s_lo > quota_bytes:
            return lo, s_lo

        for _ in range(8):
            if lo > hi:
                break
            mid = (lo + hi) // 2
            s_mid = self._export_page_size(page_num, mid, image_format, quality)
            if s_mid <= quota_bytes and s_mid > 0:
                best_h, best_s = mid, s_mid
                lo = mid + 1
            else:
                hi = mid - 1

        for step in (30, 8):
            nh = min(h_max, best_h + step)
            if nh != best_h:
                ns = self._export_page_size(page_num, nh, image_format, quality)
                if 0 < ns <= quota_bytes:
                    best_h, best_s = nh, ns
        return best_h, best_s

    def _optimize_slide(self, page_num: int, quota_bytes: int, probe: SlideProbe) -> Tuple[str, int, int, int]:
        candidates: List[Tuple[str, int]] = []

        # PNG only when JPEG doesn't save much.
        if probe.jpg_gain_ratio <= 1.15:
            candidates.append(("png", 95))

        tightness = quota_bytes / max(probe.jpg88_1080, 1)
        if tightness >= 1.35:
            jpg_qs = [95, 88]
        elif tightness >= 1.0:
            jpg_qs = [90, 82]
        else:
            jpg_qs = [82, 72]
        for q in jpg_qs:
            candidates.append(("jpg", q))

        best = ("jpg", jpg_qs[0], 480, 0, -1.0)  # fmt, q, h, size, score
        for fmt, q in candidates:
            ref1080 = None
            if fmt == "png":
                ref1080 = probe.png_1080
            elif q == 88:
                ref1080 = probe.jpg88_1080
            hq, sq = self._best_height_under_quota(page_num, quota_bytes, fmt, q, ref_size_1080=ref1080)
            if sq <= 0:
                continue
            sc = self._score_candidate(hq, fmt, q)
            if sc > best[4]:
                best = (fmt, q, hq, sq, sc)

        return best[0], best[1], best[2], best[3]

    def optimize(self, pptx_path: str, target_size_mb: float, progress_callback=None) -> OptimizationResult:
        result = OptimizationResult()
        result.target_size_mb = target_size_mb

        self.logger.info("=" * 60)
        self.logger.info("开始智能处理 V8 - 联合优化算法(格式+高度+质量)")
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

            available_for_images_mb = target_size_mb - base_volume_a_mb
            available_for_images_bytes = int(available_for_images_mb * 1024 * 1024)
            result.target_per_page_mb = available_for_images_mb / max(self.visible_slide_count, 1)

            if progress_callback:
                progress_callback("正在分析页面与压缩收益...", 18)
            probes = self._probe_slides(progress_callback)

            quotas = self._allocate_quotas(available_for_images_bytes, probes)
            self._export_call_count = 0
            self._max_export_calls = max(800, len(probes) * 90)

            if progress_callback:
                progress_callback("正在逐页联合优化...", 40)

            page_results: List[PageOptimizationResult] = []
            for idx, probe in enumerate(probes, 1):
                if progress_callback:
                    p = 40 + int((idx - 1) / max(len(probes), 1) * 50)
                    progress_callback(f"Optimizing slide {idx}/{len(probes)}...", p)
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")

                page_num = probe.page_num
                quota = quotas[page_num]
                fmt, q, h, actual_size = self._optimize_slide(page_num, quota, probe)

                pr = PageOptimizationResult(
                    page_num=page_num,
                    optimal_height=h,
                    actual_size_bytes=actual_size,
                    target_size_bytes=quota,
                    iterations=0,
                    hit_boundary=(h <= 480 or h >= 4000),
                )
                # attach per-slide export strategy for downstream exporter
                pr.image_format = fmt
                pr.image_quality = q
                page_results.append(pr)

                self.logger.info(
                    f"V8 第{page_num}页: fmt={fmt.upper()}, q={q}, h={h}px, "
                    f"size={actual_size/(1024*1024):.2f}MB, quota={quota/(1024*1024):.2f}MB"
                )

            result.page_results = page_results
            if progress_callback:
                progress_callback("V8 parameter optimization done, exporting...", 92)
            total_image_size_mb = sum(r.actual_size_bytes for r in page_results) / (1024 * 1024)
            result.estimated_final_size_mb = base_volume_a_mb + total_image_size_mb
            result.success = True
            result.message = (
                f"V8处理完成: 预计{result.estimated_final_size_mb:.2f}MB "
                f"(目标{target_size_mb:.2f}MB)"
            )
            return result

        except InterruptedError:
            result.message = "用户终止"
            return result
        except Exception as e:
            self.logger.error(f"V8优化失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            result.message = str(e)
            return result
        finally:
            self._cleanup()

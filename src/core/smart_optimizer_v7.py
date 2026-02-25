"""
Smart Optimizer V7

Core idea:
1. Build a per-slide size model from two sampled heights.
2. Allocate total image budget with bounded water-filling (prioritize hard slides).
3. Refine each slide around the modeled height with local search.

This keeps the same output contract as V6 for seamless integration.
"""

import math
from dataclasses import dataclass
from typing import List, Dict, Optional

from src.core.smart_optimizer_v6 import (
    SmartOptimizerV6,
    OptimizationResult,
    PageOptimizationResult,
)


@dataclass
class PageSizeModel:
    page_num: int
    h1: int
    h2: int
    s1: int
    s2: int
    alpha: float
    k: float
    s_min: int
    s_max: int
    weight: float


class SmartOptimizerV7(SmartOptimizerV6):
    """V7: model-driven budget allocation + local refinement."""

    def __init__(self, converter=None, default_dpi: int = 150, include_hidden_slides: bool = True):
        super().__init__(converter=converter, default_dpi=default_dpi, include_hidden_slides=include_hidden_slides)
        self._size_cache: Dict[tuple, int] = {}

    def _cached_export_size(self, page_num: int, height: int) -> int:
        key = (page_num, int(height))
        if key in self._size_cache:
            return self._size_cache[key]
        size = self._export_page_to_png(page_num, int(height))
        self._size_cache[key] = size
        return size

    def _build_page_models(self, progress_callback=None) -> List[PageSizeModel]:
        models: List[PageSizeModel] = []
        h1, h2 = 720, 1440
        h_min, h_max = 480, 4000
        visible_idx = 0
        ln_ratio = math.log(h2 / h1)

        for page_num in range(1, self.slide_count + 1):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
            if not self.include_hidden_slides and page_num in self.hidden_slides:
                continue

            visible_idx += 1
            if progress_callback:
                progress_callback(f"建模第{visible_idx}/{self.visible_slide_count}页...", int((visible_idx / max(self.visible_slide_count, 1)) * 100))

            s1 = max(self._cached_export_size(page_num, h1), 1)
            s2 = max(self._cached_export_size(page_num, h2), 1)

            ratio = max(s2 / s1, 1.01)
            alpha = math.log(ratio) / ln_ratio
            alpha = max(0.85, min(3.2, alpha))
            k = s1 / (h1 ** alpha)

            s_min = int(max(1, k * (h_min ** alpha)))
            s_max = int(max(s_min, k * (h_max ** alpha)))
            # Hard slides: high bytes + high growth exponent.
            weight = (s2 ** 0.6) * (alpha ** 1.15)

            models.append(
                PageSizeModel(
                    page_num=page_num,
                    h1=h1,
                    h2=h2,
                    s1=s1,
                    s2=s2,
                    alpha=alpha,
                    k=k,
                    s_min=s_min,
                    s_max=s_max,
                    weight=weight,
                )
            )

        return models

    def _allocate_quotas(self, total_budget_bytes: int, models: List[PageSizeModel]) -> List[int]:
        if not models:
            return []

        min_sum = sum(m.s_min for m in models)
        max_sum = sum(m.s_max for m in models)

        if total_budget_bytes <= min_sum:
            self.logger.warning("V7预算低于最小可行值，将使用最小配置")
            return [m.s_min for m in models]
        if total_budget_bytes >= max_sum:
            self.logger.info("V7预算高于最大可行值，将使用最高配置")
            return [m.s_max for m in models]

        quotas = [m.s_min for m in models]
        caps = [m.s_max - m.s_min for m in models]
        extra = total_budget_bytes - min_sum

        # Bounded water-filling with diminishing returns.
        gamma = 0.68
        rounds = 56
        for _ in range(rounds):
            if extra <= 0:
                break
            scores = []
            score_sum = 0.0
            for i, m in enumerate(models):
                rem = caps[i] - (quotas[i] - m.s_min)
                if rem <= 0:
                    scores.append(0.0)
                    continue
                gain_base = (quotas[i] - m.s_min + 1)
                score = m.weight / (gain_base ** gamma)
                scores.append(score)
                score_sum += score

            if score_sum <= 0:
                break

            used = 0
            for i, m in enumerate(models):
                if scores[i] <= 0:
                    continue
                rem = caps[i] - (quotas[i] - m.s_min)
                if rem <= 0:
                    continue
                share = int(extra * (scores[i] / score_sum))
                share = max(1, share)
                share = min(share, rem)
                quotas[i] += share
                used += share
                if used >= extra:
                    break

            if used <= 0:
                break
            extra -= used

        # Residual greedy fill.
        if extra > 0:
            order = sorted(range(len(models)), key=lambda i: models[i].weight, reverse=True)
            for i in order:
                if extra <= 0:
                    break
                rem = caps[i] - (quotas[i] - models[i].s_min)
                if rem <= 0:
                    continue
                put = min(rem, extra)
                quotas[i] += put
                extra -= put

        return quotas

    def _height_from_quota(self, model: PageSizeModel, quota_bytes: int) -> int:
        if model.k <= 0:
            return 1080
        h = int((max(quota_bytes, 1) / model.k) ** (1.0 / model.alpha))
        return max(480, min(4000, h))

    def _optimize_single_page_local(
        self,
        page_num: int,
        target_size_bytes: int,
        start_height: int,
        progress_callback=None,
    ) -> PageOptimizationResult:
        h_min, h_max = 480, 4000
        iterations = 0

        lo = max(h_min, start_height - 700)
        hi = min(h_max, start_height + 700)

        size_lo = self._cached_export_size(page_num, lo)
        size_hi = self._cached_export_size(page_num, hi)
        iterations += 2

        while lo > h_min and size_lo > target_size_bytes:
            lo = max(h_min, lo - 400)
            size_lo = self._cached_export_size(page_num, lo)
            iterations += 1

        while hi < h_max and size_hi < target_size_bytes:
            hi = min(h_max, hi + 400)
            size_hi = self._cached_export_size(page_num, hi)
            iterations += 1

        best_h = lo
        best_s = size_lo

        for _ in range(8):
            if self._is_stopped():
                raise InterruptedError("优化被用户终止")
            mid = (lo + hi) // 2
            s_mid = self._cached_export_size(page_num, mid)
            iterations += 1
            if s_mid <= target_size_bytes:
                best_h, best_s = mid, s_mid
                lo = mid + 1
            else:
                hi = mid - 1

        for step in [50, 20, 5, 1]:
            improved = True
            while improved:
                improved = False
                up = min(h_max, best_h + step)
                if up != best_h:
                    s_up = self._cached_export_size(page_num, up)
                    iterations += 1
                    if s_up <= target_size_bytes and up > best_h:
                        best_h, best_s = up, s_up
                        improved = True

                down = max(h_min, best_h - step)
                if down != best_h and best_s > target_size_bytes:
                    s_down = self._cached_export_size(page_num, down)
                    iterations += 1
                    if s_down < best_s:
                        best_h, best_s = down, s_down
                        improved = True

        if best_s <= 0:
            best_s = self._cached_export_size(page_num, best_h)
            iterations += 1

        return PageOptimizationResult(
            page_num=page_num,
            optimal_height=best_h,
            actual_size_bytes=best_s,
            target_size_bytes=target_size_bytes,
            iterations=iterations,
        )

    def optimize(self, pptx_path: str, target_size_mb: float, progress_callback=None) -> OptimizationResult:
        result = OptimizationResult()
        result.target_size_mb = target_size_mb

        self.logger.info("=" * 60)
        self.logger.info("开始智能处理 V7 - 预算驱动感知优化算法")
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
                progress_callback("正在建立页面大小模型...", 18)
            models = self._build_page_models(progress_callback)
            quotas = self._allocate_quotas(available_for_images_bytes, models)

            self.logger.info(f"V7预算: 图像可用 {available_for_images_mb:.2f}MB, 可见页 {self.visible_slide_count}")

            page_results: List[PageOptimizationResult] = []
            if progress_callback:
                progress_callback("正在逐页精修...", 40)

            for idx, (model, quota) in enumerate(zip(models, quotas), 1):
                if self._is_stopped():
                    raise InterruptedError("优化被用户终止")

                start_h = self._height_from_quota(model, quota)
                page_start = 40 + int((idx - 1) / max(len(models), 1) * 55)
                page_end = 40 + int(idx / max(len(models), 1) * 55)

                def page_progress(msg: str, prog: int, current_page=model.page_num):
                    if progress_callback:
                        p = page_start + int((prog / 100) * (page_end - page_start))
                        progress_callback(f"第{current_page}页 {msg}", p)

                page_result = self._optimize_single_page_local(
                    page_num=model.page_num,
                    target_size_bytes=quota,
                    start_height=start_h,
                    progress_callback=page_progress,
                )
                page_results.append(page_result)

            result.page_results = page_results
            total_image_size_mb = sum(r.actual_size_bytes for r in page_results) / (1024 * 1024)
            result.estimated_final_size_mb = base_volume_a_mb + total_image_size_mb

            result.success = True
            result.message = (
                f"V7处理完成: 预计{result.estimated_final_size_mb:.2f}MB "
                f"(目标{target_size_mb:.2f}MB)"
            )
            return result

        except InterruptedError:
            result.message = "用户终止"
            return result
        except Exception as e:
            self.logger.error(f"V7优化失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            result.message = str(e)
            return result
        finally:
            self._cleanup()


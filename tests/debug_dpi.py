"""调试DPI计算"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.core.smart_config import PPTExporter
import logging

logger = logging.getLogger(__name__)
exporter = PPTExporter("dummy.pptx", logger)

test_heights = [480, 1920, 2160, 4320, 8640]

print("DPI计算调试:")
print("-" * 40)
for h in test_heights:
    dpi = exporter._calculate_dpi(h)
    print(f"Height={h} -> DPI={dpi}")

print("\n公式: D = 72 + (H - 480) × 0.0259")
for h in [480, 1920, 2160, 4320, 8640]:
    calculated = 72 + (h - 480) * 0.0259
    print(f"  Height={h}: 72 + {h-480} × 0.0259 = 72 + {(h-480)*0.0259:.1f} = {calculated:.1f}")

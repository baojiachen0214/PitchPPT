"""
智能配置系统测试用例
用于验证智能配置系统的功能和准确性
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.core.smart_config import SmartConfigOptimizer, find_optimal_config


def test_compression_ratio():
    """测试压缩比计算"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试1: JPEG压缩比计算")
    print("=" * 60)
    
    test_qualities = [60, 75, 85, 95]
    for q in test_qualities:
        ratio = optimizer._compression_ratio(q)
        print(f"质量={q}: 压缩比={ratio:.4f}")
    
    print("\n预期：质量越高，压缩比越大（文件越大）")
    print()


def test_file_size_estimation():
    """测试文件大小预测"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试2: 文件大小预测")
    print("=" * 60)
    
    # 测试不同高度和质量组合
    test_configs = [
        (720, 75, 10),   # 720p, 质量75, 10张幻灯片
        (1080, 85, 10),  # 1080p, 质量85, 10张幻灯片
        (1440, 95, 10),  # 1440p, 质量95, 10张幻灯片
    ]
    
    for height, quality, slides in test_configs:
        size = optimizer.estimate_file_size(height, quality, slides)
        print(f"高度={height}px, 质量={quality}, 幻灯片={slides}: "
              f"预估大小={size:.2f}MB")
    
    print("\n预期：高度越高，文件大小越大（平方关系）")
    print()


def test_binary_search_height():
    """测试二分查找高度"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试3: 二分查找最优高度")
    print("=" * 60)
    
    # 测试不同目标大小
    test_targets = [5.0, 10.0, 20.0, 50.0]  # MB
    
    for target in test_targets:
        print(f"\n目标大小: {target}MB")
        height = optimizer.find_optimal_height(target, quality=85, slide_count=10)
        estimated = optimizer.estimate_file_size(height, 85, 10)
        print(f"  最优高度: {height}px")
        print(f"  预估大小: {estimated:.2f}MB")
        print(f"  误差: {abs(estimated - target):.2f}MB")
    
    print("\n预期：二分查找能快速找到接近目标的高度")
    print()


def test_binary_search_quality():
    """测试二分查找质量"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试4: 二分查找最优质量")
    print("=" * 60)
    
    # 测试不同目标大小
    test_targets = [5.0, 10.0, 20.0, 50.0]  # MB
    fixed_height = 1080
    
    for target in test_targets:
        print(f"\n目标大小: {target}MB, 固定高度: {fixed_height}px")
        quality = optimizer.find_optimal_quality(target, fixed_height, slide_count=10)
        estimated = optimizer.estimate_file_size(fixed_height, quality, 10)
        print(f"  最优质量: {quality}")
        print(f"  预估大小: {estimated:.2f}MB")
        print(f"  误差: {abs(estimated - target):.2f}MB")
    
    print("\n预期：二分查找能快速找到接近目标的质量")
    print()


def test_full_optimization():
    """测试完整优化流程"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试5: 完整优化流程")
    print("=" * 60)
    
    # 测试不同目标大小
    test_targets = [5.0, 10.0, 20.0, 50.0]  # MB
    slide_count = 10
    
    for target in test_targets:
        print(f"\n目标大小: {target}MB, 幻灯片: {slide_count}张")
        
        # 模拟采样数据
        sample_sizes = [500000, 800000, 1200000]  # 字节
        
        config = optimizer.optimize(target, slide_count, sample_sizes)
        
        print(f"  最优配置:")
        print(f"    质量: {config.quality}")
        print(f"    高度: {config.height}px")
        print(f"    DPI: {config.dpi}")
        print(f"    格式: {config.format.value}")
        print(f"    预估大小: {config.estimated_size_mb:.2f}MB")
        print(f"    置信度: {config.confidence:.2%}")
    
    print("\n预期：完整优化能找到平衡文件大小和质量的配置")
    print()


def test_adjust_config():
    """测试配置调整功能"""
    optimizer = SmartConfigOptimizer()
    
    print("=" * 60)
    print("测试6: 配置调整功能")
    print("=" * 60)
    
    # 初始配置
    target = 10.0  # MB
    slide_count = 10
    
    config = optimizer.optimize(target, slide_count)
    print(f"初始配置:")
    print(f"  质量: {config.quality}")
    print(f"  高度: {config.height}px")
    print(f"  预估大小: {config.estimated_size_mb:.2f}MB")
    
    # 模拟实际结果（假设实际比预估小10%）
    actual_size = config.estimated_size_mb * 0.9
    print(f"\n实际大小: {actual_size:.2f}MB (比预估小10%)")
    
    # 调整配置
    adjusted_config = optimizer.adjust_config(config, actual_size, target)
    print(f"调整后配置:")
    print(f"  质量: {adjusted_config.quality}")
    print(f"  高度: {adjusted_config.height}px")
    print(f"  预估大小: {adjusted_config.estimated_size_mb:.2f}MB")
    print(f"  置信度: {adjusted_config.confidence:.2%}")
    
    print("\n预期：调整功能能根据实际结果优化配置")
    print()


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("智能配置系统测试套件")
    print("=" * 60 + "\n")
    
    test_compression_ratio()
    test_file_size_estimation()
    test_binary_search_height()
    test_binary_search_quality()
    test_full_optimization()
    test_adjust_config()
    
    print("\n" + "=" * 60)
    print("所有测试完成！")
    print("=" * 60)


if __name__ == "__main__":
    run_all_tests()

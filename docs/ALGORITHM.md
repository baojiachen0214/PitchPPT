# PitchPPT Algorithm Documentation

**[English](ALGORITHM.md) | [中文](ALGORITHM.zh-CN.md)**

---

## Overview

PitchPPT implements three intelligent algorithms for precise PPT file size control. This document provides detailed technical documentation for each algorithm.

---

## License

This documentation is part of PitchPPT, licensed under [GNU Affero General Public License v3.0 (AGPLv3)](../LICENSE).

```
Copyright (C) 2024 Jiachen Bao

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
```

---

## Algorithm Comparison

| Feature | V4: Average Quota | V5: Dual-Round Optimization | V6: Iterative Optimization |
|---------|-------------------|----------------------------|---------------------------|
| **Strategy** | Equal quota per page | Test + Adjust | Complexity-based allocation |
| **Speed** | Fast | Medium | Slower |
| **Accuracy** | Good | Better | Best |
| **Best For** | Uniform content | Balanced needs | Complex layouts |

---

## V4: Average Quota Algorithm

### Concept

```
A = Base volume (PPT without images)
C = Target size limit
N = Number of slides
Quota per page = (C - A) / N
```

### Process Flow

1. **Calculate Base Volume (A)**
   - Remove all image content from PPT
   - Save and measure file size
   - This represents non-image element volume

2. **Check Boundary Conditions**
   - If H=480px still exceeds target → Use minimum config for all pages
   - If H=4000px still below target → Use maximum config for all pages

3. **Per-Page Optimization**
   - Start binary search from default height
   - Use "positive-negative mutation" as stop condition
   - Each page optimized independently

4. **Fine-Tuning**
   - Multi-scale adjustment: 50px → 20px → 10px → 5px → 1px
   - Ensure final error < 2%

### Advantages
- Simple and stable
- Predictable results
- Fast processing

### Limitations
- Doesn't account for content complexity differences
- May waste quota on simple pages
- May under-allocate to complex pages

---

## V5: Dual-Round Optimization Algorithm

### Concept

First round tests actual compression ratios, second round adjusts quotas based on results.

### Process Flow

1. **Round 1: Testing**
   - Export all pages at reference height (e.g., 1080px)
   - Record actual file size for each page
   - Calculate compression ratios

2. **Quota Recalculation**
   ```
   For each page:
   Adjusted Quota = Base Quota × (Page Compression Ratio / Average Ratio)
   ```

3. **Round 2: Optimization**
   - Use adjusted quotas for each page
   - Apply same optimization as V4

4. **Final Fine-Tuning**
   - Verify actual file size
   - Multi-scale adjustment if needed
   - Rollback to last valid config if exceeding limit

### Advantages
- Better allocation based on actual results
- Improved accuracy over V4
- Handles varying content complexity

### Limitations
- Requires two passes (slower)
- First round may be time-consuming

---

## V6: Iterative Optimization Algorithm

### Concept

Dynamically allocates quotas based on content complexity measured at reference height.

### Process Flow

1. **Complexity Assessment**
   - Export each page at reference height (1080px)
   - File size directly indicates content complexity
   - Larger file = More complex content

2. **Proportional Allocation**
   ```
   Total Available = C - A
   For each page:
   Page Quota = Total Available × (Page Size / Sum of All Page Sizes)
   ```

3. **Iterative Refinement**
   - Initial optimization with calculated quotas
   - Measure actual results
   - Adjust quotas based on errors
   - Repeat until convergence

4. **Final Assembly**
   - Combine all optimized pages
   - Verify total size
   - Apply final fine-tuning if needed

### Advantages
- Most accurate allocation
- Adapts to content complexity
- Best image quality for complex pages

### Limitations
- Slowest processing time
- Most computationally intensive

---

## Common Features

### Multi-Scale Fine-Tuning

All three algorithms use the same fine-tuning strategy:

```python
scales = [
    (50, 1),   # 50px adjustment, 1 iteration
    (20, 1),   # 20px adjustment, 1 iteration
    (10, 1),   # 10px adjustment, 1 iteration
    (5, 1),    # 5px adjustment, 1 iteration
    (1, 3)     # 1px adjustment, 3 iterations
]
```

### "Positive-Negative Mutation" Stop Condition

During binary search:
- Track whether current size is above or below target
- Stop when direction changes (below → above or above → below)
- Use the last "below target" configuration

### Rollback Mechanism

If final size exceeds user limit:
1. Track all valid configurations (size ≤ target)
2. If final exceeds limit, rollback to last valid config
3. Ensure final result always meets user requirements

---

## Implementation Details

### Core Classes

```python
# V4: Average Quota
class SmartOptimizerV4:
    def optimize(self, pptx_path, target_size_mb):
        # Calculate base volume A
        # Allocate equal quotas
        # Optimize each page
        pass

# V5: Dual-Round Optimization
class SmartOptimizerV5:
    def optimize(self, pptx_path, target_size_mb):
        # Round 1: Test at reference height
        # Recalculate quotas
        # Round 2: Optimize with adjusted quotas
        pass

# V6: Iterative Optimization
class SmartOptimizerV6:
    def optimize(self, pptx_path, target_size_mb):
        # Assess complexity at reference height
        # Allocate proportional quotas
        # Iteratively refine
        pass
```

### Key Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `target_size_mb` | User-specified size limit | 50 |
| `default_dpi` | Image resolution | 150 |
| `reference_height` | Height for complexity test | 1080px |
| `min_height` | Minimum image height | 480px |
| `max_height` | Maximum image height | 4000px |
| `error_threshold` | Acceptable error margin | 2% |

---

## Performance Characteristics

### Processing Time (50-page PPT)

| Algorithm | Average Time | Accuracy |
|-----------|--------------|----------|
| V4 | 2-3 minutes | ±3-5% |
| V5 | 4-6 minutes | ±2-3% |
| V6 | 6-10 minutes | ±1-2% |

### Memory Usage

- Peak memory: ~500MB (processing large PPTs)
- Temporary files: ~2x target size
- Cleanup: Automatic on completion/error

---

## Recommendations

### When to Use Each Algorithm

**Use V4 when:**
- Pages have similar content complexity
- Speed is priority
- Simple, predictable results needed

**Use V5 when:**
- Content varies significantly between pages
- Balance of speed and accuracy needed
- First-time users

**Use V6 when:**
- Maximum accuracy required
- Pages have vastly different complexity
- Time is not a constraint
- Final competition submission

---

## Future Improvements

1. **Machine Learning Integration**
   - Predict optimal parameters based on content analysis
   - Reduce iteration count

2. **Parallel Processing**
   - Process multiple pages simultaneously
   - Reduce total processing time

3. **Adaptive Scales**
   - Dynamically adjust fine-tuning scales
   - Based on convergence speed

---

## References

- [Win32 COM PowerPoint Interface](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.application)
- [Binary Search Algorithm](https://en.wikipedia.org/wiki/Binary_search_algorithm)
- [Image Compression Techniques](https://en.wikipedia.org/wiki/Image_compression)

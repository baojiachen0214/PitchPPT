# 🐼 PitchPPT 

<p align="center">
  <img src="resources/LOGO.svg" width="200" height="200" alt="PitchPPT Logo">
</p>

<p align="center">
  <strong>将PPT导出为全图PPT，完美保护内容并精准控制文件大小</strong>
</p>

<p align="center">
  <a href="README.md">English</a> | <strong>中文</strong>
</p>

<p align="center">
  <a href="https://github.com/baojiachen0214/PitchPPT.git">
    <img src="https://img.shields.io/badge/GitHub-Repository-blue?style=flat-square&logo=github" alt="GitHub">
  </a>
  <a href="https://gitee.com/bao-jiachen/PitchPPT.git">
    <img src="https://img.shields.io/badge/Gitee-Repository-red?style=flat-square" alt="Gitee">
  </a>
  <img src="https://img.shields.io/badge/Version-1.7.2-green?style=flat-square" alt="Version">
  <img src="https://img.shields.io/badge/License-AGPLv3-purple?style=flat-square" alt="License">
</p>

---

## 🎯 项目简介

**PitchPPT** 是一款专业的路演PPT保护工具，将您的PPT导出为**全图PPT**（每页转换为高清图片作为背景），在保护内容的同时**完美保留PPT的完整结构**：注释、批注、切换效果、演讲者备注、超链接等。

此外，**PitchPPT** 还提供了三种**控制全图PPT文件大小**的算法，能够满足国创赛、“挑战杯”等大学生创新创业赛事的路演材料准备与处理需求。

### 特色功能

- 🎯 **智能体积控制**：三种算法精准控制导出文件大小，误差<2%
- 🛡️ **内容保护**：导出为全图PPT，防篡改设计保护知识产权
- 🎨 **画质优化**：根据内容复杂度智能分配图片质量
- 🖼️ **多格式支持**：支持PNG、JPEG、TIFF、WebP、BMP格式；DPI预设72-600（最高16K分辨率）
- ⚡ **批量处理**：支持多个文件同时处理
- 🔧 **结构完整保留**：保留注释、切换效果、演讲者备注、超链接等

> 📖 **详细文档**：请参阅 [docs/](docs/) 目录
> - [算法说明](docs/ALGORITHM.zh-CN.md)
> - [常见问题](docs/FAQ.zh-CN.md)
> - [故障排除](docs/TROUBLESHOOTING.zh-CN.md)

---

## 📸 软件截图

<p align="center">
  <img src="resources/demo.png" alt="PitchPPT 软件截图" width="800">
</p>
<p align="center">
  <em>PitchPPT 主界面 - 支持智能优化模式和批量处理功能</em>
</p>

---

## 📦 安装

### 系统要求
- Windows 10/11（64位）
- Microsoft PowerPoint 2016或更高版本

### 下载
从 [GitHub Releases](https://github.com/baojiachen0214/PitchPPT/releases) 或 [Gitee Releases](https://gitee.com/bao-jiachen/PitchPPT/releases) 下载最新版本。

### 从源码运行
```bash
git clone https://github.com/baojiachen0214/PitchPPT.git
cd PitchPPT
pip install -r requirements.txt
python src\main.py
```

---

## 🎮 快速开始

### 两种核心模式

PitchPPT 提供**两种处理模式**，满足不同场景需求：

#### 1️⃣ 标准模式（高度自定义）
- **用途**：快速将PPT导出为全图PPT
- **特点**：统一画质设置，处理速度快
- **适用**：对文件大小没有严格要求的一般演示文稿

#### 2️⃣ 智能模式（精准控容）
- **用途**：精准控制输出文件大小
- **特点**：三种智能算法，误差<2%
- **适用**：有严格文件大小限制的竞赛、路演等场景

### 使用流程

1. **启动程序**：运行 `PitchPPT.exe`
2. **选择模式**：选择"标准模式"或"智能模式"
3. **添加文件**：
   - 单文件：拖放PPT文件或点击"选择文件"
   - 批量处理：点击"批量模式"添加多个文件
4. **配置参数**：
   - **标准模式**：选择图片画质（高/中/低）
   - **智能模式**：选择算法并设置目标文件大小（MB）
5. **开始转换**：点击"开始转换"，等待处理完成

### 算法选择（智能模式）

| 算法 | 特点 |
|------|------|
| **平均配额算法** | 每页相同配额，快速稳定 |
| **双轮优化算法** | 测试后调整，平衡精度 |
| **迭代优化算法** | 按复杂度分配，精度最高 |

### 输出特性

✅ **完整结构保留**：
- 注释和批注
- 幻灯片切换效果和动画
- 演讲者备注
- 超链接

✅ **内容保护**：
- 导出为全图PPT（每页变为背景图片）
- 内容不易被编辑或复制
- 完美保护知识产权

### 图片格式与画质选项

**支持的图片格式**

| 格式 | 说明 | 适用场景 |
|------|------|----------|
| PNG | 无损压缩，最高画质 | 图形、文字为主的幻灯片 |
| JPEG | 有损压缩，文件较小 | 照片、复杂图像 |
| TIFF | LZW压缩，专业级 | 印刷、档案保存 |
| WebP | 现代格式，压缩率高 | 网页、现代应用 |
| BMP | 无压缩，文件较大 | 兼容性需求 |

**DPI/清晰度预设**

| 预设 | DPI | 分辨率 (16:9) | 适用场景 |
|------|-----|---------------|----------|
| 屏幕 | 72 | 1920x1080 (FHD) | 屏幕显示 |
| 普通 | 150 | ~4000x2250 (4K) | 一般用途 |
| 高清 | 200 | ~5300x2980 | 高质量需求 |
| 打印 | 300 | ~8000x4500 (8K) | 印刷、投影 |
| 超高清 | 600 | ~16000x9000 (16K) | 最高画质 |

**智能模式**：根据目标文件大小自动优化分辨率，范围480px-4000px高度。

---

## 📄  许可证

本项目采用 [GNU Affero General Public License v3.0 (AGPLv3)](LICENSE) 开源协议。

---

<p align="center">
  <strong>⭐ 如果这个项目对您有帮助，请给我们一个Star！⭐</strong>
</p>

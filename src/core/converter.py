from abc import ABC, abstractmethod
from enum import Enum
from typing import List, Dict, Any, Optional, Tuple

class ConversionMode(Enum):
    """
    转换模式枚举
    """
    BACKGROUND_FILL = "background_fill"  # 图片作为背景填充
    FOREGROUND_IMAGE = "foreground_image"  # 图片作为前景对象
    SLIDE_TO_IMAGE = "slide_to_image"     # 幻灯片转图片序列
    MERGE_PRESENTATIONS = "merge_presentations"  # 合并演示文稿
    COMPRESS_IMAGES = "compress_images"     # 压缩图片
    EXTRACT_MEDIA = "extract_media"        # 提取媒体文件

class OutputFormat(Enum):
    """
    输出格式枚举
    """
    PPTX = "pptx"
    PDF = "pdf"
    JPG = "jpg"
    PNG = "png"
    TIFF = "tiff"
    BMP = "bmp"

class Resolution(Enum):
    """
    分辨率预设枚举
    """
    ORIGINAL = "original"
    HD_720P = (1280, 720)
    HD_1080P = (1920, 1080)
    UHD_4K = (3840, 2160)
    CUSTOM = "custom"


class ImageFormat(Enum):
    """
    图片格式枚举
    """
    PNG = "png"
    JPG = "jpg"
    JPEG = "jpeg"
    TIFF = "tiff"
    BMP = "bmp"
    GIF = "gif"
    WEBP = "webp"
    # 注意：PowerPoint无法直接导出真正的矢量SVG格式
    # 如需SVG，请使用其他工具将PDF转换为SVG


class DPIPreset(Enum):
    """
    DPI预设枚举
    """
    SCREEN_72 = 72      # 屏幕显示
    LOW_96 = 96         # 低质量
    NORMAL_150 = 150    # 普通质量
    HIGH_200 = 200      # 高质量
    PRINT_300 = 300     # 打印质量
    PHOTO_600 = 600     # 照片质量
    CUSTOM = "custom"   # 自定义


class CompressionLevel(Enum):
    """
    压缩级别枚举
    """
    NONE = 0            # 无压缩
    LOW = 2             # 低压缩（高质量）
    MEDIUM = 5          # 中等压缩
    HIGH = 8            # 高压缩（低质量）
    MAXIMUM = 10        # 最大压缩

class WatermarkPosition(Enum):
    """
    水印位置枚举
    """
    TOP_LEFT = "top_left"
    TOP_CENTER = "top_center"
    TOP_RIGHT = "top_right"
    CENTER_LEFT = "center_left"
    CENTER = "center"
    CENTER_RIGHT = "center_right"
    BOTTOM_LEFT = "bottom_left"
    BOTTOM_CENTER = "bottom_center"
    BOTTOM_RIGHT = "bottom_right"


class ImageExportConfig:
    """
    图片导出配置类
    提供像素级和清晰度定制
    """
    
    def __init__(self):
        # 基础格式设置
        self.format: ImageFormat = ImageFormat.JPG
        self.quality: int = 95  # JPG质量 (1-100)
        
        # DPI设置（清晰度）
        self.dpi_preset: DPIPreset = DPIPreset.PRINT_300
        self.custom_dpi: int = 300
        
        # 像素级分辨率设置
        self.use_custom_resolution: bool = False
        self.custom_width: int = 1920
        self.custom_height: int = 1080
        
        # 压缩设置
        self.compression_level: CompressionLevel = CompressionLevel.LOW
        self.optimize: bool = True  # 优化文件大小
        self.progressive: bool = True  # 渐进式JPG
        
        # 高级选项
        self.color_profile: str = "sRGB"  # sRGB, Adobe RGB, CMYK
        self.bit_depth: int = 24  # 24位或32位
        self.transparent_background: bool = False  # 透明背景（仅PNG/GIF/WebP）
        
        # 尺寸选项
        self.maintain_aspect_ratio: bool = True
        self.scale_factor: float = 1.0  # 缩放因子
        
    def get_effective_dpi(self) -> int:
        """获取有效的DPI值"""
        if self.dpi_preset == DPIPreset.CUSTOM:
            return self.custom_dpi
        elif isinstance(self.dpi_preset.value, int):
            return self.dpi_preset.value
        return 300
    
    def get_effective_resolution(self, original_width: int = 1920, original_height: int = 1080) -> Tuple[int, int]:
        """获取有效的分辨率
        
        如果 custom_width 为 0，则根据 custom_height 和原始宽高比自动计算宽度
        """
        if self.use_custom_resolution:
            # 如果宽度为0，则根据高度自动计算（保持宽高比）
            if self.custom_width == 0 and self.custom_height > 0:
                orig_ratio = original_width / original_height
                new_width = int(self.custom_height * orig_ratio)
                return (new_width, self.custom_height)
            
            if self.maintain_aspect_ratio:
                # 保持宽高比
                orig_ratio = original_width / original_height
                custom_ratio = self.custom_width / self.custom_height
                
                if custom_ratio > orig_ratio:
                    # 宽度受限
                    new_width = int(self.custom_height * orig_ratio)
                    return (new_width, self.custom_height)
                else:
                    # 高度受限
                    new_height = int(self.custom_width / orig_ratio)
                    return (self.custom_width, new_height)
            return (self.custom_width, self.custom_height)
        else:
            # 使用缩放因子
            return (
                int(original_width * self.scale_factor),
                int(original_height * self.scale_factor)
            )
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'format': self.format.value,
            'quality': self.quality,
            'dpi_preset': self.dpi_preset.value if isinstance(self.dpi_preset, DPIPreset) else self.dpi_preset,
            'custom_dpi': self.custom_dpi,
            'use_custom_resolution': self.use_custom_resolution,
            'custom_width': self.custom_width,
            'custom_height': self.custom_height,
            'compression_level': self.compression_level.value if isinstance(self.compression_level, CompressionLevel) else self.compression_level,
            'optimize': self.optimize,
            'progressive': self.progressive,
            'svg_embedded_jpeg_quality': self.svg_embedded_jpeg_quality,
            'svg_embedded_jpeg_optimize': self.svg_embedded_jpeg_optimize,
            'color_profile': self.color_profile,
            'bit_depth': self.bit_depth,
            'transparent_background': self.transparent_background,
            'maintain_aspect_ratio': self.maintain_aspect_ratio,
            'scale_factor': self.scale_factor
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ImageExportConfig':
        config = cls()
        if 'format' in data:
            config.format = ImageFormat(data['format'])
        if 'quality' in data:
            config.quality = data['quality']
        if 'dpi_preset' in data:
            if isinstance(data['dpi_preset'], int):
                config.dpi_preset = DPIPreset(data['dpi_preset'])
            else:
                config.dpi_preset = DPIPreset.CUSTOM
                config.custom_dpi = data.get('custom_dpi', 300)
        if 'custom_dpi' in data:
            config.custom_dpi = data['custom_dpi']
        if 'use_custom_resolution' in data:
            config.use_custom_resolution = data['use_custom_resolution']
        if 'custom_width' in data:
            config.custom_width = data['custom_width']
        if 'custom_height' in data:
            config.custom_height = data['custom_height']
        if 'compression_level' in data:
            if isinstance(data['compression_level'], int):
                config.compression_level = CompressionLevel(data['compression_level'])
        if 'optimize' in data:
            config.optimize = data['optimize']
        if 'progressive' in data:
            config.progressive = data['progressive']
        if 'svg_embedded_jpeg_quality' in data:
            config.svg_embedded_jpeg_quality = data['svg_embedded_jpeg_quality']
        if 'svg_embedded_jpeg_optimize' in data:
            config.svg_embedded_jpeg_optimize = data['svg_embedded_jpeg_optimize']
        if 'color_profile' in data:
            config.color_profile = data['color_profile']
        if 'bit_depth' in data:
            config.bit_depth = data['bit_depth']
        if 'transparent_background' in data:
            config.transparent_background = data['transparent_background']
        if 'maintain_aspect_ratio' in data:
            config.maintain_aspect_ratio = data['maintain_aspect_ratio']
        if 'scale_factor' in data:
            config.scale_factor = data['scale_factor']
        return config


class WatermarkConfig:
    """
    水印配置类
    """
    def __init__(self):
        self.enabled: bool = False
        self.text: str = ""
        self.image_path: str = None
        self.position: WatermarkPosition = WatermarkPosition.BOTTOM_RIGHT
        self.opacity: float = 0.3
        self.rotation: int = 0
        self.font_size: int = 24
        self.font_family: str = "Microsoft YaHei"
        self.color: str = "#000000"
        self.margin: int = 20
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'enabled': self.enabled,
            'text': self.text,
            'image_path': self.image_path,
            'position': self.position.value if isinstance(self.position, WatermarkPosition) else self.position,
            'opacity': self.opacity,
            'rotation': self.rotation,
            'font_size': self.font_size,
            'font_family': self.font_family,
            'color': self.color,
            'margin': self.margin
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WatermarkConfig':
        config = cls()
        if 'enabled' in data:
            config.enabled = data['enabled']
        if 'text' in data:
            config.text = data['text']
        if 'image_path' in data:
            config.image_path = data['image_path']
        if 'position' in data:
            config.position = WatermarkPosition(data['position'])
        if 'opacity' in data:
            config.opacity = data['opacity']
        if 'rotation' in data:
            config.rotation = data['rotation']
        if 'font_size' in data:
            config.font_size = data['font_size']
        if 'font_family' in data:
            config.font_family = data['font_family']
        if 'color' in data:
            config.color = data['color']
        if 'margin' in data:
            config.margin = data['margin']
        return config

class CropConfig:
    """
    裁剪配置类
    """
    def __init__(self):
        self.enabled: bool = False
        self.left: int = 0
        self.top: int = 0
        self.right: int = 0
        self.bottom: int = 0
        self.aspect_ratio: Optional[Tuple[int, int]] = None  # (width, height)
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'enabled': self.enabled,
            'left': self.left,
            'top': self.top,
            'right': self.right,
            'bottom': self.bottom,
            'aspect_ratio': self.aspect_ratio
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CropConfig':
        config = cls()
        if 'enabled' in data:
            config.enabled = data['enabled']
        if 'left' in data:
            config.left = data['left']
        if 'top' in data:
            config.top = data['top']
        if 'right' in data:
            config.right = data['right']
        if 'bottom' in data:
            config.bottom = data['bottom']
        if 'aspect_ratio' in data and data['aspect_ratio'] is not None:
            config.aspect_ratio = tuple(data['aspect_ratio'])
        return config

class ConversionOptions:
    """
    转换选项配置类
    """
    def __init__(self):
        self.mode: ConversionMode = ConversionMode.BACKGROUND_FILL
        self.output_format: OutputFormat = OutputFormat.PPTX
        self.image_quality: int = 95
        self.resolution_scale: float = 1.0
        self.resolution: Resolution = Resolution.ORIGINAL
        self.custom_resolution: Optional[Tuple[int, int]] = None
        self.preserve_aspect_ratio: bool = True
        self.include_hidden_slides: bool = False
        self.password: str = None
        
        # 图片导出配置（新增）
        self.image_export: ImageExportConfig = ImageExportConfig()
        
        # 高级选项
        self.watermark: WatermarkConfig = WatermarkConfig()
        self.crop: CropConfig = CropConfig()
        self.enable_compression: bool = True
        self.compression_level: int = 6
        self.remove_notes: bool = False
        self.remove_comments: bool = False
        self.optimize_for: str = "screen"  # screen, print, email
        
        # 导出选项
        self.export_notes: bool = False
        self.export_comments: bool = False
        self.export_hidden_slides: bool = False
        self.export_range: Optional[Tuple[int, int]] = None  # (start, end)
        
    def to_dict(self) -> Dict[str, Any]:
        return {
            'mode': self.mode.value,
            'output_format': self.output_format.value,
            'image_quality': self.image_quality,
            'resolution_scale': self.resolution_scale,
            'resolution': self.resolution.value if isinstance(self.resolution, Resolution) else self.resolution,
            'custom_resolution': self.custom_resolution,
            'preserve_aspect_ratio': self.preserve_aspect_ratio,
            'include_hidden_slides': self.include_hidden_slides,
            'image_export': self.image_export.to_dict(),
            'watermark': self.watermark.to_dict(),
            'crop': self.crop.to_dict(),
            'enable_compression': self.enable_compression,
            'compression_level': self.compression_level,
            'remove_notes': self.remove_notes,
            'remove_comments': self.remove_comments,
            'optimize_for': self.optimize_for,
            'export_notes': self.export_notes,
            'export_comments': self.export_comments,
            'export_hidden_slides': self.export_hidden_slides,
            'export_range': self.export_range
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ConversionOptions':
        options = cls()
        if 'mode' in data:
            options.mode = ConversionMode(data['mode'])
        if 'output_format' in data:
            options.output_format = OutputFormat(data['output_format'])
        if 'image_quality' in data:
            options.image_quality = data['image_quality']
        if 'resolution_scale' in data:
            options.resolution_scale = data['resolution_scale']
        if 'resolution' in data:
            options.resolution = Resolution(data['resolution'])
        if 'custom_resolution' in data and data['custom_resolution'] is not None:
            options.custom_resolution = tuple(data['custom_resolution'])
        if 'preserve_aspect_ratio' in data:
            options.preserve_aspect_ratio = data['preserve_aspect_ratio']
        if 'include_hidden_slides' in data:
            options.include_hidden_slides = data['include_hidden_slides']
        if 'image_export' in data:
            options.image_export = ImageExportConfig.from_dict(data['image_export'])
        if 'watermark' in data:
            options.watermark = WatermarkConfig.from_dict(data['watermark'])
        if 'crop' in data:
            options.crop = CropConfig.from_dict(data['crop'])
        if 'enable_compression' in data:
            options.enable_compression = data['enable_compression']
        if 'compression_level' in data:
            options.compression_level = data['compression_level']
        if 'remove_notes' in data:
            options.remove_notes = data['remove_notes']
        if 'remove_comments' in data:
            options.remove_comments = data['remove_comments']
        if 'optimize_for' in data:
            options.optimize_for = data['optimize_for']
        if 'export_notes' in data:
            options.export_notes = data['export_notes']
        if 'export_comments' in data:
            options.export_comments = data['export_comments']
        if 'export_hidden_slides' in data:
            options.export_hidden_slides = data['export_hidden_slides']
        if 'export_range' in data and data['export_range'] is not None:
            options.export_range = tuple(data['export_range'])
        return options
    
    def validate(self) -> Tuple[bool, str]:
        """
        验证配置的有效性
        
        Returns:
            Tuple[bool, str]: (是否有效, 错误信息)
        """
        if not 1 <= self.image_quality <= 100:
            return False, "图片质量必须在1-100之间"
        
        if not 0.1 <= self.resolution_scale <= 5.0:
            return False, "分辨率缩放必须在0.1-5.0之间"
        
        if self.custom_resolution:
            width, height = self.custom_resolution
            if width < 100 or height < 100:
                return False, "自定义分辨率不能小于100x100"
            if width > 10000 or height > 10000:
                return False, "自定义分辨率不能大于10000x10000"
        
        if self.watermark.enabled and not self.watermark.text and not self.watermark.image_path:
            return False, "启用水印时必须设置文字或图片"
        
        if self.watermark.enabled and not 0.0 <= self.watermark.opacity <= 1.0:
            return False, "水印透明度必须在0.0-1.0之间"
        
        if self.crop.enabled and self.crop.aspect_ratio:
            width, height = self.crop.aspect_ratio
            if width <= 0 or height <= 0:
                return False, "宽高比必须为正数"
        
        if self.export_range:
            start, end = self.export_range
            if start < 1 or end < start:
                return False, "导出范围无效"
        
        return True, ""
    
    def get_export_resolution(self, original_width: int, original_height: int) -> Tuple[int, int]:
        """
        计算导出分辨率
        
        Args:
            original_width: 原始宽度
            original_height: 原始高度
            
        Returns:
            Tuple[int, int]: (宽度, 高度)
        """
        if self.resolution == Resolution.ORIGINAL:
            width = original_width
            height = original_height
        elif self.resolution == Resolution.CUSTOM and self.custom_resolution:
            width, height = self.custom_resolution
        else:
            width, height = self.resolution.value
        
        # 应用缩放因子
        width = int(width * self.resolution_scale)
        height = int(height * self.resolution_scale)
        
        # 保持宽高比
        if self.preserve_aspect_ratio and self.resolution != Resolution.ORIGINAL:
            original_ratio = original_width / original_height
            new_ratio = width / height
            
            if abs(new_ratio - original_ratio) > 0.01:
                if new_ratio > original_ratio:
                    width = int(height * original_ratio)
                else:
                    height = int(width / original_ratio)
        
        return (width, height)

class PPTConverter(ABC):
    """
    PPT转换器抽象基类，定义统一接口
    """
    
    @abstractmethod
    def convert(self, input_path: str, output_path: str, options: ConversionOptions = None) -> bool:
        """
        执行PPT转换操作
        
        Args:
            input_path: 输入PPT文件路径
            output_path: 输出文件路径
            options: 转换选项配置
            
        Returns:
            bool: 转换是否成功
        """
        pass
    
    @abstractmethod
    def get_progress(self) -> float:
        """
        获取当前转换进度
        
        Returns:
            float: 进度百分比 (0.0-1.0)
        """
        pass
    
    @abstractmethod
    def get_conversion_info(self, input_path: str) -> Dict[str, Any]:
        """
        获取PPT文件的详细信息
        
        Args:
            input_path: PPT文件路径
            
        Returns:
            Dict[str, Any]: 包含文件信息的字典
        """
        pass
    
    @abstractmethod
    def batch_convert(self, input_files: List[str], output_dir: str, options: ConversionOptions = None) -> Dict[str, bool]:
        """
        批量转换多个PPT文件
        
        Args:
            input_files: 输入文件路径列表
            output_dir: 输出目录
            options: 转换选项
            
        Returns:
            Dict[str, bool]: 文件路径到转换结果的映射
        """
        pass
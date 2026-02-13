"""Redesigned PyQt6 main window with complete theme support."""

import sys
from pathlib import Path
from typing import Optional, List
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QProgressBar, QComboBox,
    QFileDialog, QMessageBox, QGroupBox, QTextEdit,
    QStatusBar, QApplication, QStackedWidget,
    QListWidget, QListWidgetItem, QFrame,
    QSplitter, QScrollArea, QToolButton,
    QMenu, QDialog, QLineEdit, QSpinBox,
    QSlider, QCheckBox, QGridLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QPalette, QAction, QIcon, QPixmap

from slide_sec.core import PPTProcessor, ProgressInfo
from slide_sec.utils import logger, get_settings_manager, get_template_manager
from slide_sec.utils.validators import FileValidator
from slide_sec.constants import APP_NAME, APP_VERSION, IMAGE_FORMAT_CONFIGS


class ConversionThread(QThread):
    """Background thread for file conversion."""
    
    progress_update = pyqtSignal(int, str)
    log_message = pyqtSignal(str)
    conversion_complete = pyqtSignal(int, Path)
    conversion_error = pyqtSignal(str)
    
    def __init__(self, files: List[Path], output_dir: Path, format_name: str, quality: int, processor: PPTProcessor):
        super().__init__()
        # 确保所有文件都是 Path 对象
        self.files = [Path(f) if isinstance(f, str) else f for f in files]
        self.output_dir = Path(output_dir) if isinstance(output_dir, str) else output_dir
        self.format_name = format_name
        self.quality = quality
        self.processor = processor
    
    def run(self):
        """Run conversion in background thread."""
        try:
            total_files = len(self.files)
            success_count = 0
            failed_count = 0
            
            self.log_message.emit(f"开始转换 {total_files} 个文件...")
            self.log_message.emit(f"输出目录: {self.output_dir}")
            self.log_message.emit(f"输出格式: {self.format_name.upper()}")
            self.log_message.emit(f"质量设置: {self.quality}%")
            
            for i, file_path in enumerate(self.files):
                # 保存文件名用于日志记录，避免后续变量被修改
                file_path = Path(file_path) if not isinstance(file_path, Path) else file_path
                file_name = file_path.name
                
                self.progress_update.emit(i, f"正在转换: {file_name}")
                self.log_message.emit(f"\n[{i+1}/{total_files}] 开始转换: {file_name}")
                
                # Generate output filename
                output_file = self.output_dir / f"{file_path.stem}_converted.pptx"
                
                try:
                    # Perform conversion
                    result = self.processor.convert_presentation(
                        input_file=file_path,
                        output_file=output_file,
                        format_name=self.format_name,
                        quality=self.quality,
                    )
                    
                    if result.get("success"):
                        success_count += 1
                        output_size = result.get('output_size_mb', 0)
                        self.log_message.emit(f"✓ 转换成功: {file_name} ({output_size:.2f} MB)")
                        logger.info(f"Converted {file_name} successfully")
                    else:
                        failed_count += 1
                        error_msg = result.get('error', '未知错误')
                        self.log_message.emit(f"✗ 转换失败: {file_name} - {error_msg}")
                        logger.error(f"Failed to convert {file_name}: {error_msg}")
                        
                except Exception as e:
                    failed_count += 1
                    self.log_message.emit(f"✗ 转换错误: {file_name} - {str(e)}")
                    logger.error(f"Error converting {file_name}: {e}")
                    
                # Update progress
                progress = int((i + 1) / total_files * 100)
                self.progress_update.emit(i + 1, f"文件 {i+1}/{total_files}: {file_name}")
                
            # Conversion complete - 根据成功/失败数量显示不同消息
            if failed_count == 0:
                self.log_message.emit(f"\n✓ 转换完成! 成功处理 {success_count}/{total_files} 个文件")
                self.conversion_complete.emit(success_count, self.output_dir)
            elif success_count == 0:
                self.log_message.emit(f"\n✗ 转换失败! 所有 {failed_count} 个文件都失败了")
                self.conversion_error.emit(f"所有 {failed_count} 个文件转换失败")
            else:
                self.log_message.emit(f"\n⚠ 部分完成: {success_count} 个成功, {failed_count} 个失败")
                self.conversion_complete.emit(success_count, self.output_dir)
            
        except Exception as e:
            self.log_message.emit(f"\n✗ 转换过程发生错误: {str(e)}")
            logger.error(f"Conversion error: {e}")
            self.conversion_error.emit(str(e))


class ThemeManager:
    """Manages application themes."""
    
    LIGHT_THEME = {
        "bg_primary": "#f5f7fa",
        "bg_secondary": "#ffffff",
        "bg_tertiary": "#f8f9fa",
        "text_primary": "#1a1a2e",
        "text_secondary": "#666666",
        "text_disabled": "#a0aec0",
        "border": "#e0e0e0",
        "accent": "#2196F3",
        "accent_hover": "#1976D2",
        "success": "#4CAF50",
        "warning": "#FF9800",
        "error": "#f44336",
        "card_bg": "#ffffff",
        "input_bg": "#ffffff",
        "shadow": "rgba(0, 0, 0, 0.1)",
    }
    
    DARK_THEME = {
        "bg_primary": "#1a1a2e",
        "bg_secondary": "#2d2d3a",
        "bg_tertiary": "#25262b",
        "text_primary": "#e0e0e0",
        "text_secondary": "#a0aec0",
        "text_disabled": "#666666",
        "border": "#3d3d3d",
        "accent": "#4CAF50",
        "accent_hover": "#45a049",
        "success": "#4CAF50",
        "warning": "#FF9800",
        "error": "#f44336",
        "card_bg": "#2d2d3a",
        "input_bg": "#1a1a2e",
        "shadow": "rgba(0, 0, 0, 0.3)",
    }
    
    def __init__(self):
        self.is_dark = False
        self.current_theme = self.LIGHT_THEME.copy()
        
    def toggle_theme(self):
        """Toggle between light and dark theme."""
        self.is_dark = not self.is_dark
        self.current_theme = self.DARK_THEME if self.is_dark else self.LIGHT_THEME
        return self.current_theme
        
    def get_theme(self):
        """Get current theme."""
        return self.current_theme


class StepIndicator(QWidget):
    """Step indicator for wizard-style interface."""
    
    def __init__(self, steps: List[str], theme: dict, parent=None):
        super().__init__(parent)
        self.steps = steps
        self.current_step = 0
        self.theme = theme
        self.setup_ui()
        
    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(20, 10, 20, 10)
        
        self.step_labels = []
        for i, step in enumerate(self.steps):
            # Step number circle
            step_widget = QWidget()
            step_layout = QVBoxLayout(step_widget)
            step_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            step_layout.setSpacing(5)
            
            number_label = QLabel(str(i + 1))
            number_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            number_label.setFixedSize(32, 32)
            number_label.setStyleSheet(f"""
                QLabel {{
                    background-color: {self.theme['bg_tertiary']};
                    color: {self.theme['text_secondary']};
                    border-radius: 16px;
                    font-weight: bold;
                    font-size: 14px;
                }}
            """)
            
            text_label = QLabel(step)
            text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            text_label.setStyleSheet(f"color: {self.theme['text_secondary']}; font-size: 12px;")
            
            step_layout.addWidget(number_label)
            step_layout.addWidget(text_label)
            
            layout.addWidget(step_widget)
            self.step_labels.append((number_label, text_label))
            
            # Add connector line
            if i < len(self.steps) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setStyleSheet(f"background-color: {self.theme['border']};")
                line.setFixedHeight(2)
                line.setMinimumWidth(60)
                layout.addWidget(line, alignment=Qt.AlignmentFlag.AlignVCenter)
                
        self.update_style()
        
    def set_current_step(self, step: int):
        self.current_step = step
        self.update_style()
        
    def update_style(self):
        for i, (num_label, text_label) in enumerate(self.step_labels):
            if i < self.current_step:
                # Completed
                num_label.setStyleSheet(f"""
                    QLabel {{
                        background-color: {self.theme['success']};
                        color: white;
                        border-radius: 16px;
                        font-weight: bold;
                        font-size: 14px;
                    }}
                """)
                text_label.setStyleSheet(f"color: {self.theme['success']}; font-size: 12px; font-weight: bold;")
            elif i == self.current_step:
                # Current
                num_label.setStyleSheet(f"""
                    QLabel {{
                        background-color: {self.theme['accent']};
                        color: white;
                        border-radius: 16px;
                        font-weight: bold;
                        font-size: 14px;
                    }}
                """)
                text_label.setStyleSheet(f"color: {self.theme['accent']}; font-size: 12px; font-weight: bold;")
            else:
                # Pending
                num_label.setStyleSheet(f"""
                    QLabel {{
                        background-color: {self.theme['bg_tertiary']};
                        color: {self.theme['text_secondary']};
                        border-radius: 16px;
                        font-weight: bold;
                        font-size: 14px;
                    }}
                """)
                text_label.setStyleSheet(f"color: {self.theme['text_disabled']}; font-size: 12px;")


class FileDropZone(QFrame):
    """Modern file drop zone with clear visual feedback."""
    
    file_dropped = pyqtSignal(str)
    
    def __init__(self, theme: dict, parent=None):
        super().__init__(parent)
        self.theme = theme
        self.setAcceptDrops(True)
        self.setMinimumHeight(200)
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Icon
        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 48px;")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)
        
        # Main text
        self.main_text = QLabel("拖拽 PowerPoint 文件到此处")
        self.main_text.setStyleSheet(f"""
            QLabel {{
                font-size: 18px;
                color: {self.theme['text_primary']};
                font-weight: 500;
            }}
        """)
        self.main_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.main_text)
        
        # Sub text
        self.sub_text = QLabel("或点击选择文件")
        self.sub_text.setStyleSheet(f"color: {self.theme['text_secondary']}; font-size: 14px;")
        self.sub_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.sub_text)
        
        # Browse button
        self.browse_btn = QPushButton("选择文件")
        self.browse_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['accent']};
                color: white;
                border: none;
                padding: 10px 30px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {self.theme['accent_hover']};
            }}
        """)
        self.browse_btn.clicked.connect(self.browse_file)
        layout.addWidget(self.browse_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.theme['bg_secondary']};
                border: 2px dashed {self.theme['border']};
                border-radius: 12px;
            }}
            FileDropZone:hover {{
                border-color: {self.theme['accent']};
                background-color: {self.theme['bg_tertiary']};
            }}
        """)
        
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 PowerPoint 文件", "",
            "PowerPoint Files (*.ppt *.pptx)"
        )
        if file_path:
            self.file_dropped.emit(file_path)
            
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(f"""
                FileDropZone {{
                    background-color: {self.theme['bg_tertiary']};
                    border: 2px dashed {self.theme['accent']};
                    border-radius: 12px;
                }}
            """)
            
    def dragLeaveEvent(self, event):
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.theme['bg_secondary']};
                border: 2px dashed {self.theme['border']};
                border-radius: 12px;
            }}
        """)
        
    def dropEvent(self, event):
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.theme['bg_secondary']};
                border: 2px dashed {self.theme['border']};
                border-radius: 12px;
            }}
        """)
        
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(('.ppt', '.pptx')):
                self.file_dropped.emit(file_path)


class QuickSettingsPanel(QWidget):
    """Simplified quick settings panel."""
    
    def __init__(self, theme: dict, parent=None):
        super().__init__(parent)
        self.theme = theme
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Template selector
        template_group = QGroupBox("快速预设")
        template_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.theme['card_bg']};
                border: 1px solid {self.theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        template_layout = QVBoxLayout(template_group)
        
        self.template_combo = QComboBox()
        self.template_combo.addItem("🎯 推荐：高质量演示", "high_quality")
        self.template_combo.addItem("🌐 网页分享", "web_share")
        self.template_combo.addItem("📧 邮件发送", "email")
        self.template_combo.addItem("🖨️ 打印输出", "print")
        self.template_combo.addItem("📱 社交媒体", "social")
        self.template_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: {self.theme['input_bg']};
                color: {self.theme['text_primary']};
                padding: 8px;
                border: 1px solid {self.theme['border']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {self.theme['card_bg']};
                color: {self.theme['text_primary']};
                border: 1px solid {self.theme['border']};
                selection-background-color: {self.theme['accent']};
                selection-color: white;
            }}
            QComboBox::drop-down {{
                border: none;
                padding-right: 8px;
            }}
        """)
        template_layout.addWidget(self.template_combo)
        
        layout.addWidget(template_group)
        
        # Format selection
        format_group = QGroupBox("输出格式")
        format_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.theme['card_bg']};
                border: 1px solid {self.theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        format_layout = QVBoxLayout(format_group)
        
        self.format_combo = QComboBox()
        format_icons = {
            "jpeg": "🖼️",
            "png": "🎨",
            "webp": "🌐",
            "tiff": "📷",
            "bmp": "📝",
            "gif": "🎬",
        }
        for fmt, config in IMAGE_FORMAT_CONFIGS.items():
            icon = format_icons.get(fmt, "📄")
            name = fmt.upper()
            self.format_combo.addItem(f"{icon} {name}", fmt)
        self.format_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: {self.theme['input_bg']};
                color: {self.theme['text_primary']};
                padding: 8px;
                border: 1px solid {self.theme['border']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {self.theme['card_bg']};
                color: {self.theme['text_primary']};
                border: 1px solid {self.theme['border']};
                selection-background-color: {self.theme['accent']};
                selection-color: white;
            }}
            QComboBox::drop-down {{
                border: none;
                padding-right: 8px;
            }}
        """)
        format_layout.addWidget(self.format_combo)
        
        layout.addWidget(format_group)
        
        # Quality slider
        quality_group = QGroupBox("图片质量")
        quality_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.theme['card_bg']};
                border: 1px solid {self.theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        quality_layout = QVBoxLayout(quality_group)
        
        quality_header = QHBoxLayout()
        quality_header.addWidget(QLabel("压缩率"))
        self.quality_value = QLabel("85%")
        self.quality_value.setStyleSheet(f"font-weight: bold; color: {self.theme['accent']};")
        quality_header.addStretch()
        quality_header.addWidget(self.quality_value)
        quality_layout.addLayout(quality_header)
        
        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(1, 100)
        self.quality_slider.setValue(85)
        self.quality_slider.valueChanged.connect(self.on_quality_changed)
        self.quality_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                background: {self.theme['bg_tertiary']};
                height: 4px;
                border-radius: 2px;
            }}
            QSlider::handle:horizontal {{
                background: {self.theme['accent']};
                width: 16px;
                border-radius: 8px;
                margin: -6px;
            }}
        """)
        quality_layout.addWidget(self.quality_slider)
        
        layout.addWidget(quality_group)
        
        # Output directory selection
        output_group = QGroupBox("输出目录")
        output_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.theme['card_bg']};
                border: 1px solid {self.theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        output_layout = QHBoxLayout(output_group)
        
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("选择输出目录...")
        self.output_path_edit.setStyleSheet(f"""
            QLineEdit {{
                background-color: {self.theme['input_bg']};
                color: {self.theme['text_primary']};
                padding: 8px;
                border: 1px solid {self.theme['border']};
                border-radius: 4px;
            }}
        """)
        output_layout.addWidget(self.output_path_edit)
        
        self.browse_output_btn = QPushButton("浏览")
        self.browse_output_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['accent']};
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['accent_hover']};
            }}
        """)
        output_layout.addWidget(self.browse_output_btn)
        
        layout.addWidget(output_group)
        
        # Advanced settings toggle
        self.advanced_btn = QPushButton("⚙️ 高级设置")
        self.advanced_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.theme['text_secondary']};
                border: 1px solid {self.theme['border']};
                padding: 8px;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['bg_tertiary']};
            }}
        """)
        layout.addWidget(self.advanced_btn)
        
        layout.addStretch()
        
    def on_quality_changed(self, value):
        self.quality_value.setText(f"{value}%")


class FileListItem(QWidget):
    """Custom file list item with status."""
    
    def __init__(self, file_path: Path, theme: dict, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.theme = theme
        self.setup_ui()
        
    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)
        
        # File icon
        icon_label = QLabel("📄")
        icon_label.setStyleSheet("font-size: 24px;")
        layout.addWidget(icon_label)
        
        # File info
        info_layout = QVBoxLayout()
        info_layout.setSpacing(2)
        
        name_label = QLabel(self.file_path.name)
        name_label.setStyleSheet(f"font-weight: 500; font-size: 13px; color: {self.theme['text_primary']};")
        info_layout.addWidget(name_label)
        
        size_label = QLabel(f"{self.file_path.stat().st_size / 1024:.1f} KB")
        size_label.setStyleSheet(f"color: {self.theme['text_disabled']}; font-size: 11px;")
        info_layout.addWidget(size_label)
        
        layout.addLayout(info_layout, stretch=1)
        
        # Status
        self.status_label = QLabel("已就绪")
        self.status_label.setStyleSheet(f"color: {self.theme['success']}; font-size: 12px;")
        layout.addWidget(self.status_label)
        
        # Remove button
        self.remove_btn = QToolButton()
        self.remove_btn.setText("✕")
        self.remove_btn.setStyleSheet(f"""
            QToolButton {{
                background-color: transparent;
                color: {self.theme['text_disabled']};
                border: none;
                font-size: 16px;
            }}
            QToolButton:hover {{
                color: {self.theme['error']};
            }}
        """)
        layout.addWidget(self.remove_btn)
        
        self.setStyleSheet(f"""
            FileListItem {{
                background-color: {self.theme['card_bg']};
                border: 1px solid {self.theme['border']};
                border-radius: 8px;
            }}
            FileListItem:hover {{
                background-color: {self.theme['bg_tertiary']};
            }}
        """)


class NewMainWindow(QMainWindow):
    """Redesigned main window with complete theme support."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} - {APP_VERSION}")
        # 固定窗口大小，不允许拉伸
        self.setFixedSize(900, 700)
        # 禁止最大化按钮
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowMaximizeButtonHint)
        
        # Initialize
        self.processor = PPTProcessor()
        self.settings_manager = get_settings_manager()
        self.template_manager = get_template_manager()
        self.files: List[Path] = []
        
        # Theme manager
        self.theme_manager = ThemeManager()
        self.current_theme = self.theme_manager.get_theme()
        
        self.setup_ui()
        self.apply_theme()
        
        logger.info("New main window initialized")
        
    def setup_ui(self):
        """Setup of user interface."""
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        
        # Main layout
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Header
        header = self.create_header()
        main_layout.addWidget(header)
        
        # Step indicator
        self.step_indicator = StepIndicator(["选择文件", "设置参数", "开始转换"], self.current_theme)
        main_layout.addWidget(self.step_indicator)
        
        # Content area
        self.content_stack = QStackedWidget()
        main_layout.addWidget(self.content_stack, stretch=1)
        
        # Step 1: File selection
        self.step1_widget = self.create_step1()
        self.content_stack.addWidget(self.step1_widget)
        
        # Step 2: Settings
        self.step2_widget = self.create_step2()
        self.content_stack.addWidget(self.step2_widget)
        
        # Step 3: Conversion
        self.step3_widget = self.create_step3()
        self.content_stack.addWidget(self.step3_widget)
        
        # Footer
        footer = self.create_footer()
        main_layout.addWidget(footer)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪 - 请选择要转换的文件")
        
    def create_header(self) -> QWidget:
        """Create application header."""
        header = QWidget()
        header.setFixedHeight(60)
        header.setStyleSheet(f"background-color: {self.current_theme['bg_secondary']}; border-bottom: 1px solid {self.current_theme['border']};")
        
        layout = QHBoxLayout(header)
        layout.setContentsMargins(20, 0, 20, 0)
        
        # Logo/Title
        title = QLabel(f"📊 {APP_NAME}")
        title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {self.current_theme['text_primary']};")
        layout.addWidget(title)
        
        layout.addStretch()
        
        # Theme toggle
        self.theme_btn = QPushButton("🌙" if not self.theme_manager.is_dark else "🌞")
        self.theme_btn.setFixedSize(36, 36)
        self.theme_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['bg_tertiary']};
                border: none;
                border-radius: 18px;
                font-size: 16px;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['border']};
            }}
        """)
        self.theme_btn.clicked.connect(self.toggle_theme)
        layout.addWidget(self.theme_btn)
        
        return header
        
    def create_step1(self) -> QWidget:
        """Create file selection step."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)
        
        # Title
        title = QLabel("选择要转换的文件")
        title.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {self.current_theme['text_primary']};")
        layout.addWidget(title)
        
        subtitle = QLabel("支持 .ppt 和 .pptx 格式的 PowerPoint 文件")
        subtitle.setStyleSheet(f"color: {self.current_theme['text_secondary']}; font-size: 14px;")
        layout.addWidget(subtitle)
        
        # Drop zone
        self.drop_zone = FileDropZone(self.current_theme)
        self.drop_zone.file_dropped.connect(self.add_file)
        layout.addWidget(self.drop_zone)
        
        # File list
        self.file_list = QListWidget()
        self.file_list.setStyleSheet(f"""
            QListWidget {{
                background-color: transparent;
                border: none;
            }}
            QListWidget::item {{
                background-color: transparent;
            }}
        """)
        layout.addWidget(self.file_list)
        
        # Navigation
        nav_layout = QHBoxLayout()
        nav_layout.addStretch()
        
        self.step1_next_btn = QPushButton("下一步 →")
        self.step1_next_btn.setEnabled(False)
        self.step1_next_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['accent']};
                color: white;
                border: none;
                padding: 12px 40px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['accent_hover']};
            }}
            QPushButton:disabled {{
                background-color: {self.current_theme['border']};
                color: {self.current_theme['text_disabled']};
            }}
        """)
        self.step1_next_btn.clicked.connect(self.go_to_step2)
        nav_layout.addWidget(self.step1_next_btn)
        
        layout.addLayout(nav_layout)
        layout.addStretch()
        
        return widget
        
    def create_step2(self) -> QWidget:
        """Create settings step."""
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(40, 20, 40, 20)
        main_layout.setSpacing(15)
        
        # Title
        title = QLabel("转换设置")
        title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {self.current_theme['text_primary']};")
        main_layout.addWidget(title)
        
        subtitle = QLabel("配置输出格式和选项")
        subtitle.setStyleSheet(f"color: {self.current_theme['text_secondary']}; font-size: 12px;")
        main_layout.addWidget(subtitle)
        
        # Settings panel - 全宽显示
        self.settings_panel = QuickSettingsPanel(self.current_theme)
        self.settings_panel.browse_output_btn.clicked.connect(self.browse_output_directory)
        self.settings_panel.advanced_btn.clicked.connect(self.show_advanced_settings)
        main_layout.addWidget(self.settings_panel)
        
        # Navigation
        nav_layout = QHBoxLayout()
        nav_layout.setContentsMargins(0, 10, 0, 0)
        
        back_btn = QPushButton("← 返回")
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {self.current_theme['text_secondary']};
                border: 1px solid {self.current_theme['border']};
                padding: 10px 24px;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['bg_tertiary']};
            }}
        """)
        back_btn.clicked.connect(self.go_to_step1)
        nav_layout.addWidget(back_btn)
        
        nav_layout.addStretch()
        
        convert_btn = QPushButton("开始转换 →")
        convert_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['success']};
                color: white;
                border: none;
                padding: 10px 32px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: #45a049;
            }}
        """)
        convert_btn.clicked.connect(self.go_to_step3)
        nav_layout.addWidget(convert_btn)
        
        main_layout.addLayout(nav_layout)
        
        return main_widget
        
    def create_step3(self) -> QWidget:
        """Create conversion step."""
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(40, 20, 40, 20)
        main_layout.setSpacing(15)
        
        # Title
        title = QLabel("转换进度")
        title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {self.current_theme['text_primary']};")
        main_layout.addWidget(title)
        
        subtitle = QLabel("正在处理您的文件，请稍候...")
        subtitle.setStyleSheet(f"color: {self.current_theme['text_secondary']}; font-size: 12px;")
        main_layout.addWidget(subtitle)
        
        # Progress card
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 12px;
                padding: 30px;
            }}
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(15)
        
        # Status icon
        self.status_icon = QLabel("⏳")
        self.status_icon.setStyleSheet("font-size: 48px;")
        self.status_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.status_icon)
        
        # Status text
        self.status_text = QLabel("正在转换...")
        self.status_text.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {self.current_theme['text_primary']};")
        self.status_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.status_text)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {self.current_theme['bg_tertiary']};
                height: 8px;
                text-align: center;
                color: {self.current_theme['text_primary']};
            }}
            QProgressBar::chunk {{
                background-color: {self.current_theme['accent']};
                border-radius: 4px;
            }}
        """)
        card_layout.addWidget(self.progress_bar)
        
        # Details
        self.progress_details = QLabel("准备中...")
        self.progress_details.setStyleSheet(f"color: {self.current_theme['text_secondary']};")
        self.progress_details.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.progress_details)
        
        main_layout.addWidget(card)
        
        # Log area
        log_group = QGroupBox("转换日志")
        log_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 12px;
                font-weight: 600;
                color: {self.current_theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {self.current_theme['bg_tertiary']};
                color: {self.current_theme['text_primary']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
            }}
        """)
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
        
        return main_widget
        
    def create_footer(self) -> QWidget:
        """Create footer with actions."""
        footer = QWidget()
        footer.setFixedHeight(50)
        footer.setStyleSheet(f"background-color: {self.current_theme['bg_tertiary']}; border-top: 1px solid {self.current_theme['border']};")
        
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(20, 0, 20, 0)
        
        # Help link
        help_label = QLabel("❓ 需要帮助?")
        help_label.setStyleSheet(f"color: {self.current_theme['text_secondary']};")
        layout.addWidget(help_label)
        
        layout.addStretch()
        
        # Version
        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setStyleSheet(f"color: {self.current_theme['text_disabled']}; font-size: 12px;")
        layout.addWidget(version_label)
        
        return footer
        
    def apply_theme(self):
        """Apply current theme to entire application."""
        # Update theme for all components
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {self.current_theme['bg_primary']};
            }}
            QWidget {{
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }}
            QStatusBar {{
                background-color: {self.current_theme['bg_secondary']};
                color: {self.current_theme['text_primary']};
            }}
        """)
        
        # Update step indicator
        self.step_indicator.theme = self.current_theme
        self.step_indicator.update_style()
        
        # Update file list
        self.file_list.clear()
        for path in self.files:
            item_widget = FileListItem(path, self.current_theme)
            item_widget.remove_btn.clicked.connect(lambda p=path: self.remove_file(p))
            
            item = QListWidgetItem()
            item.setSizeHint(item_widget.sizeHint())
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, item_widget)
        
        # Update drop zone
        self.drop_zone.theme = self.current_theme
        self.drop_zone.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.current_theme['bg_secondary']};
                border: 2px dashed {self.current_theme['border']};
                border-radius: 12px;
            }}
            FileDropZone:hover {{
                border-color: {self.current_theme['accent']};
                background-color: {self.current_theme['bg_tertiary']};
            }}
        """)
        
        # Update settings panel
        self.settings_panel.theme = self.current_theme
        self.settings_panel.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.current_theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        
        # Update buttons
        self.step1_next_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['accent']};
                color: white;
                border: none;
                padding: 12px 40px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['accent_hover']};
            }}
            QPushButton:disabled {{
                background-color: {self.current_theme['border']};
                color: {self.current_theme['text_disabled']};
            }}
        """)
        
        # Update header style
        header = self.centralWidget().layout().itemAt(0).widget()
        header.setStyleSheet(f"background-color: {self.current_theme['bg_secondary']}; border-bottom: 1px solid {self.current_theme['border']};")
        
        # Update theme button style
        self.theme_btn.setText("🌙" if not self.theme_manager.is_dark else "🌞")
        self.theme_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['bg_tertiary']};
                border: none;
                border-radius: 18px;
                font-size: 16px;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['border']};
            }}
        """)
        
        # Update title label
        title = header.findChild(QLabel)
        if title:
            title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {self.current_theme['text_primary']};")
        
        # Update step 1 navigation button
        if hasattr(self, 'step1_next_btn'):
            self.step1_next_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {self.current_theme['accent']};
                    color: white;
                    border: none;
                    padding: 12px 40px;
                    border-radius: 6px;
                    font-size: 14px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {self.current_theme['accent_hover']};
                }}
                QPushButton:disabled {{
                    background-color: {self.current_theme['border']};
                    color: {self.current_theme['text_disabled']};
                }}
            """)
        
        # Update step 2 navigation buttons
        if hasattr(self, 'step2_widget'):
            for btn in self.step2_widget.findChildren(QPushButton):
                if "返回" in btn.text():
                    btn.setStyleSheet(f"""
                        QPushButton {{
                            background-color: {self.current_theme['bg_tertiary']};
                            color: {self.current_theme['text_primary']};
                            border: 1px solid {self.current_theme['border']};
                            padding: 12px 30px;
                            border-radius: 6px;
                        }}
                        QPushButton:hover {{
                            background-color: {self.current_theme['border']};
                        }}
                    """)
                elif "开始转换" in btn.text():
                    btn.setStyleSheet(f"""
                        QPushButton {{
                            background-color: {self.current_theme['success']};
                            color: white;
                            border: none;
                            padding: 12px 40px;
                            border-radius: 6px;
                            font-size: 14px;
                            font-weight: 500;
                        }}
                        QPushButton:hover {{
                            background-color: #45a049;
                        }}
                    """)
        
        # Update step 3 navigation button
        if hasattr(self, 'step3_widget'):
            for btn in self.step3_widget.findChildren(QPushButton):
                if "完成" in btn.text():
                    btn.setStyleSheet(f"""
                        QPushButton {{
                            background-color: {self.current_theme['accent']};
                            color: white;
                            border: none;
                            padding: 12px 40px;
                            border-radius: 6px;
                            font-size: 14px;
                            font-weight: 500;
                        }}
                        QPushButton:hover {{
                            background-color: {self.current_theme['accent_hover']};
                        }}
                    """)
        
        # Update footer style
        footer = self.centralWidget().layout().itemAt(3).widget()
        footer.setStyleSheet(f"background-color: {self.current_theme['bg_tertiary']}; border-top: 1px solid {self.current_theme['border']};")
        
        # Update footer labels
        for label in footer.findChildren(QLabel):
            if "帮助" in label.text():
                label.setStyleSheet(f"color: {self.current_theme['text_secondary']};")
            elif "v" in label.text():
                label.setStyleSheet(f"color: {self.current_theme['text_disabled']}; font-size: 12px;")
        
    def toggle_theme(self):
        """Toggle between light and dark theme."""
        self.current_theme = self.theme_manager.toggle_theme()
        self.apply_theme()
        
        theme_name = "深色" if self.theme_manager.is_dark else "浅色"
        self.status_bar.showMessage(f"已切换到{theme_name}主题", 3000)
        
    def add_file(self, file_path: str):
        """Add file to list."""
        path = Path(file_path)
        
        # Validate
        result = FileValidator.validate_ppt_file(path)
        if not result.is_valid:
            QMessageBox.warning(self, "文件错误", "\n".join(result.errors))
            return
            
        # Check duplicates
        if path in self.files:
            QMessageBox.information(self, "提示", "该文件已在列表中")
            return
            
        self.files.append(path)
        
        # Add to list widget
        item_widget = FileListItem(path, self.current_theme)
        item_widget.remove_btn.clicked.connect(lambda p=path: self.remove_file(p))
        
        item = QListWidgetItem()
        item.setSizeHint(item_widget.sizeHint())
        self.file_list.addItem(item)
        self.file_list.setItemWidget(item, item_widget)
        
        # Enable next button
        self.step1_next_btn.setEnabled(True)
        self.status_bar.showMessage(f"已添加 {len(self.files)} 个文件")
        
    def remove_file(self, file_path: Path):
        """Remove file from list."""
        if file_path in self.files:
            self.files.remove(file_path)
            
        # Refresh list
        self.file_list.clear()
        for path in self.files:
            item_widget = FileListItem(path, self.current_theme)
            item_widget.remove_btn.clicked.connect(lambda p=path: self.remove_file(p))
            
            item = QListWidgetItem()
            item.setSizeHint(item_widget.sizeHint())
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, item_widget)
            
        self.step1_next_btn.setEnabled(len(self.files) > 0)
        
    def go_to_step1(self):
        """Go to step 1."""
        self.content_stack.setCurrentIndex(0)
        self.step_indicator.set_current_step(0)
        
    def go_to_step2(self):
        """Go to step 2."""
        self.content_stack.setCurrentIndex(1)
        self.step_indicator.set_current_step(1)
        self.status_bar.showMessage("请配置转换参数")
        
    def go_to_step3(self):
        """Go to step 3 and start conversion."""
        if not self.files:
            QMessageBox.warning(self, "警告", "请先选择要转换的文件")
            return
            
        self.content_stack.setCurrentIndex(2)
        self.step_indicator.set_current_step(2)
        self.status_bar.showMessage("正在转换...")
        
        # Start actual conversion in background
        self.start_conversion()
        
    def start_conversion(self):
        """Start of conversion process."""
        # Get settings
        format_name = self.settings_panel.format_combo.currentData()
        quality = self.settings_panel.quality_slider.value()
        
        # Get output directory from user selection or use default
        output_path = self.settings_panel.output_path_edit.text().strip()
        if output_path:
            output_dir = Path(output_path)
        else:
            output_dir = Path.home() / "Documents" / "SlideSec_Output"
        
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create and start conversion thread
        self.conversion_thread = ConversionThread(
            files=self.files,
            output_dir=output_dir,
            format_name=format_name,
            quality=quality,
            processor=self.processor
        )
        
        # Connect signals
        self.conversion_thread.progress_update.connect(self.on_progress_update)
        self.conversion_thread.log_message.connect(self.on_log_message)
        self.conversion_thread.conversion_complete.connect(self.on_conversion_complete)
        self.conversion_thread.conversion_error.connect(self.on_conversion_error)
        
        # Start thread
        self.conversion_thread.start()
        
    def on_progress_update(self, current_index: int, message: str):
        """Handle progress update from conversion thread."""
        total_files = len(self.files)
        # current_index is 0-based, so we need to add 1 to get the count
        progress = int((current_index + 1) / total_files * 100)
        self.progress_bar.setValue(progress)
        self.progress_details.setText(message)
        
    def on_log_message(self, message: str):
        """Handle log message from conversion thread."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        
    def on_conversion_complete(self, total_files: int, output_dir: Path):
        """Handle conversion completion."""
        self.status_text.setText("转换完成!")
        self.status_icon.setText("✅")
        self.progress_details.setText("所有文件已处理完成")
        self.status_bar.showMessage(f"转换完成! 输出目录: {output_dir}")
        
        # Show completion message
        QMessageBox.information(
            self, 
            "转换完成", 
            f"已成功转换 {total_files} 个文件!\n\n输出目录:\n{output_dir}"
        )
        
    def on_conversion_error(self, error_msg: str):
        """Handle conversion error."""
        self.status_text.setText("转换失败")
        self.status_icon.setText("❌")
        QMessageBox.critical(self, "错误", f"转换过程中发生错误:\n{error_msg}")
        
    def browse_output_directory(self):
        """Browse and select output directory."""
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择输出目录", str(Path.home() / "Documents")
        )
        if dir_path:
            self.settings_panel.output_path_edit.setText(dir_path)
            self.status_bar.showMessage(f"输出目录: {dir_path}")
    
    def show_advanced_settings(self):
        """Show advanced settings dialog."""
        dialog = QDialog(self)
        dialog.setWindowTitle("高级设置")
        dialog.setMinimumWidth(500)
        dialog.setStyleSheet(f"""
            QDialog {{
                background-color: {self.current_theme['bg_primary']};
            }}
        """)
        
        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        
        # DPI setting
        dpi_group = QGroupBox("DPI 设置")
        dpi_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.current_theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        dpi_layout = QVBoxLayout(dpi_group)
        
        dpi_slider_layout = QHBoxLayout()
        dpi_slider_layout.addWidget(QLabel("DPI:"))
        self.dpi_value_label = QLabel("150")
        self.dpi_value_label.setStyleSheet(f"font-weight: bold; color: {self.current_theme['accent']};")
        dpi_slider_layout.addStretch()
        dpi_slider_layout.addWidget(self.dpi_value_label)
        dpi_layout.addLayout(dpi_slider_layout)
        
        self.dpi_slider = QSlider(Qt.Orientation.Horizontal)
        self.dpi_slider.setRange(72, 300)
        self.dpi_slider.setValue(150)
        self.dpi_slider.valueChanged.connect(lambda v: self.dpi_value_label.setText(str(v)))
        dpi_layout.addWidget(self.dpi_slider)
        
        layout.addWidget(dpi_group)
        
        # Resolution setting
        res_group = QGroupBox("分辨率设置")
        res_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.current_theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        res_layout = QVBoxLayout(res_group)
        
        self.res_combo = QComboBox()
        self.res_combo.addItem("原始分辨率", None)
        self.res_combo.addItem("1920 x 1080 (1080p)", (1920, 1080))
        self.res_combo.addItem("1280 x 720 (720p)", (1280, 720))
        self.res_combo.addItem("3840 x 2160 (4K)", (3840, 2160))
        self.res_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: {self.current_theme['input_bg']};
                color: {self.current_theme['text_primary']};
                padding: 8px;
                border: 1px solid {self.current_theme['border']};
                border-radius: 4px;
            }}
        """)
        res_layout.addWidget(self.res_combo)
        
        layout.addWidget(res_group)
        
        # Parallel processing
        parallel_group = QGroupBox("并行处理")
        parallel_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.current_theme['card_bg']};
                border: 1px solid {self.current_theme['border']};
                border-radius: 8px;
                margin-top: 12px;
                padding: 16px;
                font-weight: 600;
                color: {self.current_theme['text_primary']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }}
        """)
        parallel_layout = QHBoxLayout(parallel_group)
        
        self.parallel_checkbox = QCheckBox("启用并行处理")
        self.parallel_checkbox.setChecked(True)
        self.parallel_checkbox.setStyleSheet(f"""
            QCheckBox {{
                color: {self.current_theme['text_primary']};
            }}
        """)
        parallel_layout.addWidget(self.parallel_checkbox)
        
        layout.addWidget(parallel_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['bg_tertiary']};
                color: {self.current_theme['text_primary']};
                border: 1px solid {self.current_theme['border']};
                padding: 8px 24px;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['border']};
            }}
        """)
        cancel_btn.clicked.connect(dialog.reject)
        button_layout.addWidget(cancel_btn)
        
        ok_btn = QPushButton("确定")
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_theme['accent']};
                color: white;
                border: none;
                padding: 8px 24px;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.current_theme['accent_hover']};
            }}
        """)
        ok_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.status_bar.showMessage("高级设置已更新")


def main():
    """Main entry point."""
    app = QApplication(sys.argv)
    
    # Set application font
    font = QFont("Segoe UI", 10)
    font.setStyleHint(QFont.StyleHint.SansSerif)
    app.setFont(font)
    
    window = NewMainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

"""Modern PyQt6 main window for SlideSec application."""

import sys
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QPushButton, QProgressBar, QComboBox, 
    QSpinBox, QSlider, QFileDialog, QMessageBox,
    QGroupBox, QTextEdit, QSplitter, QFrame,
    QStatusBar, QToolBar, QMenuBar, QMenu,
    QApplication, QGraphicsDropShadowEffect, QSizePolicy,
    QGridLayout, QTabWidget, QCheckBox, QRadioButton,
    QButtonGroup, QLineEdit, QScrollArea, QDialog,
    QListWidget, QListWidgetItem, QStackedWidget,
    QSystemTrayIcon, QStyle, QStyleFactory
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QSize,
    QPropertyAnimation, QEasingCurve, QPoint, QMimeData
)
from PyQt6.QtGui import (
    QFont, QIcon, QPalette, QColor, QDragEnterEvent,
    QDropEvent, QFontDatabase, QAction, QKeySequence,
    QPixmap, QPainter, QLinearGradient, QBrush
)

from slide_sec.core import PPTProcessor, ProgressInfo
from slide_sec.config import get_config
from slide_sec.utils import logger, get_error_manager, ErrorCategory, get_settings_manager, get_template_manager
from slide_sec.utils.validators import validate_conversion_request, FileValidator
from slide_sec.constants import (
    APP_NAME, APP_VERSION, IMAGE_FORMAT_CONFIGS,
    SUPPORTED_PPT_EXTENSIONS, DEFAULT_QUALITY, DEFAULT_DPI
)


class ModernStyleSheet:
    """Modern QSS styles for the application."""
    
    MAIN_STYLE = """
        QMainWindow {
            background-color: #f5f7fa;
        }
        
        QWidget {
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            font-size: 13px;
        }
        
        /* Cards */
        QGroupBox {
            background-color: white;
            border: 1px solid #e1e8ed;
            border-radius: 12px;
            margin-top: 12px;
            padding-top: 16px;
            padding: 16px;
            font-weight: 600;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 16px;
            padding: 0 8px;
            color: #1a1a2e;
            font-size: 14px;
        }
        
        /* Buttons */
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #667eea, stop:1 #764ba2);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            font-size: 13px;
            min-height: 40px;
        }
        
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #5a67d8, stop:1 #6b46a1);
        }
        
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #4c51bf, stop:1 #553c9a);
        }
        
        QPushButton:disabled {
            background-color: #cbd5e0;
            color: #a0aec0;
        }
        
        QPushButton#secondary {
            background-color: #edf2f7;
            color: #4a5568;
            border: 1px solid #e2e8f0;
        }
        
        QPushButton#secondary:hover {
            background-color: #e2e8f0;
        }
        
        QPushButton#danger {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #fc8181, stop:1 #f56565);
        }
        
        /* Input fields */
        QLineEdit {
            background-color: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px 14px;
            min-height: 20px;
        }
        
        QLineEdit:focus {
            border-color: #667eea;
        }
        
        QLineEdit:hover {
            border-color: #cbd5e0;
        }
        
        /* ComboBox */
        QComboBox {
            background-color: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px 14px;
            min-height: 20px;
            min-width: 120px;
        }
        
        QComboBox:hover {
            border-color: #cbd5e0;
        }
        
        QComboBox:focus {
            border-color: #667eea;
        }
        
        QComboBox::drop-down {
            border: none;
            width: 30px;
        }
        
        QComboBox::down-arrow {
            image: none;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid #4a5568;
            width: 0;
            height: 0;
        }
        
        QComboBox QAbstractItemView {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            selection-background-color: #667eea;
            selection-color: white;
            padding: 4px;
        }
        
        /* Sliders */
        QSlider::groove:horizontal {
            height: 8px;
            background: #e2e8f0;
            border-radius: 4px;
        }
        
        QSlider::sub-page:horizontal {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #667eea, stop:1 #764ba2);
            border-radius: 4px;
        }
        
        QSlider::handle:horizontal {
            background: white;
            border: 2px solid #667eea;
            width: 20px;
            height: 20px;
            margin: -6px 0;
            border-radius: 10px;
        }
        
        QSlider::handle:horizontal:hover {
            background: #667eea;
        }
        
        /* Progress Bar */
        QProgressBar {
            border: none;
            border-radius: 6px;
            background-color: #e2e8f0;
            text-align: center;
            font-weight: 600;
            color: #4a5568;
            min-height: 24px;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #667eea, stop:1 #764ba2);
            border-radius: 6px;
        }
        
        /* Text Edit */
        QTextEdit {
            background-color: #1a1a2e;
            color: #a0aec0;
            border: none;
            border-radius: 8px;
            padding: 12px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            line-height: 1.5;
        }
        
        /* Scrollbar */
        QScrollBar:vertical {
            background-color: #f7fafc;
            width: 12px;
            border-radius: 6px;
        }
        
        QScrollBar::handle:vertical {
            background-color: #cbd5e0;
            border-radius: 6px;
            min-height: 30px;
        }
        
        QScrollBar::handle:vertical:hover {
            background-color: #a0aec0;
        }
        
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        /* Labels */
        QLabel {
            color: #4a5568;
        }
        
        QLabel#title {
            font-size: 24px;
            font-weight: 700;
            color: #1a1a2e;
        }
        
        QLabel#subtitle {
            font-size: 14px;
            color: #718096;
        }
        
        QLabel#value {
            font-weight: 600;
            color: #667eea;
        }
        
        /* Status Bar */
        QStatusBar {
            background-color: white;
            border-top: 1px solid #e2e8f0;
            color: #4a5568;
        }
        
        /* Menu */
        QMenuBar {
            background-color: white;
            border-bottom: 1px solid #e2e8f0;
        }
        
        QMenuBar::item {
            padding: 8px 16px;
            background: transparent;
        }
        
        QMenuBar::item:selected {
            background-color: #edf2f7;
            border-radius: 4px;
        }
        
        QMenu {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 8px;
        }
        
        QMenu::item {
            padding: 8px 24px;
            border-radius: 4px;
        }
        
        QMenu::item:selected {
            background-color: #667eea;
            color: white;
        }
        
        /* CheckBox */
        QCheckBox {
            spacing: 8px;
        }
        
        QCheckBox::indicator {
            width: 20px;
            height: 20px;
            border: 2px solid #e2e8f0;
            border-radius: 4px;
            background-color: white;
        }
        
        QCheckBox::indicator:checked {
            background-color: #667eea;
            border-color: #667eea;
        }
        
        /* Radio Button */
        QRadioButton {
            spacing: 8px;
        }
        
        QRadioButton::indicator {
            width: 20px;
            height: 20px;
            border: 2px solid #e2e8f0;
            border-radius: 10px;
            background-color: white;
        }
        
        QRadioButton::indicator:checked {
            background-color: #667eea;
            border-color: #667eea;
        }
        
        /* Tab Widget */
        QTabWidget::pane {
            border: none;
            background-color: transparent;
        }
        
        QTabBar::tab {
            background-color: transparent;
            border: none;
            padding: 12px 24px;
            margin-right: 4px;
            color: #718096;
            font-weight: 600;
        }
        
        QTabBar::tab:selected {
            color: #667eea;
            border-bottom: 2px solid #667eea;
        }
        
        QTabBar::tab:hover {
            color: #4a5568;
        }
        
        /* List Widget */
        QListWidget {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 8px;
            outline: none;
        }
        
        QListWidget::item {
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 4px;
        }
        
        QListWidget::item:selected {
            background-color: #667eea;
            color: white;
        }
        
        QListWidget::item:hover {
            background-color: #edf2f7;
        }
        
        QListWidget::item:selected:hover {
            background-color: #5a67d8;
        }
        
        /* Spin Box */
        QSpinBox {
            background-color: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 10px;
            min-height: 20px;
        }
        
        QSpinBox:focus {
            border-color: #667eea;
        }
        
        QSpinBox::up-button, QSpinBox::down-button {
            width: 20px;
            background-color: #edf2f7;
            border: none;
        }
        
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {
            background-color: #e2e8f0;
        }
    """
    
    DARK_STYLE = """
        QMainWindow {
            background-color: #1a1a2e;
        }
        
        QWidget {
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            font-size: 13px;
            color: #e2e8f0;
        }
        
        QGroupBox {
            background-color: #16213e;
            border: 1px solid #0f3460;
            border-radius: 12px;
            margin-top: 12px;
            padding-top: 16px;
            padding: 16px;
            font-weight: 600;
            color: #e2e8f0;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 16px;
            padding: 0 8px;
            color: #e2e8f0;
        }
        
        QLineEdit {
            background-color: #16213e;
            border: 2px solid #0f3460;
            color: #e2e8f0;
            border-radius: 8px;
            padding: 10px 14px;
        }
        
        QLineEdit:focus {
            border-color: #667eea;
        }
        
        QTextEdit {
            background-color: #0f0f1e;
            color: #a0aec0;
            border: none;
            border-radius: 8px;
            padding: 12px;
        }
        
        QComboBox {
            background-color: #16213e;
            border: 2px solid #0f3460;
            color: #e2e8f0;
            border-radius: 8px;
            padding: 10px 14px;
        }
        
        QComboBox QAbstractItemView {
            background-color: #16213e;
            border: 1px solid #0f3460;
            color: #e2e8f0;
        }
        
        QProgressBar {
            border: none;
            border-radius: 6px;
            background-color: #0f3460;
            text-align: center;
            color: #e2e8f0;
        }
        
        QLabel {
            color: #e2e8f0;
        }
        
        QLabel#subtitle {
            color: #a0aec0;
        }
        
        QStatusBar {
            background-color: #16213e;
            border-top: 1px solid #0f3460;
            color: #e2e8f0;
        }
        
        QMenuBar {
            background-color: #16213e;
            border-bottom: 1px solid #0f3460;
        }
        
        QMenu {
            background-color: #16213e;
            border: 1px solid #0f3460;
        }
        
        QListWidget {
            background-color: #16213e;
            border: 1px solid #0f3460;
        }
        
        QListWidget::item:hover {
            background-color: #0f3460;
        }
        
        QSlider::groove:horizontal {
            background: #0f3460;
        }
        
        QSpinBox {
            background-color: #16213e;
            border: 2px solid #0f3460;
            color: #e2e8f0;
        }
        
        QCheckBox::indicator {
            border: 2px solid #0f3460;
            background-color: #16213e;
        }
        
        QRadioButton::indicator {
            border: 2px solid #0f3460;
            background-color: #16213e;
        }
    """


class ConversionWorker(QThread):
    """Worker thread for conversion process."""
    
    progress_signal = pyqtSignal(ProgressInfo)
    completed_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    
    def __init__(self, processor: PPTProcessor, **kwargs):
        super().__init__()
        self.processor = processor
        self.kwargs = kwargs
        self.is_running = True
        
    def run(self):
        """Run conversion in background thread."""
        try:
            def progress_callback(progress_info: ProgressInfo):
                if self.is_running:
                    self.progress_signal.emit(progress_info)
            
            self.kwargs['progress_callback'] = progress_callback
            result = self.processor.convert_presentation(**self.kwargs)
            
            if self.is_running:
                self.completed_signal.emit(result)
                
        except Exception as e:
            if self.is_running:
                self.error_signal.emit(str(e))
                
    def stop(self):
        """Stop the conversion."""
        self.is_running = False
        self.wait(1000)


class DropArea(QFrame):
    """Custom drop area for file drag and drop."""
    
    file_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(150)
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        self.setStyleSheet("""
            DropArea {
                background-color: #f7fafc;
                border: 2px dashed #cbd5e0;
                border-radius: 12px;
            }
            DropArea[dragOver="true"] {
                background-color: #edf2f7;
                border-color: #667eea;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 48px;")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)
        
        self.text_label = QLabel("拖拽 PowerPoint 文件到这里")
        self.text_label.setStyleSheet("color: #718096; font-size: 14px;")
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.text_label)
        
        self.sub_label = QLabel("或点击选择文件")
        self.sub_label.setStyleSheet("color: #a0aec0; font-size: 12px;")
        self.sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.sub_label)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setProperty("dragOver", "true")
            self.style().unpolish(self)
            self.style().polish(self)
            
    def dragLeaveEvent(self, event):
        """Handle drag leave event."""
        self.setProperty("dragOver", "false")
        self.style().unpolish(self)
        self.style().polish(self)
        
    def dropEvent(self, event: QDropEvent):
        """Handle drop event."""
        self.setProperty("dragOver", "false")
        self.style().unpolish(self)
        self.style().polish(self)
        
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.ppt', '.pptx')):
                self.file_dropped.emit(file_path)
            else:
                QMessageBox.warning(self, "不支持的文件", "请拖入 .ppt 或 .pptx 文件")
                
    def mousePressEvent(self, event):
        """Handle mouse click."""
        if event.button() == Qt.MouseButton.LeftButton:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择 PowerPoint 文件", "",
                "PowerPoint Files (*.ppt *.pptx)"
            )
            if file_path:
                self.file_dropped.emit(file_path)


class ModernMainWindow(QMainWindow):
    """Modern main window with PyQt6."""
    
    def __init__(self):
        super().__init__()
        
        # Initialize managers
        self.settings_manager = get_settings_manager()
        self.user_settings = self.settings_manager.load_settings()
        self.template_manager = get_template_manager()

        # Set window properties
        self.setWindowTitle(f"{APP_NAME} - {APP_VERSION}")
        self.setMinimumSize(1000, 750)

        # Restore window size from settings
        window_width = self.user_settings.window_width
        window_height = self.user_settings.window_height
        self.resize(window_width, window_height)

        # Initialize components
        self.config = get_config()
        self.processor = PPTProcessor()
        self.current_file: Optional[Path] = None
        self.worker: Optional[ConversionWorker] = None
        self.is_converting = False
        self.is_dark_theme = self.user_settings.theme == "dark"
        
        # Apply modern style based on theme
        self.apply_style()
        
        # Setup UI
        self.setup_ui()
        
        # Load default settings into UI
        self.load_user_settings()
        
        # Center window
        self.center_window()
        
        logger.info("Modern main window initialized")
        
    def apply_style(self):
        """Apply modern stylesheet based on current theme."""
        if self.is_dark_theme:
            self.setStyleSheet(ModernStyleSheet.DARK_STYLE)
        else:
            self.setStyleSheet(ModernStyleSheet.MAIN_STYLE)
            
    def toggle_theme(self):
        """Toggle between light and dark themes."""
        self.is_dark_theme = not self.is_dark_theme
        self.user_settings.theme = "dark" if self.is_dark_theme else "light"
        self.settings_manager.save_settings(self.user_settings)
        self.apply_style()
        
        theme_name = "深色" if self.is_dark_theme else "浅色"
        self.log_message(f"已切换到{theme_name}主题", "info")
        self.status_bar.showMessage(f"已切换到{theme_name}主题", 3000)
        
    def load_user_settings(self):
        """Load user settings into UI controls."""
        # Set default format
        format_index = self.format_combo.findData(self.user_settings.default_format)
        if format_index >= 0:
            self.format_combo.setCurrentIndex(format_index)
            
        # Set default quality
        self.quality_slider.setValue(self.user_settings.default_quality)
        self.quality_label.setText(f"{self.user_settings.default_quality}%")
        
        # Set default DPI
        dpi_index = self.dpi_combo.findData(self.user_settings.default_dpi)
        if dpi_index >= 0:
            self.dpi_combo.setCurrentIndex(dpi_index)
            
        logger.info("User settings loaded into UI")
        
    def save_window_settings(self):
        """Save window size and position."""
        size = self.size()
        self.user_settings.window_width = size.width()
        self.user_settings.window_height = size.height()
        self.settings_manager.save_settings(self.user_settings)
        
    def center_window(self):
        """Center window on screen."""
        screen = QApplication.primaryScreen().geometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )

    def load_template_list(self):
        """Load template list into combo box."""
        self.template_combo.clear()
        self.template_combo.addItem("选择预设模板...", None)

        templates = self.template_manager.get_all_templates()
        for template in templates:
            self.template_combo.addItem(f"{template.name} - {template.description}", template.name)

    def on_template_selected(self, index):
        """Handle template selection."""
        if index <= 0:
            return

        template_name = self.template_combo.currentData()
        template = self.template_manager.get_template(template_name)

        if template:
            # Apply template settings
            format_index = self.format_combo.findData(template.format_name)
            if format_index >= 0:
                self.format_combo.setCurrentIndex(format_index)

            self.quality_slider.setValue(template.quality)
            self.quality_label.setText(f"{template.quality}%")

            dpi_index = self.dpi_combo.findData(template.dpi)
            if dpi_index >= 0:
                self.dpi_combo.setCurrentIndex(dpi_index)

            if template.resolution:
                res_index = self.resolution_combo.findData(template.resolution)
                if res_index >= 0:
                    self.resolution_combo.setCurrentIndex(res_index)

            self.log_message(f"已应用模板: {template.name}", "info")
            self.status_bar.showMessage(f"已应用模板: {template.name}", 3000)

    def save_current_as_template(self):
        """Save current settings as a new template."""
        from PyQt6.QtWidgets import QInputDialog, QLineEdit

        # Get template name
        name, ok = QInputDialog.getText(
            self, "保存模板", "模板名称:",
            QLineEdit.EchoMode.Normal, ""
        )
        if not ok or not name:
            return

        # Get description
        description, ok = QInputDialog.getText(
            self, "保存模板", "模板描述:",
            QLineEdit.EchoMode.Normal, ""
        )
        if not ok:
            description = ""

        # Get current settings
        format_name = self.format_combo.currentData()
        quality = self.quality_slider.value()
        dpi = self.dpi_combo.currentData()
        resolution = self.resolution_combo.currentData()

        # Create template
        template = self.template_manager.create_template_from_current(
            name=name,
            description=description,
            format_name=format_name,
            quality=quality,
            dpi=dpi,
            resolution=resolution
        )

        # Add template
        if self.template_manager.add_template(template):
            self.load_template_list()
            self.log_message(f"模板已保存: {name}", "success")
            QMessageBox.information(self, "保存成功", f"模板 '{name}' 已保存")
        else:
            QMessageBox.warning(self, "保存失败", f"模板 '{name}' 已存在")

    def setup_ui(self):
        """Setup the user interface."""
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Left panel
        left_panel = self.create_left_panel()
        main_layout.addWidget(left_panel, 1)
        
        # Right panel
        right_panel = self.create_right_panel()
        main_layout.addWidget(right_panel, 1)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
        
        # Menu bar
        self.create_menu_bar()
        
    def create_left_panel(self) -> QWidget:
        """Create left panel with controls."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(16)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Title
        title_label = QLabel("SlideSec")
        title_label.setObjectName("title")
        layout.addWidget(title_label)
        
        subtitle = QLabel("PowerPoint 转图片演示文稿工具")
        subtitle.setObjectName("subtitle")
        layout.addWidget(subtitle)
        
        layout.addSpacing(20)

        # Template selection group
        template_group = QGroupBox("快速预设")
        template_layout = QVBoxLayout(template_group)

        self.template_combo = QComboBox()
        self.template_combo.addItem("选择预设模板...", None)
        self.load_template_list()
        self.template_combo.currentIndexChanged.connect(self.on_template_selected)
        template_layout.addWidget(self.template_combo)

        # Save current as template button
        save_template_btn = QPushButton("💾 保存当前设置为模板")
        save_template_btn.setObjectName("secondary")
        save_template_btn.clicked.connect(self.save_current_as_template)
        template_layout.addWidget(save_template_btn)

        layout.addWidget(template_group)

        # File selection group
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)

        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self.on_file_selected)
        file_layout.addWidget(self.drop_area)

        self.file_info_label = QLabel("未选择文件")
        self.file_info_label.setStyleSheet("color: #718096; padding: 8px;")
        file_layout.addWidget(self.file_info_label)

        layout.addWidget(file_group)

        # Format settings group
        format_group = QGroupBox("输出设置")
        format_layout = QGridLayout(format_group)
        format_layout.setSpacing(12)
        
        # Format selection
        format_layout.addWidget(QLabel("图片格式:"), 0, 0)
        self.format_combo = QComboBox()
        for fmt_id, fmt_config in IMAGE_FORMAT_CONFIGS.items():
            self.format_combo.addItem(fmt_config['description'], fmt_id)
        self.format_combo.currentIndexChanged.connect(self.on_format_changed)
        format_layout.addWidget(self.format_combo, 0, 1)
        
        # Quality slider
        format_layout.addWidget(QLabel("图片质量:"), 1, 0)
        quality_layout = QHBoxLayout()
        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(1, 100)
        self.quality_slider.setValue(DEFAULT_QUALITY)
        self.quality_slider.valueChanged.connect(self.on_quality_changed)
        quality_layout.addWidget(self.quality_slider)
        
        self.quality_label = QLabel(f"{DEFAULT_QUALITY}%")
        self.quality_label.setObjectName("value")
        self.quality_label.setMinimumWidth(40)
        quality_layout.addWidget(self.quality_label)
        format_layout.addLayout(quality_layout, 1, 1)
        
        # DPI selection
        format_layout.addWidget(QLabel("DPI设置:"), 2, 0)
        self.dpi_combo = QComboBox()
        for dpi in [72, 96, 150, 300, 600]:
            self.dpi_combo.addItem(f"{dpi} DPI", dpi)
        self.dpi_combo.setCurrentIndex(1)  # 96 DPI
        format_layout.addWidget(self.dpi_combo, 2, 1)
        
        # Resolution selection
        format_layout.addWidget(QLabel("分辨率:"), 3, 0)
        self.resolution_combo = QComboBox()
        self.resolution_combo.addItem("原始尺寸", None)
        self.resolution_combo.addItem("1920x1080 (Full HD)", (1920, 1080))
        self.resolution_combo.addItem("1280x720 (HD)", (1280, 720))
        self.resolution_combo.addItem("3840x2160 (4K)", (3840, 2160))
        format_layout.addWidget(self.resolution_combo, 3, 1)
        
        layout.addWidget(format_group)
        
        # Size estimation group
        size_group = QGroupBox("文件大小预估")
        size_layout = QVBoxLayout(size_group)
        
        self.size_estimate_label = QLabel("选择文件后将显示预估大小")
        self.size_estimate_label.setStyleSheet("font-size: 14px; padding: 12px;")
        size_layout.addWidget(self.size_estimate_label)
        
        layout.addWidget(size_group)
        
        layout.addStretch()
        
        # Convert button
        self.convert_btn = QPushButton("🚀 开始转换")
        self.convert_btn.setMinimumHeight(50)
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)
        
        return panel
        
    def create_right_panel(self) -> QWidget:
        """Create right panel with progress and logs."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(16)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Progress group
        progress_group = QGroupBox("转换进度")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setSpacing(12)
        
        # Overall progress
        progress_layout.addWidget(QLabel("总体进度:"))
        self.overall_progress = QProgressBar()
        self.overall_progress.setTextVisible(True)
        progress_layout.addWidget(self.overall_progress)
        
        # Stage progress
        progress_layout.addWidget(QLabel("当前阶段:"))
        self.stage_progress = QProgressBar()
        self.stage_progress.setTextVisible(True)
        progress_layout.addWidget(self.stage_progress)
        
        # Stage label
        self.stage_label = QLabel("准备就绪")
        self.stage_label.setStyleSheet("color: #718096; font-weight: 600;")
        progress_layout.addWidget(self.stage_label)
        
        # Stats
        stats_layout = QHBoxLayout()
        self.speed_label = QLabel("速度: -")
        self.eta_label = QLabel("剩余: -")
        stats_layout.addWidget(self.speed_label)
        stats_layout.addWidget(self.eta_label)
        stats_layout.addStretch()
        progress_layout.addLayout(stats_layout)
        
        # Control buttons
        btn_layout = QHBoxLayout()
        self.pause_btn = QPushButton("⏸️ 暂停")
        self.pause_btn.setEnabled(False)
        self.pause_btn.clicked.connect(self.toggle_pause)
        btn_layout.addWidget(self.pause_btn)
        
        self.cancel_btn = QPushButton("⏹️ 取消")
        self.cancel_btn.setObjectName("danger")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_conversion)
        btn_layout.addWidget(self.cancel_btn)
        
        progress_layout.addLayout(btn_layout)
        
        layout.addWidget(progress_group)
        
        # Log group
        log_group = QGroupBox("转换日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(200)
        log_layout.addWidget(self.log_text)
        
        # Log buttons
        log_btn_layout = QHBoxLayout()
        clear_log_btn = QPushButton("🗑️ 清除日志")
        clear_log_btn.setObjectName("secondary")
        clear_log_btn.clicked.connect(self.clear_logs)
        log_btn_layout.addWidget(clear_log_btn)
        log_btn_layout.addStretch()
        
        save_log_btn = QPushButton("💾 保存日志")
        save_log_btn.setObjectName("secondary")
        save_log_btn.clicked.connect(self.save_logs)
        log_btn_layout.addWidget(save_log_btn)
        
        log_layout.addLayout(log_btn_layout)
        layout.addWidget(log_group)
        
        return panel
        
    def create_menu_bar(self):
        """Create menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("文件")

        open_action = QAction("打开...", self)
        open_action.setShortcut(QKeySequence.StandardKey.Open)
        open_action.triggered.connect(self.open_file_dialog)
        file_menu.addAction(open_action)

        # Recent files submenu
        self.recent_menu = file_menu.addMenu("最近文件")
        self.update_recent_files_menu()

        file_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # View menu
        view_menu = menubar.addMenu("视图")

        self.theme_action = QAction("深色主题", self)
        self.theme_action.setCheckable(True)
        self.theme_action.setChecked(self.is_dark_theme)
        self.theme_action.triggered.connect(self.toggle_theme)
        view_menu.addAction(self.theme_action)

        # Tools menu
        tools_menu = menubar.addMenu("工具")

        # Convert action with shortcut
        convert_action = QAction("开始转换", self)
        convert_action.setShortcut(QKeySequence("Ctrl+Return"))
        convert_action.triggered.connect(self.start_conversion)
        tools_menu.addAction(convert_action)

        tools_menu.addSeparator()

        # Clear action
        clear_action = QAction("清除选择", self)
        clear_action.setShortcut(QKeySequence("Ctrl+D"))
        clear_action.triggered.connect(self.clear_selection)
        tools_menu.addAction(clear_action)

        # Help menu
        help_menu = menubar.addMenu("帮助")

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # Setup global shortcuts
        self.setup_shortcuts()

    def setup_shortcuts(self):
        """Setup keyboard shortcuts."""
        from PyQt6.QtGui import QShortcut

        # Escape to cancel conversion
        self.shortcut_escape = QShortcut(QKeySequence("Escape"), self)
        self.shortcut_escape.activated.connect(self.cancel_conversion)

        # Ctrl+T to toggle theme
        self.shortcut_theme = QShortcut(QKeySequence("Ctrl+T"), self)
        self.shortcut_theme.activated.connect(self.toggle_theme)

    def open_file_dialog(self):
        """Open file dialog."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 PowerPoint 文件", "",
            "PowerPoint Files (*.ppt *.pptx)"
        )
        if file_path:
            self.on_file_selected(file_path)
            
    def on_file_selected(self, file_path: str):
        """Handle file selection."""
        path = Path(file_path)

        # Validate file
        result = FileValidator.validate_ppt_file(path)
        if not result.is_valid:
            QMessageBox.critical(self, "文件错误", "\n".join(result.errors))
            return

        self.current_file = path
        self.file_info_label.setText(f"📄 {path.name}")
        self.convert_btn.setEnabled(True)

        # Update size estimate
        self.update_size_estimate()

        # Add to recent files
        self.settings_manager.add_recent_file(str(path))
        self.update_recent_files_menu()

        self.log_message(f"已选择文件: {path.name}", "info")
        self.status_bar.showMessage(f"已选择: {path.name}")

    def clear_selection(self):
        """Clear current file selection."""
        self.current_file = None
        self.file_info_label.setText("未选择文件")
        self.convert_btn.setEnabled(False)
        self.size_estimate_label.setText("选择文件后将显示预估大小")
        self.log_message("文件选择已清除", "info")
        self.status_bar.showMessage("就绪")

    def on_format_changed(self):
        """Handle format change."""
        self.update_size_estimate()
        
    def on_quality_changed(self, value: int):
        """Handle quality slider change."""
        self.quality_label.setText(f"{value}%")
        self.update_size_estimate()
        
    def update_size_estimate(self):
        """Update file size estimate."""
        if not self.current_file:
            return
            
        try:
            info = self.processor.get_presentation_info(self.current_file)
            slide_count = info["slide_count"]
            
            format_name = self.format_combo.currentData()
            quality = self.quality_slider.value()
            
            estimate = self.processor.estimate_output_size(
                info, format_name, quality
            )
            
            text = f"""
            <p style='font-size: 16px; margin: 0;'>
                <b>预估大小:</b> <span style='color: #667eea;'>{estimate.estimated_size_mb:.1f} MB</span>
            </p>
            <p style='font-size: 12px; color: #718096; margin: 4px 0 0 0;'>
                幻灯片数: {slide_count} | 
                置信度: {estimate.confidence}% | 
                范围: {estimate.min_size_mb:.1f} - {estimate.max_size_mb:.1f} MB
            </p>
            """
            self.size_estimate_label.setText(text)
            
        except Exception as e:
            self.size_estimate_label.setText("无法预估文件大小")
            
    def start_conversion(self):
        """Start conversion process."""
        if not self.current_file:
            QMessageBox.warning(self, "未选择文件", "请先选择要转换的PowerPoint文件")
            return
            
        # Get settings
        format_name = self.format_combo.currentData()
        quality = self.quality_slider.value()
        dpi = self.dpi_combo.currentData()
        resolution = self.resolution_combo.currentData()
        
        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = self.current_file.parent / f"{self.current_file.stem}_converted_{timestamp}.pptx"
        
        # Validate settings
        settings = {
            "format_name": format_name,
            "quality": quality,
            "resolution": resolution,
            "dpi": dpi,
        }
        
        result = validate_conversion_request(
            str(self.current_file),
            str(self.current_file.parent),
            settings
        )
        
        if not result.is_valid:
            QMessageBox.critical(self, "设置错误", "\n".join(result.errors))
            return
            
        # Confirm conversion
        reply = QMessageBox.question(
            self, "确认转换",
            f"将转换 {self.current_file.name}\n"
            f"输出格式: {IMAGE_FORMAT_CONFIGS[format_name]['description']}\n"
            f"输出位置: {output_file}\n\n"
            f"是否开始转换？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return
            
        # Start conversion
        self.is_converting = True
        self.convert_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        
        # Reset progress
        self.overall_progress.setValue(0)
        self.stage_progress.setValue(0)
        
        # Create worker
        self.worker = ConversionWorker(
            self.processor,
            input_file=self.current_file,
            output_file=output_file,
            format_name=format_name,
            quality=quality,
            resolution=resolution,
            dpi=dpi,
        )
        
        self.worker.progress_signal.connect(self.on_progress_update)
        self.worker.completed_signal.connect(self.on_conversion_complete)
        self.worker.error_signal.connect(self.on_conversion_error)
        
        self.worker.start()
        
        self.log_message("开始转换...", "info")
        self.status_bar.showMessage("正在转换...")
        
    def on_progress_update(self, progress_info: ProgressInfo):
        """Handle progress update."""
        overall_percent = int(progress_info.overall_progress * 100)
        stage_percent = int(progress_info.stage_progress * 100)
        
        self.overall_progress.setValue(overall_percent)
        self.stage_progress.setValue(stage_percent)
        
        stage_names = {
            "initialize": "初始化",
            "analyze_presentation": "分析演示文稿",
            "export_images": "导出图片",
            "process_images": "处理图片",
            "create_presentation": "创建演示文稿",
            "finalize": "完成",
        }
        
        stage_text = stage_names.get(progress_info.current_stage, progress_info.current_stage)
        if progress_info.current_slide > 0:
            stage_text += f" ({progress_info.current_slide}/{progress_info.total_slides})"
            
        self.stage_label.setText(stage_text)
        
        if progress_info.processing_speed:
            self.speed_label.setText(f"速度: {progress_info.processing_speed:.1f} 张/秒")
            
        if progress_info.overall_eta_seconds:
            minutes = int(progress_info.overall_eta_seconds / 60)
            seconds = int(progress_info.overall_eta_seconds % 60)
            self.eta_label.setText(f"剩余: {minutes}:{seconds:02d}")
            
    def on_conversion_complete(self, result: dict):
        """Handle conversion completion."""
        self.is_converting = False
        self.convert_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        
        if result.get("success"):
            output_file = Path(result["output_file"])
            processing_time = result.get("processing_time", 0)
            
            self.log_message(f"✅ 转换完成！用时 {processing_time:.1f} 秒", "success")
            
            QMessageBox.information(
                self, "转换完成",
                f"转换成功完成！\n\n"
                f"输出文件: {output_file.name}\n"
                f"处理时间: {processing_time:.1f} 秒\n\n"
                f"是否打开输出文件夹？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            self.status_bar.showMessage("转换完成", 5000)
        else:
            error_msg = result.get("error", "未知错误")
            self.log_message(f"❌ 转换失败: {error_msg}", "error")
            QMessageBox.critical(self, "转换失败", f"转换过程中发生错误:\n{error_msg}")
            self.status_bar.showMessage("转换失败")
            
    def on_conversion_error(self, error_msg: str):
        """Handle conversion error."""
        self.is_converting = False
        self.convert_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        
        self.log_message(f"❌ 错误: {error_msg}", "error")
        QMessageBox.critical(self, "错误", f"转换失败:\n{error_msg}")
        self.status_bar.showMessage("错误")
        
    def toggle_pause(self):
        """Toggle pause/resume."""
        # TODO: Implement pause functionality
        pass
        
    def cancel_conversion(self):
        """Cancel conversion."""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.log_message("转换已取消", "warning")
            
        self.is_converting = False
        self.convert_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.status_bar.showMessage("已取消")
        
    def log_message(self, message: str, level: str = "info"):
        """Add message to log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        colors = {
            "info": "#a0aec0",
            "success": "#48bb78",
            "warning": "#ed8936",
            "error": "#f56565",
        }
        
        color = colors.get(level, "#a0aec0")
        html = f"<span style='color: {color};'>[{timestamp}] {message}</span>"
        
        self.log_text.append(html)
        
    def clear_logs(self):
        """Clear log text."""
        self.log_text.clear()
        
    def save_logs(self):
        """Save logs to file."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存日志", "slide_sec_log.txt",
            "Text Files (*.txt)"
        )
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.log_text.toPlainText())
            QMessageBox.information(self, "保存成功", f"日志已保存到:\n{file_path}")
            
    def show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self, f"关于 {APP_NAME}",
            f"<h2>{APP_NAME} {APP_VERSION}</h2>"
            f"<p>专业的 PowerPoint 转图片演示文稿工具</p>"
            f"<p>支持多种图片格式，高质量输出</p>"
            f"<p>© 2024 SlideSec. All rights reserved.</p>"
        )
        
    def update_recent_files_menu(self):
        """Update the recent files menu."""
        self.recent_menu.clear()

        recent_files = self.user_settings.recent_files

        if not recent_files:
            no_recent_action = QAction("无最近文件", self)
            no_recent_action.setEnabled(False)
            self.recent_menu.addAction(no_recent_action)
        else:
            for file_path in recent_files:
                action = QAction(Path(file_path).name, self)
                action.setData(file_path)
                action.triggered.connect(lambda checked, path=file_path: self.on_file_selected(path))
                self.recent_menu.addAction(action)

            self.recent_menu.addSeparator()
            clear_action = QAction("清除历史", self)
            clear_action.triggered.connect(self.clear_recent_files)
            self.recent_menu.addAction(clear_action)

    def clear_recent_files(self):
        """Clear recent files list."""
        self.settings_manager.clear_recent_files()
        self.update_recent_files_menu()
        self.log_message("最近文件列表已清除", "info")

    def closeEvent(self, event):
        """Handle window close event."""
        if self.is_converting:
            reply = QMessageBox.question(
                self, "确认退出",
                "转换正在进行中，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                if self.worker:
                    self.worker.stop()
                self.save_window_settings()
                event.accept()
            else:
                event.ignore()
        else:
            self.save_window_settings()
            event.accept()


def main():
    """Main entry point for PyQt6 application."""
    app = QApplication(sys.argv)
    
    # Set application font
    font = QFont("Segoe UI", 10)
    font.setStyleHint(QFont.StyleHint.SansSerif)
    app.setFont(font)
    
    # Create and show window
    window = ModernMainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

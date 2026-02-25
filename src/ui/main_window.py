import sys
import os
import shutil

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QProgressBar, QTextEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, 
                             QComboBox, QSlider, QMessageBox, QCheckBox, QAbstractItemView,
                             QGroupBox, QFormLayout, QLineEdit, QFrame, QStackedWidget,
                             QListWidget, QListWidgetItem, QToolButton, QStatusBar,
                             QTabWidget, QSplitter, QMenu, QSpinBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QTimer
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont

from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
from src.utils.logger import Logger
from src.utils.config_manager import ConfigManager
from src.utils.history_manager import HistoryManager
from src.ui.smart_config_widget import SmartConfigWidget
from pathlib import Path
from typing import List


class ModernStyleSheet:
    """现代化样式表"""
    
    MAIN_STYLE = """
        QMainWindow {
            background-color: #f5f7fa;
        }
        
        QWidget {
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            font-size: 13px;
        }
        
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
        
        QPushButton {
            background-color: #2563eb;
            color: white;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 500;
            font-size: 12px;
            min-height: 32px;
            max-height: 36px;
        }
        
        QPushButton:hover {
            background-color: #1d4ed8;
        }
        
        QPushButton:pressed {
            background-color: #1e40af;
        }
        
        QPushButton:disabled {
            background-color: #cbd5e0;
            color: #a0aec0;
        }
        
        QPushButton#secondary {
            background-color: #f3f4f6;
            color: #374151;
            border: 1px solid #d1d5db;
        }
        
        QPushButton#secondary:hover {
            background-color: #e5e7eb;
        }
        
        QPushButton#danger {
            background-color: #ef4444;
        }
        
        QPushButton#danger:hover {
            background-color: #dc2626;
        }
        
        QPushButton#success {
            background-color: #10b981;
        }
        
        QPushButton#success:hover {
            background-color: #059669;
        }
        
        QLineEdit {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 6px 10px;
            min-height: 16px;
            font-size: 12px;
        }
        
        QLineEdit:focus {
            border-color: #667eea;
        }
        
        QLineEdit:hover {
            border-color: #cbd5e0;
        }
        
        QComboBox {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 6px 10px;
            min-height: 16px;
            min-width: 100px;
            font-size: 12px;
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
        
        QComboBox QAbstractItemView {
            background-color: white;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            selection-background-color: #2563eb;
            selection-color: white;
            padding: 4px;
        }
        
        QSlider::groove:horizontal {
            height: 6px;
            background: #e5e7eb;
            border-radius: 3px;
        }
        
        QSlider::sub-page:horizontal {
            background-color: #2563eb;
            border-radius: 3px;
        }
        
        QSlider::handle:horizontal {
            background: white;
            border: 2px solid #2563eb;
            width: 16px;
            height: 16px;
            margin: -5px 0;
            border-radius: 8px;
        }
        
        QSlider::handle:horizontal:hover {
            background: #2563eb;
            border-color: #1d4ed8;
        }
        
        QProgressBar {
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            background-color: #f7fafc;
            text-align: center;
            font-weight: 500;
            font-size: 11px;
            color: #2d3748;
            min-height: 20px;
            max-height: 20px;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #667eea, stop:1 #764ba2);
            border-radius: 5px;
        }
        
        QProgressBar::chunk:disabled {
            background-color: #cbd5e0;
        }
        
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
        
        QCheckBox {
            spacing: 6px;
            font-size: 12px;
            color: #374151;
        }
        
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
            border: 1px solid #d1d5db;
            border-radius: 3px;
            background-color: white;
        }
        
        QCheckBox::indicator:checked {
            background-color: #2563eb;
            border-color: #2563eb;
            image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHZpZXdCb3g9IjAgMCAxMiAxMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEwIDNMNC41IDguNUwyIDYiIHN0cm9rZT0id2hpdGUiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIi8+Cjwvc3ZnPgo=);
        }
        
        QCheckBox::indicator:hover {
            border-color: #2563eb;
        }
        
        QListWidget {
            background-color: white;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            padding: 6px;
            outline: none;
        }
        
        QListWidget::item {
            padding: 8px;
            border-radius: 4px;
            margin-bottom: 2px;
        }
        
        QListWidget::item:selected {
            background-color: #2563eb;
            color: white;
        }
        
        QListWidget::item:hover {
            background-color: #f3f4f6;
        }
        
        QStatusBar {
            background-color: white;
            border-top: 1px solid #e5e7eb;
            color: #6b7280;
        }
    """


class StepIndicator(QWidget):
    """步骤指示器"""
    
    def __init__(self, steps: List[str], parent=None):
        super().__init__(parent)
        self.steps = steps
        self.current_step = 0
        self.setFixedHeight(70)
        self.setup_ui()
        
    def setup_ui(self):
        # 使用水平布局，整体居中
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.setAlignment(Qt.AlignCenter)
        
        # 创建内容容器
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(20, 8, 20, 8)
        container_layout.setSpacing(0)
        container_layout.setAlignment(Qt.AlignCenter)
        
        self.step_labels = []
        for i, step in enumerate(self.steps):
            # 步骤项（数字+文字）
            step_item = QWidget()
            step_item.setFixedWidth(80)
            step_layout = QVBoxLayout(step_item)
            step_layout.setContentsMargins(0, 0, 0, 0)
            step_layout.setSpacing(4)
            step_layout.setAlignment(Qt.AlignCenter)
            
            # 数字圆圈
            number_label = QLabel(str(i + 1))
            number_label.setAlignment(Qt.AlignCenter)
            number_label.setFixedSize(26, 26)
            number_label.setStyleSheet("""
                QLabel {
                    background-color: #e5e7eb;
                    color: #6b7280;
                    border-radius: 13px;
                    font-weight: bold;
                    font-size: 11px;
                }
            """)
            
            # 文字标签
            text_label = QLabel(step)
            text_label.setAlignment(Qt.AlignCenter)
            text_label.setStyleSheet("color: #6b7280; font-size: 10px;")
            
            step_layout.addWidget(number_label, alignment=Qt.AlignCenter)
            step_layout.addWidget(text_label, alignment=Qt.AlignCenter)
            
            container_layout.addWidget(step_item)
            self.step_labels.append((number_label, text_label))
            
            # 添加连接线
            if i < len(self.steps) - 1:
                line_container = QWidget()
                line_container.setFixedWidth(60)
                line_layout = QHBoxLayout(line_container)
                line_layout.setContentsMargins(0, 0, 0, 12)  # 底部偏移，与圆圈中心对齐
                line_layout.setAlignment(Qt.AlignCenter)
                
                line = QFrame()
                line.setFrameShape(QFrame.HLine)
                line.setStyleSheet("background-color: #e5e7eb;")
                line.setFixedHeight(2)
                line.setFixedWidth(40)
                
                line_layout.addWidget(line)
                container_layout.addWidget(line_container)
        
        main_layout.addWidget(container)
        self.update_style()
        
    def set_current_step(self, step: int):
        self.current_step = step
        self.update_style()
        
    def update_style(self):
        for i, (num_label, text_label) in enumerate(self.step_labels):
            if i < self.current_step:
                # 已完成 - 使用绿色
                num_label.setStyleSheet("""
                    QLabel {
                        background-color: #10b981;
                        color: white;
                        border-radius: 13px;
                        font-weight: bold;
                        font-size: 11px;
                    }
                """)
                text_label.setStyleSheet("color: #10b981; font-size: 10px; font-weight: bold;")
            elif i == self.current_step:
                # 当前步骤 - 使用蓝色
                num_label.setStyleSheet("""
                    QLabel {
                        background-color: #2563eb;
                        color: white;
                        border-radius: 13px;
                        font-weight: bold;
                        font-size: 11px;
                    }
                """)
                text_label.setStyleSheet("color: #2563eb; font-size: 10px; font-weight: bold;")
            else:
                # 待处理 - 使用灰色
                num_label.setStyleSheet("""
                    QLabel {
                        background-color: #e5e7eb;
                        color: #9ca3af;
                        border-radius: 13px;
                        font-weight: bold;
                        font-size: 11px;
                    }
                """)
                text_label.setStyleSheet("color: #9ca3af; font-size: 10px;")


class FileDropZone(QFrame):
    """文件拖放区域"""
    
    file_dropped = pyqtSignal(str)
    folder_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(200)
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(12)
        
        # 图标
        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 48px;")
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)
        
        # 主文本
        self.main_text = QLabel("拖拽 PowerPoint 文件或文件夹到此处")
        self.main_text.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #1a1a2e;
                font-weight: 500;
            }
        """)
        self.main_text.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.main_text)
        
        # 副文本
        self.sub_text = QLabel("支持 .ppt 和 .pptx 格式，可拖拽文件夹")
        self.sub_text.setStyleSheet("color: #718096; font-size: 12px;")
        self.sub_text.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.sub_text)
        
        self.setStyleSheet("""
            FileDropZone {
                background-color: white;
                border: 2px dashed #cbd5e0;
                border-radius: 10px;
            }
            FileDropZone:hover {
                border-color: #667eea;
                background-color: #f7fafc;
            }
        """)
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 PowerPoint 文件", "",
            "PowerPoint 文件 (*.ppt *.pptx)"
        )
        if file_path:
            self.file_dropped.emit(file_path)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                FileDropZone {
                    background-color: #f7fafc;
                    border: 2px dashed #667eea;
                    border-radius: 12px;
                }
            """)
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            FileDropZone {
                background-color: white;
                border: 2px dashed #cbd5e0;
                border-radius: 12px;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet("""
            FileDropZone {
                background-color: white;
                border: 2px dashed #cbd5e0;
                border-radius: 12px;
            }
        """)
        
        urls = event.mimeData().urls()
        if urls:
            for url in urls:
                file_path = url.toLocalFile()
                if os.path.isdir(file_path):
                    # 文件夹
                    self.folder_dropped.emit(file_path)
                elif file_path.endswith(('.ppt', '.pptx')):
                    # 文件
                    self.file_dropped.emit(file_path)


class ConversionWorker(QThread):
    """转换工作线程 - 支持暂停和终止"""

    progress_updated = pyqtSignal(float, str)
    conversion_finished = pyqtSignal(bool, str)

    def __init__(self, input_path, output_path, options):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.options = options
        self.logger = Logger().get_logger()
        self.converter = None  # 在线程中创建，避免COM线程问题

        # 控制标志
        self._paused = False
        self._stopped = False
        self._pause_condition = __import__('threading').Condition()

    def pause(self):
        """暂停转换"""
        self._paused = True
        self.logger.info("转换已暂停")

    def resume(self):
        """继续转换"""
        with self._pause_condition:
            self._paused = False
            self._pause_condition.notify_all()
        self.logger.info("转换继续")

    def stop(self):
        """终止转换"""
        self._stopped = True
        self.resume()  # 如果处于暂停状态，先唤醒线程
        self.logger.info("转换已终止")

    def _check_pause(self):
        """检查是否需要暂停"""
        with self._pause_condition:
            while self._paused and not self._stopped:
                self._pause_condition.wait(0.1)

    def run(self):
        try:
            self.logger.info(f"[Worker] 开始转换任务")
            self.logger.info(f"[Worker] 输入: {self.input_path}")
            self.logger.info(f"[Worker] 输出: {self.output_path}")
            
            # 在线程中创建converter实例（避免COM线程问题）
            self.converter = Win32PPTConverter()
            self.logger.info(f"[Worker] Converter实例已创建")
            
            # 创建包装后的进度回调
            original_callback = self.converter.progress_tracker.callback

            def progress_wrapper(progress, message=""):
                # 检查是否暂停
                self._check_pause()
                # 检查是否终止
                if self._stopped:
                    raise InterruptedError("转换被用户终止")
                # 更新进度
                self.progress_updated.emit(progress, message)
                # 调用原始回调
                if original_callback:
                    original_callback(progress, message)

            # 替换callback
            self.converter.progress_tracker.callback = progress_wrapper

            self.logger.info(f"[Worker] 开始调用convert方法")
            success = self.converter.convert(
                self.input_path,
                self.output_path,
                self.options
            )
            self.logger.info(f"[Worker] convert方法返回: {success}")

            # 恢复原始callback
            self.converter.progress_tracker.callback = original_callback

            # 检查文件是否真的生成
            if success and not os.path.exists(self.output_path):
                self.logger.error(f"[Worker] 转换返回成功但文件不存在: {self.output_path}")
                success = False

            if self._stopped:
                self.conversion_finished.emit(False, "用户终止")
            else:
                self.conversion_finished.emit(success, self.output_path)

        except InterruptedError:
            self.logger.info("[Worker] 转换被用户终止")
            self.conversion_finished.emit(False, "用户终止")
        except Exception as e:
            self.logger.error(f"[Worker] 转换线程异常: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.conversion_finished.emit(False, str(e))
        finally:
            # 清理converter
            if self.converter:
                try:
                    self.converter._cleanup()
                except:
                    pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logger = Logger().get_logger()
        self.logger.info("Initializing PitchPPT application")
        
        # 初始化转换器
        self.converter = Win32PPTConverter()
        self.current_input_file = None
        self.current_output_dir = None
        self._smart_optimization_target_size = None  # 记录智能处理的目标大小
        self._smart_optimization_retry_count = 0  # 记录智能处理重试次数
        self.files: List[Path] = []
        
        # 初始化配置和历史记录管理器
        self.config_manager = ConfigManager()
        self.history_manager = HistoryManager()
        
        # 最小化到托盘标志
        self._minimize_to_tray = False
        
        # 应用样式
        self.setStyleSheet(ModernStyleSheet.MAIN_STYLE)
        
        # 设置窗口大小（不固定，允许调整）
        self.setMinimumSize(900, 650)
        self.resize(950, 700)
        
        self.init_ui()
        self._init_tray_icon()
    
    def init_ui(self):
        self.setWindowTitle("PitchPPT - 专业路演PPT处理工具")
        
        # 设置窗口图标（任务栏显示）- 使用多尺寸ICO文件
        try:
            ico_paths = [
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO_256x256.ico"),
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO_128x128.ico"),
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO_64x64.ico"),
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO_48x48.ico"),
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO_32x32.ico"),
                os.path.join(os.path.dirname(__file__), "../../resources/LOGO.png"),
            ]
            
            for logo_path in ico_paths:
                if os.path.exists(logo_path):
                    from PyQt5.QtGui import QIcon
                    icon = QIcon(logo_path)
                    if not icon.isNull():
                        self.setWindowIcon(icon)
                        self.logger.info(f"Window icon loaded: {logo_path}")
                        break
        except Exception as e:
            self.logger.warning(f"Failed to set window icon: {e}")
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 创建头部
        header = self.create_header()
        main_layout.addWidget(header)
        
        # 创建主页面堆栈（用于切换主页、历史记录、设置）
        self.main_stack = QStackedWidget()
        main_layout.addWidget(self.main_stack, stretch=1)
        
        # 创建转换页面（包含步骤指示器和步骤页面）
        self.conversion_page = self.create_conversion_page()
        self.main_stack.addWidget(self.conversion_page)
        
        # 创建历史记录页面
        self.history_page = self.create_history_page()
        self.main_stack.addWidget(self.history_page)
        
        # 创建设置页面
        self.settings_page = self.create_settings_page()
        self.main_stack.addWidget(self.settings_page)
        
        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪 - 请选择要转换的文件")

        # 添加"遇到问题?"链接和版本号标签到状态栏右侧
        status_right_widget = QWidget()
        status_right_layout = QHBoxLayout(status_right_widget)
        status_right_layout.setContentsMargins(0, 0, 10, 0)
        status_right_layout.setSpacing(10)
        
        # "遇到问题?"链接
        help_label = QLabel("<a href='#' style='color: #3b82f6; text-decoration: underline;'>遇到问题?</a>")
        help_label.setStyleSheet("font-size: 11px;")
        help_label.setCursor(Qt.PointingHandCursor)
        help_label.setOpenExternalLinks(False)
        help_label.linkActivated.connect(self._show_help_dialog)
        status_right_layout.addWidget(help_label)
        
        # 版本号
        version_label = QLabel("v1.7.2")
        version_label.setStyleSheet("color: #9ca3af; font-size: 11px;")
        status_right_layout.addWidget(version_label)
        
        self.status_bar.addPermanentWidget(status_right_widget)

        self.logger.info("Main window UI initialized with wizard-style interface")
    
    def _init_tray_icon(self):
        """初始化系统托盘图标"""
        try:
            from PyQt5.QtWidgets import QSystemTrayIcon, QMenu
            from PyQt5.QtGui import QIcon
            
            # 创建托盘图标
            icon_path = os.path.join(os.path.dirname(__file__), "../../resources/LOGO_64x64.ico")
            if not os.path.exists(icon_path):
                icon_path = os.path.join(os.path.dirname(__file__), "../../resources/LOGO.png")
            
            if os.path.exists(icon_path):
                icon = QIcon(icon_path)
            else:
                icon = self.windowIcon()
            
            self.tray_icon = QSystemTrayIcon(icon, self)
            
            # 创建托盘菜单
            tray_menu = QMenu()
            
            show_action = tray_menu.addAction("显示窗口")
            show_action.triggered.connect(self.show_and_activate)
            
            quit_action = tray_menu.addAction("退出")
            quit_action.triggered.connect(self.quit_application)
            
            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.activated.connect(self._on_tray_activated)
            self.tray_icon.show()
            
            self.logger.info("系统托盘图标已初始化")
        except Exception as e:
            self.logger.warning(f"初始化系统托盘失败: {e}")
            self.tray_icon = None
    
    def _on_tray_activated(self, reason):
        """托盘图标激活事件"""
        from PyQt5.QtWidgets import QSystemTrayIcon
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_and_activate()
    
    def _show_help_dialog(self):
        """显示帮助对话框，包含作者信息和开源仓库链接"""
        from PyQt5.QtWidgets import QMessageBox
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("帮助 & 联系")
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setText("""
        <p style='font-size: 14px; font-weight: bold; margin-bottom: 10px;'>PitchPPT 智能PPT处理工具</p>
        <p style='margin-bottom: 8px;'><b>作者：</b>Jiachen Bao</p>
        <p style='margin-bottom: 8px;'><b>联系邮箱：</b>thestein@foxmail.com</p>
        <hr style='margin: 12px 0;'>
        <p style='margin-bottom: 8px;'>遇到问题或有建议？欢迎通过以下方式反馈：</p>
        <p style='margin-bottom: 5px;'>• <a href='https://github.com/baojiachen0214/PitchPPT.git'>GitHub 开源仓库</a></p>
        <p style='margin-bottom: 5px;'>• <a href='https://gitee.com/bao-jiachen/PitchPPT.git'>Gitee 开源仓库</a></p>
        <p style='margin-top: 10px; color: #666; font-size: 12px;'>您可以在开源仓库中提交 Issue 或发送邮件联系开发者</p>
        """)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
    
    def show_and_activate(self):
        """显示并激活窗口"""
        self.show()
        self.activateWindow()
        self.raise_()
    
    def quit_application(self):
        """退出应用"""
        self._minimize_to_tray = False
        self.close()
    
    def closeEvent(self, event):
        """重写关闭事件"""
        if self._minimize_to_tray and self.tray_icon:
            # 最小化到托盘
            event.ignore()
            self.hide()
            if self.tray_icon:
                self.tray_icon.showMessage(
                    "PitchPPT",
                    "程序已最小化到系统托盘，双击图标可恢复窗口",
                    self.tray_icon.Information,
                    2000
                )
        else:
            # 正常关闭
            event.accept()
            if self.tray_icon:
                self.tray_icon.hide()
    
    def create_conversion_page(self) -> QWidget:
        """创建转换页面（包含步骤指示器和步骤堆栈）"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 创建步骤指示器
        self.step_indicator = StepIndicator(["选择文件", "设置参数", "开始转换"])
        layout.addWidget(self.step_indicator)
        
        # 创建步骤内容堆栈
        self.content_stack = QStackedWidget()
        layout.addWidget(self.content_stack, stretch=1)
        
        # 第一步：文件选择
        self.step1_widget = self.create_step1()
        self.content_stack.addWidget(self.step1_widget)
        
        # 第二步：设置参数
        self.step2_widget = self.create_step2()
        self.content_stack.addWidget(self.step2_widget)
        
        # 第三步：开始转换
        self.step3_widget = self.create_step3()
        self.content_stack.addWidget(self.step3_widget)
        
        return page
    
    def create_header(self) -> QWidget:
        """创建头部"""
        header = QWidget()
        header.setFixedHeight(60)
        header.setStyleSheet("background-color: white; border-bottom: 1px solid #e2e8f0;")
        
        layout = QHBoxLayout(header)
        layout.setContentsMargins(20, 0, 20, 0)
        
        # Logo和标题
        logo_layout = QHBoxLayout()
        logo_layout.setSpacing(10)

        # 加载并显示Logo - 优先使用SVG矢量格式以获得最佳清晰度
        # 使用HiDPI感知尺寸：显示尺寸52px，但渲染尺寸为8倍(416px)以支持高分屏
        # 8倍渲染可以确保在300%甚至400%缩放的高分屏上也能保持清晰边缘
        logo_display_size = 52  # 显示尺寸
        logo_render_size = 416  # 渲染尺寸（8倍缩放以获得极致清晰的效果）
        logo_scale_factor = logo_render_size / logo_display_size  # 缩放因子 = 8.0
        
        logo_label = QLabel()
        logo_label.setFixedSize(logo_display_size, logo_display_size)
        # 禁用ScaledContents以避免不必要的缩放导致的模糊
        logo_label.setScaledContents(False)
        # 清除任何可能的样式，避免阴影和圆角
        logo_label.setStyleSheet("background: transparent; border: none; border-radius: 0px;")

        logo_loaded = False

        # 方法1: 优先使用SVG矢量格式（无限缩放不失真）
        try:
            from PyQt5.QtSvg import QSvgRenderer
            from PyQt5.QtGui import QPainter, QPixmap
            from PyQt5.QtCore import QRectF

            svg_path = os.path.join(os.path.dirname(__file__), "../../resources/LOGO.svg")
            if os.path.exists(svg_path):
                renderer = QSvgRenderer(svg_path)
                if renderer.isValid():
                    # 使用2倍尺寸渲染以获得更清晰的显示效果
                    pixmap = QPixmap(logo_render_size, logo_render_size)
                    pixmap.fill(Qt.transparent)

                    # 使用QPainter渲染SVG - 保持原始比例
                    painter = QPainter(pixmap)
                    # 启用所有高质量渲染选项
                    painter.setRenderHint(QPainter.Antialiasing, True)
                    painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
                    painter.setRenderHint(QPainter.HighQualityAntialiasing, True)
                    painter.setRenderHint(QPainter.TextAntialiasing, True)
                    # 注意：VerticalSubpixelPositioning在某些PyQt5版本中不可用

                    # 计算保持比例的缩放
                    view_box = renderer.viewBoxF()
                    if view_box.isValid():
                        # 计算缩放因子，保持宽高比，留一点边距
                        margin = 4 * logo_scale_factor  # 按缩放因子调整边距
                        available_size = logo_render_size - 2 * margin
                        scale = min(available_size / view_box.width(), available_size / view_box.height())
                        new_width = view_box.width() * scale
                        new_height = view_box.height() * scale
                        # 居中显示
                        x_offset = (logo_render_size - new_width) / 2
                        y_offset = (logo_render_size - new_height) / 2
                        target_rect = QRectF(x_offset, y_offset, new_width, new_height)
                    else:
                        target_rect = QRectF(0, 0, logo_render_size, logo_render_size)
                    renderer.render(painter, target_rect)
                    painter.end()
                    
                    # 缩放到显示尺寸 - 使用最高质量缩放
                    pixmap = pixmap.scaled(logo_display_size, logo_display_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    # 额外优化：如果系统支持，启用更高质量的图像处理
                    pixmap.setDevicePixelRatio(1.0)

                    if not pixmap.isNull():
                        logo_label.setPixmap(pixmap)
                        logo_loaded = True
                        self.logger.info(f"Logo loaded from SVG with HiDPI support ({logo_render_size}x{logo_render_size} rendered)")
        except Exception as e:
            self.logger.warning(f"SVG loading failed: {e}")

        # 方法2: 如果SVG加载失败，使用PIL加载ICO文件
        if not logo_loaded:
            try:
                from PIL import Image
                from PyQt5.QtGui import QImage, QPixmap

                ico_paths = [
                    os.path.join(os.path.dirname(__file__), "../../resources/LOGO_256x256.ico"),
                    os.path.join(os.path.dirname(__file__), "../../resources/LOGO_128x128.ico"),
                    os.path.join(os.path.dirname(__file__), "../../resources/LOGO_64x64.ico"),
                    os.path.join(os.path.dirname(__file__), "../../resources/LOGO_48x48.ico"),
                ]

                for ico_path in ico_paths:
                    if os.path.exists(ico_path):
                        # 使用PIL打开ICO文件
                        img = Image.open(ico_path)
                        # 转换为RGBA模式
                        if img.mode != 'RGBA':
                            img = img.convert('RGBA')

                        margin = 4 * logo_scale_factor
                        available_size = logo_render_size - 2 * margin

                        # 调整为可用尺寸，保持比例（使用高质量LANCZOS重采样）
                        img.thumbnail((available_size, available_size), Image.Resampling.LANCZOS)
                        # 创建透明背景
                        final_img = Image.new('RGBA', (logo_render_size, logo_render_size), (0, 0, 0, 0))
                        # 居中粘贴
                        x = (logo_render_size - img.width) // 2
                        y = (logo_render_size - img.height) // 2
                        final_img.paste(img, (x, y))

                        # 转换为QPixmap
                        data = final_img.tobytes('raw', 'RGBA')
                        qimage = QImage(data, logo_render_size, logo_render_size, QImage.Format_ARGB32)
                        pixmap = QPixmap.fromImage(qimage)
                        
                        # 缩放到显示尺寸 - 使用最高质量缩放
                        pixmap = pixmap.scaled(logo_display_size, logo_display_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        pixmap.setDevicePixelRatio(1.0)

                        if not pixmap.isNull():
                            logo_label.setPixmap(pixmap)
                            logo_loaded = True
                            self.logger.info(f"Logo loaded from ICO with HiDPI support ({ico_path})")
                            break
            except Exception as e:
                self.logger.warning(f"PIL ICO loading failed: {e}")

        # 方法3: 如果ICO加载失败，尝试直接使用QPixmap加载PNG
        if not logo_loaded:
            try:
                png_path = os.path.join(os.path.dirname(__file__), "../../resources/LOGO.png")
                if os.path.exists(png_path):
                    pixmap = QPixmap(png_path)
                    if not pixmap.isNull():
                        margin = 4 * logo_scale_factor
                        available_size = logo_render_size - 2 * margin

                        # 保持比例缩放并居中（使用SmoothTransformation）
                        scaled = pixmap.scaled(available_size, available_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        # 创建透明背景
                        final_pixmap = QPixmap(logo_render_size, logo_render_size)
                        final_pixmap.fill(Qt.transparent)
                        # 居中绘制
                        painter = QPainter(final_pixmap)
                        painter.setRenderHint(QPainter.Antialiasing, True)
                        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
                        painter.setRenderHint(QPainter.HighQualityAntialiasing, True)
                        painter.setRenderHint(QPainter.TextAntialiasing, True)
                        x = (logo_render_size - scaled.width()) // 2
                        y = (logo_render_size - scaled.height()) // 2
                        painter.drawPixmap(x, y, scaled)
                        painter.end()
                        
                        # 缩放到显示尺寸 - 使用最高质量缩放
                        final_pixmap = final_pixmap.scaled(logo_display_size, logo_display_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        final_pixmap.setDevicePixelRatio(1.0)
                        logo_label.setPixmap(final_pixmap)
                        logo_loaded = True
            except Exception as e:
                self.logger.warning(f"PNG loading failed: {e}")

        if not logo_loaded:
            # 如果所有图标都加载失败，显示文字Logo
            logo_label.setText("📊")
            logo_label.setStyleSheet("font-size: 32px; background: transparent; border: none;")
            logo_label.setAlignment(Qt.AlignCenter)
        
        title = QLabel("PitchPPT")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1a1a2e;")
        
        logo_layout.addWidget(logo_label)
        logo_layout.addWidget(title)
        layout.addLayout(logo_layout)
        
        layout.addStretch()
        
        # 导航按钮
        nav_btn_style = """
            QPushButton {
                background-color: transparent;
                color: #4a5568;
                border: none;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 600;
                border-radius: 4px;
            }
            QPushButton:hover {
                color: #667eea;
                background-color: #f7fafc;
            }
            QPushButton:checked {
                color: #667eea;
                background-color: #edf2f7;
            }
        """
        
        # 主页按钮
        self.home_btn = QPushButton("🏠 主页")
        self.home_btn.setStyleSheet(nav_btn_style)
        self.home_btn.setCheckable(True)
        self.home_btn.setChecked(True)
        self.home_btn.clicked.connect(self.show_main_page)
        layout.addWidget(self.home_btn)
        
        # 历史记录按钮
        self.history_btn = QPushButton("📜 历史记录")
        self.history_btn.setStyleSheet(nav_btn_style)
        self.history_btn.setCheckable(True)
        self.history_btn.clicked.connect(self.show_history_page)
        layout.addWidget(self.history_btn)
        
        # 设置按钮
        self.settings_btn = QPushButton("⚙️ 设置")
        self.settings_btn.setStyleSheet(nav_btn_style)
        self.settings_btn.setCheckable(True)
        self.settings_btn.clicked.connect(self.show_settings_page)
        layout.addWidget(self.settings_btn)
        
        return header
    
    def create_step1(self) -> QWidget:
        """创建第一步：文件选择"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)
        
        # 标题
        title = QLabel("选择要转换的文件")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a2e;")
        layout.addWidget(title)
        
        subtitle = QLabel("支持 .ppt 和 .pptx 格式的 PowerPoint 文件，可拖拽文件或文件夹")
        subtitle.setStyleSheet("color: #718096; font-size: 12px;")
        layout.addWidget(subtitle)
        
        # 拖放区域
        self.drop_zone = FileDropZone()
        self.drop_zone.file_dropped.connect(self.add_file)
        self.drop_zone.folder_dropped.connect(self.add_folder)
        layout.addWidget(self.drop_zone)
        
        # 文件列表（点击可重新选择）
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(120)
        self.file_list.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 6px;
                outline: none;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 8px;
                border-radius: 4px;
                margin-bottom: 2px;
            }
            QListWidget::item:selected {
                background-color: #667eea;
                color: white;
            }
        """)
        # 双击文件列表项才打开文件选择对话框
        self.file_list.itemDoubleClicked.connect(self.on_file_item_double_clicked)
        layout.addWidget(self.file_list)
        
        # 删除文件按钮
        delete_btn = QPushButton("🗑️ 删除文件")
        delete_btn.setFixedHeight(36)
        delete_btn.setEnabled(False)
        delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #fee2e2;
                color: #dc2626;
                border: 1px solid #fecaca;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #fecaca;
            }
            QPushButton:pressed {
                background-color: #fca5a5;
            }
            QPushButton:disabled {
                background-color: #f3f4f6;
                color: #9ca3af;
                border: 1px solid #e5e7eb;
            }
        """)
        delete_btn.clicked.connect(self.delete_file)
        self.delete_file_btn = delete_btn
        layout.addWidget(delete_btn)
        
        # 导航按钮
        nav_layout = QHBoxLayout()
        nav_layout.addStretch()

        self.step1_next_btn = QPushButton("👉 下一步")
        self.step1_next_btn.setEnabled(False)
        self.step1_next_btn.setFixedHeight(40)
        self.step1_next_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 24px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
            QPushButton:disabled {
                background-color: #9ca3af;
            }
        """)
        self.step1_next_btn.clicked.connect(self.go_to_step2)
        nav_layout.addWidget(self.step1_next_btn)

        layout.addLayout(nav_layout)
        
        return widget
    
    def create_step2(self) -> QWidget:
        """创建第二步：设置参数 - 使用标签页组织"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(10)
        
        # 标题
        title = QLabel("⚙️ 转换设置")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a2e;")
        layout.addWidget(title)
        
        subtitle = QLabel("配置详细的转换参数和选项")
        subtitle.setStyleSheet("color: #718096; font-size: 12px;")
        layout.addWidget(subtitle)
        
        # 创建标签页 - 修复对齐问题
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-top: none;
                border-radius: 0 0 8px 8px;
                background-color: white;
                padding: 12px;
            }
            QTabBar::tab {
                background-color: #f9fafb;
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 12px;
                font-weight: 600;
                color: #6b7280;
                border: 1px solid transparent;
                border-bottom: 1px solid #e5e7eb;
            }
            QTabBar::tab:selected {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-bottom: 1px solid white;
                color: #2563eb;
                font-weight: 700;
            }
            QTabBar::tab:hover:!selected {
                background-color: #f3f4f6;
            }
        """)
        
        # 基础设置标签页
        basic_tab = self.create_basic_settings_tab()
        tabs.addTab(basic_tab, "📋 基础设置")
        
        # 图片设置标签页（包含智能处理功能）
        image_tab = self.create_image_settings_tab()
        tabs.addTab(image_tab, "🖼️ 图片设置")
        
        layout.addWidget(tabs)
        
        # 导航按钮 - 蓝色底白色字，加粗，使用emoji
        nav_layout = QHBoxLayout()
        nav_layout.setContentsMargins(0, 16, 0, 0)
        
        back_btn = QPushButton("👈 返回")
        back_btn.setFixedHeight(36)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
        """)
        back_btn.clicked.connect(self.go_to_step1)
        nav_layout.addWidget(back_btn)
        
        nav_layout.addStretch()
        
        convert_btn = QPushButton("👉 开始转换")
        convert_btn.setFixedHeight(36)
        convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 24px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:pressed {
                background-color: #047857;
            }
        """)
        convert_btn.clicked.connect(self.go_to_step3)
        nav_layout.addWidget(convert_btn)
        
        layout.addLayout(nav_layout)
        
        return widget
    
    def create_basic_settings_tab(self) -> QWidget:
        """创建基础设置标签页"""
        tab = QWidget()
        layout = QFormLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        
        # 输出目录
        output_dir_layout = QHBoxLayout()
        output_dir_layout.setSpacing(8)
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("默认与源文件相同目录")
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setFixedHeight(32)
        self.output_dir_btn = QPushButton("📂 浏览")
        self.output_dir_btn.setFixedHeight(32)
        self.output_dir_btn.setFixedWidth(70)
        self.output_dir_btn.setStyleSheet("""
            QPushButton {
                background-color: #f3f4f6;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 4px 12px;
                font-size: 12px;
                color: #374151;
            }
            QPushButton:hover {
                background-color: #e5e7eb;
                border-color: #9ca3af;
            }
        """)
        self.output_dir_btn.clicked.connect(self.select_output_dir)
        output_dir_layout.addWidget(self.output_dir_edit)
        output_dir_layout.addWidget(self.output_dir_btn)
        layout.addRow("输出目录:", output_dir_layout)
        
        # 文件命名
        self.filename_edit = QLineEdit()
        self.filename_edit.setPlaceholderText("留空使用原文件名")
        layout.addRow("文件命名:", self.filename_edit)
        
        # 输出格式 - 固定使用PPTX格式
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItem("📊 PowerPoint (PPTX)", "pptx")
        self.output_format_combo.setVisible(False)
        
        # 转换模式 - 支持背景填充/前景覆盖
        self.mode_combo = QComboBox()
        self.mode_combo.addItem("🎨 背景填充模式", "background_fill")
        self.mode_combo.addItem("🖼️ 前景覆盖模式（完全覆盖）", "foreground_image")
        self.mode_combo.setToolTip("背景填充：将图片设为幻灯片背景\n前景覆盖：将图片作为前景对象并完全覆盖幻灯片")
        layout.addRow("图片覆盖方式:", self.mode_combo)
        
        # 幻灯片范围 - 修复显示逻辑
        slide_range_layout = QHBoxLayout()
        slide_range_layout.setSpacing(8)
        self.slide_range_combo = QComboBox()
        self.slide_range_combo.addItem("全部幻灯片", "all")
        self.slide_range_combo.addItem("自定义范围", "custom")
        self.slide_range_combo.currentIndexChanged.connect(self.on_slide_range_changed)
        self.slide_range_combo.setFixedWidth(120)
        slide_range_layout.addWidget(self.slide_range_combo)
        
        # 自定义范围输入框 - 默认隐藏
        self.slide_range_edit = QLineEdit()
        self.slide_range_edit.setPlaceholderText("如: 1-5,8,10-12")
        self.slide_range_edit.setVisible(False)
        self.slide_range_edit.setFixedWidth(140)
        self.slide_range_edit.setFixedHeight(28)
        slide_range_layout.addWidget(self.slide_range_edit)
        slide_range_layout.addStretch()
        layout.addRow("幻灯片范围:", slide_range_layout)
        
        # 高级选项行 - 包含隐藏幻灯片
        advanced_options_layout = QHBoxLayout()
        advanced_options_layout.setSpacing(12)
        
        advanced_label = QLabel("高级选项:")
        advanced_options_layout.addWidget(advanced_label)
        
        self.include_hidden_checkbox = QCheckBox("包含隐藏的幻灯片")
        advanced_options_layout.addWidget(self.include_hidden_checkbox)
        advanced_options_layout.addStretch()
        layout.addRow(advanced_options_layout)
        
        # 处理完成后操作
        self.after_completion_combo = QComboBox()
        self.after_completion_combo.addItem("无操作", "none")
        self.after_completion_combo.addItem("打开输出文件", "open_file")
        self.after_completion_combo.addItem("打开输出目录", "open_folder")
        self.after_completion_combo.addItem("显示通知", "notify")
        layout.addRow("完成后操作:", self.after_completion_combo)
        
        return tab
    
    def create_image_settings_tab(self) -> QWidget:
        """创建图片设置标签页 - 集成智能处理模式"""
        tab = QWidget()
        layout = QFormLayout(tab)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        
        # ========== 智能处理模式开关 ==========
        self.smart_mode_checkbox = QCheckBox("启用智能处理模式")
        self.smart_mode_checkbox.setChecked(False)
        self.smart_mode_checkbox.stateChanged.connect(self.on_smart_mode_changed)
        layout.addRow(self.smart_mode_checkbox)
        
        # 智能模式控制选项（初始隐藏）
        self.smart_options_widget = QWidget()
        smart_options_layout = QFormLayout(self.smart_options_widget)
        smart_options_layout.setContentsMargins(0, 0, 0, 0)
        smart_options_layout.setSpacing(10)
        smart_options_layout.setLabelAlignment(Qt.AlignLeft)
        smart_options_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        
        # 目标文件大小
        target_size_widget = QWidget()
        target_size_layout = QHBoxLayout(target_size_widget)
        target_size_layout.setContentsMargins(0, 0, 0, 0)
        target_size_layout.setSpacing(8)
        
        self.target_size_spinbox = QSpinBox()
        self.target_size_spinbox.setRange(1, 2048)  # 1MB - 2048MB
        self.target_size_spinbox.setValue(50)
        self.target_size_spinbox.setFixedWidth(80)
        target_size_layout.addWidget(self.target_size_spinbox)
        
        target_size_unit = QLabel("MB")
        target_size_layout.addWidget(target_size_unit)
        target_size_layout.addStretch()
        
        smart_options_layout.addRow("目标文件大小:", target_size_widget)
        
        # 处理算法选择
        self.algorithm_combo = QComboBox()
        self.algorithm_combo.addItem("平均配额算法", "v4")
        self.algorithm_combo.addItem("双轮优化算法", "v5")
        self.algorithm_combo.addItem("迭代优化算法", "v6")
        self.algorithm_combo.addItem("预算驱动画质算法（V7）", "v7")
        self.algorithm_combo.addItem("联合优化算法（V8）", "v8")
        self.algorithm_combo.setToolTip(
            "平均配额算法: 每页使用相同配额，快速完成\n"
            "双轮优化算法: 第一轮平均配额，第二轮根据实际结果调整配额\n"
            "迭代优化算法: 根据内容复杂度分配配额，画质更均衡\n"
            "预算驱动画质算法(V7): 二点建模+预算水位分配+局部精修，在限定体积内优先提升复杂页清晰度\n"
            "联合优化算法(V8): 每页联合优化格式(PNG/JPEG)+高度+JPEG质量，提升有限体积下可读性"
        )
        smart_options_layout.addRow("处理算法:", self.algorithm_combo)
        
        self.smart_options_widget.setVisible(False)
        layout.addRow(self.smart_options_widget)
        
        # ========== 手动参数设置区域 ==========
        self.manual_settings_group = QWidget()
        manual_layout = QFormLayout(self.manual_settings_group)
        manual_layout.setContentsMargins(0, 0, 0, 0)
        manual_layout.setSpacing(12)
        manual_layout.setLabelAlignment(Qt.AlignLeft)
        manual_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        
        # 图片格式 - 添加切换联动
        self.image_format_combo = QComboBox()
        self.image_format_combo.addItem("PNG (无损压缩)", "png")
        self.image_format_combo.addItem("JPEG (有损压缩)", "jpg")
        self.image_format_combo.addItem("TIFF (LZW压缩)", "tiff")
        self.image_format_combo.addItem("BMP (无压缩)", "bmp")
        self.image_format_combo.addItem("WebP (现代格式)", "webp")
        self.image_format_combo.currentIndexChanged.connect(self.on_image_format_changed)
        manual_layout.addRow("图片格式:", self.image_format_combo)

        # JPEG质量容器 - 仅适用于JPEG/WebP等有损格式
        self.quality_container = QWidget()
        quality_container_layout = QHBoxLayout(self.quality_container)
        quality_container_layout.setContentsMargins(0, 0, 0, 0)
        quality_container_layout.setSpacing(12)
        
        self.quality_label_text = QLabel("JPEG质量:")
        quality_container_layout.addWidget(self.quality_label_text)

        self.quality_slider = QSlider(Qt.Horizontal)
        self.quality_slider.setRange(50, 100)
        self.quality_slider.setValue(95)
        self.quality_slider.setFixedHeight(20)
        self.quality_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 6px;
                background: #e5e7eb;
                border-radius: 3px;
            }
            QSlider::sub-page:horizontal {
                background: #3b82f6;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                width: 16px;
                height: 16px;
                margin: -5px 0;
                background: white;
                border: 2px solid #3b82f6;
                border-radius: 8px;
            }
            QSlider::handle:horizontal:hover {
                background: #eff6ff;
            }
        """)
        self.quality_slider.valueChanged.connect(self.on_quality_changed)

        self.quality_label = QLabel("95%")
        self.quality_label.setStyleSheet("font-weight: 500; color: #2563eb; min-width: 40px;")
        self.quality_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        quality_container_layout.addWidget(self.quality_slider, stretch=1)
        quality_container_layout.addWidget(self.quality_label)
        manual_layout.addRow(self.quality_container)
        
        # 输出清晰度 (DPI)
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItem("屏幕 (72 DPI)", 72)
        self.dpi_combo.addItem("普通 (150 DPI)", 150)
        self.dpi_combo.addItem("高清 (200 DPI)", 200)
        self.dpi_combo.addItem("打印 (300 DPI)", 300)
        self.dpi_combo.addItem("超高清 (600 DPI)", 600)
        self.dpi_combo.setCurrentIndex(3)
        manual_layout.addRow("输出清晰度:", self.dpi_combo)
        
        # PNG选项容器 - 包含label和选项，整体隐藏
        self.png_options_row = QWidget()
        png_row_layout = QHBoxLayout(self.png_options_row)
        png_row_layout.setContentsMargins(0, 0, 0, 0)
        png_row_layout.setSpacing(12)
        
        png_label = QLabel("PNG选项:")
        png_row_layout.addWidget(png_label)
        
        self.png_optimize_checkbox = QCheckBox("优化PNG大小")
        self.png_optimize_checkbox.setChecked(True)
        png_row_layout.addWidget(self.png_optimize_checkbox)
        
        self.transparent_bg_checkbox = QCheckBox("透明PNG背景")
        self.transparent_bg_checkbox.toggled.connect(self.on_transparent_bg_changed)
        png_row_layout.addWidget(self.transparent_bg_checkbox)
        
        png_row_layout.addStretch()
        manual_layout.addRow(self.png_options_row)
        
        # 图片高度 - 使用滑块 (100-8640, 16K标准)
        height_widget = QWidget()
        height_layout = QHBoxLayout(height_widget)
        height_layout.setContentsMargins(0, 0, 0, 0)
        height_layout.setSpacing(12)

        # 滑块 (100-8640)
        self.image_height_slider = QSlider(Qt.Horizontal)
        self.image_height_slider.setRange(100, 8640)
        self.image_height_slider.setValue(1080)
        self.image_height_slider.setFixedHeight(24)
        self.image_height_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 8px;
                background: #e5e7eb;
                border-radius: 4px;
            }
            QSlider::sub-page:horizontal {
                background: #3b82f6;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                width: 20px;
                height: 20px;
                margin: -6px 0;
                background: white;
                border: 2px solid #3b82f6;
                border-radius: 10px;
            }
            QSlider::handle:horizontal:hover {
                background: #eff6ff;
            }
        """)

        height_layout.addWidget(self.image_height_slider, stretch=1)

        # 显示当前数值的标签
        self.image_height_label = QLabel("1080 px")
        self.image_height_label.setStyleSheet("""
            QLabel {
                color: #374151;
                font-size: 13px;
                font-weight: 500;
                min-width: 60px;
            }
        """)
        height_layout.addWidget(self.image_height_label)

        # 连接滑块值变化到标签更新
        self.image_height_slider.valueChanged.connect(self.on_height_slider_changed)

        manual_layout.addRow("图片高度:", height_widget)
        
        # 预设按钮行
        preset_widget = QWidget()
        preset_layout = QHBoxLayout(preset_widget)
        preset_layout.setContentsMargins(0, 0, 0, 0)
        preset_layout.setSpacing(8)
        
        preset_label = QLabel("快速选择:")
        preset_label.setStyleSheet("color: #6b7280; font-size: 12px;")
        preset_layout.addWidget(preset_label)
        
        presets = [
            ("720p", 720),
            ("1080p", 1080),
            ("2K", 1440),
            ("4K", 2160),
            ("8K", 4320),
            ("16K", 8640),
        ]

        for name, value in presets:
            btn = QPushButton(name)
            btn.setProperty("preset_value", value)
            btn.setFixedWidth(50)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #f3f4f6;
                    border: 1px solid #d1d5db;
                    border-radius: 4px;
                    padding: 4px 8px;
                    font-size: 11px;
                    color: #374151;
                }
                QPushButton:hover {
                    background-color: #e5e7eb;
                    border-color: #9ca3af;
                }
                QPushButton:pressed {
                    background-color: #d1d5db;
                }
            """)
            btn.clicked.connect(lambda checked, v=value: self.set_image_height(v))
            preset_layout.addWidget(btn)
        
        preset_layout.addStretch()
        manual_layout.addRow("", preset_widget)
        
        # 将手动设置区域添加到主布局
        layout.addRow(self.manual_settings_group)
        
        # 初始化图片格式相关的显示状态
        self.on_image_format_changed(self.image_format_combo.currentIndex())
        
        return tab
    
    def on_height_slider_changed(self, value):
        """滑块值改变时更新标签显示"""
        self.image_height_label.setText(f"{value} px")

    def set_image_height(self, value):
        """设置图片高度（通过预设按钮）"""
        self.image_height_slider.setValue(value)
        self.image_height_label.setText(f"{value} px")

    def on_smart_mode_changed(self, state):
        """智能模式开关状态改变时更新UI"""
        is_smart_mode = (state == Qt.Checked)
        
        # 显示/隐藏智能选项
        self.smart_options_widget.setVisible(is_smart_mode)
        
        # 显示/隐藏手动设置区域
        self.manual_settings_group.setVisible(not is_smart_mode)
        
        # 智能模式下隐藏DPI设置（因为DPI不影响智能处理的文件大小）
        if hasattr(self, 'dpi_combo'):
            self.dpi_combo.setEnabled(not is_smart_mode)
        
        if is_smart_mode:
            self.logger.info("智能处理模式已启用")
        else:
            self.logger.info("智能处理模式已禁用，使用手动参数")

    def on_transparent_bg_changed(self, checked):
        """透明背景选项改变时，自动切换到PNG格式"""
        if checked:
            # 检查当前格式是否为PNG
            current_format = self.image_format_combo.currentData()
            if current_format != "png":
                # 自动切换到PNG格式
                self.image_format_combo.setCurrentIndex(0)  # PNG是第一个选项
                self.logger.info("透明背景需要PNG格式，已自动切换")
                # 显示提示
                QMessageBox.information(
                    self, 
                    "格式自动切换", 
                    "透明背景功能需要PNG格式支持，已自动为您切换。"
                )
    
    def on_image_format_changed(self, index):
        """图片格式改变时更新可用选项"""
        format_data = self.image_format_combo.currentData()
        
        is_jpeg = (format_data == "jpg")
        is_png = (format_data == "png")
        is_webp = (format_data == "webp")

        # JPEG质量容器 - JPEG和WebP格式显示
        if hasattr(self, 'quality_container'):
            self.quality_container.setVisible(is_jpeg or is_webp)
            # 更新标签文本
            if is_webp:
                self.quality_label_text.setText("WebP质量:")
            else:
                self.quality_label_text.setText("JPEG质量:")
        if hasattr(self, 'quality_label'):
            self.quality_label.setText(f"{self.quality_slider.value()}%")

        # PNG选项行 - 仅PNG格式显示（包含label和选项）
        if hasattr(self, 'png_options_row'):
            self.png_options_row.setVisible(is_png)

        # 如果当前不是PNG但勾选了透明背景，取消勾选
        if not is_png and hasattr(self, 'transparent_bg_checkbox'):
            if self.transparent_bg_checkbox.isChecked():
                self.transparent_bg_checkbox.setChecked(False)
    
    def on_format_changed(self, index):
        """输出格式改变时更新可用选项"""
        format_data = self.output_format_combo.currentData()
        # PDF格式时禁用图片相关选项和转换模式选择
        is_pdf = (format_data == "pdf")
        
        # 禁用/启用转换模式选择
        self.mode_combo.setEnabled(not is_pdf)
        if is_pdf:
            self.mode_combo.setToolTip("PDF模式使用特殊的导出逻辑，不需要选择转换模式")
        else:
            self.mode_combo.setToolTip("背景填充：将图片设为幻灯片背景\n前景图片：将图片作为对象插入\n图片序列：导出为独立图片文件")
        
        # 禁用/启用图片相关选项
        self.image_format_combo.setEnabled(not is_pdf)
        self.dpi_combo.setEnabled(not is_pdf)
        self.image_height_slider.setEnabled(not is_pdf)
        self.image_height_label.setEnabled(not is_pdf)

        # 根据当前图片格式更新选项状态
        if not is_pdf:
            self.on_image_format_changed(self.image_format_combo.currentIndex())
    
    def on_slide_range_changed(self, index):
        """幻灯片范围选择改变 - 控制输入框显示/隐藏"""
        is_custom = (self.slide_range_combo.currentData() == "custom")
        self.slide_range_edit.setVisible(is_custom)
    
    def create_step3(self) -> QWidget:
        """创建第三步：开始转换 - 重新设计版"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(12)

        # 顶部区域：标题 + 进度概览 - 移除边框
        header_card = QWidget()
        header_card.setStyleSheet("""
            QWidget {
                background-color: white;
                border: none;
                border-radius: 8px;
            }
        """)
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(16, 12, 16, 12)
        header_layout.setSpacing(16)

        # 左侧：状态图标和文本
        status_layout = QHBoxLayout()
        status_layout.setSpacing(12)

        self.status_icon = QLabel("⏳")
        self.status_icon.setStyleSheet("font-size: 36px;")
        status_layout.addWidget(self.status_icon)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        self.status_text = QLabel("正在转换...")
        self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #1a1a2e;")
        text_layout.addWidget(self.status_text)

        self.progress_details = QLabel("准备中...")
        self.progress_details.setStyleSheet("color: #718096; font-size: 12px;")
        text_layout.addWidget(self.progress_details)

        status_layout.addLayout(text_layout)
        status_layout.addStretch()
        header_layout.addLayout(status_layout, stretch=1)

        # 右侧：控制按钮（暂停、终止、完成）
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        # 暂停按钮
        self.pause_btn = QPushButton("⏸️ 暂停")
        self.pause_btn.setVisible(True)
        self.pause_btn.setStyleSheet("""
            QPushButton {
                background-color: #f59e0b;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: 600;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #d97706;
            }
            QPushButton:pressed {
                background-color: #b45309;
            }
        """)
        self.pause_btn.clicked.connect(self.toggle_pause_conversion)
        btn_layout.addWidget(self.pause_btn)

        # 终止按钮
        self.stop_btn = QPushButton("⏹️ 终止")
        self.stop_btn.setVisible(True)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #ef4444;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: 600;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #dc2626;
            }
            QPushButton:pressed {
                background-color: #b91c1c;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_conversion)
        btn_layout.addWidget(self.stop_btn)

        # 完成/重试按钮（转换完成后显示）
        self.finish_btn = QPushButton("✓ 完成")
        self.finish_btn.setVisible(False)
        self.finish_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #10b981, stop:1 #059669);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 13px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #059669, stop:1 #047857);
            }
        """)
        self.finish_btn.clicked.connect(self.finish_conversion)
        btn_layout.addWidget(self.finish_btn)

        # 返回按钮（转换失败时显示）
        self.back_btn = QPushButton("👈 返回")
        self.back_btn.setVisible(False)
        self.back_btn.setStyleSheet("""
            QPushButton {
                background-color: #6b7280;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 13px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #4b5563;
            }
            QPushButton:pressed {
                background-color: #374151;
            }
        """)
        self.back_btn.clicked.connect(self.go_back_from_conversion)
        btn_layout.addWidget(self.back_btn)

        header_layout.addLayout(btn_layout)
        layout.addWidget(header_card)

        # 进度条区域 - 移除边框
        progress_widget = QWidget()
        progress_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border: none;
                border-radius: 8px;
            }
        """)
        progress_layout = QVBoxLayout(progress_widget)
        progress_layout.setContentsMargins(16, 12, 16, 12)
        progress_layout.setSpacing(8)

        # 进度条标题和百分比
        progress_header = QHBoxLayout()
        progress_label = QLabel("转换进度")
        progress_label.setStyleSheet("font-size: 13px; font-weight: 500; color: #374151;")
        progress_header.addWidget(progress_label)

        self.progress_percent = QLabel("0%")
        self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #667eea;")
        progress_header.addWidget(self.progress_percent)
        progress_header.addStretch()
        progress_layout.addLayout(progress_header)

        # 进度条 - 使用更细的样式
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 0.1px;
                background-color: #e5e7eb;
                height: 8px;
            }
            QProgressBar::chunk {
                background-color: #667eea;
                border-radius: 0.1px;
            }
        """)
        progress_layout.addWidget(self.progress_bar)

        layout.addWidget(progress_widget)

        # 文件列表进度表格（批量模式显示）
        self.file_progress_table = QTableWidget()
        self.file_progress_table.setVisible(False)
        self.file_progress_table.setColumnCount(7)
        self.file_progress_table.setHorizontalHeaderLabels([
            "文件名", "状态", "进度", "原始大小", "处理后大小", "误差", "导出位置"
        ])
        self.file_progress_table.horizontalHeader().setStretchLastSection(True)
        self.file_progress_table.verticalHeader().setVisible(False)
        self.file_progress_table.setAlternatingRowColors(True)
        self.file_progress_table.setMaximumHeight(120)  # 只显示3行的高度
        self.file_progress_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                gridline-color: #f3f4f6;
            }
            QTableWidget::item {
                padding: 2px;
                font-size: 10px;
            }
            QHeaderView::section {
                background-color: #f9fafb;
                padding: 4px;
                border: none;
                border-bottom: 1px solid #e5e7eb;
                font-size: 10px;
                font-weight: 600;
                color: #374151;
            }
            QTableWidget::item:selected {
                background-color: #667eea;
                color: white;
            }
        """)
        # 设置列宽（新顺序：文件名、状态、进度、原始大小、处理后大小、误差、导出位置）
        self.file_progress_table.setColumnWidth(0, 150)  # 文件名
        self.file_progress_table.setColumnWidth(1, 50)   # 状态
        self.file_progress_table.setColumnWidth(2, 50)   # 进度
        self.file_progress_table.setColumnWidth(3, 70)   # 原始大小
        self.file_progress_table.setColumnWidth(4, 80)   # 处理后大小
        self.file_progress_table.setColumnWidth(5, 50)   # 误差（改窄）
        # 导出位置列自动拉伸
        layout.addWidget(self.file_progress_table)

        # 日志区域（增大高度）- 移除边框
        log_widget = QWidget()
        log_widget.setStyleSheet("""
            QWidget {
                background-color: #0f172a;
                border: none;
                border-radius: 8px;
            }
        """)
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(12, 10, 12, 10)
        log_layout.setSpacing(6)

        # 日志标题栏
        log_header_layout = QHBoxLayout()
        log_header = QLabel("📋 转换日志")
        log_header.setStyleSheet("color: #94a3b8; font-size: 11px; font-weight: 500;")
        log_header_layout.addWidget(log_header)
        log_header_layout.addStretch()

        # 清空日志按钮
        clear_btn = QPushButton("清空")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #64748b;
                border: 1px solid #334155;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 10px;
            }
            QPushButton:hover {
                color: #94a3b8;
                border-color: #475569;
            }
        """)
        clear_btn.clicked.connect(lambda: self.log_text.clear())
        log_header_layout.addWidget(clear_btn)
        log_layout.addLayout(log_header_layout)

        # 日志文本（减小高度，给进度表格留空间）
        self.log_text = QTextEdit()
        self.log_text.setMinimumHeight(120)
        self.log_text.setMaximumHeight(180)
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: transparent;
                border: none;
                color: #e2e8f0;
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 12px;
                line-height: 1.5;
            }
        """)
        log_layout.addWidget(self.log_text)

        layout.addWidget(log_widget, stretch=1)

        return widget
    
    def add_file(self, file_path: str):
        """添加文件"""
        # 检查是否已存在
        if file_path not in [str(f) for f in self.files]:
            self.files.append(Path(file_path))
            
            # 添加到列表
            item = QListWidgetItem(f"📄 {os.path.basename(file_path)}")
            self.file_list.addItem(item)
        
        # 自动设置输出目录
        self.current_output_dir = os.path.dirname(file_path)
        self.output_dir_edit.setText(self.current_output_dir)
        
        # 启用下一步按钮和删除按钮
        self.step1_next_btn.setEnabled(True)
        self.delete_file_btn.setEnabled(True)
        
        # 设置当前输入文件（第一个文件）
        if len(self.files) > 0:
            self.current_input_file = str(self.files[0])
        
        self.logger.info(f"已添加文件: {file_path} (共 {len(self.files)} 个文件)")
        self.status_bar.showMessage(f"已添加: {os.path.basename(file_path)} (共 {len(self.files)} 个文件)")
    
    def add_folder(self, folder_path: str):
        """添加文件夹中的所有PPT文件"""
        import glob
        
        # 查找文件夹中的所有PPT文件
        ppt_files = glob.glob(os.path.join(folder_path, "*.ppt")) + glob.glob(os.path.join(folder_path, "*.pptx"))
        
        if not ppt_files:
            QMessageBox.warning(self, "警告", f"文件夹中没有找到PPT文件！")
            return
        
        # 添加所有文件
        added_count = 0
        for file_path in ppt_files:
            if file_path not in [str(f) for f in self.files]:
                self.files.append(Path(file_path))
                
                # 添加到列表
                item = QListWidgetItem(f"📄 {os.path.basename(file_path)}")
                self.file_list.addItem(item)
                added_count += 1
        
        # 自动设置输出目录
        self.current_output_dir = folder_path
        self.output_dir_edit.setText(self.current_output_dir)
        
        # 启用下一步按钮和删除按钮
        self.step1_next_btn.setEnabled(True)
        self.delete_file_btn.setEnabled(True)
        
        # 设置当前输入文件（第一个文件）
        if len(self.files) > 0:
            self.current_input_file = str(self.files[0])
        
        self.logger.info(f"已从文件夹添加 {added_count} 个文件 (共 {len(self.files)} 个文件)")
        self.status_bar.showMessage(f"已从文件夹添加 {added_count} 个文件 (共 {len(self.files)} 个文件)")
    
    def delete_file(self):
        """删除已选择的文件"""
        current_item = self.file_list.currentItem()
        if current_item:
            row = self.file_list.row(current_item)
            if 0 <= row < len(self.files):
                self.files.pop(row)
                self.file_list.takeItem(row)
                
                if len(self.files) == 0:
                    self.step1_next_btn.setEnabled(False)
                    self.delete_file_btn.setEnabled(False)
                    self.current_input_file = None
                else:
                    # 更新当前输入文件
                    self.current_input_file = str(self.files[0])
                
                self.logger.info(f"已删除文件 (剩余 {len(self.files)} 个文件)")
                self.status_bar.showMessage(f"已删除文件 (剩余 {len(self.files)} 个文件)")
    
    def select_output_dir(self):
        """选择输出目录"""
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if output_dir:
            self.current_output_dir = output_dir
            self.output_dir_edit.setText(output_dir)
            self.logger.info(f"选择了输出目录: {output_dir}")
            self.status_bar.showMessage(f"输出目录: {output_dir}")
    
    def on_file_item_double_clicked(self, item):
        """双击文件列表项时添加更多文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择 PowerPoint 文件", "",
            "PowerPoint 文件 (*.ppt *.pptx)"
        )
        if file_paths:
            for file_path in file_paths:
                self.add_file(file_path)
    
    def on_quality_changed(self):
        """质量滑块值改变"""
        quality = self.quality_slider.value()
        self.quality_label.setText(f"{quality}%")
    
    def go_to_step2(self):
        """进入第二步"""
        self.content_stack.setCurrentIndex(1)
        self.step_indicator.set_current_step(1)
        self.status_bar.showMessage("设置转换参数")
        
        # 设置智能配置组件的PPTX路径
        if hasattr(self, 'smart_config_tab') and self.current_input_file:
            self.smart_config_tab.set_pptx_path(self.current_input_file)
    
    def go_to_step3(self):
        """进入第三步并开始转换"""
        try:
            self.logger.info("进入第三步，开始转换流程")
            
            # 重置转换成功标志
            self._conversion_success = False
            
            # 检查是否有文件
            if len(self.files) == 0:
                QMessageBox.warning(self, "警告", "请先选择要转换的文件！")
                return
            
            # 判断是否是批处理模式（多个文件）
            is_batch_mode = len(self.files) > 1
            
            if is_batch_mode:
                # 批处理模式：启动批处理流程
                self._start_batch_conversion()
                return
            
            # 单文件模式
            if not self.current_input_file:
                QMessageBox.warning(self, "警告", "请先选择输入文件！")
                return
            
            # 检查是否启用智能处理模式
            if hasattr(self, 'smart_mode_checkbox') and self.smart_mode_checkbox.isChecked():
                # 启动智能处理流程
                self._start_smart_optimization()
                return
            
            # 确定输出目录
            output_dir = self.current_output_dir or os.path.dirname(self.current_input_file)
            self.logger.info(f"输出目录: {output_dir}")
            self.logger.info(f"current_output_dir: {self.current_output_dir}")
            self.logger.info(f"current_input_file: {self.current_input_file}")
            
            # 确保输出目录存在
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                self.logger.info(f"创建输出目录: {output_dir}")
            
            # 确定输出文件名
            input_name = os.path.splitext(os.path.basename(self.current_input_file))[0]
            custom_name = self.filename_edit.text().strip()
            if custom_name:
                output_name = custom_name
            else:
                output_name = f"{input_name}_converted"
            
            # 设置转换模式
            mode_data = self.mode_combo.currentData()
            output_ext = self.output_format_combo.currentData() or "pptx"
            options = ConversionOptions()
            
            # PDF模式特殊处理：不使用转换模式，使用图片序列模式
            if output_ext == "pdf":
                options.mode = ConversionMode.SLIDE_TO_IMAGE  # PDF模式内部使用图片序列模式
                options.output_format = OutputFormat.PDF
            else:
                options.mode = {
                    "background_fill": ConversionMode.BACKGROUND_FILL,
                    "foreground_image": ConversionMode.FOREGROUND_IMAGE,
                    "slide_to_image": ConversionMode.SLIDE_TO_IMAGE
                }[mode_data]
            
            # 图片序列模式特殊处理：输出路径是文件夹
            if options.mode == ConversionMode.SLIDE_TO_IMAGE and output_ext != "pdf":
                output_path = os.path.join(output_dir, output_name)
                # 确保文件夹名称不冲突
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(output_dir, f"{output_name}_{counter}")
                    counter += 1
                
                # 图片序列模式不使用output_format，而是使用image_export.format
                options.output_format = OutputFormat.JPG  # 默认值，实际不使用
            else:
                # 其他模式：输出路径是文件
                # 使用固定命名，避免递增数字
                if hasattr(self, '_smart_optimization_target_size') and self._smart_optimization_target_size is not None:
                    # 智能模式：使用固定命名
                    output_path = os.path.join(output_dir, f"{output_name}_optimized.{output_ext}")
                else:
                    # 普通模式：检查文件是否存在
                    output_path = os.path.join(output_dir, f"{output_name}.{output_ext}")
                    
                    # 检查文件是否存在
                    counter = 1
                    while os.path.exists(output_path):
                        output_path = os.path.join(output_dir, f"{output_name}_{counter}.{output_ext}")
                        counter += 1
                
                options.output_format = {
                    "pptx": OutputFormat.PPTX,
                    "pdf": OutputFormat.PDF,
                }[output_ext]
            
            options.image_quality = self.quality_slider.value()
            options.include_hidden_slides = self.include_hidden_checkbox.isChecked()
        
            # 设置图片导出配置
            try:
                # 图片格式
                img_format_data = self.image_format_combo.currentData()
                if img_format_data:
                    from src.core.converter import ImageFormat
                    options.image_export.format = ImageFormat(img_format_data)
                
                # DPI设置
                dpi_value = self.dpi_combo.currentData()
                if dpi_value:
                    from src.core.converter import DPIPreset
                    options.image_export.dpi_preset = DPIPreset(dpi_value)
                
                # 自定义分辨率
                if hasattr(self, 'custom_resolution_checkbox') and self.custom_resolution_checkbox.isChecked():
                    options.image_export.use_custom_resolution = True
                    try:
                        options.image_export.custom_width = int(self.resolution_width.text())
                        options.image_export.custom_height = int(self.resolution_height.text())
                    except ValueError:
                        pass
                
                # PNG优化选项（仅PNG格式）
                if hasattr(self, 'png_optimize_checkbox'):
                    options.image_export.optimize = self.png_optimize_checkbox.isChecked()
                
                # 其他选项 - 宽高比默认保持
                options.image_export.maintain_aspect_ratio = True
                if hasattr(self, 'transparent_bg_checkbox'):
                    options.image_export.transparent_background = self.transparent_bg_checkbox.isChecked()
                
                # 幻灯片范围
                if hasattr(self, 'slide_range_combo'):
                    range_data = self.slide_range_combo.currentData()
                    if range_data == "custom" and hasattr(self, 'slide_range_edit'):
                        # 解析范围字符串，如 "1-5,8,10-12"
                        range_str = self.slide_range_edit.text().strip()
                        if range_str:
                            try:
                                start, end = self._parse_slide_range(range_str)
                                options.export_range = (start, end)
                            except ValueError:
                                pass
                
                # 图片高度配置（关键：控制输出图片尺寸）
                if hasattr(self, 'image_height_slider'):
                    target_height = self.image_height_slider.value()
                    options.image_export.use_custom_resolution = True
                    options.image_export.custom_height = target_height
                    # 宽度将根据原始PPT的宽高比自动计算
                    options.image_export.custom_width = 0  # 0表示自动计算
                    options.image_export.maintain_aspect_ratio = True
                    self.logger.info(f"设置图片高度: {target_height}px")
                
            except Exception as e:
                self.logger.warning(f"设置高级选项时出错: {e}")
            
            # 切换到第三步
            self.content_stack.setCurrentIndex(2)
            self.step_indicator.set_current_step(2)
            self.status_bar.showMessage("正在转换...")

            # 重置进度显示
            self.status_icon.setText("⏳")
            self.status_icon.setStyleSheet("font-size: 36px;")
            self.status_text.setText("正在转换...")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #1a1a2e;")
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_percent.setText("0%")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #667eea;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 8px;
                }
                QProgressBar::chunk {
                    background-color: #667eea;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("准备中...")
            self.log_text.clear()

            # 显示暂停和终止按钮，隐藏完成按钮和返回按钮
            self.pause_btn.setVisible(True)
            self.pause_btn.setText("⏸️ 暂停")
            self.pause_btn.setEnabled(True)
            self.stop_btn.setVisible(True)
            self.stop_btn.setEnabled(True)
            self.finish_btn.setVisible(False)
            self.back_btn.setVisible(False)

            # 清理之前的worker
            if hasattr(self, 'worker') and self.worker is not None:
                if self.worker.isRunning():
                    self.worker.quit()
                    self.worker.wait(1000)
                self.worker.deleteLater()
                self.worker = None
            
            # 记录日志
            self.logger.info(f"开始转换任务")
            self.logger.info(f"输入: {self.current_input_file}")
            self.logger.info(f"输出: {output_path}")
            self.log_message(f"🚀 开始转换任务")
            self.log_message(f"📁 输入文件: {os.path.basename(self.current_input_file)}")
            self.log_message(f"💾 输出文件: {os.path.basename(output_path)}")
            self.log_message(f"📊 输出格式: {self.output_format_combo.currentText()}")
            self.log_message(f"🎨 转换模式: {self.mode_combo.currentText()}")
            self.log_message(f"🖼️ 图片质量: {self.quality_slider.value()}%")
            
            # 创建工作线程（converter在线程中创建，避免COM线程问题）
            self.worker = ConversionWorker(
                self.current_input_file, 
                output_path, 
                options
            )
            self.worker.progress_updated.connect(self._update_progress)
            self.worker.conversion_finished.connect(self._on_conversion_finished)
            self.worker.start()
        
        except Exception as e:
            self.logger.error(f"进入第三步时发生异常: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            QMessageBox.critical(self, "错误", f"启动转换失败: {e}")
            self.go_to_step2()
    
    def go_to_step1(self):
        """返回第一步"""
        self.content_stack.setCurrentIndex(0)
        self.step_indicator.set_current_step(0)
        self.status_bar.showMessage("选择要转换的文件")
        
        # 清理转换器状态
        if hasattr(self, 'converter') and self.converter:
            try:
                self.converter._cleanup(force_kill=False)
            except:
                pass
        
        # 重新初始化转换器
        self.converter = Win32PPTConverter()
        
        # 清理worker
        if hasattr(self, 'worker') and self.worker is not None:
            if self.worker.isRunning():
                self.worker.quit()
                self.worker.wait(1000)
            self.worker.deleteLater()
            self.worker = None
        
        # 清理batch_worker
        if hasattr(self, 'batch_worker') and self.batch_worker is not None:
            if self.batch_worker.isRunning():
                self.batch_worker.quit()
                self.batch_worker.wait(1000)
            self.batch_worker.deleteLater()
            self.batch_worker = None
        
        # 重置UI状态
        self._conversion_success = False
        self._last_output_path = None
        
        # 清理进度表格
        if hasattr(self, 'file_progress_table'):
            self.file_progress_table.setRowCount(0)
            self.file_progress_table.setVisible(False)
        
        # 清理日志
        if hasattr(self, 'log_text'):
            self.log_text.clear()
        
        # 重置进度条
        if hasattr(self, 'progress_bar'):
            self.progress_bar.setValue(0)
        if hasattr(self, 'progress_percent'):
            self.progress_percent.setText("0%")
        if hasattr(self, 'progress_details'):
            self.progress_details.setText("")
        
        # 隐藏暂停和终止按钮
        if hasattr(self, 'pause_btn'):
            self.pause_btn.setVisible(False)
        if hasattr(self, 'stop_btn'):
            self.stop_btn.setVisible(False)
        if hasattr(self, 'finish_btn'):
            self.finish_btn.setVisible(False)
    
    def _update_progress(self, value: float, task: str):
        """更新进度 - 显示百分比"""
        progress_int = int(value * 100)
        self.progress_bar.setValue(progress_int)
        self.progress_percent.setText(f"{progress_int}%")
        self.progress_details.setText(task)
        self.log_message(f"⏳ {task}")
    
    def _on_conversion_finished(self, success: bool, output_path: str):
        """转换完成 - 适配新布局"""
        # 记录转换状态
        self._conversion_success = success
        self._last_output_path = output_path
        
        self.logger.info(f"[Main] 转换完成回调: success={success}, output_path={output_path}")

        # 隐藏暂停和终止按钮
        self.pause_btn.setVisible(False)
        self.stop_btn.setVisible(False)

        # 检查是否是用户终止
        if not success and output_path == "用户终止":
            self.log_message("⏹️ 已终止转换")
            self.status_icon.setText("⏹️")
            self.status_icon.setStyleSheet("font-size: 36px;")
            self.status_text.setText("已终止")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #ef4444;")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #ef4444;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 10px;
                }
                QProgressBar::chunk {
                    background-color: #ef4444;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("转换已终止")
            self.status_bar.showMessage("转换已终止")
            
            # 显示重试和返回按钮
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(True)
            self.finish_btn.setText("🔄 重试")
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #f59e0b, stop:1 #d97706);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #d97706, stop:1 #b45309);
                }
            """)
            return

        if success:
            # 获取文件大小
            size_mb = 0
            if os.path.exists(output_path):
                # 判断是文件夹还是文件
                if os.path.isdir(output_path):
                    # 图片序列模式：计算文件夹总大小
                    total_size = 0
                    file_count = 0
                    for dirpath, dirnames, filenames in os.walk(output_path):
                        for f in filenames:
                            fp = os.path.join(dirpath, f)
                            total_size += os.path.getsize(fp)
                            file_count += 1
                    size_mb = total_size / 1024 / 1024
                    self.log_message(f"🎉 转换成功！")
                    self.log_message(f"📊 文件夹大小: {size_mb:.2f} MB (共 {file_count} 个文件)")
                else:
                    # 其他模式：计算单个文件大小
                    size_mb = os.path.getsize(output_path) / 1024 / 1024
                    self.log_message(f"🎉 转换成功！")
                    self.log_message(f"📊 文件大小: {size_mb:.2f} MB")
                self.log_message(f"💾 保存位置: {output_path}")

                # 添加到历史记录
                self.history_manager.add_record(
                    input_path=self.current_input_file,
                    output_path=output_path,
                    mode=self.mode_combo.currentText(),
                    output_format=self.output_format_combo.currentData(),
                    success=True,
                    file_size=size_mb
                )

            # 成功状态（进度条已在ProgressTracker.complete中设置为100%）
            self.status_icon.setText("🥳")
            self.status_icon.setStyleSheet("font-size: 36px; color: #10b981;")
            self.status_text.setText("转换成功！")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #10b981;")
            # 不再重复设置进度条值，避免跳变
            self.progress_percent.setText("100%")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #10b981;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 8px;
                }
                QProgressBar::chunk {
                    background-color: #10b981;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("文件已保存到指定位置")

            # 成功时显示完成按钮，隐藏返回按钮
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(False)
            self.finish_btn.setText("🤗 回到开始")
            # 成功时使用绿色样式
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #10b981, stop:1 #059669);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #059669, stop:1 #047857);
                }
            """)
            
            self.status_bar.showMessage("转换完成")
            
            # 执行完成后操作
            self._execute_after_completion(output_path)
        else:
            # 失败状态 - 使用伤心emoji
            self.status_icon.setText("😢")
            self.status_icon.setStyleSheet("font-size: 36px;")
            self.status_text.setText("转换失败")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #ef4444;")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #ef4444;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 10px;
                }
                QProgressBar::chunk {
                    background-color: #ef4444;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("转换过程中发生错误，请查看日志")
            self.log_message(f"❌ 转换失败: {output_path}")
            
            # 添加失败记录到历史记录
            self.history_manager.add_record(
                input_path=self.current_input_file,
                output_path="",  # 失败时没有输出文件
                mode=self.mode_combo.currentText(),
                output_format=self.output_format_combo.currentData(),
                success=False
            )
            
            # 失败时显示重试和返回按钮
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(True)
            self.finish_btn.setText("🔄 重试")
            # 失败时使用橙色样式
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #f59e0b, stop:1 #d97706);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #d97706, stop:1 #b45309);
                }
            """)
            self.status_bar.showMessage("转换失败")

    def _execute_after_completion(self, output_path: str):
        """执行完成后操作"""
        try:
            action = self.after_completion_combo.currentData()
            
            if action == "open_file":
                # 打开输出文件
                if output_path and os.path.exists(output_path):
                    if os.path.isdir(output_path):
                        # 图片序列模式，打开目录
                        os.startfile(output_path)
                    else:
                        os.startfile(output_path)
                    self.log_message(f"📂 已打开输出文件: {output_path}")
            
            elif action == "open_folder":
                # 打开输出目录
                if output_path:
                    output_dir = os.path.dirname(output_path) if os.path.isfile(output_path) else output_path
                    if os.path.exists(output_dir):
                        os.startfile(output_dir)
                        self.log_message(f"📂 已打开输出目录: {output_dir}")
            
            elif action == "notify":
                # 显示桌面通知
                try:
                    from PyQt5.QtWidgets import QSystemTrayIcon
                    if hasattr(self, 'tray_icon') and self.tray_icon:
                        self.tray_icon.showMessage(
                            "PitchPPT 转换完成",
                            f"文件已成功保存到: {output_path}",
                            QSystemTrayIcon.Information,
                            3000
                        )
                except Exception as e:
                    self.logger.warning(f"显示通知失败: {e}")
        
        except Exception as e:
            self.logger.error(f"执行完成后操作失败: {e}")

    def finish_conversion(self):
        """完成转换或重试"""
        # 检查是否是重试操作
        if hasattr(self, '_conversion_success') and not self._conversion_success:
            # 重试转换 - 先清理之前的worker线程
            self.log_message("🔄 重新启动转换...")
            
            # 断开信号连接，防止重复调用
            try:
                if hasattr(self, 'worker') and self.worker is not None:
                    self.worker.conversion_finished.disconnect()
                if hasattr(self, 'smart_worker') and self.smart_worker is not None:
                    self.smart_worker.error_occurred.disconnect()
                if hasattr(self, 'batch_worker') and self.batch_worker is not None:
                    self.batch_worker.batch_finished.disconnect()
            except:
                pass
            
            # 清理之前的普通worker
            if hasattr(self, 'worker') and self.worker is not None:
                if self.worker.isRunning():
                    self.worker.stop()
                    if not self.worker.wait(5000):
                        self.worker.terminate()
                        self.worker.wait(1000)
                self.worker.deleteLater()
                self.worker = None
            
            # 清理之前的智能处理worker
            if hasattr(self, 'smart_worker') and self.smart_worker is not None:
                if self.smart_worker.isRunning():
                    self.smart_worker.stop()
                    if not self.smart_worker.wait(5000):
                        self.smart_worker.terminate()
                        self.smart_worker.wait(1000)
                self.smart_worker.deleteLater()
                self.smart_worker = None
            
            # 清理之前的批处理worker
            if hasattr(self, 'batch_worker') and self.batch_worker is not None:
                if self.batch_worker.isRunning():
                    self.batch_worker.stop()
                    if not self.batch_worker.wait(5000):
                        self.batch_worker.terminate()
                        self.batch_worker.wait(1000)
                self.batch_worker.deleteLater()
                self.batch_worker = None
            
            # 清理全局converter状态
            if hasattr(self, 'converter') and self.converter:
                try:
                    self.converter._cleanup(force_kill=False)
                except:
                    pass
            
            # 等待一段时间让资源释放
            import time
            time.sleep(0.5)
            
            # 重新调用go_to_step3
            self.go_to_step3()
        else:
            # 正常完成，返回第一步，但不清空文件列表
            self.go_to_step1()
            # 不清空文件列表，让用户可以继续处理已选择的文件
            # self.file_list.clear()
            # self.current_input_file = None
            # self.step1_next_btn.setEnabled(False)
            self.status_bar.showMessage("转换完成 - 可以继续处理其他文件或重新选择")

    def go_back_from_conversion(self):
        """从转换页面返回到第二步（设置页面）"""
        # 隐藏按钮
        self.finish_btn.setVisible(False)
        self.back_btn.setVisible(False)
        # 返回到第二步
        self.go_to_step2()
        self.status_bar.showMessage("返回设置页面")

    def toggle_pause_conversion(self):
        """暂停/继续转换"""
        # 支持普通模式、智能模式和批处理模式
        worker = None
        if hasattr(self, 'worker') and self.worker is not None and self.worker.isRunning():
            worker = self.worker
        elif hasattr(self, 'smart_worker') and self.smart_worker is not None and self.smart_worker.isRunning():
            worker = self.smart_worker
        elif hasattr(self, 'batch_worker') and self.batch_worker is not None and self.batch_worker.isRunning():
            worker = self.batch_worker
        
        if worker is None:
            return

        if self.pause_btn.text() == "⏸️ 暂停":
            # 暂停转换
            worker.pause()
            self.pause_btn.setText("▶️ 继续")
            self.status_icon.setText("⏸️")
            self.status_text.setText("转换已暂停")
            self.log_message("⏸️ 转换已暂停")
        else:
            # 继续转换
            worker.resume()
            self.pause_btn.setText("⏸️ 暂停")
            self.status_icon.setText("⏳")
            self.status_text.setText("正在转换...")
            self.log_message("▶️ 转换继续")

    def stop_conversion(self):
        """终止转换"""
        # 支持普通模式、智能模式和批处理模式
        worker = None
        if hasattr(self, 'worker') and self.worker is not None and self.worker.isRunning():
            worker = self.worker
        elif hasattr(self, 'smart_worker') and self.smart_worker is not None and self.smart_worker.isRunning():
            worker = self.smart_worker
        elif hasattr(self, 'batch_worker') and self.batch_worker is not None and self.batch_worker.isRunning():
            worker = self.batch_worker
        
        if worker is None:
            return

        reply = QMessageBox.question(
            self,
            "确认终止",
            "确定要终止当前转换吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.log_message("⏹️ 正在终止转换...")
            
            # 断开信号连接，防止重复调用
            try:
                if hasattr(self, 'worker') and worker == self.worker:
                    self.worker.conversion_finished.disconnect()
                elif hasattr(self, 'smart_worker') and worker == self.smart_worker:
                    self.smart_worker.error_occurred.disconnect()
                elif hasattr(self, 'batch_worker') and worker == self.batch_worker:
                    self.batch_worker.batch_finished.disconnect()
            except:
                pass
            
            worker.stop()
            
            # 等待线程完全停止
            if not worker.wait(5000):  # 等待5秒
                self.logger.warning("线程未能在超时时间内停止，强制终止")
                worker.terminate()
                worker.wait(1000)
            
            # 等待一段时间让资源释放
            import time
            time.sleep(0.5)
            
            # 根据worker类型调用不同的完成处理
            if hasattr(self, 'worker') and worker == self.worker:
                # 先清理worker
                self.worker.deleteLater()
                self.worker = None
                self._on_conversion_finished(False, "用户终止")
            elif hasattr(self, 'smart_worker') and worker == self.smart_worker:
                # 先清理worker
                self.smart_worker.deleteLater()
                self.smart_worker = None
                self._on_smart_optimization_error("用户终止")
            elif hasattr(self, 'batch_worker') and worker == self.batch_worker:
                # 先清理worker
                self.batch_worker.deleteLater()
                self.batch_worker = None
                # 批处理模式：直接调用完成处理
                self._on_batch_finished({
                    'total': len(self.files),
                    'success': 0,
                    'failed': len(self.files),
                    'results': [],
                    'error': '用户终止'
                })
    
    def _parse_slide_range(self, range_str: str) -> tuple:
        """解析幻灯片范围字符串，如 '1-5,8,10-12' -> (1, 12)"""
        import re
        numbers = []
        for part in range_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = part.split('-', 1)
                numbers.extend(range(int(start.strip()), int(end.strip()) + 1))
            else:
                numbers.append(int(part))
        if not numbers:
            raise ValueError("无效的范围字符串")
        return (min(numbers), max(numbers))

    def _start_smart_optimization(self):
        """启动智能处理流程 - 使用新的SmartRenderer系统"""
        try:
            target_size_mb = self.target_size_spinbox.value()
            
            self.logger.info(f"启动智能处理: 目标大小={target_size_mb}MB")
            
            # 确定输出路径
            output_dir = self.current_output_dir or os.path.dirname(self.current_input_file)
            input_name = os.path.splitext(os.path.basename(self.current_input_file))[0]
            custom_name = self.filename_edit.text().strip()
            if custom_name:
                output_name = custom_name
            else:
                # 使用设置中的文件名后缀
                suffix = self.filename_suffix_edit.text().strip() if hasattr(self, 'filename_suffix_edit') else "_converted"
                output_name = f"{input_name}{suffix}"
            
            output_path = os.path.join(output_dir, f"{output_name}.pptx")
            
            # 检查文件是否存在，添加序号
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(output_dir, f"{output_name}_{counter}.pptx")
                counter += 1
            
            # 切换到第三步界面
            self.content_stack.setCurrentIndex(2)
            self.step_indicator.set_current_step(2)
            self.status_bar.showMessage("正在智能处理...")
            
            # 重置进度显示
            self.status_icon.setText("🎯")
            self.status_icon.setStyleSheet("font-size: 36px;")
            self.status_text.setText("正在智能处理...")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #0369a1;")
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_percent.setText("0%")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #0369a1;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 8px;
                }
                QProgressBar::chunk {
                    background-color: #0ea5e9;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("准备智能处理...")
            self.log_text.clear()
            
            # 显示暂停和终止按钮，隐藏完成按钮和返回按钮
            self.pause_btn.setVisible(True)
            self.pause_btn.setText("⏸️ 暂停")
            self.pause_btn.setEnabled(True)
            self.stop_btn.setVisible(True)
            self.stop_btn.setEnabled(True)
            self.finish_btn.setVisible(False)
            self.back_btn.setVisible(False)
            
            # 记录日志
            self.log_message(f"🚀 启动智能处理 (新系统)")
            self.log_message(f"📁 输入文件: {os.path.basename(self.current_input_file)}")
            self.log_message(f"💾 输出文件: {os.path.basename(output_path)}")
            self.log_message(f"🎯 目标大小: {target_size_mb}MB")
            
            # 清理之前的worker
            if hasattr(self, 'smart_worker') and self.smart_worker is not None:
                if self.smart_worker.isRunning():
                    self.smart_worker.quit()
                    self.smart_worker.wait(1000)
                self.smart_worker.deleteLater()
                self.smart_worker = None
            
            # 创建新的智能处理工作线程（V4架构 - 每页独立调优）
            from src.ui.smart_optimization_worker_v4 import SmartOptimizationWorkerV4
            
            # 获取选择的算法
            algorithm = "v4"
            if hasattr(self, 'algorithm_combo'):
                algorithm = self.algorithm_combo.currentData()
            
            self.smart_worker = SmartOptimizationWorkerV4(
                self.current_input_file,
                output_path,
                target_size_mb,
                algorithm=algorithm,
                conversion_mode={
                    "background_fill": ConversionMode.BACKGROUND_FILL,
                    "foreground_image": ConversionMode.FOREGROUND_IMAGE,
                    "slide_to_image": ConversionMode.SLIDE_TO_IMAGE
                }[self.mode_combo.currentData()],
                include_hidden_slides=self.include_hidden_checkbox.isChecked(),
                logger=self.logger
            )
            self.smart_worker.progress_updated.connect(self._update_smart_progress)
            self.smart_worker.result_ready.connect(self._on_smart_optimization_finished)
            self.smart_worker.error_occurred.connect(self._on_smart_optimization_error)
            self.smart_worker.start()
            
        except Exception as e:
            self.logger.error(f"启动智能处理失败: {e}")
            QMessageBox.critical(self, "错误", f"启动智能处理失败: {e}")
            self.go_to_step2()

    def _update_smart_progress(self, message: str, progress: int):
        """更新智能处理进度"""
        self.progress_bar.setValue(progress)
        self.progress_percent.setText(f"{progress}%")
        self.progress_details.setText(message)
        self.log_message(f"⏳ {message}")

    def _start_batch_conversion(self):
        """启动批处理转换流程"""
        try:
            self.logger.info(f"启动批处理转换，共 {len(self.files)} 个文件")
            
            # 确定输出目录
            output_dir = self.current_output_dir or os.path.dirname(str(self.files[0]))
            
            # 检查是否启用智能处理模式
            is_smart_mode = hasattr(self, 'smart_mode_checkbox') and self.smart_mode_checkbox.isChecked()
            
            # 构建转换选项
            options = ConversionOptions()
            mode_data = self.mode_combo.currentData()
            options.mode = {
                "background_fill": ConversionMode.BACKGROUND_FILL,
                "foreground_image": ConversionMode.FOREGROUND_IMAGE,
                "slide_to_image": ConversionMode.SLIDE_TO_IMAGE
            }[mode_data]
            
            if not is_smart_mode:
                # 普通模式：设置转换选项
                output_ext = self.output_format_combo.currentData() or "pptx"
                
                if output_ext == "pdf":
                    options.mode = ConversionMode.SLIDE_TO_IMAGE
                    options.output_format = OutputFormat.PDF
                else:
                    options.mode = {
                        "background_fill": ConversionMode.BACKGROUND_FILL,
                        "foreground_image": ConversionMode.FOREGROUND_IMAGE,
                        "slide_to_image": ConversionMode.SLIDE_TO_IMAGE
                    }[mode_data]
                    options.output_format = {
                        "pptx": OutputFormat.PPTX,
                        "pdf": OutputFormat.PDF,
                    }[output_ext]
                
                options.image_quality = self.quality_slider.value()
                options.include_hidden_slides = self.include_hidden_checkbox.isChecked()
                
                try:
                    img_format_data = self.image_format_combo.currentData()
                    if img_format_data:
                        from src.core.converter import ImageFormat
                        options.image_export.format = ImageFormat(img_format_data)
                    
                    if hasattr(self, 'image_height_slider'):
                        target_height = self.image_height_slider.value()
                        options.image_export.use_custom_resolution = True
                        options.image_export.custom_height = target_height
                        options.image_export.custom_width = 0
                        options.image_export.maintain_aspect_ratio = True
                except Exception as e:
                    self.logger.warning(f"设置高级选项时出错: {e}")
            
            # 切换到第三步界面
            self.content_stack.setCurrentIndex(2)
            self.step_indicator.set_current_step(2)
            self.status_bar.showMessage("正在批量转换...")
            
            # 重置进度显示
            self.status_icon.setText("📦")
            self.status_icon.setStyleSheet("font-size: 36px;")
            self.status_text.setText("正在批量转换...")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #1a1a2e;")
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_percent.setText("0%")
            self.progress_percent.setStyleSheet("font-size: 13px; font-weight: 600; color: #667eea;")
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: none;
                    border-radius: 0.1px;
                    background-color: #e5e7eb;
                    height: 8px;
                }
                QProgressBar::chunk {
                    background-color: #667eea;
                    border-radius: 0.1px;
                }
            """)
            self.progress_details.setText("准备批量转换...")
            self.log_text.clear()
            
            # 显示暂停和终止按钮
            self.pause_btn.setVisible(True)
            self.pause_btn.setText("⏸️ 暂停")
            self.pause_btn.setEnabled(True)
            self.stop_btn.setVisible(True)
            self.stop_btn.setEnabled(True)
            self.finish_btn.setVisible(False)
            self.back_btn.setVisible(False)  # 确保隐藏返回按钮
            
            # 显示文件进度表格
            self.file_progress_table.setVisible(True)
            self.file_progress_table.setRowCount(0)
            
            # 检查是否是智能模式
            is_smart_mode = hasattr(self, 'smart_mode_checkbox') and self.smart_mode_checkbox.isChecked()
            
            # 智能模式下显示误差列，非智能模式下隐藏
            self.file_progress_table.setColumnHidden(5, not is_smart_mode)
            
            # 初始化文件进度表格，使用完整路径作为数据
            for file_path in self.files:
                row = self.file_progress_table.rowCount()
                self.file_progress_table.insertRow(row)
                
                # 获取原始文件大小
                try:
                    original_size = os.path.getsize(str(file_path)) / (1024 * 1024)  # MB
                    original_size_str = f"{original_size:.2f}MB"
                except:
                    original_size_str = "-"
                
                # 文件名列存储完整路径作为数据
                name_item = QTableWidgetItem(os.path.basename(str(file_path)))
                name_item.setData(Qt.UserRole, str(file_path))  # 存储完整路径
                self.file_progress_table.setItem(row, 0, name_item)
                self.file_progress_table.setItem(row, 1, QTableWidgetItem("⏳"))
                self.file_progress_table.setItem(row, 2, QTableWidgetItem("0%"))
                self.file_progress_table.setItem(row, 3, QTableWidgetItem(original_size_str))  # 原始大小
                self.file_progress_table.setItem(row, 4, QTableWidgetItem("-"))  # 处理后大小
                self.file_progress_table.setItem(row, 5, QTableWidgetItem("-"))  # 误差
                self.file_progress_table.setItem(row, 6, QTableWidgetItem("-"))  # 导出位置
            
            # 记录日志
            self.log_message(f"🚀 启动批处理转换")
            self.log_message(f"📁 文件数量: {len(self.files)}")
            self.log_message(f"💾 输出目录: {output_dir}")
            if is_smart_mode:
                target_size_mb = self.target_size_spinbox.value()
                self.log_message(f"🎯 目标大小: {target_size_mb}MB")
            
            # 清理之前的worker
            if hasattr(self, 'batch_worker') and self.batch_worker is not None:
                if self.batch_worker.isRunning():
                    self.batch_worker.quit()
                    self.batch_worker.wait(1000)
                self.batch_worker.deleteLater()
                self.batch_worker = None
            
            # 创建批处理工作线程
            from src.ui.batch_conversion_worker import BatchConversionWorker
            
            # 获取选择的算法
            algorithm = "v4"
            if is_smart_mode and hasattr(self, 'algorithm_combo'):
                algorithm = self.algorithm_combo.currentData()
            
            self.batch_worker = BatchConversionWorker(
                [str(f) for f in self.files],
                options,
                output_dir,
                is_smart_mode=is_smart_mode,
                target_size_mb=self.target_size_spinbox.value() if is_smart_mode else 10.0,
                algorithm=algorithm,
                logger=self.logger
            )
            self.batch_worker.file_started.connect(self._on_batch_file_started)
            self.batch_worker.file_progress.connect(self._on_batch_file_progress)
            self.batch_worker.file_finished.connect(self._on_batch_file_finished)
            self.batch_worker.batch_progress.connect(self._on_batch_progress)
            self.batch_worker.batch_finished.connect(self._on_batch_finished)
            self.batch_worker.start()
            
        except Exception as e:
            self.logger.error(f"启动批处理转换失败: {e}")
            QMessageBox.critical(self, "错误", f"启动批处理转换失败: {e}")
            self.go_to_step2()

    def _on_batch_file_started(self, file_path: str):
        """批处理：文件开始处理"""
        file_name = os.path.basename(file_path)
        self.log_message(f"📄 开始处理: {file_name}")
        
        # 更新文件进度表格，使用完整路径匹配
        for row in range(self.file_progress_table.rowCount()):
            item = self.file_progress_table.item(row, 0)
            if item and item.data(Qt.UserRole) == file_path:
                self.file_progress_table.setItem(row, 1, QTableWidgetItem("🔄"))
                break

    def _on_batch_file_progress(self, file_path: str, progress: float, message: str):
        """批处理：文件进度更新"""
        file_name = os.path.basename(file_path)
        file_progress_percent = int(progress * 100)
        
        # 更新文件进度表格，使用完整路径匹配
        for row in range(self.file_progress_table.rowCount()):
            item = self.file_progress_table.item(row, 0)
            if item and item.data(Qt.UserRole) == file_path:
                self.file_progress_table.setItem(row, 2, QTableWidgetItem(f"{file_progress_percent}%"))
                break
        
        # 计算总进度：已完成文件数 + 当前文件进度
        total_files = len(self.files)
        completed_files = 0
        for row in range(self.file_progress_table.rowCount()):
            status = self.file_progress_table.item(row, 1).text()
            if status in ["✓ 完成", "✗ 失败"]:
                completed_files += 1
        
        # 总进度 = (已完成文件数 + 当前文件进度) / 总文件数
        total_progress = ((completed_files + progress) / total_files) * 100
        total_progress = min(99, int(total_progress))  # 最大99%，完成后才显示100%
        
        self.progress_bar.setValue(total_progress)
        self.progress_percent.setText(f"{total_progress}%")
        self.progress_details.setText(f"处理 {file_name}: {message} ({file_progress_percent}%)")

    def _on_batch_file_finished(self, file_path: str, success: bool, output_path: str):
        """批处理：文件处理完成"""
        file_name = os.path.basename(file_path)
        
        # 检查是否是智能模式
        is_smart_mode = hasattr(self, 'smart_mode_checkbox') and self.smart_mode_checkbox.isChecked()
        target_size_mb = self.target_size_spinbox.value() if is_smart_mode else 0
        
        if success:
            # 获取处理后文件大小
            processed_size_str = "-"
            error_percent_str = "-"
            
            if output_path and os.path.exists(output_path):
                # 判断是文件夹还是文件
                if os.path.isdir(output_path):
                    total_size = 0
                    for dirpath, dirnames, filenames in os.walk(output_path):
                        for f in filenames:
                            fp = os.path.join(dirpath, f)
                            total_size += os.path.getsize(fp)
                    processed_size_mb = total_size / (1024 * 1024)
                else:
                    processed_size_mb = os.path.getsize(output_path) / (1024 * 1024)
                
                processed_size_str = f"{processed_size_mb:.2f}MB"
                
                # 计算误差（仅智能模式）
                if is_smart_mode and target_size_mb > 0:
                    error_percent = ((processed_size_mb - target_size_mb) / target_size_mb) * 100
                    error_percent_str = f"{error_percent:+.1f}%"
            
            # 显示完整导出路径
            self.log_message(f"✓ 完成: {file_name}")
            self.log_message(f"  保存到: {output_path}")
            self.log_message(f"  处理后大小: {processed_size_str}")
            
            # 更新文件进度表格，使用完整路径匹配
            # 列顺序：文件名(0)、状态(1)、进度(2)、原始大小(3)、处理后大小(4)、误差(5)、导出位置(6)
            for row in range(self.file_progress_table.rowCount()):
                item = self.file_progress_table.item(row, 0)
                if item and item.data(Qt.UserRole) == file_path:
                    self.file_progress_table.setItem(row, 1, QTableWidgetItem("✅"))
                    self.file_progress_table.setItem(row, 2, QTableWidgetItem("100%"))
                    self.file_progress_table.setItem(row, 4, QTableWidgetItem(processed_size_str))
                    if is_smart_mode:
                        self.file_progress_table.setItem(row, 5, QTableWidgetItem(error_percent_str))
                    self.file_progress_table.setItem(row, 6, QTableWidgetItem(output_path))
                    break
        else:
            self.log_message(f"✗ 失败: {file_name}")
            
            # 更新文件进度表格，使用完整路径匹配
            for row in range(self.file_progress_table.rowCount()):
                item = self.file_progress_table.item(row, 0)
                if item and item.data(Qt.UserRole) == file_path:
                    self.file_progress_table.setItem(row, 1, QTableWidgetItem("❌"))
                    self.file_progress_table.setItem(row, 2, QTableWidgetItem("-"))
                    break

    def _on_batch_progress(self, progress: float):
        """批处理：整体进度更新（文件层面）"""
        # 这个回调只在文件完成时触发，所以这里不做进度条更新
        # 实际进度更新在 _on_batch_file_progress 中处理
        pass

    def _on_batch_finished(self, result: dict):
        """批处理：全部完成"""
        total = result.get('total', 0)
        success = result.get('success', 0)
        failed = result.get('failed', 0)
        results = result.get('results', [])
        error = result.get('error', '')
        
        self.logger.info(f"批处理完成: 成功 {success}/{total}, 失败 {failed}/{total}, error={error}")
        
        # 添加到历史记录
        for r in results:
            if r.get('success') and r.get('output'):
                try:
                    output_path = r.get('output', '')
                    file_size = 0
                    if os.path.exists(output_path):
                        if os.path.isdir(output_path):
                            total_size = 0
                            for dirpath, dirnames, filenames in os.walk(output_path):
                                for f in filenames:
                                    fp = os.path.join(dirpath, f)
                                    total_size += os.path.getsize(fp)
                            file_size = total_size / (1024 * 1024)
                        else:
                            file_size = os.path.getsize(output_path) / (1024 * 1024)
                    
                    self.history_manager.add_record(
                        input_path=r.get('file', ''),
                        output_path=output_path,
                        mode='智能模式' if hasattr(self, 'smart_mode_checkbox') and self.smart_mode_checkbox.isChecked() else self.mode_combo.currentText(),
                        output_format=self.output_format_combo.currentData(),
                        success=True,
                        file_size=file_size
                    )
                except Exception as e:
                    self.logger.warning(f"添加历史记录失败: {e}")
        
        # 隐藏暂停和终止按钮
        self.pause_btn.setVisible(False)
        self.stop_btn.setVisible(False)
        
        # 设置转换成功标志
        self._conversion_success = (failed == 0 and not error)
        
        # 检查是否是用户终止
        if error == '用户终止':
            self.status_icon.setText("⏹️")
            self.status_text.setText("已终止")
            self.status_text.setStyleSheet("font-size: 16px; font-weight: 600; color: #ef4444;")
            self.log_message(f"⏹️ 批量转换已终止")
            self.status_bar.showMessage("批量转换已终止")
            
            # 显示重试和返回按钮
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(True)
            self.finish_btn.setText("🔄 重试")
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #f59e0b, stop:1 #d97706);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #d97706, stop:1 #b45309);
                }
            """)
            return
        
        # 显示完成按钮
        self.finish_btn.setVisible(True)
        self.back_btn.setVisible(False)
        self.finish_btn.setText("🤗 回到开始")
        self.finish_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #10b981, stop:1 #059669);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 24px;
                font-size: 13px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #059669, stop:1 #047857);
            }
        """)
        
        # 更新状态显示
        if failed == 0:
            self.status_icon.setText("🥳")
            self.status_text.setText("批量转换完成!")
            self.log_message(f"✅ 批量转换完成: 成功 {success}/{total}")
        else:
            self.status_icon.setText("⚠️")
            self.status_text.setText(f"批量转换完成 ({failed}个失败)")
            self.log_message(f"⚠️ 批量转换完成: 成功 {success}/{total}, 失败 {failed}/{total}")
        
        self.progress_bar.setValue(100)
        self.progress_percent.setText("100%")
        self.status_bar.showMessage(f"批量转换完成: 成功 {success}/{total}")
        
        # 执行完成后操作
        if failed == 0 and results:
            # 获取最后一个成功文件的输出路径
            last_output = None
            for r in reversed(results):
                if r.get('success') and r.get('output'):
                    last_output = r.get('output')
                    break
            if last_output:
                self._execute_after_completion(last_output)

    def _on_smart_optimization_finished(self, result: dict):
        """智能处理完成 - V4系统（每页独立调优）"""
        if result.get('success'):
            self.logger.info(f"智能处理成功: {result}")
            
            # 设置转换成功标志
            self._conversion_success = True
            
            # V4系统已经完成了转换，不需要再次转换
            # 清除目标大小标记，避免触发微调逻辑
            self._smart_optimization_target_size = None
            self._smart_optimization_retry_count = 0
            
            # 记录优化结果
            self.log_message(f"✅ 智能处理完成!")
            self.log_message(f"   基准体积A: {result.get('base_volume_a_mb', 0):.2f}MB")
            self.log_message(f"   目标每页: {result.get('target_per_page_mb', 0):.2f}MB")
            self.log_message(f"   总页数: {result.get('total_pages', 0)}页")
            self.log_message(f"   预估大小: {result.get('estimated_size_mb', 0):.2f}MB")
            self.log_message(f"   实际大小: {result.get('actual_size_mb', 0):.2f}MB")
            
            # 显示极端边界警告（如果有）
            boundary_warning = result.get('boundary_warning', '')
            if boundary_warning:
                self.log_message(f"⚠️ {boundary_warning}")
            
            # 显示每页的高度（前5页和最后5页）
            page_heights = result.get('page_heights', [])
            if page_heights:
                total_pages = len(page_heights)
                if total_pages <= 10:
                    height_str = ", ".join([f"{i+1}:{h}px" for i, h in enumerate(page_heights)])
                else:
                    first_5 = [f"{i+1}:{h}px" for i, h in enumerate(page_heights[:5])]
                    last_5 = [f"{i+1}:{h}px" for i, h in enumerate(page_heights[-5:], start=total_pages-5)]
                    height_str = ", ".join(first_5) + ", ..., " + ", ".join(last_5)
                self.log_message(f"   每页高度: {height_str}")
            
            # 更新状态显示
            self.status_icon.setText("🥳")
            self.status_text.setText("智能处理完成!")
            self.progress_bar.setValue(100)
            self.progress_percent.setText("100%")
            
            # 显示完成按钮
            self.pause_btn.setVisible(False)
            self.stop_btn.setVisible(False)
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(False)
            self.finish_btn.setText("🤗 回到开始")
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #10b981, stop:1 #059669);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #059669, stop:1 #047857);
                }
            """)
            
            # 记录输出路径
            self._last_output_path = result.get('output_path', '')
            
            # 更新当前输出文件信息
            if self._last_output_path:
                try:
                    self.current_output_file = self._last_output_path
                    self.current_output_dir = os.path.dirname(self._last_output_path)
                    self.status_bar.showMessage(f"智能处理完成: {os.path.basename(self._last_output_path)}")
                    
                    # 添加到历史记录
                    try:
                        actual_size_mb = result.get('actual_size_mb', 0)
                        if actual_size_mb == 0 and os.path.exists(self._last_output_path):
                            actual_size_mb = os.path.getsize(self._last_output_path) / (1024 * 1024)
                        
                        self.history_manager.add_record(
                            input_path=self.current_input_file,
                            output_path=self._last_output_path,
                            mode='智能模式',
                            output_format='pptx',
                            success=True,
                            file_size=actual_size_mb
                        )
                        self.logger.info(f"已添加历史记录: {self._last_output_path}")
                    except Exception as e:
                        self.logger.warning(f"添加历史记录失败: {e}")
                    
                    # 执行完成后操作
                    self._execute_after_completion(self._last_output_path)
                except Exception as e:
                    self.logger.warning(f"更新输出文件信息失败: {e}")
            
        else:
            self._on_smart_optimization_error(result.get('message', '优化失败'))

    def _on_smart_optimization_error(self, error_message: str):
        """智能处理出错"""
        self.logger.error(f"智能处理失败: {error_message}")
        
        # 隐藏暂停和终止按钮
        self.pause_btn.setVisible(False)
        self.stop_btn.setVisible(False)
        
        # 如果是用户终止，不显示警告对话框，显示重试按钮
        if error_message == "用户终止":
            self.log_message("⏹️ 已终止转换")
            self.status_icon.setText("⏹️")
            self.status_text.setText("已终止")
            self.status_bar.showMessage("转换已终止")
            
            # 设置转换失败标志，以便重试按钮能正常工作
            self._conversion_success = False
            
            # 显示重试和返回按钮
            self.finish_btn.setVisible(True)
            self.back_btn.setVisible(True)
            self.finish_btn.setText("🔄 重试")
            self.finish_btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #f59e0b, stop:1 #d97706);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 10px 24px;
                    font-size: 13px;
                    font-weight: 600;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 #d97706, stop:1 #b45309);
                }
            """)
        else:
            QMessageBox.warning(self, "智能处理失败", f"优化过程出错:\n{error_message}")
            self.go_to_step2()
    
    def on_smart_config_applied(self, config: dict):
        """处理智能配置应用"""
        self.logger.info(f"应用智能配置: {config}")
        
        # 应用到图片设置
        if hasattr(self, 'quality_slider'):
            self.quality_slider.setValue(config.get('quality', 85))
            self.on_quality_changed(config.get('quality', 85))
        
        if hasattr(self, 'image_height_slider'):
            self.image_height_slider.setValue(config.get('height', 1080))
            self.on_height_slider_changed(config.get('height', 1080))
        
        if hasattr(self, 'dpi_combo'):
            # 找到最接近的DPI选项
            target_dpi = config.get('dpi', 96)
            closest_index = 0
            closest_diff = float('inf')
            for i in range(self.dpi_combo.count()):
                dpi = self.dpi_combo.itemData(i)
                if dpi is not None:
                    diff = abs(dpi - target_dpi)
                    if diff < closest_diff:
                        closest_diff = diff
                        closest_index = i
            self.dpi_combo.setCurrentIndex(closest_index)
        
        # 更新状态栏
        self.status_bar.showMessage(
            f"智能配置已应用: 高度={config.get('height')}px, "
            f"质量={config.get('quality')}%, DPI={config.get('dpi')}"
        )
        
        # 记录日志
        self.logger.info(
            f"智能配置应用成功 - 高度: {config.get('height')}px, "
            f"质量: {config.get('quality')}%, DPI: {config.get('dpi')}, "
            f"预估大小: {config.get('estimated_size_mb', 0):.1f}MB"
        )
    
    def log_message(self, message: str):
        """记录日志消息"""
        timestamp = __import__('datetime').datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    # ==================== 页面切换方法 ====================
    
    def show_main_page(self):
        """显示主页"""
        self.main_stack.setCurrentIndex(0)
        self.home_btn.setChecked(True)
        self.history_btn.setChecked(False)
        self.settings_btn.setChecked(False)
        self.status_bar.showMessage("主页 - 选择要转换的文件")
    
    def show_history_page(self):
        """显示历史记录页面"""
        self.main_stack.setCurrentIndex(1)
        self.home_btn.setChecked(False)
        self.history_btn.setChecked(True)
        self.settings_btn.setChecked(False)
        self.refresh_history_table()
        self.status_bar.showMessage("历史记录")
    
    def show_settings_page(self):
        """显示设置页面"""
        self.main_stack.setCurrentIndex(2)
        self.home_btn.setChecked(False)
        self.history_btn.setChecked(False)
        self.settings_btn.setChecked(True)
        self.status_bar.showMessage("设置")
    
    # ==================== 历史记录页面 ====================
    
    def create_history_page(self) -> QWidget:
        """创建历史记录页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("📜 转换历史记录")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: #1a1a2e;")
        layout.addWidget(title)
        
        subtitle = QLabel("查看和管理您的转换历史")
        subtitle.setStyleSheet("color: #718096; font-size: 13px;")
        layout.addWidget(subtitle)
        
        # 统计信息卡片 - 透明背景
        stats_group = QGroupBox("统计信息")
        stats_group.setStyleSheet("""
            QGroupBox {
                background-color: transparent;
                border: 1px solid #e2e8f0;
                border-radius: 4px;
                margin-top: 10px;
                font-weight: bold;
                font-size: 13px;
                color: #2d3748;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        stats_layout = QHBoxLayout(stats_group)
        
        self.stat_total_label = QLabel("总转换: 0")
        self.stat_success_label = QLabel("成功: 0")
        self.stat_failed_label = QLabel("失败: 0")
        self.stat_rate_label = QLabel("成功率: 0%")
        
        for label in [self.stat_total_label, self.stat_success_label, 
                      self.stat_failed_label, self.stat_rate_label]:
            label.setStyleSheet("font-size: 14px; font-weight: 600; color: #4a5568;")
            stats_layout.addWidget(label)
        
        stats_layout.addStretch()
        layout.addWidget(stats_group)
        
        # 历史记录表格
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels([
            "时间", "输入文件", "输出文件", "格式", "状态"
        ])

        # 设置列宽策略 - 时间列固定宽度，格式列较窄
        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # 时间
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 输入文件
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 输出文件
        header.setSectionResizeMode(3, QHeaderView.Fixed)  # 格式
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # 状态

        self.history_table.setColumnWidth(0, 130)  # 时间列
        self.history_table.setColumnWidth(3, 60)   # 格式列
        
        # 隐藏行号（垂直表头）
        self.history_table.verticalHeader().setVisible(False)
        
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 4px;
                gridline-color: #e2e8f0;
            }
            QTableWidget::item {
                padding: 6px 8px;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #f7fafc;
                padding: 10px;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                font-weight: 600;
                color: #4a5568;
                font-size: 13px;
            }
            QTableWidget::item:selected {
                background-color: #dbeafe;
                color: #1e40af;
            }
        """)
        layout.addWidget(self.history_table)
        
        # 操作按钮
        btn_layout = QHBoxLayout()
        
        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.setFixedHeight(32)
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #f3f4f6;
                color: #374151;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e5e7eb;
            }
            QPushButton:pressed {
                background-color: #d1d5db;
            }
        """)
        refresh_btn.clicked.connect(self.refresh_history_table)
        btn_layout.addWidget(refresh_btn)
        
        clear_btn = QPushButton("🗑️ 清空历史")
        clear_btn.setFixedHeight(32)
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #fee2e2;
                color: #dc2626;
                border: 1px solid #fecaca;
                border-radius: 6px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #fecaca;
            }
            QPushButton:pressed {
                background-color: #fca5a5;
            }
        """)
        clear_btn.clicked.connect(self.clear_history)
        btn_layout.addWidget(clear_btn)
        
        btn_layout.addStretch()
        
        back_btn = QPushButton("👈 返回主页")
        back_btn.setFixedHeight(36)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
        """)
        back_btn.clicked.connect(self.show_main_page)
        btn_layout.addWidget(back_btn)
        
        layout.addLayout(btn_layout)
        
        return page
    
    def refresh_history_table(self):
        """刷新历史记录表格"""
        history = self.history_manager.get_all()
        self.history_table.setRowCount(len(history))
        
        for row, record in enumerate(history):
            # 时间
            timestamp = record.get("timestamp", "")
            if timestamp:
                try:
                    from datetime import datetime
                    dt = datetime.fromisoformat(timestamp)
                    time_str = dt.strftime("%Y-%m-%d %H:%M")
                    full_time_str = dt.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    time_str = timestamp
                    full_time_str = timestamp
            else:
                time_str = "未知"
                full_time_str = "未知"
            time_item = QTableWidgetItem(time_str)
            time_item.setToolTip(f"完整时间: {full_time_str}")
            self.history_table.setItem(row, 0, time_item)
            
            # 输入文件
            input_path = record.get("input_path", "")
            input_name = os.path.basename(input_path) if input_path else "未知"
            input_item = QTableWidgetItem(input_name)
            input_item.setToolTip(f"完整路径: {input_path}")
            self.history_table.setItem(row, 1, input_item)
            
            # 输出文件
            output_path = record.get("output_path", "")
            output_name = os.path.basename(output_path) if output_path else "未知"
            output_item = QTableWidgetItem(output_name)
            output_item.setToolTip(f"完整路径: {output_path}")
            self.history_table.setItem(row, 2, output_item)
            
            # 格式
            fmt = record.get("output_format", "未知")
            fmt_item = QTableWidgetItem(fmt)
            fmt_item.setToolTip(f"输出格式: {fmt}")
            self.history_table.setItem(row, 3, fmt_item)
            
            # 状态 - 使用更深的绿色
            success = record.get("success", False)
            status_text = "✅ 成功" if success else "❌ 失败"
            status_item = QTableWidgetItem(status_text)
            # 使用更深的绿色 (#059669) 和红色 (#dc2626)
            from PyQt5.QtGui import QColor
            status_color = QColor("#059669") if success else QColor("#dc2626")
            status_item.setForeground(status_color)
            status_item.setToolTip("转换成功" if success else "转换失败")
            self.history_table.setItem(row, 4, status_item)
        
        # 更新统计信息
        self.update_history_stats()
    
    def update_history_stats(self):
        """更新历史记录统计信息"""
        stats = self.history_manager.get_statistics()
        self.stat_total_label.setText(f"总转换: {stats['total_records']}")
        self.stat_success_label.setText(f"成功: {stats['successful']}")
        self.stat_failed_label.setText(f"失败: {stats['failed']}")
        success_rate = stats['success_rate'] * 100
        self.stat_rate_label.setText(f"成功率: {success_rate:.1f}%")
    
    def clear_history(self):
        """清空历史记录"""
        reply = QMessageBox.question(
            self, "确认", "确定要清空所有历史记录吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            # 清空历史记录文件
            try:
                import json
                with open(self.history_manager.history_file, 'w', encoding='utf-8') as f:
                    json.dump([], f)
                self.history_manager._history = []
                self.refresh_history_table()
                self.status_bar.showMessage("历史记录已清空")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"清空历史记录失败: {e}")
    
    def open_file(self, file_path: str):
        """打开文件"""
        if file_path and os.path.exists(file_path):
            import subprocess
            subprocess.Popen(['start', file_path], shell=True)
        else:
            QMessageBox.warning(self, "错误", "文件不存在或路径无效")
    
    def open_folder(self, file_path: str):
        """打开文件夹"""
        if file_path:
            folder = os.path.dirname(file_path)
            if os.path.exists(folder):
                import subprocess
                subprocess.Popen(['explorer', folder])
            else:
                QMessageBox.warning(self, "错误", "文件夹不存在")
        else:
            QMessageBox.warning(self, "错误", "路径无效")
    
    def delete_history_record(self, row: int):
        """删除单条历史记录"""
        reply = QMessageBox.question(
            self, "确认", "确定要删除这条历史记录吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                history = self.history_manager.get_all()
                if 0 <= row < len(history):
                    history.pop(row)
                    # 保存更新后的历史记录
                    import json
                    with open(self.history_manager.history_file, 'w', encoding='utf-8') as f:
                        json.dump(history, f, ensure_ascii=False, indent=2)
                    self.history_manager._history = history
                    self.refresh_history_table()
                    self.status_bar.showMessage("记录已删除")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"删除记录失败: {e}")
    
    # ==================== 设置页面 ====================
    
    def create_settings_page(self) -> QWidget:
        """创建设置页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setSpacing(15)
        
        # 标题
        title = QLabel("⚙️ 设置")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: #1a1a2e;")
        layout.addWidget(title)
        
        subtitle = QLabel("配置应用程序的全局设置")
        subtitle.setStyleSheet("color: #718096; font-size: 13px;")
        layout.addWidget(subtitle)
        
        # 创建标签页
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-top: none;
                border-radius: 0 0 8px 8px;
                background-color: white;
                padding: 12px;
            }
            QTabBar::tab {
                background-color: #f9fafb;
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 12px;
                font-weight: 600;
                color: #6b7280;
                border: 1px solid transparent;
                border-bottom: 1px solid #e5e7eb;
            }
            QTabBar::tab:selected {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-bottom: 1px solid white;
                color: #2563eb;
                font-weight: 700;
            }
            QTabBar::tab:hover:!selected {
                background-color: #f3f4f6;
            }
        """)

        # 常规设置
        general_tab = self.create_general_settings_tab()
        tabs.addTab(general_tab, "常规")
        
        # 高级设置
        advanced_tab = self.create_advanced_settings_tab()
        tabs.addTab(advanced_tab, "高级")
        
        layout.addWidget(tabs)
        
        # 底部按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("💾 保存设置")
        save_btn.setFixedHeight(36)
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:pressed {
                background-color: #047857;
            }
        """)
        save_btn.clicked.connect(self.save_settings)
        btn_layout.addWidget(save_btn)

        reset_btn = QPushButton("🔄 重置为默认")
        reset_btn.setFixedHeight(36)
        reset_btn.setStyleSheet("""
            QPushButton {
                background-color: #f3f4f6;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: bold;
                color: #374151;
            }
            QPushButton:hover {
                background-color: #e5e7eb;
                border-color: #9ca3af;
            }
        """)
        reset_btn.clicked.connect(self.reset_settings)
        btn_layout.addWidget(reset_btn)

        back_btn = QPushButton("👈 返回主页")
        back_btn.setFixedHeight(36)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
        """)
        back_btn.clicked.connect(self.show_main_page)
        btn_layout.addWidget(back_btn)
        
        layout.addLayout(btn_layout)
        
        return page
    
    def create_general_settings_tab(self) -> QWidget:
        """创建常规设置标签页"""
        tab = QWidget()
        layout = QFormLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 默认输出目录
        self.default_output_dir_edit = QLineEdit()
        self.default_output_dir_edit.setPlaceholderText("默认与源文件相同目录")
        self.default_output_dir_edit.setReadOnly(True)
        
        output_dir_btn = QPushButton("浏览...")
        output_dir_btn.setFixedHeight(32)
        output_dir_btn.setStyleSheet("""
            QPushButton {
                background-color: #f3f4f6;
                color: #374151;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e5e7eb;
                border-color: #9ca3af;
            }
            QPushButton:pressed {
                background-color: #d1d5db;
            }
        """)
        output_dir_btn.clicked.connect(self.select_default_output_dir)
        
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(self.default_output_dir_edit)
        output_dir_layout.addWidget(output_dir_btn)
        layout.addRow("默认输出目录:", output_dir_layout)
        
        # 文件名后缀（迁移到常规）
        self.filename_suffix_edit = QLineEdit("_converted")
        layout.addRow("默认文件名后缀:", self.filename_suffix_edit)
        
        # 最小化到托盘
        self.minimize_to_tray_checkbox = QCheckBox("关闭时最小化到系统托盘")
        layout.addRow("", self.minimize_to_tray_checkbox)
        
        layout.addRow("", QLabel())  # 占位
        
        return tab
    
    def create_advanced_settings_tab(self) -> QWidget:
        """创建高级设置标签页"""
        tab = QWidget()
        layout = QFormLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 日志级别
        self.log_level_combo = QComboBox()
        self.log_level_combo.addItem("调试 (Debug)", "DEBUG")
        self.log_level_combo.addItem("信息 (Info)", "INFO")
        self.log_level_combo.addItem("警告 (Warning)", "WARNING")
        self.log_level_combo.addItem("错误 (Error)", "ERROR")
        self.log_level_combo.setCurrentIndex(1)  # 默认 INFO
        layout.addRow("日志级别:", self.log_level_combo)
        
        # 最大历史记录数
        self.max_history_spin = QLineEdit("50")
        layout.addRow("最大历史记录数:", self.max_history_spin)
        
        # PowerPoint 可见性
        self.ppt_visible_checkbox = QCheckBox("显示 PowerPoint 窗口（调试用）")
        layout.addRow("", self.ppt_visible_checkbox)
        
        layout.addRow("", QLabel())  # 占位
        
        return tab
    
    def select_default_output_dir(self):
        """选择默认输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择默认输出目录", "")
        if dir_path:
            self.default_output_dir_edit.setText(dir_path)
    
    def save_settings(self):
        """保存设置"""
        try:
            settings = {
                "general": {
                    "default_output_dir": self.default_output_dir_edit.text(),
                    "filename_suffix": self.filename_suffix_edit.text(),
                    "minimize_to_tray": self.minimize_to_tray_checkbox.isChecked()
                },
                "advanced": {
                    "log_level": self.log_level_combo.currentData(),
                    "max_history": int(self.max_history_spin.text() or 50),
                    "ppt_visible": self.ppt_visible_checkbox.isChecked()
                }
            }
            
            # 保存到配置文件
            self.config_manager.update_config(settings)
            
            # 更新运行时设置
            self._minimize_to_tray = self.minimize_to_tray_checkbox.isChecked()
            
            # 设置日志级别
            log_level = self.log_level_combo.currentData()
            if log_level:
                import logging
                logging.getLogger().setLevel(getattr(logging, log_level, logging.INFO))
            
            # 设置最大历史记录数
            max_history = int(self.max_history_spin.text() or 50)
            if hasattr(self, 'history_manager'):
                self.history_manager.max_records = max_history
            
            # 设置 PowerPoint 可见性
            ppt_visible = self.ppt_visible_checkbox.isChecked()
            if hasattr(self, 'converter') and self.converter and self.converter.powerpoint:
                try:
                    self.converter.powerpoint.Visible = ppt_visible
                except:
                    pass
            
            QMessageBox.information(self, "成功", "设置已保存！")
            self.status_bar.showMessage("设置已保存")
            self.logger.info("设置已保存")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存设置失败: {e}")
            self.logger.error(f"保存设置失败: {e}")
    
    def reset_settings(self):
        """重置设置为默认值"""
        reply = QMessageBox.question(
            self, "确认", "确定要重置所有设置为默认值吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            # 重置所有控件
            self.default_output_dir_edit.clear()
            self.filename_suffix_edit.setText("_converted")
            self.minimize_to_tray_checkbox.setChecked(False)
            self.log_level_combo.setCurrentIndex(1)
            self.max_history_spin.setText("50")
            self.ppt_visible_checkbox.setChecked(False)
            
            QMessageBox.information(self, "成功", "设置已重置为默认值！")
            self.status_bar.showMessage("设置已重置")

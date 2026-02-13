"""
智能配置UI组件

提供用户友好的界面来配置智能导出参数：
1. 核心约束输入（文件大小上限、导出范围）
2. 画质倾向策略（清晰度优先/色彩平衡优先）
3. 参数阈值设置（高级选项）
4. 实时状态反馈（进度条、预测结果）
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QProgressBar, QSpinBox, QComboBox,
                             QRadioButton, QButtonGroup, QGroupBox, QFormLayout,
                             QLineEdit, QCheckBox, QFrame, QSlider, QMessageBox,
                             QSizePolicy, QGridLayout, QToolButton, QMenu)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont
from typing import Optional, Callable
import logging


class SmartConfigWorker(QThread):
    """智能配置优化工作线程"""
    
    progress_updated = pyqtSignal(str, int)  # 状态消息, 进度百分比
    result_ready = pyqtSignal(dict)  # 优化结果
    error_occurred = pyqtSignal(str)  # 错误信息
    
    def __init__(self, pptx_path: str, target_size_mb: float, 
                 priority_mode: str = "balanced", logger=None, parent=None):
        super().__init__(parent)
        self.pptx_path = pptx_path
        self.target_size_mb = target_size_mb
        self.priority_mode = priority_mode  # "resolution", "quality", "balanced"
        self.logger = logger or logging.getLogger(__name__)
        self._stopped = False
    
    def stop(self):
        """停止优化"""
        self._stopped = True
    
    def run(self):
        """执行智能配置优化"""
        try:
            from src.core.smart_config import SmartConfigOptimizer
            
            self.progress_updated.emit("正在初始化优化器...", 5)
            
            # 创建优化器，传入主logger
            optimizer = SmartConfigOptimizer(logger=self.logger)
            
            # 定义进度回调
            def progress_callback(message: str, progress: int):
                if not self._stopped:
                    self.progress_updated.emit(message, progress)
            
            self.progress_updated.emit("正在进行样本采样...", 15)
            
            # 执行优化
            result = optimizer.optimize(
                self.pptx_path, 
                self.target_size_mb,
                progress_callback=progress_callback
            )
            
            if not self._stopped:
                self.progress_updated.emit("优化完成!", 100)
                
                # 构建结果字典
                result_dict = {
                    "success": result.success,
                    "quality": result.quality,
                    "height": result.height,
                    "dpi": result.dpi,
                    "estimated_size_mb": result.estimated_size_mb,
                    "confidence": result.confidence,
                    "message": result.message,
                    "iterations": result.iterations,
                    "total_time_seconds": result.total_time_seconds,
                    "sample_pages": result.sample_pages,
                    "priority_mode": self.priority_mode
                }
                
                self.result_ready.emit(result_dict)
                
        except Exception as e:
            self.logger.error(f"智能配置优化失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            if not self._stopped:
                self.error_occurred.emit(str(e))


class SmartConfigWidget(QWidget):
    """智能配置组件"""
    
    # 信号
    config_applied = pyqtSignal(dict)  # 配置已应用
    optimization_started = pyqtSignal()  # 开始优化
    optimization_finished = pyqtSignal(bool, str)  # 优化完成 (成功/失败, 消息)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.current_pptx_path = None
        self.optimization_worker = None
        self.last_result = None
        
        self.setup_ui()
        self.apply_styles()
    
    def setup_ui(self):
        """设置UI布局"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(16)
        
        # ========== 标题区域 ==========
        title_layout = QHBoxLayout()
        
        title_label = QLabel("🎯 智能配置")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: 600;
                color: #1a1a2e;
            }
        """)
        title_layout.addWidget(title_label)
        
        # 帮助按钮
        help_btn = QToolButton()
        help_btn.setText("?")
        help_btn.setToolTip("智能配置说明")
        help_btn.setStyleSheet("""
            QToolButton {
                background-color: #e5e7eb;
                color: #6b7280;
                border: none;
                border-radius: 10px;
                width: 20px;
                height: 20px;
                font-size: 12px;
                font-weight: bold;
            }
            QToolButton:hover {
                background-color: #d1d5db;
            }
        """)
        help_btn.clicked.connect(self.show_help)
        title_layout.addWidget(help_btn)
        title_layout.addStretch()
        
        main_layout.addLayout(title_layout)
        
        # ========== 1. 核心约束输入 ==========
        constraint_group = QGroupBox("核心约束")
        constraint_group.setStyleSheet(self._get_groupbox_style())
        constraint_layout = QFormLayout(constraint_group)
        constraint_layout.setSpacing(12)
        
        # 文件大小上限
        size_layout = QHBoxLayout()
        size_layout.setSpacing(8)
        
        self.size_spinbox = QSpinBox()
        self.size_spinbox.setRange(1, 1000)
        self.size_spinbox.setValue(100)
        self.size_spinbox.setSuffix(" MB")
        self.size_spinbox.setFixedWidth(120)
        self.size_spinbox.setStyleSheet(self._get_input_style())
        
        # 预设快捷按钮
        preset_menu = QMenu(self)
        presets = [
            ("微信传输 (100MB)", 100),
            ("邮件附件 (20MB)", 20),
            ("网盘分享 (50MB)", 50),
            ("高质量 (200MB)", 200),
            ("无损 (500MB)", 500),
        ]
        for name, size in presets:
            action = preset_menu.addAction(name)
            action.triggered.connect(lambda checked, s=size: self.size_spinbox.setValue(s))
        
        preset_btn = QPushButton("📋 预设")
        preset_btn.setMenu(preset_menu)
        preset_btn.setStyleSheet(self._get_preset_button_style())
        preset_btn.setFixedWidth(80)
        
        size_layout.addWidget(self.size_spinbox)
        size_layout.addWidget(preset_btn)
        size_layout.addStretch()
        
        constraint_layout.addRow("文件大小上限:", size_layout)
        
        # 导出范围
        range_layout = QHBoxLayout()
        range_layout.setSpacing(12)
        
        self.range_group = QButtonGroup(self)
        
        self.range_all_radio = QRadioButton("全部幻灯片")
        self.range_all_radio.setChecked(True)
        self.range_group.addButton(self.range_all_radio, 0)
        range_layout.addWidget(self.range_all_radio)
        
        self.range_current_radio = QRadioButton("当前页")
        self.range_group.addButton(self.range_current_radio, 1)
        range_layout.addWidget(self.range_current_radio)
        
        self.range_custom_radio = QRadioButton("自定义")
        self.range_group.addButton(self.range_custom_radio, 2)
        range_layout.addWidget(self.range_custom_radio)
        
        self.range_custom_edit = QLineEdit()
        self.range_custom_edit.setPlaceholderText("如: 1-5,8,10-12")
        self.range_custom_edit.setFixedWidth(140)
        self.range_custom_edit.setStyleSheet(self._get_input_style())
        self.range_custom_edit.setEnabled(False)
        range_layout.addWidget(self.range_custom_edit)
        
        range_layout.addStretch()
        constraint_layout.addRow("导出范围:", range_layout)
        
        # 连接信号
        self.range_custom_radio.toggled.connect(self.range_custom_edit.setEnabled)
        
        main_layout.addWidget(constraint_group)
        
        # ========== 2. 画质倾向策略 ==========
        priority_group = QGroupBox("画质优化策略")
        priority_group.setStyleSheet(self._get_groupbox_style())
        priority_layout = QVBoxLayout(priority_group)
        priority_layout.setSpacing(12)
        
        # 策略说明
        desc_label = QLabel("选择算法优化的优先方向:")
        desc_label.setStyleSheet("color: #6b7280; font-size: 12px;")
        priority_layout.addWidget(desc_label)
        
        # 策略选项
        self.priority_group = QButtonGroup(self)
        
        # 清晰度优先
        self.resolution_radio = QRadioButton("📐 清晰度优先 (Resolution First)")
        self.resolution_radio.setChecked(True)
        self.resolution_radio.setToolTip("优先保持高分辨率，适合文字密集的文档")
        self.priority_group.addButton(self.resolution_radio, 0)
        priority_layout.addWidget(self.resolution_radio)
        
        resolution_desc = QLabel("      算法会尽量拉高高度和DPI，在文件超标时优先牺牲压缩质量。适用于文字密集的文档。")
        resolution_desc.setStyleSheet("color: #9ca3af; font-size: 11px; margin-left: 20px;")
        resolution_desc.setWordWrap(True)
        priority_layout.addWidget(resolution_desc)
        
        # 色彩平衡优先
        self.quality_radio = QRadioButton("🎨 色彩平衡优先 (Quality First)")
        self.quality_radio.setToolTip("优先保持高压缩质量，适合摄影作品或渐变较多的PPT")
        self.priority_group.addButton(self.quality_radio, 1)
        priority_layout.addWidget(self.quality_radio)
        
        quality_desc = QLabel("      算法会保持较高的压缩质量，在文件超标时优先降低高度。适用于摄影作品或渐变较多的PPT。")
        quality_desc.setStyleSheet("color: #9ca3af; font-size: 11px; margin-left: 20px;")
        quality_desc.setWordWrap(True)
        priority_layout.addWidget(quality_desc)
        
        # 平衡模式
        self.balanced_radio = QRadioButton("⚖️ 平衡模式 (Balanced)")
        self.balanced_radio.setToolTip("在清晰度和色彩之间取得平衡")
        self.priority_group.addButton(self.balanced_radio, 2)
        priority_layout.addWidget(self.balanced_radio)
        
        balanced_desc = QLabel("      算法会综合考虑清晰度和色彩质量，自动寻找最佳平衡点。")
        balanced_desc.setStyleSheet("color: #9ca3af; font-size: 11px; margin-left: 20px;")
        balanced_desc.setWordWrap(True)
        priority_layout.addWidget(balanced_desc)
        
        main_layout.addWidget(priority_group)
        
        # ========== 3. 高级设置（可折叠） ==========
        self.advanced_group = QGroupBox("高级设置")
        self.advanced_group.setCheckable(True)
        self.advanced_group.setChecked(False)
        self.advanced_group.setStyleSheet(self._get_groupbox_style())
        advanced_layout = QFormLayout(self.advanced_group)
        advanced_layout.setSpacing(10)
        
        # 最小可接受高度
        self.min_height_spinbox = QSpinBox()
        self.min_height_spinbox.setRange(480, 8640)
        self.min_height_spinbox.setValue(480)
        self.min_height_spinbox.setSuffix(" px")
        self.min_height_spinbox.setStyleSheet(self._get_input_style())
        advanced_layout.addRow("最小可接受高度:", self.min_height_spinbox)
        
        # 最大高度限制
        self.max_height_spinbox = QSpinBox()
        self.max_height_spinbox.setRange(480, 8640)
        self.max_height_spinbox.setValue(8640)
        self.max_height_spinbox.setSuffix(" px")
        self.max_height_spinbox.setStyleSheet(self._get_input_style())
        advanced_layout.addRow("最大高度限制:", self.max_height_spinbox)
        
        # DPI上限
        self.max_dpi_spinbox = QSpinBox()
        self.max_dpi_spinbox.setRange(72, 600)
        self.max_dpi_spinbox.setValue(300)
        self.max_dpi_spinbox.setSuffix(" DPI")
        self.max_dpi_spinbox.setStyleSheet(self._get_input_style())
        advanced_layout.addRow("DPI上限:", self.max_dpi_spinbox)
        
        # 质量上限
        self.max_quality_spinbox = QSpinBox()
        self.max_quality_spinbox.setRange(60, 100)
        self.max_quality_spinbox.setValue(100)
        self.max_quality_spinbox.setSuffix(" %")
        self.max_quality_spinbox.setStyleSheet(self._get_input_style())
        advanced_layout.addRow("质量上限:", self.max_quality_spinbox)
        
        main_layout.addWidget(self.advanced_group)
        
        # ========== 4. 操作按钮 ==========
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        
        # 开始优化按钮
        self.optimize_btn = QPushButton("🚀 开始智能优化")
        self.optimize_btn.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 14px;
                font-weight: 600;
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
        """)
        self.optimize_btn.clicked.connect(self.start_optimization)
        button_layout.addWidget(self.optimize_btn)
        
        # 应用配置按钮
        self.apply_btn = QPushButton("✓ 应用配置")
        self.apply_btn.setEnabled(False)
        self.apply_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:pressed {
                background-color: #047857;
            }
            QPushButton:disabled {
                background-color: #cbd5e0;
                color: #a0aec0;
            }
        """)
        self.apply_btn.clicked.connect(self.apply_config)
        button_layout.addWidget(self.apply_btn)
        
        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        
        # ========== 5. 实时状态反馈 ==========
        self.status_frame = QFrame()
        self.status_frame.setStyleSheet("""
            QFrame {
                background-color: #f3f4f6;
                border-radius: 8px;
                padding: 12px;
            }
        """)
        status_layout = QVBoxLayout(self.status_frame)
        status_layout.setSpacing(8)
        
        # 状态标签
        self.status_label = QLabel("就绪 - 点击\"开始智能优化\"进行参数优化")
        self.status_label.setStyleSheet("color: #6b7280; font-size: 13px;")
        status_layout.addWidget(self.status_label)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                background-color: white;
                text-align: center;
                font-size: 11px;
                color: #4a5568;
                min-height: 20px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #667eea, stop:1 #764ba2);
                border-radius: 5px;
            }
        """)
        status_layout.addWidget(self.progress_bar)
        
        # 预测结果标签
        self.prediction_label = QLabel("")
        self.prediction_label.setStyleSheet("""
            QLabel {
                color: #2563eb;
                font-size: 13px;
                font-weight: 500;
            }
        """)
        self.prediction_label.setVisible(False)
        status_layout.addWidget(self.prediction_label)
        
        main_layout.addWidget(self.status_frame)
        
        # 添加弹性空间
        main_layout.addStretch()
    
    def apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
            
            QRadioButton {
                spacing: 8px;
                font-size: 13px;
                color: #374151;
            }
            
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #d1d5db;
                border-radius: 9px;
                background-color: white;
            }
            
            QRadioButton::indicator:checked {
                background-color: #2563eb;
                border-color: #2563eb;
            }
            
            QRadioButton::indicator:hover {
                border-color: #2563eb;
            }
        """)
    
    def _get_groupbox_style(self):
        """获取GroupBox样式"""
        return """
            QGroupBox {
                background-color: white;
                border: 1px solid #e1e8ed;
                border-radius: 10px;
                margin-top: 8px;
                padding-top: 12px;
                padding: 16px;
                font-weight: 600;
                font-size: 13px;
                color: #1a1a2e;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
            
            QGroupBox::indicator {
                width: 16px;
                height: 16px;
            }
        """
    
    def _get_input_style(self):
        """获取输入框样式"""
        return """
            QSpinBox, QLineEdit {
                background-color: white;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 6px 10px;
                font-size: 13px;
            }
            
            QSpinBox:focus, QLineEdit:focus {
                border-color: #2563eb;
            }
            
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                border: none;
                background-color: #f3f4f6;
            }
            
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background-color: #e5e7eb;
            }
        """
    
    def _get_preset_button_style(self):
        """获取预设按钮样式"""
        return """
            QPushButton {
                background-color: #f3f4f6;
                color: #374151;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
            }
            
            QPushButton:hover {
                background-color: #e5e7eb;
                border-color: #9ca3af;
            }
            
            QPushButton::menu-indicator {
                image: none;
                width: 0px;
            }
        """
    
    def show_help(self):
        """显示帮助信息"""
        help_text = """
        <h3>🎯 智能配置使用说明</h3>
        
        <p><b>智能配置</b>会自动分析您的PPT内容，通过二分搜索算法找到最优的导出参数组合，
        确保输出文件大小符合您的要求，同时保持最佳的视觉效果。</p>
        
        <h4>📋 使用步骤：</h4>
        <ol>
            <li><b>设置文件大小上限</b> - 输入您希望的目标文件大小（MB）</li>
            <li><b>选择导出范围</b> - 全部幻灯片、当前页或自定义范围</li>
            <li><b>选择画质策略</b> - 根据PPT内容类型选择优化方向</li>
            <li><b>点击"开始智能优化"</b> - 算法会自动计算最优参数</li>
            <li><b>应用配置</b> - 将优化结果应用到导出设置</li>
        </ol>
        
        <h4>🎨 画质策略说明：</h4>
        <ul>
            <li><b>清晰度优先</b> - 适合文字密集、图表较多的商务PPT</li>
            <li><b>色彩平衡优先</b> - 适合摄影作品、渐变背景的设计类PPT</li>
            <li><b>平衡模式</b> - 在清晰度和色彩之间自动寻找最佳平衡</li>
        </ul>
        
        <h4>⚙️ 算法原理：</h4>
        <p>算法采用二分搜索策略，通过采样代表性页面（第1页、中间页、最后一页）
        快速估算全量导出的大小，然后迭代调整高度和质量参数，直到找到满足
        文件大小限制的最优配置。</p>
        """
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("智能配置帮助")
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setText(help_text)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
    
    def set_pptx_path(self, path: str):
        """设置当前PPTX文件路径"""
        self.current_pptx_path = path
        self.logger.info(f"智能配置: 设置PPTX路径 {path}")
    
    def start_optimization(self):
        """开始智能优化"""
        if not self.current_pptx_path:
            QMessageBox.warning(self, "警告", "请先选择PPT文件")
            return
        
        if not os.path.exists(self.current_pptx_path):
            QMessageBox.warning(self, "警告", "PPT文件不存在")
            return
        
        # 获取参数
        target_size = self.size_spinbox.value()
        priority_mode = self._get_priority_mode()
        
        # 更新UI状态
        self.optimize_btn.setEnabled(False)
        self.apply_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.prediction_label.setVisible(False)
        
        self.optimization_started.emit()
        
        # 创建工作线程
        self.optimization_worker = SmartConfigWorker(
            self.current_pptx_path,
            target_size,
            priority_mode
        )
        
        self.optimization_worker.progress_updated.connect(self.on_progress_updated)
        self.optimization_worker.result_ready.connect(self.on_optimization_finished)
        self.optimization_worker.error_occurred.connect(self.on_optimization_error)
        
        self.optimization_worker.start()
    
    def _get_priority_mode(self) -> str:
        """获取当前选择的优先模式"""
        if self.resolution_radio.isChecked():
            return "resolution"
        elif self.quality_radio.isChecked():
            return "quality"
        else:
            return "balanced"
    
    def on_progress_updated(self, message: str, progress: int):
        """进度更新回调"""
        self.status_label.setText(message)
        self.progress_bar.setValue(progress)
    
    def on_optimization_finished(self, result: dict):
        """优化完成回调"""
        self.last_result = result
        
        if result.get("success"):
            # 显示预测结果
            prediction_text = (
                f"✓ 优化完成! 预计导出: "
                f"高度={result['height']}px, "
                f"质量={result['quality']}%, "
                f"DPI={result['dpi']}, "
                f"预估大小={result['estimated_size_mb']:.1f}MB "
                f"(置信度: {result['confidence']:.0%})"
            )
            self.prediction_label.setText(prediction_text)
            self.prediction_label.setVisible(True)
            
            self.status_label.setText("优化完成 - 点击\"应用配置\"使用此配置")
            self.apply_btn.setEnabled(True)
            
            self.optimization_finished.emit(True, result.get("message", ""))
        else:
            self.status_label.setText(f"优化失败: {result.get('message', '未知错误')}")
            self.optimization_finished.emit(False, result.get("message", "优化失败"))
        
        self.optimize_btn.setEnabled(True)
    
    def on_optimization_error(self, error_msg: str):
        """优化错误回调"""
        self.status_label.setText(f"错误: {error_msg}")
        self.progress_bar.setValue(0)
        self.optimize_btn.setEnabled(True)
        self.optimization_finished.emit(False, error_msg)
    
    def apply_config(self):
        """应用配置到主设置"""
        if not self.last_result:
            return
        
        config = {
            "quality": self.last_result["quality"],
            "height": self.last_result["height"],
            "dpi": self.last_result["dpi"],
            "estimated_size_mb": self.last_result["estimated_size_mb"],
            "target_size_mb": self.size_spinbox.value(),
            "priority_mode": self.last_result.get("priority_mode", "balanced"),
            "slide_range": self._get_slide_range(),
            "confidence": self.last_result.get("confidence", 0)
        }
        
        self.config_applied.emit(config)
        
        QMessageBox.information(
            self,
            "配置已应用",
            f"智能配置已成功应用到导出设置!\n\n"
            f"导出参数:\n"
            f"  • 图像高度: {config['height']}px\n"
            f"  • JPEG质量: {config['quality']}%\n"
            f"  • DPI: {config['dpi']}\n"
            f"  • 预估大小: {config['estimated_size_mb']:.1f}MB\n\n"
            f"现在可以开始导出了。"
        )
    
    def _get_slide_range(self) -> str:
        """获取幻灯片范围设置"""
        if self.range_all_radio.isChecked():
            return "all"
        elif self.range_current_radio.isChecked():
            return "current"
        else:
            return self.range_custom_edit.text() or "all"
    
    def get_current_config(self) -> dict:
        """获取当前配置（用于外部调用）"""
        return {
            "target_size_mb": self.size_spinbox.value(),
            "priority_mode": self._get_priority_mode(),
            "slide_range": self._get_slide_range(),
            "min_height": self.min_height_spinbox.value(),
            "max_height": self.max_height_spinbox.value(),
            "max_dpi": self.max_dpi_spinbox.value(),
            "max_quality": self.max_quality_spinbox.value(),
        }
    
    def reset(self):
        """重置所有设置"""
        self.size_spinbox.setValue(100)
        self.range_all_radio.setChecked(True)
        self.resolution_radio.setChecked(True)
        self.min_height_spinbox.setValue(480)
        self.max_height_spinbox.setValue(8640)
        self.max_dpi_spinbox.setValue(300)
        self.max_quality_spinbox.setValue(100)
        
        self.progress_bar.setValue(0)
        self.status_label.setText("就绪 - 点击\"开始智能优化\"进行参数优化")
        self.prediction_label.setVisible(False)
        self.apply_btn.setEnabled(False)
        self.last_result = None


# 便捷导入
__all__ = ['SmartConfigWidget', 'SmartConfigWorker']

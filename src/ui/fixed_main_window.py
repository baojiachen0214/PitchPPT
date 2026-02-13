import sys
import os

# 添加项目根目录到sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QProgressBar, QTextEdit, QTabWidget,
                             QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, 
                             QComboBox, QSlider, QMessageBox, QCheckBox, QAbstractItemView,
                             QSplitter, QGroupBox, QFormLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
from src.utils.logger import Logger
from src.utils.config_manager import ConfigManager
from src.utils.history_manager import HistoryManager

class ConversionWorker(QThread):
    """
    转换工作线程，避免阻塞UI
    """
    progress_updated = pyqtSignal(float, str)
    conversion_finished = pyqtSignal(bool, str)
    
    def __init__(self, converter, input_path, output_path, options):
        super().__init__()
        self.converter = converter
        self.input_path = input_path
        self.output_path = output_path
        self.options = options
        self.logger = Logger().get_logger()
    
    def run(self):
        try:
            success = self.converter.convert(
                self.input_path, 
                self.output_path, 
                self.options
            )
            self.conversion_finished.emit(success, self.output_path)
        except Exception as e:
            self.logger.error(f"转换线程异常: {e}")
            self.conversion_finished.emit(False, str(e))


class BatchConversionWorker(QThread):
    """
    批量转换工作线程
    """
    progress_updated = pyqtSignal(int, int, str)
    file_finished = pyqtSignal(int, bool, str, str)
    conversion_finished = pyqtSignal(int, int)
    
    def __init__(self, converter, file_list, output_dir, options):
        super().__init__()
        self.converter = converter
        self.file_list = file_list
        self.output_dir = output_dir
        self.options = options
        self.logger = Logger().get_logger()
        self.success_count = 0
        self.fail_count = 0
    
    def run(self):
        total = len(self.file_list)
        
        for index, input_path in enumerate(self.file_list):
            self.progress_updated.emit(index, total, f"正在转换: {os.path.basename(input_path)}")
            
            try:
                # 生成输出路径
                input_name = os.path.splitext(os.path.basename(input_path))[0]
                output_ext = {
                    OutputFormat.PPTX: "pptx",
                    OutputFormat.PDF: "pdf",
                    OutputFormat.JPG: "jpg"
                }.get(self.options.output_format, "pptx")
                
                if self.options.output_format == OutputFormat.JPG:
                    output_path = os.path.join(self.output_dir, f"{input_name}")
                else:
                    output_path = os.path.join(self.output_dir, f"{input_name}_converted.{output_ext}")
                
                # 执行转换
                success = self.converter.convert(input_path, output_path, self.options)
                
                if success:
                    self.success_count += 1
                    self.file_finished.emit(index, True, output_path, "")
                else:
                    self.fail_count += 1
                    self.file_finished.emit(index, False, "", "转换失败")
                    
            except Exception as e:
                self.fail_count += 1
                self.logger.error(f"批量转换文件失败 {input_path}: {e}")
                self.file_finished.emit(index, False, "", str(e))
        
        self.progress_updated.emit(total, total, "批量转换完成")
        self.conversion_finished.emit(self.success_count, self.fail_count)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logger = Logger().get_logger()
        self.logger.info("Initializing PitchPPT application")
        
        # 初始化转换器
        self.converter = Win32PPTConverter()
        self.current_input_file = None
        self.current_output_dir = None
        
        # 初始化配置和历史记录管理器
        self.config_manager = ConfigManager()
        self.history_manager = HistoryManager()
        
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("PitchPPT - 专业路演PPT处理工具")
        self.setGeometry(100, 100, 1000, 700)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 创建选项卡控件
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # 创建主功能页面
        main_tab = self.create_main_tab()
        tab_widget.addTab(main_tab, "主功能")
        
        # 创建批量转换页面
        batch_tab = self.create_batch_tab()
        tab_widget.addTab(batch_tab, "批量转换")
        
        self.logger.info("Main window UI initialized with full component set")
    
    def create_main_tab(self):
        """创建主功能页面"""
        main_tab = QWidget()
        layout = QVBoxLayout(main_tab)
        layout.setSpacing(20)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QHBoxLayout(file_group)
        
        self.file_label = QLabel("未选择文件")
        self.file_label.setMinimumWidth(300)
        self.select_btn = QPushButton("选择PPT文件")
        self.select_btn.clicked.connect(self.select_input_file)
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_btn)
        
        layout.addWidget(file_group)
        
        # 转换选项区域
        options_group = QGroupBox("转换选项")
        options_layout = QFormLayout(options_group)
        
        # 输出格式选择
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItems(["PPTX", "PDF"])
        options_layout.addRow("输出格式:", self.output_format_combo)
        
        # 转换模式选择
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["背景填充", "前景图片", "幻灯片转图片序列"])
        options_layout.addRow("转换模式:", self.mode_combo)
        
        # 图片质量滑块
        self.quality_slider = QSlider(Qt.Horizontal)
        self.quality_slider.setRange(1, 100)
        self.quality_slider.setValue(95)
        self.quality_slider.valueChanged.connect(self.on_quality_changed)
        self.quality_label = QLabel(f"图片质量: {self.quality_slider.value()}")
        quality_layout = QHBoxLayout()
        quality_layout.addWidget(self.quality_slider)
        quality_layout.addWidget(self.quality_label)
        options_layout.addRow("图片质量:", quality_layout)
        
        # 包含隐藏幻灯片选项
        self.include_hidden_checkbox = QCheckBox("包含隐藏幻灯片")
        self.include_hidden_checkbox.setChecked(False)
        options_layout.addRow("", self.include_hidden_checkbox)
        
        layout.addWidget(options_group)
        
        # 控制按钮区域
        control_layout = QHBoxLayout()
        
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_conversion)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        self.status_label = QLabel("就绪")
        
        control_layout.addWidget(self.convert_btn)
        layout.addLayout(control_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.status_label)
        
        return main_tab
    
    def create_batch_tab(self):
        """创建批量转换页面"""
        batch_tab = QWidget()
        layout = QVBoxLayout(batch_tab)
        
        # 文件列表
        list_group = QGroupBox("文件列表")
        list_layout = QVBoxLayout(list_group)
        
        self.batch_file_table = QTableWidget()
        self.batch_file_table.setColumnCount(3)
        self.batch_file_table.setHorizontalHeaderLabels(["#", "文件名", "路径"])
        header = self.batch_file_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.batch_file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        
        list_layout.addWidget(self.batch_file_table)
        
        layout.addWidget(list_group)
        
        # 批量选项
        batch_options_group = QGroupBox("批量转换选项")
        batch_options_layout = QFormLayout(batch_options_group)
        
        # 输出格式
        self.batch_output_format_combo = QComboBox()
        self.batch_output_format_combo.addItems(["PPTX", "PDF"])
        batch_options_layout.addRow("输出格式:", self.batch_output_format_combo)
        
        # 图片质量
        self.batch_quality_slider = QSlider(Qt.Horizontal)
        self.batch_quality_slider.setRange(1, 100)
        self.batch_quality_slider.setValue(95)
        self.batch_quality_label = QLabel(f"图片质量: {self.batch_quality_slider.value()}")
        batch_quality_layout = QHBoxLayout()
        batch_quality_layout.addWidget(self.batch_quality_slider)
        batch_quality_layout.addWidget(self.batch_quality_label)
        batch_options_layout.addRow("图片质量:", batch_quality_layout)
        
        # 输出目录
        batch_output_layout = QHBoxLayout()
        self.batch_output_dir_label = QLabel("输出目录: 未选择")
        self.batch_output_dir_btn = QPushButton("选择输出目录")
        self.batch_output_dir_btn.clicked.connect(self.select_batch_output_dir)
        batch_output_layout.addWidget(self.batch_output_dir_label)
        batch_output_layout.addWidget(self.batch_output_dir_btn)
        batch_options_layout.addRow("输出目录:", batch_output_layout)
        
        layout.addWidget(batch_options_group)
        
        # 批量控制按钮
        batch_control_layout = QHBoxLayout()
        
        self.add_files_btn = QPushButton("添加文件")
        self.add_files_btn.clicked.connect(self.add_batch_files)
        self.remove_selected_btn = QPushButton("移除选中")
        self.remove_selected_btn.clicked.connect(self.remove_selected_batch_files)
        self.clear_list_btn = QPushButton("清空列表")
        self.clear_list_btn.clicked.connect(self.clear_batch_file_list)
        self.batch_convert_btn = QPushButton("开始批量转换")
        self.batch_convert_btn.setEnabled(False)
        self.batch_convert_btn.clicked.connect(self.start_batch_conversion)
        
        batch_control_layout.addWidget(self.add_files_btn)
        batch_control_layout.addWidget(self.remove_selected_btn)
        batch_control_layout.addWidget(self.clear_list_btn)
        batch_control_layout.addWidget(self.batch_convert_btn)
        layout.addLayout(batch_control_layout)
        
        # 批量进度
        self.batch_progress_bar = QProgressBar()
        self.batch_progress_bar.setVisible(False)
        self.batch_status_label = QLabel("就绪")
        
        layout.addWidget(self.batch_progress_bar)
        layout.addWidget(self.batch_status_label)
        
        return batch_tab
    
    def select_input_file(self):
        """选择输入文件"""
        filename, _ = QFileDialog.getOpenFileName(
            self, "选择PPT文件", "", "PowerPoint文件 (*.ppt *.pptx)"
        )
        
        if filename:
            self.current_input_file = filename
            display_name = os.path.basename(filename)
            self.file_label.setText(display_name)
            self.file_label.setToolTip(filename)
            
            # 自动设置输出目录
            self.current_output_dir = os.path.dirname(filename)
            
            self.logger.info(f"已选择输入文件: {filename}")
            self.log_message(f"已选择文件: {display_name}")
            self.convert_btn.setEnabled(True)
    
    def on_quality_changed(self):
        """质量滑块值改变"""
        self.quality_label.setText(f"图片质量: {self.quality_slider.value()}")
        self.batch_quality_label.setText(f"图片质量: {self.batch_quality_slider.value()}")
    
    def start_conversion(self):
        """开始转换过程"""
        if not self.current_input_file:
            QMessageBox.warning(self, "警告", "请先选择输入文件！")
            return
            
        # 获取输出路径
        input_dir = os.path.dirname(self.current_input_file)
        input_name = os.path.splitext(os.path.basename(self.current_input_file))[0]
        output_ext = self.output_format_combo.currentText().lower()
        
        output_filename = f"{input_name}_converted.{output_ext}"
        output_path = os.path.join(input_dir, output_filename)
        
        # 检查文件是否存在
        counter = 1
        while os.path.exists(output_path):
            output_filename = f"{input_name}_converted({counter}).{output_ext}"
            output_path = os.path.join(input_dir, output_filename)
            counter += 1
        
        # 创建转换选项
        options = ConversionOptions()
        options.output_format = {
            "PPTX": OutputFormat.PPTX,
            "PDF": OutputFormat.PDF,
        }[self.output_format_combo.currentText()]
        options.image_quality = self.quality_slider.value()
        options.include_hidden_slides = self.include_hidden_checkbox.isChecked()
        
        # 设置转换模式
        mode_text = self.mode_combo.currentText()
        options.mode = {
            "背景填充": ConversionMode.BACKGROUND_FILL,
            "前景图片": ConversionMode.FOREGROUND_IMAGE,
            "幻灯片转图片序列": ConversionMode.SLIDE_TO_IMAGE
        }[mode_text]
        
        self.logger.info(f"开始转换任务")
        self.logger.info(f"输入: {self.current_input_file}")
        self.logger.info(f"输出: {output_path}")
        
        # 禁用按钮
        self.convert_btn.setEnabled(False)
        self.status_label.setText("正在转换...")
        
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # 创建工作线程
        self.worker = ConversionWorker(
            self.converter, 
            self.current_input_file, 
            output_path, 
            options
        )
        self.worker.progress_updated.connect(self._update_progress_bar)
        self.worker.conversion_finished.connect(self._on_conversion_finished)
        self.worker.start()
    
    def _update_progress_bar(self, value: float, task: str):
        """更新进度条和状态"""
        self.progress_bar.setValue(int(value * 100))
        self.status_label.setText(task)
    
    def _on_conversion_finished(self, success: bool, output_path: str):
        """转换完成处理"""
        self.progress_bar.setVisible(False)
        
        if success:
            self.status_label.setText("转换成功！")
            self.log_message(f"🎉 转换成功：{output_path}")
            
            # 获取文件大小
            if os.path.exists(output_path):
                size_mb = os.path.getsize(output_path) / 1024 / 1024
                self.log_message(f"文件大小: {size_mb:.2f} MB")
                
                # 添加到历史记录
                self.history_manager.add_record(
                    input_path=self.current_input_file,
                    output_path=output_path,
                    mode=self.mode_combo.currentText(),
                    output_format=self.output_format_combo.currentText(),
                    success=True,
                    file_size=size_mb
                )
            
            QMessageBox.information(self, "成功", f"转换成功！\n输出文件：{output_path}")
        else:
            self.status_label.setText("转换失败")
            self.log_message(f"❌ 转换失败：{output_path}")
            QMessageBox.critical(self, "错误", f"转换失败！\n详细信息请查看日志")
        
        # 重新启用按钮
        self.convert_btn.setEnabled(True)
    
    def add_batch_files(self):
        """添加批量文件"""
        filenames, _ = QFileDialog.getOpenFileNames(
            self, "选择PPT文件", "", "PowerPoint文件 (*.ppt *.pptx)"
        )
        
        if filenames:
            start_row = self.batch_file_table.rowCount()
            self.batch_file_table.setRowCount(start_row + len(filenames))
            
            for i, filename in enumerate(filenames):
                row = start_row + i
                self.batch_file_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
                self.batch_file_table.setItem(row, 1, QTableWidgetItem(os.path.basename(filename)))
                self.batch_file_table.setItem(row, 2, QTableWidgetItem(filename))
            
            self.batch_convert_btn.setEnabled(self.batch_file_table.rowCount() > 0)
            self.logger.info(f"添加了 {len(filenames)} 个批量文件")
    
    def remove_selected_batch_files(self):
        """移除选中的批量文件"""
        selected_rows = []
        for item in self.batch_file_table.selectedItems():
            if item.row() not in selected_rows:
                selected_rows.append(item.row())
        
        # 从后往前删除，避免索引变化问题
        for row in sorted(selected_rows, reverse=True):
            self.batch_file_table.removeRow(row)
        
        # 重新编号
        for row in range(self.batch_file_table.rowCount()):
            self.batch_file_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        
        self.batch_convert_btn.setEnabled(self.batch_file_table.rowCount() > 0)
        self.logger.info(f"移除了 {len(selected_rows)} 个批量文件")
    
    def clear_batch_file_list(self):
        """清空批量文件列表"""
        self.batch_file_table.setRowCount(0)
        self.batch_convert_btn.setEnabled(False)
        self.logger.info("清空了批量文件列表")
    
    def select_batch_output_dir(self):
        """选择批量输出目录"""
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if output_dir:
            self.current_output_dir = output_dir
            self.batch_output_dir_label.setText(f"输出目录: {os.path.basename(output_dir)}")
            self.logger.info(f"选择了批量输出目录: {output_dir}")
    
    def start_batch_conversion(self):
        """开始批量转换"""
        if self.batch_file_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请先添加要转换的文件！")
            return
        
        output_dir = self.current_output_dir or os.path.dirname(self.batch_file_table.item(0, 2).text())
        
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录！")
            return
        
        # 获取文件列表
        file_list = []
        for row in range(self.batch_file_table.rowCount()):
            file_path = self.batch_file_table.item(row, 2).text()
            if os.path.exists(file_path):
                file_list.append(file_path)
        
        if not file_list:
            QMessageBox.warning(self, "警告", "没有有效的文件需要转换！")
            return
        
        # 创建转换选项
        options = ConversionOptions()
        options.output_format = {
            "PPTX": OutputFormat.PPTX,
            "PDF": OutputFormat.PDF,
        }[self.batch_output_format_combo.currentText()]
        options.image_quality = self.batch_quality_slider.value()
        
        # 设置转换模式
        mode_text = self.mode_combo.currentText()  # 使用主页面的模式选择
        options.mode = {
            "背景填充": ConversionMode.BACKGROUND_FILL,
            "前景图片": ConversionMode.FOREGROUND_IMAGE,
            "幻灯片转图片序列": ConversionMode.SLIDE_TO_IMAGE
        }[mode_text]
        
        self.logger.info(f"开始批量转换，文件数: {len(file_list)}")
        
        # 禁用按钮
        self.batch_convert_btn.setEnabled(False)
        self.add_files_btn.setEnabled(False)
        
        # 显示进度条
        self.batch_progress_bar.setVisible(True)
        self.batch_progress_bar.setRange(0, len(file_list))
        self.batch_progress_bar.setValue(0)
        self.batch_status_label.setText("准备开始转换...")
        
        # 创建批量转换工作线程
        self.batch_worker = BatchConversionWorker(
            self.converter,
            file_list,
            output_dir,
            options
        )
        self.batch_worker.progress_updated.connect(self._update_batch_progress)
        self.batch_worker.file_finished.connect(self._on_batch_file_finished)
        self.batch_worker.conversion_finished.connect(self._on_batch_finished)
        self.batch_worker.start()
    
    def _update_batch_progress(self, current: int, total: int, status: str):
        """更新批量转换进度"""
        self.batch_progress_bar.setMaximum(total)
        self.batch_progress_bar.setValue(current)
        self.batch_status_label.setText(status)
    
    def _on_batch_file_finished(self, index: int, success: bool, output_path: str, error: str):
        """批量转换单个文件完成"""
        if success:
            self.logger.info(f"批量转换文件成功: {output_path}")
            self.log_message(f"✅ 批量转换成功: {os.path.basename(output_path)}")
        else:
            self.logger.error(f"批量转换文件失败: {error}")
            self.log_message(f"❌ 批量转换失败: {error}")
    
    def _on_batch_finished(self, success_count: int, fail_count: int):
        """批量转换完成"""
        self.batch_progress_bar.setVisible(False)
        self.batch_status_label.setText(f"批量转换完成: {success_count} 成功, {fail_count} 失败")
        
        # 重新启用按钮
        self.batch_convert_btn.setEnabled(True)
        self.add_files_btn.setEnabled(True)
        
        QMessageBox.information(
            self, 
            "批量转换完成", 
            f"批量转换完成！\n成功: {success_count}\n失败: {fail_count}"
        )
    
    def log_message(self, message: str):
        """记录日志消息"""
        timestamp = __import__('datetime').datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")

def main():
    # 设置应用属性（必须在创建QApplication之前）
    from PyQt5.QtWidgets import QApplication
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    # 设置全局字体
    font = QFont("Microsoft YaHei")
    font.setPointSize(9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
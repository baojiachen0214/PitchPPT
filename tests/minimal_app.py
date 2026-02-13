import sys
import os

# 添加项目根目录到sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QProgressBar, QGroupBox, QFormLayout
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import tempfile

from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat
from src.utils.logger import Logger
from src.utils.config_manager import ConfigManager
from src.utils.history_manager import HistoryManager

class SimpleMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logger = Logger().get_logger()
        self.config_manager = ConfigManager()
        self.history_manager = HistoryManager()
        
        # 初始化转换器
        try:
            self.converter = Win32PPTConverter()
            self.logger.info("转换器初始化成功")
        except Exception as e:
            self.logger.error(f"转换器初始化失败: {e}")
            self.converter = None
        
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("PitchPPT - 简化版")
        self.setGeometry(100, 100, 800, 600)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建布局
        layout = QVBoxLayout(central_widget)
        
        # 标题
        title_label = QLabel("PitchPPT - PPT转换工具")
        title_label.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 功能区域
        tabs = QTabWidget()
        
        # 单文件转换标签页
        single_tab = self.create_single_tab()
        tabs.addTab(single_tab, "单文件转换")
        
        # 批量转换标签页
        batch_tab = self.create_batch_tab()
        tabs.addTab(batch_tab, "批量转换")
        
        layout.addWidget(tabs)
        
        # 状态栏
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("就绪")
        
        self.logger.info("UI初始化完成")
    
    def create_single_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 选择文件
        select_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.select_btn = QPushButton("选择PPT文件")
        self.select_btn.clicked.connect(self.select_file)
        select_layout.addWidget(self.file_label)
        select_layout.addWidget(self.select_btn)
        layout.addLayout(select_layout)
        
        # 转换按钮
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.clicked.connect(self.start_single_conversion)
        layout.addWidget(self.convert_btn)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        return widget
    
    def create_batch_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 文件列表
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(2)
        self.file_table.setHorizontalHeaderLabels(["文件名", "路径"])
        header = self.file_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.file_table)
        
        # 控制按钮
        button_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("添加文件")
        self.add_files_btn.clicked.connect(self.add_batch_files)
        self.remove_file_btn = QPushButton("移除选中")
        self.remove_file_btn.clicked.connect(self.remove_selected_files)
        self.clear_all_btn = QPushButton("清空列表")
        self.clear_all_btn.clicked.connect(self.clear_file_list)
        self.batch_convert_btn = QPushButton("批量转换")
        self.batch_convert_btn.clicked.connect(self.start_batch_conversion)
        
        button_layout.addWidget(self.add_files_btn)
        button_layout.addWidget(self.remove_file_btn)
        button_layout.addWidget(self.clear_all_btn)
        button_layout.addWidget(self.batch_convert_btn)
        layout.addLayout(button_layout)
        
        # 进度条
        self.batch_progress_bar = QProgressBar()
        self.batch_progress_bar.setVisible(False)
        layout.addWidget(self.batch_progress_bar)
        
        return widget
    
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PPT文件", "", "PowerPoint文件 (*.ppt *.pptx)"
        )
        if file_path:
            self.input_file = file_path
            self.file_label.setText(file_path)
            self.convert_btn.setEnabled(True)
            self.status_bar.showMessage(f"已选择: {os.path.basename(file_path)}")
    
    def add_batch_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择PPT文件", "", "PowerPoint文件 (*.ppt *.pptx)"
        )
        if file_paths:
            for file_path in file_paths:
                row_position = self.file_table.rowCount()
                self.file_table.insertRow(row_position)
                self.file_table.setItem(row_position, 0, QTableWidgetItem(os.path.basename(file_path)))
                self.file_table.setItem(row_position, 1, QTableWidgetItem(file_path))
            
            self.status_bar.showMessage(f"添加了 {len(file_paths)} 个文件")
    
    def remove_selected_files(self):
        selected_rows = []
        for item in self.file_table.selectedItems():
            if item.row() not in selected_rows:
                selected_rows.append(item.row())
        
        # 从后往前删除，避免索引变化问题
        for row in sorted(selected_rows, reverse=True):
            self.file_table.removeRow(row)
        
        self.status_bar.showMessage("已移除选中文件")
    
    def clear_file_list(self):
        self.file_table.setRowCount(0)
        self.status_bar.showMessage("文件列表已清空")
    
    def start_single_conversion(self):
        if hasattr(self, 'input_file'):
            self.status_bar.showMessage("正在转换...")
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # 创建转换选项
            options = ConversionOptions()
            options.mode = ConversionMode.BACKGROUND_FILL
            options.output_format = OutputFormat.PPTX
            options.image_quality = 95
            
            try:
                # 执行转换
                output_path = self.input_file.replace('.ppt', '_converted.ppt').replace('.pptx', '_converted.pptx')
                success = self.converter.convert(self.input_file, output_path, options)
                
                if success:
                    self.status_bar.showMessage(f"转换成功: {output_path}")
                    self.logger.info(f"转换成功: {self.input_file} -> {output_path}")
                    
                    # 添加到历史记录
                    self.history_manager.add_record(
                        input_path=self.input_file,
                        output_path=output_path,
                        mode=options.mode.value,
                        output_format=options.output_format.value,
                        success=True
                    )
                else:
                    self.status_bar.showMessage("转换失败")
                    self.logger.error(f"转换失败: {self.input_file}")
            except Exception as e:
                self.status_bar.showMessage(f"转换错误: {str(e)}")
                self.logger.error(f"转换过程中出错: {e}")
            
            self.progress_bar.setVisible(False)
        else:
            self.status_bar.showMessage("请先选择文件")
    
    def start_batch_conversion(self):
        if self.file_table.rowCount() == 0:
            self.status_bar.showMessage("请先添加文件")
            return
        
        self.status_bar.showMessage("批量转换中...")
        self.batch_progress_bar.setVisible(True)
        self.batch_progress_bar.setRange(0, self.file_table.rowCount())
        self.batch_progress_bar.setValue(0)
        
        # 创建转换选项
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        options.output_format = OutputFormat.PPTX
        options.image_quality = 95
        
        success_count = 0
        total_count = self.file_table.rowCount()
        
        for row in range(total_count):
            item = self.file_table.item(row, 1)  # 获取路径列
            if item:
                input_path = item.text()
                output_path = input_path.replace('.ppt', '_converted.ppt').replace('.pptx', '_converted.pptx')
                
                try:
                    success = self.converter.convert(input_path, output_path, options)
                    if success:
                        success_count += 1
                        self.logger.info(f"批量转换成功: {input_path}")
                        
                        # 添加到历史记录
                        self.history_manager.add_record(
                            input_path=input_path,
                            output_path=output_path,
                            mode=options.mode.value,
                            output_format=options.output_format.value,
                            success=True
                        )
                    else:
                        self.logger.error(f"批量转换失败: {input_path}")
                except Exception as e:
                    self.logger.error(f"批量转换出错 {input_path}: {e}")
            
            # 更新进度
            self.batch_progress_bar.setValue(row + 1)
        
        self.status_bar.showMessage(f"批量转换完成: {success_count}/{total_count} 成功")
        self.batch_progress_bar.setVisible(False)

def main():
    # 设置应用属性（必须在创建QApplication之前）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    # 设置全局字体
    font = QFont("Microsoft YaHei")
    font.setPointSize(9)  # 使用较小的字体
    app.setFont(font)
    
    window = SimpleMainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
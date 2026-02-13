import unittest
import os
import tempfile
from unittest.mock import MagicMock, patch
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt, QMimeData, QUrl
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from src.ui.main_window import MainWindow

class TestMainWindow(unittest.TestCase):
    """
    MainWindow单元测试
    """
    
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication([])
    
    def setUp(self):
        self.window = MainWindow()
    
    def tearDown(self):
        self.window.close()
        
    def test_window_initialization(self):
        """
        测试窗口初始化
        """
        self.assertEqual(self.window.windowTitle(), "PitchPPT - 专业路演PPT处理工具")
        self.assertGreater(self.window.width(), 0)
        self.assertGreater(self.window.height(), 0)
        
    def test_file_label_initial_state(self):
        """
        测试文件标签初始状态
        """
        self.assertEqual(self.window.file_label.text(), "未选择文件")
        
    def test_quality_slider_initial_value(self):
        """
        测试质量滑块初始值
        """
        self.assertEqual(self.window.quality_slider.value(), 95)
        self.assertEqual(self.window.quality_value.text(), "95%")
        
    def test_quality_slider_change(self):
        """
        测试质量滑块值变化
        """
        self.window.quality_slider.setValue(80)
        self.assertEqual(self.window.quality_value.text(), "80%")
        
    def test_select_input_file_dialog(self):
        """
        测试选择输入文件对话框
        """
        with patch('src.ui.main_window.QFileDialog.getOpenFileName') as mock_dialog:
            mock_dialog.return_value = ("/path/to/test.pptx", "*.pptx")
            
            # 模拟按钮点击
            select_btn = self.window.findChild(type(self.window.convert_btn), "selectButton")
            if select_btn:
                QTest.mouseClick(select_btn, Qt.LeftButton)
            
            mock_dialog.assert_called_once()
            self.assertEqual(self.window.current_input_file, "/path/to/test.pptx")
            self.assertIn("test.pptx", self.window.file_label.text())
            
    def test_convert_button_initial_state(self):
        """
        测试转换按钮初始状态
        """
        self.assertFalse(self.window.convert_btn.isEnabled())
        
    def test_log_message(self):
        """
        测试日志消息功能
        """
        initial_text = self.window.log_text.toPlainText()
        self.window.log_message("Test message")
        final_text = self.window.log_text.toPlainText()
        
        self.assertIn("Test message", final_text)
        self.assertNotEqual(initial_text, final_text)
        
    def test_drag_enter_event_accepted(self):
        """
        测试拖拽进入事件接受
        """
        from PyQt5.QtCore import QMimeData, QPoint
        from PyQt5.QtGui import QDragEnterEvent
        
        mime_data = QMimeData()
        mime_data.setUrls([QUrl.fromLocalFile("test.pptx")])
        event = QDragEnterEvent(QPoint(0, 0), Qt.CopyAction, mime_data, Qt.LeftButton, Qt.NoModifier)
        
        with patch.object(event, 'acceptProposedAction') as mock_accept:
            self.window.dragEnterEvent(event)
            mock_accept.assert_called_once()
        
    def test_drop_event_ppt_file(self):
        """
        测试拖放PPT文件
        """
        from PyQt5.QtCore import QMimeData, QPoint
        from PyQt5.QtGui import QDropEvent
        
        ppt_path = os.path.join(tempfile.gettempdir(), "test.pptx")
        with open(ppt_path, 'w') as f:
            f.write("Mock PPT content")
        
        try:
            mime_data = QMimeData()
            mime_data.setUrls([QUrl.fromLocalFile(ppt_path)])
            event = QDropEvent(QPoint(0, 0), Qt.CopyAction, mime_data, Qt.LeftButton, Qt.NoModifier)
            
            with patch('src.ui.main_window.QMessageBox.warning') as mock_warning:
                self.window.dropEvent(event)
                mock_warning.assert_not_called()
                
            self.assertEqual(os.path.normpath(self.window.current_input_file), os.path.normpath(ppt_path))
            self.assertIn("test.pptx", self.window.file_label.text())
            
        finally:
            if os.path.exists(ppt_path):
                os.remove(ppt_path)
        
    def test_drop_event_non_ppt_file(self):
        """
        测试拖放非PPT文件
        """
        from PyQt5.QtCore import QMimeData, QPoint
        from PyQt5.QtGui import QDropEvent
        
        txt_path = os.path.join(tempfile.gettempdir(), "test.txt")
        with open(txt_path, 'w') as f:
            f.write("Test content")
        
        try:
            mime_data = QMimeData()
            mime_data.setUrls([QUrl.fromLocalFile(txt_path)])
            event = QDropEvent(QPoint(0, 0), Qt.CopyAction, mime_data, Qt.LeftButton, Qt.NoModifier)
            
            with patch('src.ui.main_window.QMessageBox.warning') as mock_warning:
                self.window.dropEvent(event)
                mock_warning.assert_called_once()
                
        finally:
            if os.path.exists(txt_path):
                os.remove(txt_path)
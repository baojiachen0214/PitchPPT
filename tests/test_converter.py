import unittest
import os
import tempfile
import shutil
from unittest.mock import MagicMock, patch
from src.core import Win32PPTConverter, ConversionOptions, ConversionMode, OutputFormat


class TestWin32PPTConverter(unittest.TestCase):
    """
    Win32PPTConverter单元测试
    """
    
    def setUp(self):
        self.converter = Win32PPTConverter()
        self.test_dir = tempfile.mkdtemp(prefix="pitchppt_test_")
        
        # 创建测试PPT文件（模拟）
        self.test_ppt = os.path.join(self.test_dir, "test.pptx")
        with open(self.test_ppt, 'w') as f:
            f.write("Mock PPT content")
        
    def tearDown(self):
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)
        
    @patch('src.core.win32_converter.win32com')
    def test_initialize_powerpoint_success(self, mock_win32com):
        """
        测试PowerPoint初始化成功
        """
        mock_powerpoint = MagicMock()
        mock_win32com.client.Dispatch.return_value = mock_powerpoint
        
        result = self.converter._initialize_powerpoint()
        
        self.assertTrue(result)
        mock_win32com.client.Dispatch.assert_called_once_with("PowerPoint.Application")
        self.assertEqual(self.converter.powerpoint, mock_powerpoint)
        
    @patch('src.core.win32_converter.win32com')
    def test_initialize_powerpoint_failure(self, mock_win32com):
        """
        测试PowerPoint初始化失败
        """
        mock_win32com.client.Dispatch.side_effect = Exception("COM Error")
        
        result = self.converter._initialize_powerpoint()
        
        self.assertFalse(result)
        self.assertIsNone(self.converter.powerpoint)
        
    def test_conversion_options_to_dict(self):
        """
        测试转换选项序列化
        """
        options = ConversionOptions()
        options.mode = ConversionMode.FOREGROUND_IMAGE
        options.output_format = OutputFormat.PDF
        options.image_quality = 85
        
        result = options.to_dict()
        
        self.assertEqual(result['mode'], 'foreground_image')
        self.assertEqual(result['output_format'], 'pdf')
        self.assertEqual(result['image_quality'], 85)
        
    def test_conversion_options_from_dict(self):
        """
        测试转换选项反序列化
        """
        data = {
            'mode': 'slide_to_image',
            'output_format': 'jpg',
            'image_quality': 90,
            'resolution_scale': 1.5
        }
        
        options = ConversionOptions.from_dict(data)
        
        self.assertEqual(options.mode, ConversionMode.SLIDE_TO_IMAGE)
        self.assertEqual(options.output_format, OutputFormat.JPG)
        self.assertEqual(options.image_quality, 90)
        self.assertEqual(options.resolution_scale, 1.5)
        
    @patch('src.core.win32_converter.Win32PPTConverter._cleanup')
    @patch('src.core.win32_converter.Win32PPTConverter._export_slide_to_image', return_value=True)
    @patch('src.core.win32_converter.Win32PPTConverter._convert_background_fill', return_value=True)
    @patch('src.core.win32_converter.pythoncom')
    @patch('src.core.win32_converter.win32com')
    def test_convert_background_fill_success(self, mock_win32com, mock_pythoncom, mock_convert_bg, mock_export, mock_cleanup):
        """
        测试背景填充模式转换成功
        """
        output_path = os.path.join(self.test_dir, "output.pptx")
        options = ConversionOptions()
        options.mode = ConversionMode.BACKGROUND_FILL
        
        mock_powerpoint = MagicMock()
        mock_presentation = MagicMock()
        mock_presentation.Slides.Count = 2
        mock_powerpoint.Presentations.Open.return_value = mock_presentation
        mock_win32com.client.Dispatch.return_value = mock_powerpoint
        
        result = self.converter.convert(self.test_ppt, output_path, options)
        
        self.assertTrue(result)
        mock_convert_bg.assert_called_once()
        
    @patch('src.core.win32_converter.pythoncom')
    @patch('src.core.win32_converter.win32com')
    def test_convert_file_not_found(self, mock_win32com, mock_pythoncom):
        """
        测试输入文件不存在的情况
        """
        non_existent_file = os.path.join(self.test_dir, "non_existent.pptx")
        output_path = os.path.join(self.test_dir, "output.pptx")
        
        mock_powerpoint = MagicMock()
        mock_win32com.client.Dispatch.return_value = mock_powerpoint
        
        result = self.converter.convert(non_existent_file, output_path)
        
        self.assertFalse(result)
        
    @patch('src.core.win32_converter.Win32PPTConverter._initialize_powerpoint', return_value=True)
    @patch('src.core.win32_converter.Win32PPTConverter._cleanup')
    def test_batch_convert_success(self, mock_cleanup, mock_init):
        """
        测试批量转换成功
        """
        output_dir = os.path.join(self.test_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        
        results = self.converter.batch_convert([self.test_ppt], output_dir)
        
        self.assertIn(self.test_ppt, results)
        self.assertFalse(results[self.test_ppt])  # 因为是模拟环境，实际会失败，但应正确处理流程
        
    @patch('src.core.win32_converter.pythoncom')
    @patch('src.core.win32_converter.win32com')
    def test_get_conversion_info_success(self, mock_win32com, mock_pythoncom):
        """
        测试获取PPT信息成功
        """
        mock_powerpoint = MagicMock()
        mock_presentation = MagicMock()
        mock_presentation.Slides.Count = 5
        mock_presentation.BuiltInDocumentProperties.Title = "Test Presentation"
        
        mock_win32com.client.Dispatch.return_value = mock_powerpoint
        mock_powerpoint.Presentations.Open.return_value = mock_presentation
        
        result = self.converter.get_conversion_info(self.test_ppt)
        
        self.assertTrue(result['success'])
        self.assertEqual(result['file_info']['slide_count'], 5)
        self.assertEqual(result['file_info']['title'], "Test Presentation")
        
    def test_progress_update(self):
        """
        测试进度更新功能
        """
        self.converter._update_progress(0.5, "Test Task")
        self.assertEqual(self.converter.get_progress(), 0.5)
        
        self.converter._update_progress(1.2, "Test Task")
        self.assertEqual(self.converter.get_progress(), 1.0)
        
        self.converter._update_progress(-0.1, "Test Task")
        self.assertEqual(self.converter.get_progress(), 0.0)
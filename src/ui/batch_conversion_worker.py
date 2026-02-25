﻿import os
import logging
import threading
from typing import List, Callable, Optional
from PyQt5.QtCore import QThread, pyqtSignal
from PIL import Image

from src.core.win32_converter import Win32PPTConverter
from src.core.smart_optimizer_v4 import SmartOptimizerV4
from src.core.smart_optimizer_v5 import SmartOptimizerV5
from src.core.smart_optimizer_v6 import SmartOptimizerV6
from src.core.smart_optimizer_v7 import SmartOptimizerV7
from src.core.smart_optimizer_v8 import SmartOptimizerV8
from src.core.converter import ConversionOptions, ConversionMode


class BatchConversionWorker(QThread):
    """Batch conversion worker thread."""

    file_started = pyqtSignal(str)
    file_progress = pyqtSignal(str, float, str)
    file_finished = pyqtSignal(str, bool, str)
    batch_progress = pyqtSignal(float)
    batch_finished = pyqtSignal(dict)

    def __init__(self, 
                 files: List[str],
                 options: ConversionOptions,
                 output_dir: str,
                 is_smart_mode: bool = False,
                 target_size_mb: float = 10.0,
                 algorithm: str = "v4",
                 logger: Optional[logging.Logger] = None):
        super().__init__()
        
        self.files = files
        self.options = options
        self.output_dir = output_dir
        self.is_smart_mode = is_smart_mode
        self.target_size_mb = target_size_mb
        self.algorithm = algorithm
        self.logger = logger or logging.getLogger(__name__)
        
        self._paused = False
        self._stopped = False
        self._pause_condition = threading.Condition()
        
        # 鍏变韩鐨凱owerPoint瀹炰緥锛堢敤浜庢壒澶勭悊澶嶇敤锛?        self._shared_converter = None
        self._shared_optimizer = None

    def pause(self):
        self._paused = True
        self.logger.info("鎵瑰鐞嗗凡鏆傚仠")

    def resume(self):
        with self._pause_condition:
            self._paused = False
            self._pause_condition.notify_all()
        self.logger.info("Batch conversion resumed")

    def stop(self):
        self._stopped = True
        self.resume()
        self.logger.info("鎵瑰鐞嗗凡缁堟")

    def _check_pause(self):
        with self._pause_condition:
            while self._paused and not self._stopped:
                self._pause_condition.wait(0.1)

    def run(self):
        try:
            results = []
            total_files = len(self.files)
            
            self.logger.info(f"Batch conversion started, total {total_files} files")
            
            if self.is_smart_mode:
                # Smart mode: initialize shared converter and optimizer
                self._shared_converter = Win32PPTConverter()
                self._shared_converter._initialize_powerpoint()
                
                # 鏍规嵁绠楁硶閫夋嫨鍒涘缓瀵瑰簲鐨刼ptimizer
                include_hidden = self.options.include_hidden_slides
                if self.algorithm == "v8":
                    self._shared_optimizer = SmartOptimizerV8(
                        converter=self._shared_converter,
                        include_hidden_slides=include_hidden
                    )
                    self.logger.info("Using V8 joint optimization algorithm")
                elif self.algorithm == "v7":
                    self._shared_optimizer = SmartOptimizerV7(
                        converter=self._shared_converter,
                        include_hidden_slides=include_hidden
                    )
                    self.logger.info("浣跨敤棰勭畻椹卞姩鎰熺煡浼樺寲绠楁硶 V7")
                elif self.algorithm == "v6":
                    self._shared_optimizer = SmartOptimizerV6(
                        converter=self._shared_converter,
                        include_hidden_slides=include_hidden
                    )
                    self.logger.info("浣跨敤澶嶆潅搴﹁嚜閫傚簲绠楁硶 V6")
                elif self.algorithm == "v5":
                    self._shared_optimizer = SmartOptimizerV5(
                        converter=self._shared_converter,
                        include_hidden_slides=include_hidden
                    )
                    self.logger.info("浣跨敤鍧囪　鐢昏川璐績绠楁硶 V5")
                else:
                    self._shared_optimizer = SmartOptimizerV4(
                        logger=self.logger,
                        converter=self._shared_converter,
                        include_hidden_slides=include_hidden
                    )
                    self.logger.info("浣跨敤鍧囧垎閰嶉璐績绠楁硶 V4")
            else:
                self._shared_converter = Win32PPTConverter()
                self._shared_converter._initialize_powerpoint()
            
            self.logger.info("PowerPoint瀹炰緥宸查鍒濆鍖栵紝寮€濮嬪鐞嗘枃浠?..")
            
            for i, input_file in enumerate(self.files):
                if self._stopped:
                    self.logger.info("鎵瑰鐞嗚鐢ㄦ埛缁堟")
                    break
                
                self._check_pause()
                
                self.file_started.emit(input_file)
                self.logger.info(f"[{i+1}/{total_files}] 寮€濮嬪鐞? {os.path.basename(input_file)}")
                
                input_name = os.path.splitext(os.path.basename(input_file))[0]
                output_ext = self.options.output_format.value if self.options.output_format else "pptx"
                output_path = os.path.join(self.output_dir, f"{input_name}_converted.{output_ext}")
                
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(self.output_dir, f"{input_name}_converted_{counter}.{output_ext}")
                    counter += 1
                
                try:
                    if self.is_smart_mode:
                        success = self._process_smart_mode(input_file, output_path)
                    else:
                        success = self._process_normal_mode(input_file, output_path)
                    
                    self.file_finished.emit(input_file, success, output_path if success else "")
                    results.append({
                        'file': input_file,
                        'success': success,
                        'output': output_path if success else ""
                    })
                    
                    if success:
                        self.logger.info(f"[{i+1}/{total_files}] 鉁?鎴愬姛: {os.path.basename(input_file)}")
                    else:
                        self.logger.warning(f"[{i+1}/{total_files}] 鉁?澶辫触: {os.path.basename(input_file)}")
                    
                except Exception as e:
                    self.logger.error(f"[{i+1}/{total_files}] 寮傚父: {os.path.basename(input_file)} - {e}")
                    self.file_finished.emit(input_file, False, "")
                    results.append({
                        'file': input_file,
                        'success': False,
                        'output': "",
                        'error': str(e)
                    })
                
                progress = (i + 1) / total_files
                self.batch_progress.emit(progress)
            
            success_count = sum(1 for r in results if r['success'])
            failed_count = sum(1 for r in results if not r['success'])
            
            self.logger.info(f"鎵瑰鐞嗗畬鎴? 鎴愬姛 {success_count}/{total_files}, 澶辫触 {failed_count}/{total_files}")
            
            self.batch_finished.emit({
                'total': total_files,
                'success': success_count,
                'failed': failed_count,
                'results': results
            })
            
        except Exception as e:
            self.logger.error(f"鎵瑰鐞嗙嚎绋嬪紓甯? {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.batch_finished.emit({
                'total': len(self.files),
                'success': 0,
                'failed': len(self.files),
                'results': [],
                'error': str(e)
            })
        finally:
            # 鎵瑰鐞嗗畬鎴愬悗缁熶竴娓呯悊
            self._cleanup_shared_instances()

    def _cleanup_shared_instances(self):
        """娓呯悊鍏变韩瀹炰緥"""
        if self._shared_converter:
            try:
                self._shared_converter._cleanup(force_kill=False)
                self.logger.info("Shared converter cleaned")
            except Exception as e:
                self.logger.warning(f"娓呯悊鍏变韩Converter鏃跺嚭閿? {e}")
            finally:
                self._shared_converter = None
        
        if self._shared_optimizer:
            try:
                self._shared_optimizer._cleanup()
                self.logger.info("Shared optimizer cleaned")
            except Exception as e:
                self.logger.warning(f"娓呯悊鍏变韩Optimizer鏃跺嚭閿? {e}")
            finally:
                self._shared_optimizer = None

    def _process_normal_mode(self, input_file: str, output_path: str) -> bool:
        """浣跨敤鍏变韩Converter澶勭悊鍗曚釜鏂囦欢"""
        def progress_callback(progress: float, message: str):
            self.file_progress.emit(input_file, progress, message)
        
        # 浣跨敤鍏变韩鐨刢onverter锛屼絾璁剧疆涓存椂鐨勮繘搴﹀洖璋?        original_callback = self._shared_converter._progress_callback
        original_tracker_callback = self._shared_converter._progress_tracker.callback
        try:
            self._shared_converter._progress_callback = progress_callback
            self._shared_converter._progress_tracker.callback = progress_callback
            success = self._shared_converter.convert(input_file, output_path, self.options)
            return success
        finally:
            # 鎭㈠鍘熷鍥炶皟
            self._shared_converter._progress_callback = original_callback
            self._shared_converter._progress_tracker.callback = original_tracker_callback

    def _process_smart_mode(self, input_file: str, output_path: str) -> bool:
        """浣跨敤鍏变韩Optimizer澶勭悊鍗曚釜鏂囦欢"""
        def progress_callback(message: str, progress: int):
            self.file_progress.emit(input_file, progress / 100.0, message)
        
        # 淇濆瓨鍘熷鍥炶皟骞惰缃柊鐨勫洖璋?        original_callback = self._shared_converter._progress_callback
        original_tracker_callback = self._shared_converter._progress_tracker.callback
        
        def converter_progress_callback(progress: float, message: str):
            self.file_progress.emit(input_file, progress, message)
        
        try:
            # 璁剧疆converter鐨勮繘搴﹀洖璋?            self._shared_converter._progress_callback = converter_progress_callback
            self._shared_converter._progress_tracker.callback = converter_progress_callback
            
            # Step 1: 璁＄畻浼樺寲鍙傛暟
            self.logger.info(f"[鏅鸿兘妯″紡] 寮€濮嬭绠椾紭鍖栧弬鏁? {input_file}")
            result = self._shared_optimizer.optimize(
                input_file, 
                self.target_size_mb,
                progress_callback
            )
            
            if not result.success:
                self.logger.error(f"[鏅鸿兘妯″紡] 浼樺寲鍙傛暟璁＄畻澶辫触: {result.message}")
                return False
            
            self.logger.info(f"[Smart] optimization finished, generating final file")
            
            # Step 2: 浣跨敤浼樺寲鍙傛暟鐢熸垚鏈€缁圥PT鏂囦欢
            if result.page_results:
                page_heights = [r.optimal_height for r in result.page_results]
                page_image_settings = [
                    {
                        'format': getattr(r, 'image_format', 'png'),
                        'quality': int(getattr(r, 'image_quality', 95))
                    }
                    for r in result.page_results
                ]
                aspect_ratio = result.aspect_ratio
                
                success = self._export_with_page_heights(
                    input_file,
                    output_path,
                    page_heights,
                    aspect_ratio,
                    progress_callback,
                    page_image_settings
                )
                
                if success:
                    self.logger.info(f"[鏅鸿兘妯″紡] 鏈€缁堟枃浠跺凡鐢熸垚: {output_path}")
                else:
                    self.logger.error(f"[Smart] final file generation failed")
                
                return success
            else:
                self.logger.error(f"[Smart] no valid optimization result")
                return False
                
        except Exception as e:
            self.logger.error(f"鏅鸿兘浼樺寲澶辫触: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
        finally:
            # 鎭㈠鍘熷鍥炶皟
            self._shared_converter._progress_callback = original_callback
            self._shared_converter._progress_tracker.callback = original_tracker_callback
    
    def _export_with_page_heights(self, input_file: str, output_path: str,
                                   page_heights: list, aspect_ratio: float,
                                   progress_callback=None, page_image_settings: Optional[List[dict]] = None) -> bool:
        """浣跨敤姣忛〉鐨勬渶浼橀珮搴﹀鍑哄畬鏁碢PT"""
        import tempfile
        import shutil
        import time
        
        try:
            self.logger.info(f"[瀵煎嚭] 浣跨敤浼樺寲楂樺害瀵煎嚭PPT: {output_path}")
            
            # 浣跨敤鍏变韩鐨刢onverter
            converter = self._shared_converter
            
            # 纭繚PowerPoint瀹炰緥鍙敤
            if not converter.powerpoint:
                self.logger.error("[Export] PowerPoint instance unavailable")
                return False
            
            # 绛夊緟涓€灏忔鏃堕棿锛岀‘淇濅箣鍓嶇殑鎿嶄綔宸插畬鎴?            time.sleep(0.3)
            
            # 鎵撳紑鍘熷PPT
            abs_path = os.path.abspath(input_file)
            self.logger.info(f"[瀵煎嚭] 鎵撳紑鏂囦欢: {abs_path}")
            pres = converter.powerpoint.Presentations.Open(abs_path)
            
            try:
                slide_count = pres.Slides.Count
                slide_aspect_ratio = pres.PageSetup.SlideWidth / pres.PageSetup.SlideHeight
                self.logger.info(f"[瀵煎嚭] 骞荤伅鐗囨暟閲? {slide_count}")
                
                # 鍒涘缓涓存椂鐩綍
                temp_dir = converter._create_temp_dir(prefix="smart_export_")
                self.logger.info(f"[瀵煎嚭] 涓存椂鐩綍: {temp_dir}")
                
                # 瀵煎嚭姣忎竴椤典负PNG
                image_files = []
                for page_num in range(1, slide_count + 1):
                    # 妫€鏌ユ槸鍚﹁鍋滄
                    if self._stopped:
                        return False
                    
                    if progress_callback:
                        progress = 80 + int((page_num / slide_count) * 10)
                        progress_callback(f"Exporting slide {page_num}/{slide_count}...", progress)
                    
                    optimal_height = page_heights[page_num - 1]
                    image_setting = None
                    if page_image_settings and (page_num - 1) < len(page_image_settings):
                        image_setting = page_image_settings[page_num - 1]
                    optimal_width = int(optimal_height * slide_aspect_ratio)

                    image_path = os.path.join(temp_dir, f"slide_{page_num:03d}.png")
                    slide = pres.Slides(page_num)
                    exported_ok = False

                    if image_setting:
                        try:
                            from src.core.converter import ImageExportConfig, ImageFormat
                            fmt = (image_setting.get('format') or "png").lower()
                            quality = int(image_setting.get('quality', 95))
                            ext = "jpg" if fmt in ("jpg", "jpeg") else fmt
                            image_path = os.path.join(temp_dir, f"slide_{page_num:03d}.{ext}")

                            cfg = ImageExportConfig()
                            cfg.format = ImageFormat(fmt)
                            cfg.quality = max(1, min(100, quality))
                            cfg.use_custom_resolution = True
                            cfg.custom_height = optimal_height
                            cfg.custom_width = 0
                            cfg.maintain_aspect_ratio = True
                            cfg.optimize = True
                            cfg.progressive = True
                            exported_ok = converter._export_slide_to_image(slide, image_path, cfg, pres)
                        except Exception as e:
                            self.logger.warning(f"第{page_num}页按V8参数导出失败，回退PNG: {e}")

                    if not exported_ok:
                        slide.Export(image_path, "PNG", optimal_width, optimal_height)

                    image_files.append(image_path)
                    self.logger.info(f"  第{page_num}页: H={optimal_height}px, file={os.path.basename(image_path)}")

                use_foreground_cover = (self.options.mode == ConversionMode.FOREGROUND_IMAGE)
                new_pres = None

                try:
                    if use_foreground_cover:
                        # 鍓嶆櫙瑕嗙洊妯″紡锛氭柊寤虹┖鐧絇PT锛岄伩鍏嶇户鎵垮師鏂囦欢姣嶇増鏁版嵁
                        new_pres = converter.powerpoint.Presentations.Add()
                        try:
                            try:
                                new_pres.PageSetup.SlideSize = pres.PageSetup.SlideSize
                            except Exception as e:
                                self.logger.warning(f"璁剧疆SlideSize澶辫触锛屽皢浠呭悓姝ュ楂? {e}")
                            new_pres.PageSetup.SlideWidth = pres.PageSetup.SlideWidth
                            new_pres.PageSetup.SlideHeight = pres.PageSetup.SlideHeight
                            self.logger.info(
                                f"[鍓嶆櫙瑕嗙洊] 婧愬昂瀵?{pres.PageSetup.SlideWidth:.2f}x{pres.PageSetup.SlideHeight:.2f}, "
                                f"鐩爣灏哄={new_pres.PageSetup.SlideWidth:.2f}x{new_pres.PageSetup.SlideHeight:.2f}"
                            )
                        except Exception as e:
                            self.logger.warning(f"璁剧疆鏂癙PT灏哄澶辫触锛屼娇鐢ㄩ粯璁ゅ昂瀵? {e}")
                    else:
                        # 鑳屾櫙妯″紡锛氫繚鐣欏師妯℃澘缁撴瀯
                        template_path = os.path.join(temp_dir, "template.pptx")
                        self.logger.info(f"[瀵煎嚭] 淇濆瓨妯℃澘: {template_path}")
                        try:
                            pres.SaveAs(template_path)
                        except Exception as e:
                            self.logger.error(f"淇濆瓨妯℃澘PPT澶辫触: {e}")
                            return False
                        time.sleep(0.3)
                        self.logger.info(f"[瀵煎嚭] 鎵撳紑妯℃澘: {template_path}")
                        new_pres = converter.powerpoint.Presentations.Open(template_path)

                    for i, image_file in enumerate(image_files, 1):
                        if self._stopped:
                            return False

                        if progress_callback:
                            progress = 90 + int((i / len(image_files)) * 8)
                            mode_text = "鍓嶆櫙瑕嗙洊" if use_foreground_cover else "鑳屾櫙"
                            progress_callback(f"正在设置第{i}/{len(image_files)}页{mode_text}...", progress)

                        if use_foreground_cover:
                            if i > new_pres.Slides.Count:
                                slide = new_pres.Slides.Add(i, 12)  # ppLayoutBlank
                            else:
                                slide = new_pres.Slides(i)
                                slide.Layout = 12
                        elif i <= new_pres.Slides.Count:
                            slide = new_pres.Slides(i)
                        else:
                            continue

                        try:
                            for shape_idx in range(slide.Shapes.Count, 0, -1):
                                try:
                                    slide.Shapes(shape_idx).Delete()
                                except Exception:
                                    pass

                            abs_image_path = os.path.abspath(image_file)
                            if use_foreground_cover:
                                slide_width = new_pres.PageSetup.SlideWidth
                                slide_height = new_pres.PageSetup.SlideHeight
                                with Image.open(abs_image_path) as img:
                                    img_width, img_height = img.size
                                    img_ratio = img_width / img_height
                                    slide_ratio = slide_width / slide_height

                                if img_ratio > slide_ratio:
                                    scaled_height = slide_height
                                    scaled_width = int(scaled_height * img_ratio)
                                else:
                                    scaled_width = slide_width
                                    scaled_height = int(scaled_width / img_ratio)

                                left = (slide_width - scaled_width) // 2
                                top = (slide_height - scaled_height) // 2

                                picture = slide.Shapes.AddPicture(
                                    FileName=abs_image_path,
                                    LinkToFile=False,
                                    SaveWithDocument=True,
                                    Left=left,
                                    Top=top,
                                    Width=scaled_width,
                                    Height=scaled_height
                                )
                                try:
                                    picture.ZOrder(1)
                                except Exception:
                                    pass
                            else:
                                slide.FollowMasterBackground = False
                                slide.Background.Fill.UserPicture(abs_image_path)
                        except Exception as e:
                            if use_foreground_cover:
                                self.logger.warning(f"Set foreground for slide {i} failed: {e}")
                            else:
                                self.logger.warning(f"Set background for slide {i} failed: {e}")
                    
                    # 淇濆瓨鏂癙PT
                    if progress_callback:
                        progress_callback("姝ｅ湪淇濆瓨鏂囦欢...", 98)
                    
                    save_path = os.path.abspath(output_path)
                    self.logger.info(f"[瀵煎嚭] 淇濆瓨鏈€缁堟枃浠? {save_path}")
                    new_pres.SaveAs(save_path)
                    self.logger.info(f"[瀵煎嚭] PPT宸蹭繚瀛? {save_path}")
                    
                    # 楠岃瘉鏂囦欢鏄惁鍒涘缓
                    time.sleep(0.5)
                    
                    if os.path.exists(save_path):
                        file_size = os.path.getsize(save_path)
                        self.logger.info(f"[瀵煎嚭] 鉁?鏂囦欢宸插垱寤? {save_path} ({file_size} bytes)")
                        return True
                    else:
                        self.logger.error(f"[瀵煎嚭] 鉁?鏂囦欢鏈垱寤? {save_path}")
                        return False
                    
                finally:
                    try:
                        if new_pres is not None:
                            new_pres.Close()
                            self.logger.info(f"[Export] new presentation closed")
                    except Exception as e:
                        self.logger.warning(f"鍏抽棴鏂癙PT澶辫触: {e}")

                    try:
                        pres.Close()
                    except Exception:
                        pass
                
            except Exception as e:
                self.logger.error(f"瀵煎嚭PPT澶辫触: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
                return False
                
        except Exception as e:
            self.logger.error(f"瀵煎嚭PPT澶辫触: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

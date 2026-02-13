import sys
import os

# 添加项目根目录到sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

# Windows任务栏图标修复 - 必须在导入Qt之前设置
# 使用Windows API设置应用程序ID和图标
try:
    import ctypes
    from ctypes import wintypes
    
    # 设置应用程序用户模型ID（Windows 7+）
    app_id = "PitchPPT.Application.1.0"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    
    # 获取资源文件夹中的ICO文件路径
    ico_path = os.path.join(project_root, "resources", "LOGO_256x256.ico")
    if not os.path.exists(ico_path):
        ico_path = os.path.join(project_root, "resources", "LOGO_128x128.ico")
    if not os.path.exists(ico_path):
        ico_path = os.path.join(project_root, "resources", "LOGO_64x64.ico")
    if not os.path.exists(ico_path):
        ico_path = os.path.join(project_root, "resources", "LOGO_48x48.ico")
    
    if os.path.exists(ico_path):
        # 使用Windows API加载图标
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040
        
        hicon = ctypes.windll.user32.LoadImageW(
            None, ico_path, IMAGE_ICON, 
            0, 0,  # 使用默认大小
            LR_LOADFROMFILE | LR_DEFAULTSIZE
        )
        
        if hicon:
            # 设置窗口图标
            ctypes.windll.user32.SetClassLongPtrW(
                ctypes.windll.kernel32.GetConsoleWindow(),
                -14,  # GCLP_HICON
                hicon
            )
            print(f"Windows icon loaded via API: {ico_path}")
except Exception as e:
    print(f"Windows API icon setup failed: {e}")

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt, QRectF
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWinExtras import QtWin

from src.ui.main_window import MainWindow

def load_icon_from_svg(svg_path):
    """从SVG文件加载图标，返回QIcon（支持多尺寸）"""
    try:
        renderer = QSvgRenderer(svg_path)
        if not renderer.isValid():
            return None

        icon = QIcon()

        # 添加多个尺寸的图标以适应不同场景
        sizes = [16, 24, 32, 48, 64, 128, 256]
        for size in sizes:
            pixmap = QPixmap(size, size)
            pixmap.fill(Qt.transparent)

            painter = QPainter(pixmap)
            renderer.render(painter, QRectF(0, 0, size, size))
            painter.end()

            if not pixmap.isNull():
                icon.addPixmap(pixmap)

        return icon if not icon.isNull() else None
    except Exception as e:
        print(f"Failed to load SVG: {e}")
        return None

def load_icon_from_ico(ico_path):
    """从ICO文件加载图标，返回QIcon"""
    try:
        # 尝试使用QPixmap加载ICO文件
        pixmap = QPixmap(ico_path)
        if not pixmap.isNull():
            return QIcon(pixmap)
    except Exception as e:
        print(f"Failed to load ICO with QPixmap: {e}")
    return None

if __name__ == "__main__":
    # 设置应用属性（必须在创建QApplication之前）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)

    # 设置全局字体
    font = QFont("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)

    # 加载QSS样式表
    try:
        style_file = os.path.join(os.path.dirname(__file__), "../resources/style.qss")
        with open(style_file, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
    except Exception as e:
        print(f"Failed to load stylesheet: {e}")

    # 设置应用图标 - 优先使用SVG矢量格式
    icon = None

    # 方法1: 优先加载SVG矢量格式
    svg_path = os.path.join(project_root, "resources", "LOGO.svg")
    if os.path.exists(svg_path):
        icon = load_icon_from_svg(svg_path)
        if icon:
            print(f"Icon loaded from SVG: {svg_path}")

    # 方法2: 如果SVG加载失败，尝试加载ICO文件
    if not icon:
        ico_paths = [
            os.path.join(project_root, "resources", "LOGO_256x256.ico"),
            os.path.join(project_root, "resources", "LOGO_128x128.ico"),
            os.path.join(project_root, "resources", "LOGO_64x64.ico"),
            os.path.join(project_root, "resources", "LOGO_48x48.ico"),
            os.path.join(project_root, "resources", "LOGO_32x32.ico"),
        ]

        for ico_path in ico_paths:
            if os.path.exists(ico_path):
                icon = load_icon_from_ico(ico_path)
                if icon:
                    print(f"Icon loaded from ICO: {ico_path}")
                    break

    # 方法3: 如果ICO加载失败，使用PNG
    if not icon:
        png_path = os.path.join(project_root, "resources", "LOGO.png")
        if os.path.exists(png_path):
            pixmap = QPixmap(png_path)
            if not pixmap.isNull():
                icon = QIcon(pixmap)
                print(f"Icon loaded from PNG: {png_path}")

    if icon:
        app.setWindowIcon(icon)
    else:
        print("Warning: Could not load any icon file")

    window = MainWindow()

    # 再次设置窗口图标（确保任务栏显示）
    if icon:
        window.setWindowIcon(icon)

    window.show()
    sys.exit(app.exec_())

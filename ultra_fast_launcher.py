"""
极速启动器 - 使用新的SplashScreen类
"""
import os
import sys
import time

# 导入必要的Qt模块
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer

def minimal_environment():
    """最小化环境设置"""
    # 1. 禁用所有日志
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.devnull, 'w')

    # 2. 极端Qt优化
    os.environ.update({
        "QT_QPA_PLATFORM": "windows",
        "QT_LOGGING_RULES": "*.debug=false;*.info=false;*.warning=false;qt.*=false",
        "QT_AUTO_SCREEN_SCALE_FACTOR": "0",
        "QT_ENABLE_HIGHDPI_SCALING": "0",
        "QT_DISABLE_FONTCONFIG": "1",
        "PYTHONUNBUFFERED": "1",
        "PYTHONDONTWRITEBYTECODE": "1",
    })

    # 3. Windows进程优化
    if sys.platform == 'win32':
        try:
            import ctypes
            # 设置进程优先级为正常，避免系统调度开销
            ctypes.windll.kernel32.SetPriorityClass(
                ctypes.windll.kernel32.GetCurrentProcess(),
                0x00000020  # NORMAL_PRIORITY_CLASS
            )
        except:
            pass

if __name__ == "__main__":
    start_time = time.perf_counter()

    # 1. 极简环境
    minimal_environment()

    # 2. 创建应用程序
    app = QApplication(sys.argv)

    # 3. 设置全局字体
    from PyQt5.QtGui import QFont
    font = QFont("Microsoft YaHei", 10)
    app.setStyle('Fusion')
    app.setFont(font)

    # 4. 创建并显示启动屏幕
    from correction import SplashScreen
    splash = SplashScreen()
    splash.show_and_animate()

    # 强制处理事件，确保启动屏幕显示
    QApplication.processEvents()

    # 5. 创建主窗口，传递启动屏幕引用
    from correction import MainWindow
    window = MainWindow(splash_screen=splash)

    # 6. 记录启动时间
    def log_startup_time():
        elapsed = time.perf_counter() - start_time
        print(f"🚀 应用程序启动时间: {elapsed:.2f}秒")

    QTimer.singleShot(2000, log_startup_time)

    sys.exit(app.exec_())

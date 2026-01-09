# Automation Tool.py - 顶部导入优化
import sys
import os
import json

# 第一步：只导入绝对必要的模块
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit, QGroupBox, QGridLayout, QHBoxLayout, QProgressBar, QMenu, QAction, QDialog, QMessageBox, QFileDialog, QLineEdit, QCheckBox, QFormLayout, QStyle
from PyQt5.QtGui import QFont, QIcon, QCursor
from PyQt5.QtCore import pyqtSignal, QThread

# 第四步：优化文件路径处理
from pathlib import Path
current_script = Path(__file__).resolve()
project_root = current_script.parent.parent
sys.path.append(str(project_root))

# 快捷方式配置管理
def get_app_config_dir():
    """获取应用配置目录"""
    if getattr(sys, 'frozen', False):
        # 如果是打包的exe，使用exe目录
        app_data_dir = os.path.dirname(sys.executable)
    else:
        # 如果是开发模式，使用脚本目录
        app_data_dir = os.path.dirname(os.path.abspath(__file__))
    return app_data_dir

def get_app_config_path():
    """获取应用配置文件路径"""
    config_dir = get_app_config_dir()
    return os.path.join(config_dir, ".app_config.json")

def load_app_config():
    """加载应用配置"""
    config_path = get_app_config_path()
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载配置文件失败: {e}")
    return {}

def save_app_config(config):
    """保存应用配置"""
    config_path = get_app_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {e}")

def should_show_shortcut_dialog():
    """检查是否应该显示快捷方式对话框"""
    config = load_app_config()
    return not config.get("shortcut_dialog_shown", False)

def mark_shortcut_dialog_shown(dont_show_again=False):
    """标记快捷方式对话框已显示"""
    config = load_app_config()
    if dont_show_again:
        config["shortcut_dialog_shown"] = True
        config["shortcut_choice"] = "no"
    save_app_config(config)

def get_shortcut_choice():
    """获取用户的快捷方式选择"""
    config = load_app_config()
    return config.get("shortcut_choice", None)

def save_shortcut_choice(choice):
    """保存用户的快捷方式选择"""
    config = load_app_config()
    config["shortcut_choice"] = choice
    save_app_config(config)

def get_app_name():
    """获取应用名称（从exe文件名读取）"""
    if getattr(sys, 'frozen', False):
        # 如果是打包的exe，从exe文件名获取
        exe_path = sys.executable
        exe_name = os.path.splitext(os.path.basename(exe_path))[0]
        print(f"原始exe名称: {exe_name}")

        # 确保正确处理中文字符
        try:
            # 尝试多种编码方式处理
            if isinstance(exe_name, str):
                # 如果已经是unicode字符串，直接使用
                processed_name = exe_name
            else:
                # 如果是bytes，尝试解码
                for encoding in ['utf-8', 'gb18030', 'gbk', 'cp936']:
                    try:
                        processed_name = exe_name.decode(encoding)
                        break
                    except (UnicodeDecodeError, AttributeError):
                        continue
                else:
                    processed_name = str(exe_name)
        except Exception as e:
            print(f"处理exe名称时出错: {e}")
            processed_name = "Automation Tool"

        print(f"处理后的exe名称: {processed_name}")
        return processed_name
    else:
        # 开发模式，使用脚本文件名
        script_path = os.path.abspath(__file__)
        script_name = os.path.splitext(os.path.basename(script_path))[0]
        print(f"开发模式脚本名称: {script_name}")
        return script_name

def create_desktop_shortcut():
    """在桌面创建快捷方式"""
    try:
        import winshell
        import tempfile
        import subprocess

        # 获取当前exe路径或脚本路径
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
            exe_dir = os.path.dirname(exe_path)
            # 在打包exe的根目录下查找图标
            icon_path = os.path.join(exe_dir, "tool_icon.ico")
        else:
            exe_path = os.path.abspath(__file__)
            exe_dir = os.path.dirname(exe_path)
            # 在脚本所在目录查找图标
            icon_path = os.path.join(exe_dir, "tool_icon.ico")

        # 获取应用名称（从exe文件名读取）
        app_name = get_app_name()

        print(f"图标路径: {icon_path}")
        print(f"图标是否存在: {os.path.exists(icon_path)}")

        # 获取桌面路径
        desktop = winshell.desktop()

        # 快捷方式路径 - 使用动态名称，确保与exe文件名一致
        shortcut_name = f"{app_name}.lnk"
        shortcut_path = os.path.join(desktop, shortcut_name)

        print(f"创建快捷方式: {shortcut_path}")
        print(f"目标路径: {exe_path}")
        print(f"工作目录: {exe_dir}")
        print(f"应用名称: {app_name}")

        # 首先尝试使用winshell直接创建（更简单）
        try:
            with winshell.shortcut(shortcut_path) as shortcut:
                shortcut.path = exe_path
                shortcut.working_directory = exe_dir
                shortcut.description = app_name
                if icon_path and os.path.exists(icon_path):
                    print(f"设置图标: {icon_path}")
                    shortcut.icon_location = (icon_path, 0)
                else:
                    print("图标文件不存在，使用默认图标")
                shortcut.write()

            print(f"成功创建桌面快捷方式: {shortcut_path}")
            return True

        except Exception as winshell_error:
            print(f"winshell方法失败: {winshell_error}")

            # 如果winshell失败，尝试VBScript方法
            try:
                # 创建临时VBScript文件，使用ASCII编码避免中文问题
                vb_script = f'''Set WshShell = WScript.CreateObject("WScript.Shell")
Set shortcut = WshShell.CreateShortcut("{shortcut_path}")
shortcut.TargetPath = "{exe_path}"
shortcut.WorkingDirectory = "{exe_dir}"
shortcut.Description = "{app_name}"'''

                if icon_path and os.path.exists(icon_path):
                    # 确保图标路径使用绝对路径
                    abs_icon_path = os.path.abspath(icon_path)
                    print(f"VBScript设置图标: {abs_icon_path}")
                    vb_script += f'\nshortcut.IconLocation = "{abs_icon_path}"'

                vb_script += '\nshortcut.Save'

                # 写入临时VBS文件
                with tempfile.NamedTemporaryFile(suffix='.vbs', delete=False, mode='w', encoding='ascii', errors='ignore') as temp_vbs:
                    temp_vbs.write(vb_script)
                    temp_vbs_path = temp_vbs.name

                print(f"执行VBScript: {temp_vbs_path}")
                # 执行VBScript
                result = subprocess.run(['cscript', '//Nologo', temp_vbs_path], shell=True, capture_output=True, text=True)

                # 清理临时文件
                os.unlink(temp_vbs_path)

                if result.returncode == 0:
                    print(f"成功创建桌面快捷方式: {shortcut_path}")
                    return True
                else:
                    print(f"VBScript执行失败: {result.stderr}")
                    return False

            except Exception as vb_error:
                print(f"VBScript方法也失败: {vb_error}")
                return False

    except ImportError:
        print("缺少创建快捷方式所需的库: winshell 或 subprocess")
        return False
    except Exception as e:
        print(f"创建桌面快捷方式失败: {e}")
        return False

# 延迟导入函数 - 仅在需要时加载
def get_memo_generator():
    from modules.memo_generator import generate_memo
    return generate_memo

def get_pdf_extractor():
    from modules.pdf_extractor import PdfTableExtractor
    return PdfTableExtractor

def get_outlook_email_thread():
    from modules.outlook_automation import OutlookEmailThread
    return OutlookEmailThread

def get_folder_creator():
    from modules.folder_creation import FolderCreator
    return FolderCreator

def get_file_converter():
    from modules.file_converter import FileConverter
    return FileConverter

def get_file_converter_ui():
    from modules.file_converter_ui import FileConverterUI
    return FileConverterUI

def get_find_files_with_progress():
    from modules.findfile import find_files_with_progress
    return find_files_with_progress

def get_find_excel_file():
    from utils.file_utils import find_excel_file
    return find_excel_file

# -------------------------- 启动窗口类 --------------------------
class SplashScreen(QWidget):
    """启动屏幕 - 显示在加载主界面时"""
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置为无边框、置顶窗口
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)

        # 设置透明背景
        self.setAttribute(Qt.WA_TranslucentBackground)

        # 设置窗口大小
        self.setFixedSize(300, 300)

        # 居中显示
        self.center_on_screen()

        # 创建布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 图标显示
        self.icon_label = QLabel()
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)

        # 加载进度标签
        self.status_label = QLabel('启动中...')
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont("Microsoft YaHei", 10))
        layout.addWidget(self.status_label)

        # 设置窗口样式
        self.setStyleSheet("""
            background-color: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            border: 2px solid #27AE60;
        """)

        # 加载图标
        self._load_icon()

    def _load_icon(self):
        """加载图标"""
        try:
            # 使用根目录下的tool_icon.ico
            icon_path = "tool_icon.ico"
            if os.path.exists(icon_path):
                pixmap = QIcon(icon_path).pixmap(200, 200)
                self.icon_label.setPixmap(pixmap)
            else:
                self.icon_label.setText("🛠️")
                self.icon_label.setFont(QFont("Arial", 80))
        except Exception as e:
            print(f"加载图标失败: {e}")
            self.icon_label.setText("🛠️")
            self.icon_label.setFont(QFont("Arial", 80))

    def center_on_screen(self):
        """将窗口居中显示在屏幕上"""
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def update_status(self, message):
        """更新状态信息"""
        self.status_label.setText(message)
        QApplication.processEvents()  # 强制更新UI

    def show_and_animate(self):
        """显示窗口并添加淡入动画"""
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(500)
        self.animation.setStartValue(0)
        self.animation.setEndValue(1)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)

        self.show()
        self.animation.start()

    def hide_and_animate(self):
        """隐藏窗口并添加淡出动画"""
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)
        self.animation.setStartValue(1)
        self.animation.setEndValue(0)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.animation.finished.connect(self.close)
        self.animation.start()

# -------------------------- 快捷方式询问对话框类 --------------------------
class ShortcutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        self.setWindowTitle('桌面快捷方式')
        self.setModal(True)
        self.setFixedSize(400, 250)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # 图标和标题区域
        title_layout = QHBoxLayout()
        title_layout.setSpacing(15)

        # 应用图标
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(48, 48)
        self.icon_label.setScaledContents(True)
        title_layout.addWidget(self.icon_label)

        # 标题和描述
        text_layout = QVBoxLayout()
        self.title_label = QLabel('Automation Tool')
        self.title_label.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        text_layout.addWidget(self.title_label)

        self.desc_label = QLabel('是否要在桌面创建快捷方式？')
        self.desc_label.setFont(QFont("Microsoft YaHei", 10))
        text_layout.addWidget(self.desc_label)

        title_layout.addLayout(text_layout)
        title_layout.addStretch()
        layout.addLayout(title_layout)

        # 分隔线
        separator = QLabel()
        separator.setStyleSheet("background-color: #ddd; margin: 10px 0;")
        separator.setFixedHeight(1)
        layout.addWidget(separator)

        # 复选框 - 下次不再弹出
        self.dont_ask_checkbox = QCheckBox("下次不再弹出此提示")
        self.dont_ask_checkbox.setFont(QFont("Microsoft YaHei", 9))
        layout.addWidget(self.dont_ask_checkbox)

        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        button_layout.addStretch()

        self.cancel_btn = QPushButton('取消')
        self.cancel_btn.setFont(QFont("Microsoft YaHei", 10))
        self.cancel_btn.setFixedWidth(80)
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)

        self.create_btn = QPushButton('创建')
        self.create_btn.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        self.create_btn.setFixedWidth(80)
        self.create_btn.clicked.connect(self.accept)
        self.create_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: white;
                border: none;
                border-radius: 2px;
                padding: 5px 15px;
            }
            QPushButton:hover {
                background-color: #106EBE;
            }
            QPushButton:pressed {
                background-color: #005A9E;
            }
        """)
        button_layout.addWidget(self.create_btn)

        layout.addLayout(button_layout)

        # 设置默认焦点
        self.create_btn.setDefault(True)

        # 延迟加载图标，避免阻塞UI创建
        QTimer.singleShot(10, self._load_icon)

    def _load_icon(self):
        """加载应用图标"""
        try:
            # 检查当前工作目录
            current_dir = os.getcwd()
            print(f"当前工作目录: {current_dir}")

            # 检查tool_icon.ico文件
            icon_path = "tool_icon.ico"
            abs_icon_path = os.path.abspath(icon_path)
            print(f"图标文件路径: {abs_icon_path}")
            print(f"图标文件是否存在: {os.path.exists(icon_path)}")
            print(f"图标文件大小: {os.path.getsize(icon_path) if os.path.exists(icon_path) else 'N/A'}")

            # 首先尝试加载同目录下的图标文件（无论开发模式还是打包模式）
            icon_path = "tool_icon.ico"
            print(f"尝试加载图标: {os.path.abspath(icon_path)}")
            if os.path.exists(icon_path):
                print(f"图标文件存在，加载中...")
                pixmap = QIcon(icon_path).pixmap(48, 48)
                if not pixmap.isNull():
                    self.icon_label.setPixmap(pixmap)
                    self.setWindowIcon(QIcon(icon_path))
                    print("图标加载成功")
                    return
                else:
                    print("图标文件存在但加载失败")
            else:
                print(f"图标文件不存在: {os.path.abspath(icon_path)}")

            # 如果没有图标文件，使用默认图标
            print("尝试使用系统默认图标...")
            if hasattr(QStyle, 'SP_ComputerIcon'):
                icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
                pixmap = icon.pixmap(48, 48)
                if not pixmap.isNull():
                    print("使用系统默认图标成功")
                    self.icon_label.setPixmap(pixmap)
                    self.setWindowIcon(icon)
                else:
                    print("系统默认图标加载失败")
            else:
                print("QStyle.SP_ComputerIcon不可用")

                # 备用方案
                print("使用emoji图标")
                self.icon_label.setText("🛠️")
                self.icon_label.setFont(QFont("Arial", 24))

        except Exception as e:
            print(f"加载图标失败: {e}")
            self.icon_label.setText("🛠️")
            self.icon_label.setFont(QFont("Arial", 24))

    def set_app_name(self, name):
        """设置应用名称"""
        self.title_label.setText(name)

# -------------------------- 文件搜索对话框类 --------------------------
class FileSearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_thread = None
        self.initUI()
        # 延迟加载图标，避免阻塞启动
        QTimer.singleShot(50, self._load_icon)
        # 设置默认值
        self.search_dir_input.setText(r"C:\Zeiss\CMM_Tools\FW_C99\backup")
        self.search_content_input.setText("Install_version = V47.04")
        self.file_names_input.setText("config.kmg")

    def initUI(self):
        self.setWindowTitle('文件内容搜索工具')
        self.setGeometry(300, 300, 800, 600)

        layout = QVBoxLayout(self)

        # 控制按钮区域
        control_layout = QHBoxLayout()

        self.select_dir_btn = QPushButton('选择目录...')
        self.select_dir_btn.clicked.connect(self.select_directory)
        control_layout.addWidget(self.select_dir_btn)

        self.start_search_btn = QPushButton('开始搜索')
        self.start_search_btn.clicked.connect(self.start_search)
        control_layout.addWidget(self.start_search_btn)

        self.clear_results_btn = QPushButton('清空结果')
        self.clear_results_btn.clicked.connect(self.clear_results)
        control_layout.addWidget(self.clear_results_btn)

        self.cancel_search_btn = QPushButton('取消搜索')
        self.cancel_search_btn.clicked.connect(self.cancel_search)
        self.cancel_search_btn.setEnabled(False)
        control_layout.addWidget(self.cancel_search_btn)

        layout.addLayout(control_layout)

        # 参数设置区域
        params_group = QGroupBox("搜索参数设置")
        params_layout = QFormLayout()

        # 搜索目录
        self.search_dir_input = QLineEdit()
        self.search_dir_input.setPlaceholderText("请输入或选择要搜索的目录")
        params_layout.addRow("搜索目录:", self.search_dir_input)

        # 搜索内容
        self.search_content_input = QLineEdit()
        self.search_content_input.setPlaceholderText("请输入要搜索的内容")
        params_layout.addRow("搜索内容:", self.search_content_input)

        # 文件名
        self.file_names_input = QLineEdit()
        self.file_names_input.setPlaceholderText("请输入要搜索的文件名，用逗号分隔（留空则搜索常见配置文件）")
        params_layout.addRow("文件名:", self.file_names_input)

        # 区分大小写选项
        self.case_sensitive_checkbox = QCheckBox("区分大小写")
        params_layout.addRow("", self.case_sensitive_checkbox)

        params_group.setLayout(params_layout)
        layout.addWidget(params_group)

        # 结果显示区域
        results_group = QGroupBox("搜索结果")
        results_layout = QVBoxLayout()

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #fafafa;
                font-family: Consolas, monospace;
                font-size: 12px;
            }
        """)
        results_layout.addWidget(self.result_text)
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #bbb;
                border-radius: 4px;
                text-align: center;
                height: 12px;
            }
            QProgressBar::chunk {
                background-color: #90caf9;
                border-radius: 3px;
            }
        """)
        layout.addWidget(self.progress_bar)
    
    
    def select_directory(self):
        """选择搜索目录"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择搜索目录",
            self.search_dir_input.text() or os.path.expanduser("~")
        )
        if dir_path:
            self.search_dir_input.setText(dir_path)

    def start_search(self):
        """开始搜索"""
        search_dir = self.search_dir_input.text().strip()
        search_content = self.search_content_input.text().strip()
        file_names_text = self.file_names_input.text().strip()

        if not search_dir:
            QMessageBox.warning(self, "警告", "请选择搜索目录")
            return

        if not search_content:
            QMessageBox.warning(self, "警告", "请输入搜索内容")
            return

        if not os.path.exists(search_dir):
            QMessageBox.warning(self, "警告", f"目录不存在：{search_dir}")
            return

        # 解析文件名
        file_names = None
        if file_names_text:
            file_names = [name.strip() for name in file_names_text.split(',') if name.strip()]

        # 确认搜索
        message = f"搜索目录：{search_dir}\n搜索内容：{search_content}\n搜索文件：{file_names or '常见配置文件'}"
        reply = QMessageBox.question(
            self, "确认搜索", f"确认开始搜索？\n\n{message}",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply == QMessageBox.No:
            return

        # 开始搜索
        self.start_search_btn.setEnabled(False)
        self.cancel_search_btn.setEnabled(True)
        self.clear_results()
        self.add_result(f"开始搜索内容: '{search_content}' (文件名: {file_names or 'config.kmg'})")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 创建搜索线程
        self.current_thread = FileSearchThread(
            search_dir, search_content, file_names,
            self.case_sensitive_checkbox.isChecked()
        )
        self.current_thread.result_signal.connect(self.add_result)
        self.current_thread.progress_signal.connect(self.update_progress)
        self.current_thread.finished.connect(self.on_search_finished)
        self.current_thread.start()

    def add_result(self, message):
        """向结果区域添加文本"""
        self.result_text.append(message)
        # 自动滚动到底部
        cursor = self.result_text.textCursor()
        cursor.movePosition(cursor.End)
        self.result_text.setTextCursor(cursor)

    def clear_results(self):
        """清空结果区域"""
        self.result_text.clear()

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def on_search_finished(self, success):
        """搜索完成回调"""
        self.progress_bar.setVisible(False)
        self.start_search_btn.setEnabled(True)
        self.cancel_search_btn.setEnabled(False)
        if success:
            self.add_result("\n" + "="*50)
            self.add_result("搜索完成！")
        else:
            self.add_result("\n" + "="*50)
            self.add_result("搜索过程中出现错误！")

        self.current_thread = None

    def cancel_search(self):
        """取消搜索"""
        if self.current_thread and self.current_thread.isRunning():
            self.current_thread.cancel()
            self.add_result("正在取消搜索...")
            self.start_search_btn.setEnabled(True)
            self.cancel_search_btn.setEnabled(False)

# -------------------------- 文件搜索线程类 --------------------------
class FileSearchThread(QThread):
    result_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished = pyqtSignal(bool)

    def __init__(self, root_dir, search_content, file_names=None, case_sensitive=False):
        super().__init__()
        self.root_dir = root_dir
        self.search_content = search_content
        self.file_names = file_names
        self.case_sensitive = case_sensitive
        self.is_canceled = False

    def run(self):
        try:
            self.search_files()
            self.finished.emit(True)
        except Exception as e:
            self.result_signal.emit(f"搜索过程中出错: {str(e)}")
            self.finished.emit(False)

    def search_files(self):
        """执行文件搜索"""
        found_files = []
        total_files = 0
        processed_files = 0

        # 1. 如果没有指定文件名，默认搜索常见配置文件
        if self.file_names is None:
            self.file_names = ['config.kmg']

        # 2. 统计总文件数
        self.result_signal.emit("正在扫描文件...")
        for root, _, files in os.walk(self.root_dir):
            if self.is_canceled:
                self.result_signal.emit("搜索已取消")
                return
            for file in files:
                if any(file.lower() == name.lower() for name in self.file_names):
                    total_files += 1

        if total_files == 0:
            self.result_signal.emit(f"在指定目录中未找到目标文件类型: {', '.join(self.file_names)}")
            return

        self.result_signal.emit(f"找到 {total_files} 个目标文件，开始搜索内容...")

        # 3. 开始搜索
        for root, _, files in os.walk(self.root_dir):
            for file in files:
                if self.is_canceled:
                    self.result_signal.emit("搜索已取消")
                    return

                file_lower = file.lower()
                # 检查文件名是否匹配（不区分大小写）
                if any(file_lower == name.lower() for name in self.file_names):
                    file_path = os.path.join(root, file)
                    processed_files += 1

                    # 更新进度
                    progress = int((processed_files / total_files) * 100)
                    self.progress_signal.emit(progress)

                    try:
                        # 尝试多种编码读取文件
                        encodings = ['utf-8', 'gb18030', 'gbk', 'latin-1']
                        content_found = False

                        for encoding in encodings:
                            if self.is_canceled:
                                self.result_signal.emit("搜索已取消")
                                return

                            try:
                                with open(file_path, 'r', encoding=encoding) as f:
                                    for line_num, line in enumerate(f, 1):
                                        if self.is_canceled:
                                            self.result_signal.emit("搜索已取消")
                                            return

                                        line_to_check = line if self.case_sensitive else line.lower()
                                        search_to_check = self.search_content if self.case_sensitive else self.search_content.lower()

                                        if search_to_check in line_to_check:
                                            # 找到匹配内容
                                            self.result_signal.emit(f"\n 找到匹配文件: {file_path}")
                                            self.result_signal.emit(f"   行号: {line_num}, 匹配行: {line.strip()}")
                                            found_files.append((file_path, line_num, line.strip()))
                                            content_found = True
                                            break  # 找到后跳出编码循环
                                if content_found:
                                    break
                            except UnicodeDecodeError:
                                continue  # 尝试下一种编码
                            except Exception as e:
                                self.result_signal.emit(f"\n 读取文件 {file_path} 时出错: {e}")
                                break

                    except Exception as e:
                        self.result_signal.emit(f"\n访问文件 {file_path} 时发生错误: {e}")

        # 4. 完成搜索
        self.progress_signal.emit(100)
        self.result_signal.emit("\n" + "="*50)
        if found_files:
            self.result_signal.emit(f"搜索完成! 共找到 {len(found_files)} 个匹配项。")
        else:
            self.result_signal.emit(f"搜索完成! 未找到包含 '{self.search_content}' 的文件。")

    def cancel(self):
        """取消搜索"""
        self.is_canceled = True

# -------------------------- 线程类 --------------------------
class FolderThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self.folder_creator = None
        self.is_canceled = False

    def run(self):
        try:
            # 延迟导入FolderCreator
            FolderCreatorClass = get_folder_creator()
            self.folder_creator = FolderCreatorClass()
            self.folder_creator.log_signal.connect(self.progress)
            self.folder_creator.finished.connect(self.on_finished)
            if not self.is_canceled:
                self.folder_creator.create_folders()
            else:
                self.progress.emit("文件夹创建任务已被取消")
                self.finished.emit(False)
        except Exception as e:
            self.progress.emit(f"文件夹线程出错：{str(e)}")
            self.finished.emit(False)

    def cancel(self):
        self.is_canceled = True
        self.progress.emit("正在取消文件夹创建任务...")

    def on_finished(self, success):
        self.finished.emit(success)

class MemoThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, excel_path=None):
        super().__init__()
        self.excel_path = excel_path
        self.is_canceled = False

    def run(self):
        try:
            self.progress.emit("📋 启动MEMO生成任务...")
            # 延迟导入memo_generator
            generate_memo = get_memo_generator()
            success, msg, output_path = generate_memo(
                excel_path=self.excel_path,
                progress_callback=lambda log: self.progress.emit(log)
            )
            self.finished.emit(success, msg)
        except Exception as e:
            err_msg = f"MEMO线程出错：{str(e)}"
            self.progress.emit(f"❌ {err_msg}")
            self.finished.emit(False, err_msg)

    def cancel(self):
        self.is_canceled = True
        self.progress.emit("⏹️  正在取消MEMO生成任务...")

class PdfExtractThread(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool)

    def __init__(self, input_dir, output_dir):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.extractor = None

    def run(self):
        try:
            # 延迟导入PdfTableExtractor
            PdfTableExtractor = get_pdf_extractor()
            self.extractor = PdfTableExtractor()
            self.extractor.log_signal.connect(self.log)
            self.extractor.progress_signal.connect(self.progress)
            self.extractor.finished_signal.connect(self.finished)
            self.extractor.set_paths(self.input_dir, self.output_dir)
            self.extractor.batch_extract()
        except Exception as e:
            self.log.emit(f"PDF提取线程出错：{str(e)}")
            self.finished.emit(False)

    def cancel(self):
        if self.extractor and hasattr(self.extractor, 'cancel_extract'):
            self.extractor.cancel_extract()

# -------------------------- 主窗口类 --------------------------
class MainWindow(QMainWindow):
    def __init__(self, splash_screen=None):
        super().__init__()
        self.excel_path = None
        self.splash_screen = splash_screen  # 保存启动屏幕引用

        # 初始化线程变量
        self.outlook_thread = None
        self.memo_thread = None
        self.pdf_thread = None
        self.folder_thread = None
        self.current_thread = None

        # 初始化PDF目录变量
        self.pdf_input_dir = ""
        self.pdf_output_dir = ""

        # 极简第一阶段：仅设置窗口属性
        self.setWindowTitle('Automation Tool')
        self.center_window()


        # 更新启动屏幕状态
        if self.splash_screen:
            self.splash_screen.update_status('创建界面...')

        # 创建绝对简单的占位界面
        self._create_simple_placeholder()

        # 不立即显示窗口，等待加载完成
        # self.show()  # 注释掉这行，我们将在加载完成后显示

        # 延迟加载所有其他组件
        QTimer.singleShot(10, self._phase1_load)

    def _create_simple_placeholder(self):
        """创建极简占位界面"""
        central = QWidget()
        layout = QVBoxLayout(central)

        # 只有标题
        title = QLabel('Automation Tool')
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 简单状态
        self.status_label = QLabel('启动中...')
        layout.addWidget(self.status_label)

        self.setCentralWidget(central)

    def _phase1_load(self):
        """第一阶段：加载核心UI"""
        if self.splash_screen:
            self.splash_screen.update_status('加载界面...')
        self.status_label.setText('加载界面...')
        
        # 重新创建完整UI
        self._recreate_ui()

        QTimer.singleShot(50, self._phase2_load)
    def _phase2_load(self):
        """第二阶段：加载功能模块"""
        if self.splash_screen:
            self.splash_screen.update_status('初始化功能模块...')
        
        self.update_log('初始化功能模块...')

        try:
            # 延迟导入config
            from modules import config
            self.pdf_input_dir = config.PDF_INPUT_DIR
            self.pdf_output_dir = config.PDF_OUTPUT_DIR
            
            # 更新UI标签
            self.pdf_input_label.setText(self.pdf_input_dir)
            self.pdf_output_label.setText(self.pdf_output_dir)
            
            # 延迟查找Excel
            QTimer.singleShot(50, self._phase3_load)
        except ImportError as e:
            self.update_log(f"❌ 导入配置模块失败: {str(e)}")
            self.update_log("⚠️  请确保modules/config.py文件存在且配置正确")
            self._finalize_loading()

    def _phase3_load(self):
        """第三阶段：查找Excel文件"""
        if self.splash_screen:
            self.splash_screen.update_status('查找Excel文件...')
        
        self.find_and_display_excel()
        self._finalize_loading()

    def _finalize_loading(self):
        """完成加载，关闭启动屏幕并显示主窗口"""
        if self.splash_screen:
            # 延迟关闭启动屏幕，确保用户能看到"已完成初始化"消息
            QTimer.singleShot(500, self._close_splash_and_show)
        else:
            # 如果没有启动屏幕，直接显示主窗口
            self.show()

    def _close_splash_and_show(self):
        """关闭启动屏幕并显示主窗口"""
        if self.splash_screen:
            self.splash_screen.hide_and_animate()
            QTimer.singleShot(350, self.show)  # 等待动画完成再显示主窗口
        else:
            self.show()
    def center_window(self):
        """将主窗口居中显示"""
        screen = QApplication.primaryScreen().geometry()
        window_width = 950
        window_height = 700
        
        x = (screen.width() - window_width) // 2
        y = (screen.height() - window_height) // 2
        
        self.setGeometry(x, y, window_width, window_height)

    def _recreate_ui(self):
        """从原来的initUI复制，但分阶段"""
        # 创建完整UI
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)

        # 顶部Help按钮 - 移到最上方靠左位置
        # 创建一个水平布局来容纳Help按钮
        top_status_layout = QHBoxLayout()

        # Help按钮 - 靠左放置，无下拉箭头
        self.help_btn = QPushButton('Help')
        self.help_btn.setFont(QFont("Arial", 9, QFont.Bold))
        self.help_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                color: #2c3e50;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 6px 12px;
                margin: 5px;
                min-width: 60px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
            QPushButton::menu-indicator {
                image: none;
                width: 0px;
            }
        """)
        self.help_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.help_btn.setMenu(self.create_help_menu())
        top_status_layout.addWidget(self.help_btn, alignment=Qt.AlignLeft)

        # 添加弹性空间将Help按钮推到左侧
        top_status_layout.addStretch()

        # 将顶部布局添加到主布局
        self.layout.addLayout(top_status_layout)

        # 添加分隔线
        separator = QLabel()
        separator.setStyleSheet("""
            background-color: #dee2e6;
            height: 1px;
            margin: 0;
        """)
        separator.setFixedHeight(1)
        self.layout.addWidget(separator)

        # 合并Excel和PDF设置到一行 - 去除标题文字
        settings_row = QHBoxLayout()
        settings_row.setSpacing(15)
        settings_row.setContentsMargins(0, 10, 0, 10)

        # Excel文件信息 - 左侧（去除标题文字）
        excel_group = QGroupBox("")  # 删除标题文字
        excel_group.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 12, QFont.Bold))
        excel_group.setStyleSheet("""
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 10px;
                margin-top: 0px;
            }
            QGroupBox::title {
                height: 0px;
                padding: 0px;
                margin: 0px;
                subcontrol-origin: margin;
            }
        """)
        excel_layout = QVBoxLayout()
        excel_layout.setSpacing(8)

        self.refresh_excel_btn = QPushButton('刷新Excel数据')
        self.refresh_excel_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.refresh_excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cabb8;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a8a96;
                color: white;
            }
            QPushButton:pressed {
                background-color: #386a74;
                color: white;
            }
        """)
        self.refresh_excel_btn.clicked.connect(self.refresh_excel_data)
        excel_layout.addWidget(self.refresh_excel_btn)

        self.excel_label = QLabel('正在查找Excel文件...')
        self.excel_label.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.excel_label.setWordWrap(True)
        self.excel_label.setStyleSheet("""
            color: #495057;
            padding: 4px;
            background-color: #F8F9FA;
            border: 1px solid #DEE2E6;
            border-radius: 4px;
        """)
        excel_layout.addWidget(self.excel_label)
        excel_group.setLayout(excel_layout)
        # PDF路径选择 - 左侧（占2/3空间）
        pdf_group = QGroupBox("")  # 删除标题文字
        pdf_group.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 12, QFont.Bold))
        pdf_group.setStyleSheet("""
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 10px;
                margin-top: 0px;
            }
            QGroupBox::title {
                height: 0px;
                padding: 0px;
                margin: 0px;
                subcontrol-origin: margin;
            }
        """)
        pdf_layout = QHBoxLayout()
        pdf_layout.setSpacing(10)

        self.pdf_input_btn = QPushButton('查看PDF输入文件夹')
        self.pdf_input_btn.setFont(QFont("Arial", 9))
        self.pdf_input_btn.setStyleSheet("""
            QPushButton {
                background-color: #98FB98;
                color: #333;
                border: 1px solid #90EE90;
                padding: 6px 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #90EE90;
            }
        """)
        self.pdf_input_btn.clicked.connect(self.show_pdf_input_dir)
        self.pdf_input_label = QLabel('PDF输入目录')
        self.pdf_input_label.setWordWrap(True)
        self.pdf_input_label.setStyleSheet("color: #7f8c8d; font-size: 13px;")

        self.pdf_output_btn = QPushButton('选择TXT输出文件夹')
        self.pdf_output_btn.setFont(QFont("Arial", 9))
        self.pdf_output_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFFACD;
                color: #333;
                border: 1px solid #EEE8AA;
                padding: 6px 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #EEE8AA;
            }
        """)
        self.pdf_output_btn.clicked.connect(self.select_pdf_output_dir)
        self.pdf_output_label = QLabel('TXT输出目录')
        self.pdf_output_label.setWordWrap(True)
        self.pdf_output_label.setStyleSheet("color: #7f8c8d; font-size: 13px;")

        pdf_left_col = QVBoxLayout()
        pdf_left_col.addWidget(self.pdf_input_btn)
        pdf_left_col.addWidget(self.pdf_input_label)
        pdf_right_col = QVBoxLayout()
        pdf_right_col.addWidget(self.pdf_output_btn)
        pdf_right_col.addWidget(self.pdf_output_label)
        pdf_layout.addLayout(pdf_left_col)
        pdf_layout.addLayout(pdf_right_col)
        pdf_group.setLayout(pdf_layout)
        settings_row.addWidget(pdf_group, stretch=2)  # PDF占2/3空间

        # Excel文件信息 - 右侧（占1/3空间）
        excel_group = QGroupBox("")  # 删除标题文字
        excel_group.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 12, QFont.Bold))
        excel_group.setStyleSheet("""
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 10px;
                margin-top: 0px;
            }
            QGroupBox::title {
                height: 0px;
                padding: 0px;
                margin: 0px;
                subcontrol-origin: margin;
            }
        """)
        excel_layout = QVBoxLayout()
        excel_layout.setSpacing(8)

        self.refresh_excel_btn = QPushButton('刷新Excel数据')
        self.refresh_excel_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.refresh_excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cabb8;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a8a96;
                color: white;
            }
            QPushButton:pressed {
                background-color: #386a74;
                color: white;
            }
        """)
        self.refresh_excel_btn.clicked.connect(self.refresh_excel_data)
        excel_layout.addWidget(self.refresh_excel_btn)

        self.excel_label = QLabel('正在查找Excel文件...')
        self.excel_label.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 9))
        self.excel_label.setWordWrap(True)
        self.excel_label.setStyleSheet("""
            color: #495057;
            padding: 4px;
            background-color: #F8F9FA;
            border: 1px solid #DEE2E6;
            border-radius: 4px;
        """)
        excel_layout.addWidget(self.excel_label)
        excel_group.setLayout(excel_layout)
        settings_row.addWidget(excel_group, stretch=1)  # Excel占1/3空间

        self.layout.addLayout(settings_row)

        # 功能按钮组 - 删除标题
        button_group = QGroupBox("")  # 删除标题文字
        button_group.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 12, QFont.Bold))
        button_group.setStyleSheet("""
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                margin-top: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                height: 0px;
                padding: 0px;
                margin: 0px;
                subcontrol-origin: margin;
            }
        """)
        button_layout = QVBoxLayout()
        button_layout.setSpacing(10)

        # 创建网格布局
        button_grid = QGridLayout()
        button_grid.setSpacing(10)
        button_grid.setContentsMargins(0, 0, 0, 0)

        # 第一行按钮
        self.outlook_btn = QPushButton('生成Outlook邮件')
        self.outlook_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.outlook_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cabb8;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a8a96;
                color: white;
            }
            QPushButton:pressed {
                background-color: #386a74;
                color: white;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.outlook_btn.clicked.connect(self.run_outlook)
        button_grid.addWidget(self.outlook_btn, 0, 0)

        self.memo_btn = QPushButton('生成MEMO')
        self.memo_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.memo_btn.setStyleSheet("""
            QPushButton {
                background-color: #1ABC9C;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #16A085;
            }
            QPushButton:pressed {
                background-color: #117A65;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.memo_btn.clicked.connect(self.run_memo)
        button_grid.addWidget(self.memo_btn, 0, 1)

        self.pdf_btn = QPushButton('收集云盘步距规数据')
        self.pdf_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.pdf_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cabb8;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a8a96;
                color: white;
            }
            QPushButton:pressed {
                background-color: #386a74;
                color: white;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.pdf_btn.clicked.connect(self.run_pdf_extract)
        button_grid.addWidget(self.pdf_btn, 0, 2)

        self.file_search_btn = QPushButton('搜索文件内容')
        self.file_search_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.file_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #1ABC9C;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #16A085;
            }
            QPushButton:pressed {
                background-color: #117A65;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.file_search_btn.clicked.connect(self.run_file_search)
        button_grid.addWidget(self.file_search_btn, 1, 0)

        self.file_converter_btn = QPushButton('文件转换器')
        self.file_converter_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.file_converter_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cabb8;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a8a96;
                color: white;
            }
            QPushButton:pressed {
                background-color: #386a74;
                color: white;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.file_converter_btn.clicked.connect(self.run_file_converter)
        button_grid.addWidget(self.file_converter_btn, 1, 1)

        self.folder_btn = QPushButton('创建DATA文件夹')
        self.folder_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.folder_btn.setStyleSheet("""
            QPushButton {
                background-color: #1ABC9C;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #16A085;
            }
            QPushButton:pressed {
                background-color: #117A65;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.folder_btn.clicked.connect(self.run_folder_creation)
        button_grid.addWidget(self.folder_btn, 1, 2)

        self.cancel_btn = QPushButton('取消任务')
        self.cancel_btn.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 10))
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
            QPushButton:pressed {
                background-color: #992E22;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
                color: #95A5A6;
            }
        """)
        self.cancel_btn.clicked.connect(self.cancel_task)
        self.cancel_btn.setEnabled(False)
        button_grid.addWidget(self.cancel_btn, 2, 1)

        button_layout.addLayout(button_grid)
        button_group.setLayout(button_layout)
        self.layout.addWidget(button_group)

        # 操作日志 - 给予更多空间（删除标题）
        log_group = QGroupBox("")  # 删除标题文字
        log_group.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 12, QFont.Bold))
        log_group.setStyleSheet("""
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                margin-top: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                height: 0px;
                padding: 0px;
                margin: 0px;
                subcontrol-origin: margin;
            }
        """)
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto", 9))
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #DEE2E6;
                border-radius: 4px;
                background-color: #F8F9FA;
                padding: 10px;
                color: #495057;
                line-height: 1.5;
            }
        """)
        
        # 添加日志到日志组
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        
        # 将日志组添加到主布局
        self.layout.addWidget(log_group, stretch=1)  # 使用stretch参数让日志区域占据更多空间

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #DEE2E6;
                border-radius: 4px;
                height: 12px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #90caf9;
                border-radius: 3px;
            }
        """)
        self.layout.addWidget(self.progress_bar)

    def _phase2_load(self):
        """第二阶段：加载功能模块"""
        if self.splash_screen:
            self.splash_screen.update_status('初始化功能模块...')
        
        self.update_log('初始化功能模块...')

        try:
            # 延迟导入config
            from modules import config
            self.pdf_input_dir = config.PDF_INPUT_DIR
            self.pdf_output_dir = config.PDF_OUTPUT_DIR

            # 更新UI标签
            self.pdf_input_label.setText(self.pdf_input_dir)
            self.pdf_output_label.setText(self.pdf_output_dir)

            # 延迟查找Excel
            QTimer.singleShot(50, self._phase3_load)
        except ImportError as e:
            self.update_log(f"❌ 导入配置模块失败: {str(e)}")
            self.update_log("⚠️  请确保modules/config.py文件存在且配置正确")
            self._finalize_loading()

    def _phase3_load(self):
        """第三阶段：查找Excel文件"""
        if self.splash_screen:
            self.splash_screen.update_status('查找Excel文件...')
        
        self.find_and_display_excel()
        self._finalize_loading()

    def _finalize_loading(self):
        """完成加载，关闭启动屏幕并显示主窗口"""
        if self.splash_screen:
            # 延迟关闭启动屏幕，确保用户能看到"已完成初始化"消息
            QTimer.singleShot(500, self._close_splash_and_show)
        else:
            # 如果没有启动屏幕，直接显示主窗口
            self.show()

    def _close_splash_and_show(self):
        """关闭启动屏幕并显示主窗口"""
        if self.splash_screen:
            self.splash_screen.hide_and_animate()
            QTimer.singleShot(350, self.show)  # 等待动画完成再显示主窗口
        else:
            self.show()

    # -------------------------- Help菜单功能 --------------------------
    def create_help_menu(self):
        """创建问号按钮的下拉菜单"""
        help_menu = QMenu(self)
        help_menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #dcdcdc;
                font-family: Arial;
                font-size: 10pt;
                font-weight: normal;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 25px 6px 25px;
                padding: 4px 20px;
            }
            QMenu::item:selected {
            background-color: #e3f2fd;
            color: #1976d2;
        }
            QMenu::item:pressed {
            background-color: #bbdefb;
            color: #0d47a1;
        """)
        # Version菜单项
        version_action = QAction("Version", self)
        version_action.triggered.connect(self.show_version)
        help_menu.addAction(version_action)

        # Manual菜单项
        manual_action = QAction("Manual", self)
        manual_action.triggered.connect(self.open_manual)
        help_menu.addAction(manual_action)

        return help_menu

    def show_version(self):
        """显示版本号弹窗"""
        QMessageBox.information(self, "版本信息", "Version: V5.6\n "
                                "更新内容\n"
                                "优化启动速度\n"
                                "更新获取txt最新文件方法\n"
                                "支持多个邮件创建\n"
                                "增加快捷方式", QMessageBox.Ok)

    def open_manual(self):
        """打开使用说明（exe文件同级目录的说明文件）"""
        # 检测是否为PyInstaller打包的exe文件
        if getattr(sys, 'frozen', False):
            # 如果是exe文件运行，获取exe文件所在目录
            exe_dir = os.path.dirname(sys.executable)
            manual_path = os.path.join(exe_dir, "Automation tool使用说明.pdf")
        else:
            # 如果是Python脚本运行，使用脚本所在目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            manual_path = os.path.join(current_dir, "Automation tool使用说明.pdf")

        if os.path.exists(manual_path):
            os.startfile(manual_path)  # 用系统默认程序打开
        else:
            # 尝试其他可能的文件名
            alternative_names = [
                "Automation tool使用说明.pdf",
                "Automation tool使用说明.docx",
                "使用说明.pdf",
                "使用说明.docx",
                "manual.pdf",
                "manual.docx"
            ]

            found = False
            for alt_name in alternative_names:
                alt_path = os.path.join(exe_dir if getattr(sys, 'frozen', False) else current_dir, alt_name)
                if os.path.exists(alt_path):
                    os.startfile(alt_path)
                    found = True
                    break

            if not found:
                exe_dir_info = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else "脚本所在目录"
                QMessageBox.warning(
                    self, "文件缺失",
                    f"未找到使用说明文件！\n请将使用说明文件放在exe文件同级目录下：\n{exe_dir_info}\n\n支持的文件名：\n- Automation tool使用说明.pdf\n- Automation tool使用说明.docx\n- 使用说明.pdf\n- 使用说明.docx"
                )

    # -------------------------- 辅助方法 --------------------------
    def find_and_display_excel(self):
        """查找并显示Excel文件信息"""
        try:
            # 延迟导入find_excel_file
            find_excel_file_func = get_find_excel_file()
            self.excel_path, message = find_excel_file_func()
            self.excel_label.setText(message)
            excel_exists = self.excel_path is not None
            self.outlook_btn.setEnabled(excel_exists)
            self.memo_btn.setEnabled(excel_exists)
            self.folder_btn.setEnabled(True)
            self.pdf_btn.setEnabled(True)
            self.update_log('已完成初始化')
        except Exception as e:
            self.update_log(f"❌ 查找Excel文件失败: {str(e)}")
            self.update_log('已完成初始化')

    def refresh_excel_data(self):
        """重新读取Excel文件，刷新数据"""
        self.update_log("正在刷新Excel数据...")
        try:
            # 延迟导入find_excel_file
            find_excel_file_func = get_find_excel_file()
            self.excel_path, message = find_excel_file_func()
            self.excel_label.setText(message)

            excel_exists = self.excel_path is not None
            self.outlook_btn.setEnabled(excel_exists)
            self.memo_btn.setEnabled(excel_exists)

            if excel_exists:
                self.update_log(" Excel数据已刷新（修改内容已生效）")
            else:
                self.update_log(" 未找到Excel文件，刷新失败")
        except Exception as e:
            self.update_log(f"❌ 刷新Excel数据失败: {str(e)}")

    def _prepare_task(self, disable_all_buttons=True):
        """准备任务：禁用按钮、启用取消按钮、显示进度条、清空日志"""
        if disable_all_buttons:
            self.outlook_btn.setEnabled(False)
            self.memo_btn.setEnabled(False)
            self.folder_btn.setEnabled(False)
            self.pdf_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()

    def _reset_task_state(self):
        """重置任务状态：恢复按钮可用状态"""
        excel_exists = self.excel_path is not None
        self.outlook_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)
        self.folder_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)

        # 更新取消按钮状态
        self._update_cancel_button_state()
        self.progress_bar.setVisible(False)
        
    def _update_cancel_button_state(self):
        """更新取消按钮状态：检查是否有任何任务正在运行"""
        any_task_running = False

        # 检查所有任务线程
        if hasattr(self, 'outlook_thread') and self.outlook_thread and self.outlook_thread.isRunning():
            any_task_running = True
        if hasattr(self, 'memo_thread') and self.memo_thread and self.memo_thread.isRunning():
            any_task_running = True
        if hasattr(self, 'pdf_thread') and self.pdf_thread and self.pdf_thread.isRunning():
            any_task_running = True
        if hasattr(self, 'folder_thread') and self.folder_thread and self.folder_thread.isRunning():
            any_task_running = True

        self.cancel_btn.setEnabled(any_task_running)
        return any_task_running

    def update_log(self, message):
        """更新日志显示"""
        self.log_text.append(message)
        self.statusBar().showMessage(message)
        self.log_text.moveCursor(self.log_text.textCursor().End)

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    # -------------------------- PDF相关方法 --------------------------
    def show_pdf_input_dir(self):
        """显示PDF输入文件夹"""
        if self.pdf_input_dir and os.path.exists(self.pdf_input_dir):
            os.startfile(self.pdf_input_dir)
        else:
            QMessageBox.warning(self, "路径不存在", f"PDF输入文件夹不存在：\n{self.pdf_input_dir}")

    def select_pdf_output_dir(self):
        """选择PDF输出文件夹"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择TXT输出文件夹", self.pdf_output_dir
        )
        if dir_path:
            self.pdf_output_dir = dir_path
            self.pdf_output_label.setText(dir_path)

    def run_pdf_extract(self):
        """运行PDF提取任务"""
        if not os.path.exists(self.pdf_input_dir):
            QMessageBox.warning(self, "路径错误", f"PDF输入文件夹不存在：\n{self.pdf_input_dir}")
            return

        # 检查是否有PDF任务正在运行
        if hasattr(self, 'pdf_thread') and self.pdf_thread and self.pdf_thread.isRunning():
            QMessageBox.warning(self, "警告", "PDF提取任务正在运行，请等待完成后再启动新任务。")
            return

        self._prepare_task(disable_all_buttons=False)
        self.pdf_btn.setEnabled(False)  # Only disable the specific button
        self.update_log("开始执行PDF表格提取任务...")
        self.update_log(f"PDF输入路径：{self.pdf_input_dir}")
        self.update_log(f"TXT输出路径：{self.pdf_output_dir}")

        self.pdf_thread = PdfExtractThread(
            input_dir=self.pdf_input_dir,
            output_dir=self.pdf_output_dir
        )
        self.current_thread = self.pdf_thread  # Keep reference for cancel functionality
        self.pdf_thread.log.connect(self.update_log)
        self.pdf_thread.progress.connect(self.update_progress)
        self.pdf_thread.finished.connect(self.on_pdf_finished)
        self.pdf_thread.start()

    def on_pdf_finished(self, success):
        """PDF提取任务完成回调"""
        self._reset_task_state()
        if success:
            self.update_log("PDF提取任务已完成！")
            self.statusBar().showMessage("PDF提取任务已完成")
            reply = QMessageBox.question(
                self, "完成",
                f"PDF提取任务已完成，是否打开输出文件夹？\n{self.pdf_output_dir}",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                os.startfile(self.pdf_output_dir)
        else:
            self.update_log("PDF提取任务失败！")
            self.statusBar().showMessage("PDF提取任务失败")
        self.pdf_thread = None
        self.current_thread = None

    # -------------------------- 其他功能方法 --------------------------
    def run_outlook(self):
        if not self.excel_path:
            QMessageBox.warning(self, "错误", "未找到Excel文件，请检查tool文件夹")
            return

        # 检查是否有Outlook任务正在运行
        if hasattr(self, 'outlook_thread') and self.outlook_thread and self.outlook_thread.isRunning():
            QMessageBox.warning(self, "警告", "Outlook任务正在运行，请等待完成后再启动新任务。")
            return

        self._prepare_task(disable_all_buttons=False)
        self.outlook_btn.setEnabled(False)  # Only disable the specific button
        self.update_log("开始生成Outlook邮件...")
        OutlookEmailThread = get_outlook_email_thread()
        self.outlook_thread = OutlookEmailThread(self.excel_path)
        self.current_thread = self.outlook_thread  # Keep reference for cancel functionality
        
        # 注意：OutlookEmailThread需要实现progress和finished信号
        if hasattr(self.outlook_thread, 'progress'):
            self.outlook_thread.progress.connect(self.update_log)
        if hasattr(self.outlook_thread, 'finished'):
            self.outlook_thread.finished.connect(self.on_outlook_finished)
        
        self.outlook_thread.start()

    def on_outlook_finished(self, success):
        self._reset_task_state()
        if success:
            self.update_log("Outlook邮件生成完成！")
            self.statusBar().showMessage("Outlook邮件生成完成")
        else:
            self.update_log("Outlook邮件生成失败！")
            self.statusBar().showMessage("Outlook邮件生成失败")
        self.outlook_thread = None
        self.current_thread = None

    def run_folder_creation(self):
        # 检查是否有文件夹任务正在运行
        if hasattr(self, 'folder_thread') and self.folder_thread and self.folder_thread.isRunning():
            QMessageBox.warning(self, "警告", "文件夹创建任务正在运行，请等待完成后再启动新任务。")
            return

        self._prepare_task(disable_all_buttons=False)
        self.folder_btn.setEnabled(False)  # Only disable the specific button
        self.update_log("开始执行文件夹创建+文件检索流程...")
        self.folder_thread = FolderThread()
        self.current_thread = self.folder_thread  # Keep reference for cancel functionality
        self.folder_thread.progress.connect(self.update_log)
        self.folder_thread.finished.connect(self.on_folder_finished)
        self.folder_thread.start()

    def on_folder_finished(self, success):
        self._reset_task_state()
        if success:
            self.update_log("文件夹创建+文件检索流程完成！")
            self.statusBar().showMessage("文件夹流程完成")
        else:
            self.update_log("文件夹创建+文件检索流程失败！")
            self.statusBar().showMessage("文件夹流程失败")
        self.folder_thread = None
        self.current_thread = None

    def run_memo(self):
        if not self.excel_path:
            QMessageBox.warning(self, "错误", "未找到Excel文件，请检查tool文件夹")
            return

        template_path = os.path.join(os.path.expanduser("~"), "Desktop", "tool", "MemoTemplate.docx")
        if not os.path.exists(template_path):
            QMessageBox.warning(
                self, "模板缺失",
                f"未找到MEMO模板：{template_path}\n请将MemoTemplate.docx放入tool文件夹后重试"
            )
            return

        # 检查是否有MEMO任务正在运行
        if hasattr(self, 'memo_thread') and self.memo_thread and self.memo_thread.isRunning():
            QMessageBox.warning(self, "警告", "MEMO生成任务正在运行，请等待完成后再启动新任务。")
            return

        self._prepare_task(disable_all_buttons=False)
        self.memo_btn.setEnabled(False)  # Only disable the specific button
        self.update_log("开始生成MEMO...")

        self.memo_thread = MemoThread(excel_path=self.excel_path)
        self.current_thread = self.memo_thread  # Keep reference for cancel functionality
        self.memo_thread.progress.connect(self.update_log)
        self.memo_thread.finished.connect(self.on_memo_finished)
        self.memo_thread.start()

    def on_memo_finished(self, success, msg):
        self._reset_task_state()
        self.update_log(f"\n{msg}")
        self.statusBar().showMessage(msg)
        if success:
            # Extract file path from message if present
            file_path = ""
            if "（" in msg and "）" in msg:
                file_path = msg.split("（")[1].split("）")[0]

            if file_path and os.path.exists(file_path):
                # Create a custom dialog with clickable file path
                dialog = QDialog(self)
                dialog.setWindowTitle("生成成功")
                dialog.setMinimumWidth(400)

                layout = QVBoxLayout(dialog)

                # Success icon and main text
                icon_label = QLabel()
                icon_label.setPixmap(QApplication.style().standardIcon(QStyle.SP_MessageBoxInformation).pixmap(32, 32))
                layout.addWidget(icon_label, alignment=Qt.AlignCenter)

                title_label = QLabel("MEMO生成成功！")
                title_label.setFont(QFont("Arial", 12, QFont.Bold))
                layout.addWidget(title_label, alignment=Qt.AlignCenter)

                # File path display
                path_label = QLabel(f"文件已保存：{file_path}")
                path_label.setWordWrap(True)
                path_label.setStyleSheet("color: #2C3E50; margin: 10px 0;")
                layout.addWidget(path_label)

                # Clickable link
                file_path_forward = file_path.replace("\\", "/")
                link_label = QLabel(f'<a href="file:///{file_path_forward}">点击打开文件</a>')
                link_label.setOpenExternalLinks(True)
                link_label.setStyleSheet("color: #3498DB; text-decoration: underline;")
                link_label.setAlignment(Qt.AlignCenter)
                link_label.setCursor(QCursor(Qt.PointingHandCursor))
                layout.addWidget(link_label)

                # OK button
                ok_button = QPushButton("确定")
                ok_button.setStyleSheet("""
                    QPushButton {
                        background-color: #;
                        color: white;
                        border: none;
                        padding: 8px 16px;
                        border-radius: 4px;
                        font-weight: 500;
                    }
                    QPushButton:hover {
                        background-color: #229954;
                    }
                """)
                ok_button.clicked.connect(dialog.accept)
                layout.addWidget(ok_button, alignment=Qt.AlignCenter)

                dialog.exec_()
            else:
                # Fallback to simple message box
                QMessageBox.information(self, "生成成功", msg)
        # 确保线程变量被正确清理
        self.memo_thread = None
        self.current_thread = None
        # 更新取消按钮状态
        self._update_cancel_button_state()

    def run_file_search(self):
        """运行文件搜索功能 - 弹出独立窗口"""
        # 创建独立的搜索窗口
        search_window = FileSearchDialog(self)
        search_window.exec_()

    def run_file_converter(self):
        """运行文件转换器功能 - 使用独立的UI界面"""
        self.update_log("🚀 启动文件转换器...")

        try:
            # 创建文件转换器UI窗口
            FileConverterUI = get_file_converter_ui()
            self.file_converter_ui = FileConverterUI()

            # 设置为模态对话框
            self.file_converter_ui.setWindowModality(Qt.ApplicationModal)

            # 显示窗口
            self.file_converter_ui.show()

            self.update_log(" 文件转换器UI已启动")
        except Exception as e:
            self.update_log(f"❌ 启动文件转换器失败: {str(e)}")

    def cancel_task(self):
        # 检查是否有任何任务正在运行
        if not self._update_cancel_button_state():
            QMessageBox.information(self, "提示", "当前没有正在执行的任务")
            return

        reply = QMessageBox.question(
            self, "确认取消", "确定要取消当前任务吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            # 尝试取消所有可能正在运行的任务
            if hasattr(self, 'outlook_thread') and self.outlook_thread and self.outlook_thread.isRunning():
                if hasattr(self.outlook_thread, 'cancel'):
                    self.outlook_thread.cancel()
                self.outlook_thread = None

            if hasattr(self, 'memo_thread') and self.memo_thread and self.memo_thread.isRunning():
                if hasattr(self.memo_thread, 'cancel'):
                    self.memo_thread.cancel()
                self.memo_thread = None

            if hasattr(self, 'pdf_thread') and self.pdf_thread and self.pdf_thread.isRunning():
                if hasattr(self.pdf_thread, 'cancel'):
                    self.pdf_thread.cancel()
                self.pdf_thread = None

            if hasattr(self, 'folder_thread') and self.folder_thread and self.folder_thread.isRunning():
                if hasattr(self.folder_thread, 'cancel'):
                    self.folder_thread.cancel()
                self.folder_thread = None

            self.current_thread = None
            self.update_log("所有任务已取消")
            self._reset_task_state()  # 恢复所有按钮状态

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    import sys
    import time

    # 记录启动时间
    start_time = time.perf_counter()

    # 1. 提前设置环境变量，优化Qt启动
    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = ""
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

    # 2. 禁用Qt调试信息（显著加速）
    os.environ["QT_LOGGING_RULES"] = "*.debug=false;*.info=false;*.warning=false"

    # 3. 设置Windows进程优先级（仅Windows）
    if sys.platform == 'win32':
        import ctypes
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("automation.tool.Automation Tool.v3.0")
            # 设置进程为高优先级
            ctypes.windll.kernel32.SetPriorityClass(-1, 0x00000080)  # HIGH_PRIORITY_CLASS
        except:
            pass
            
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

     # 设置全局字体（所有控件都会继承这个字体）
    font = QFont("Microsoft YaHei", 10)  # 使用微软雅黑字体
    app.setFont(font)
    
    # 立即加载应用程序图标
    def load_app_icon():
        # 使用根目录下的tool_icon.ico
        icon_path = "tool_icon.ico"
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
    load_app_icon()  # 立即加载，不延迟

    # 创建并显示启动屏幕
    splash = SplashScreen()
    splash.show_and_animate()

    # 强制处理事件，确保启动屏幕显示
    QApplication.processEvents()

    # 创建主窗口，传递启动屏幕引用
    window = MainWindow(splash_screen=splash)

    # 在主窗口显示前检查是否需要显示快捷方式对话框
    def check_shortcut_dialog():
        if should_show_shortcut_dialog():
            # 获取应用名称
            app_name = get_app_name()

            # 创建快捷方式对话框
            dialog = ShortcutDialog()
            dialog.set_app_name(app_name)

            # 显示对话框并等待用户响应
            result = dialog.exec_()

            if result == QDialog.Accepted:  # 用户点击"创建"
                print(f"用户选择创建桌面快捷方式...")
                if create_desktop_shortcut():
                    save_shortcut_choice("yes")
                    print("桌面快捷方式创建成功")
                else:
                    print("桌面快捷方式创建失败")
            else:  # 用户点击"取消"
                print("用户选择不创建桌面快捷方式")

            # 标记对话框已显示
            dont_show_again = dialog.dont_ask_checkbox.isChecked()
            mark_shortcut_dialog_shown(dont_show_again)

    # 在主窗口显示后延迟执行快捷方式检查
    QTimer.singleShot(100, check_shortcut_dialog)

    # 记录启动时间
    def log_startup_time():
        elapsed = time.perf_counter() - start_time
        print(f"🚀 应用程序启动时间: {elapsed:.2f}秒")

    QTimer.singleShot(2000, log_startup_time)

    sys.exit(app.exec_())

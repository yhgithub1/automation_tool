import sys
from pathlib import Path

# 关键：添加项目根目录到 Python 搜索路径
# 1. 获取当前脚本（correction.py）的绝对路径
current_script = Path(__file__).resolve()
# 2. 项目根目录是 current_script 的父目录的父目录（automation_tool → pythonProject）
project_root = current_script.parent.parent
# 3. 将根目录添加到搜索路径
sys.path.append(str(project_root))

import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout,
                             QWidget, QLabel, QMessageBox, QTextEdit, QProgressBar, QGroupBox,
                             QFileDialog, QMenu, QAction, QDialog, QFormLayout, QLineEdit, 
                             QCheckBox, QScrollArea, QFrame, QInputDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QCursor
from tkinter import simpledialog

# 导入模块
from modules.outlook_automation import OutlookEmailThread
from modules.folder_creation import FolderCreator
from modules.memo_generator import generate_memo
from utils.file_utils import find_excel_file
from modules.pdf_extractor import PdfTableExtractor
from modules.findfile import find_files_with_progress
from modules.file_converter import FileConverter
from modules.file_converter_ui import FileConverterUI


# -------------------------- 文件搜索对话框类 --------------------------
class FileSearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_thread = None
        
        # 设置窗口图标（在initUI之前设置）
        # 获取桌面tool文件夹路径
        desktop_path = os.path.expanduser("~/Desktop")
        tool_folder = os.path.join(desktop_path, "tool")
        
        # 确保tool文件夹存在
        if not os.path.exists(tool_folder):
            os.makedirs(tool_folder)
            
        icon_path = os.path.join(tool_folder, 'robot-solid-full.svg')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"图标文件不存在: {icon_path}")
            
        self.initUI()
        # 设置默认值
        self.search_dir_input.setText(r"C:\Zeiss\CMM_Tools\FW_C99\backup")
        self.search_content_input.setText("Install_version = V47.04")
        self.file_names_input.setText("config.kmg")

    def initUI(self):
        self.setWindowTitle('文件内容搜索工具')
        self.setGeometry(300, 300, 800, 600)
        
        # 获取桌面tool文件夹路径
        desktop_path = os.path.expanduser("~/Desktop")
        tool_folder = os.path.join(desktop_path, "tool")
        
        # 确保tool文件夹存在
        if not os.path.exists(tool_folder):
            os.makedirs(tool_folder)
            
        icon_path = os.path.join(tool_folder, 'robot-solid-full.svg')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"图标文件不存在: {icon_path}")

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
        self.folder_creator = FolderCreator()
        self.is_canceled = False

    def run(self):
        self.folder_creator.log_signal.connect(self.progress)
        self.folder_creator.finished.connect(self.on_finished)
        if not self.is_canceled:
            self.folder_creator.create_folders()
        else:
            self.progress.emit("文件夹创建任务已被取消")
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
            self.progress.emit("📋 启动备忘录生成任务...")
            success, msg, output_path = generate_memo(
                excel_path=self.excel_path,
                progress_callback=lambda log: self.progress.emit(log)
            )
            self.finished.emit(success, msg)
        except Exception as e:
            err_msg = f"备忘录线程出错：{str(e)}"
            self.progress.emit(f"❌ {err_msg}")
            self.finished.emit(False, err_msg)

    def cancel(self):
        self.is_canceled = True
        self.progress.emit("⏹️  正在取消备忘录生成任务...")


class PdfExtractThread(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool)

    def __init__(self, input_dir, output_dir):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.extractor = PdfTableExtractor()

    def run(self):
        self.extractor.log_signal.connect(self.log)
        self.extractor.progress_signal.connect(self.progress)
        self.extractor.finished_signal.connect(self.finished)
        self.extractor.set_paths(self.input_dir, self.output_dir)
        self.extractor.batch_extract()

    def cancel(self):
        if hasattr(self.extractor, 'cancel_extract'):
            self.extractor.cancel_extract()


# -------------------------- 主窗口类 --------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.current_thread = None
        self.pdf_input_dir = PdfTableExtractor.DEFAULT_INPUT_DIR
        self.pdf_output_dir = PdfTableExtractor.DEFAULT_OUTPUT_DIR
        self.initUI()
        self.find_and_display_excel()

    def initUI(self):
        self.setWindowTitle('自动化工具集')
        self.setGeometry(300, 300, 900, 600)
        
        # 直接从桌面tool文件夹查找图标文件
        desktop_path = os.path.expanduser("~/Desktop")
        tool_folder = os.path.join(desktop_path, "tool")
        icon_path = os.path.join(tool_folder, 'robot-solid-full.svg')
        
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"主窗口图标文件不存在: {icon_path}")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 标题栏：标题 + 问号帮助按钮
        title_bar_layout = QHBoxLayout()

        # 标题
        title_label = QLabel('自动化小工具')
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50; margin: 10px 0;")
        title_bar_layout.addWidget(title_label, stretch=1)

        # 问号帮助按钮（带下拉菜单）
        self.help_btn = QPushButton('?')
        self.help_btn.setFont(QFont("Arial", 10, QFont.Bold))
        self.help_btn.setStyleSheet("""
            QPushButton { 
                background-color: #f8f9fa; 
                color: #2c3e50; 
                border: 1px solid #dee2e6; 
                border-radius: 50%; 
                width: 30px; 
                height: 30px; 
                margin: 10px 10px 10px 0;
            }
            QPushButton:hover { 
                background-color: #e9ecef; 
            }
        """)
        self.help_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.help_btn.setMenu(self.create_help_menu())
        title_bar_layout.addWidget(self.help_btn, alignment=Qt.AlignRight)

        layout.addLayout(title_bar_layout)

        # Excel文件信息组
        file_group = QGroupBox("Excel文件信息")
        file_layout = QVBoxLayout()
        self.refresh_excel_btn = QPushButton('刷新Excel数据')
        self.refresh_excel_btn.setFont(QFont("Arial", 9))
        self.refresh_excel_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                padding: 5px; 
                border-radius: 4px;
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
        """)
        self.refresh_excel_btn.clicked.connect(self.refresh_excel_data)
        file_layout.addWidget(self.refresh_excel_btn)

        self.excel_label = QLabel('正在查找Excel文件...')
        self.excel_label.setWordWrap(True)
        file_layout.addWidget(self.excel_label)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # PDF路径选择组
        pdf_path_group = QGroupBox("PDF提取路径设置")
        pdf_path_layout = QHBoxLayout()

        self.pdf_input_btn = QPushButton('查看PDF输入文件夹')
        self.pdf_input_btn.setFont(QFont("Arial", 9))
        self.pdf_input_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                padding: 8px; 
                margin: 5px; 
                border-radius: 6px;
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
        """)
        self.pdf_input_btn.clicked.connect(self.show_pdf_input_dir)
        self.pdf_input_label = QLabel(self.pdf_input_dir)
        self.pdf_input_label.setWordWrap(True)
        self.pdf_input_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")

        self.pdf_output_btn = QPushButton('选择TXT输出文件夹')
        self.pdf_output_btn.setFont(QFont("Arial", 9))
        self.pdf_output_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                padding: 8px; 
                margin: 5px; 
                border-radius: 6px;
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
        """)
        self.pdf_output_btn.clicked.connect(self.select_pdf_output_dir)
        self.pdf_output_label = QLabel(self.pdf_output_dir)
        self.pdf_output_label.setWordWrap(True)
        self.pdf_output_label.setStyleSheet("color: #7f8c8d; font-size: 15px;")

        pdf_input_col = QVBoxLayout()
        pdf_input_col.addWidget(self.pdf_input_btn)
        pdf_input_col.addWidget(self.pdf_input_label)
        pdf_output_col = QVBoxLayout()
        pdf_output_col.addWidget(self.pdf_output_btn)
        pdf_output_col.addWidget(self.pdf_output_label)
        pdf_path_layout.addLayout(pdf_input_col)
        pdf_path_layout.addLayout(pdf_output_col)
        pdf_path_group.setLayout(pdf_path_layout)
        layout.addWidget(pdf_path_group)

        # 功能按钮组
        button_group = QGroupBox("功能")
        button_layout = QVBoxLayout()

        top_btn_layout = QHBoxLayout()
        self.outlook_btn = QPushButton('生成Outlook邮件')
        self.outlook_btn.setFont(QFont("Arial", 9))
        self.outlook_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #90caf9; 
                padding: 10px; 
                margin: 3px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.outlook_btn.clicked.connect(self.run_outlook)
        top_btn_layout.addWidget(self.outlook_btn)

        self.memo_btn = QPushButton('生成MEMO')
        self.memo_btn.setFont(QFont("Arial", 9))
        self.memo_btn.setStyleSheet("""
            QPushButton { 
                background-color: #f3e5f5; 
                color: #6a1b9a; 
                border: 1px solid #ce93d8; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #ce93d8; 
                color: white;
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.memo_btn.clicked.connect(self.run_memo)
        top_btn_layout.addWidget(self.memo_btn)

        self.pdf_btn = QPushButton('收集云盘步距规数据')
        self.pdf_btn.setFont(QFont("Arial", 9))
        self.pdf_btn.setStyleSheet("""
            QPushButton { 
                background-color: #fff3e0; 
                color: #e65100; 
                border: 1px solid #ffe0b2; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #ffe0b2; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.pdf_btn.clicked.connect(self.run_pdf_extract)
        top_btn_layout.addWidget(self.pdf_btn)

        self.file_search_btn = QPushButton('搜索文件内容')
        self.file_search_btn.setFont(QFont("Arial", 9))
        self.file_search_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e8f5e8; 
                color: #2e7d32; 
                border: 1px solid #c8e6c9; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #c8e6c9; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.file_search_btn.clicked.connect(self.run_file_search)
        top_btn_layout.addWidget(self.file_search_btn)

        self.file_converter_btn = QPushButton('文件转换器')
        self.file_converter_btn.setFont(QFont("Arial", 9))
        self.file_converter_btn.setStyleSheet("""
            QPushButton { 
                background-color: #fff8e1; 
                color: #f57f17; 
                border: 1px solid #ffecb3; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #ffecb3; 
            }
        """)
        self.file_converter_btn.clicked.connect(self.run_file_converter)
        top_btn_layout.addWidget(self.file_converter_btn)

        button_layout.addLayout(top_btn_layout)

        bottom_btn_layout = QHBoxLayout()
        self.folder_btn = QPushButton('创建DATA文件夹&检索tool文件')
        self.folder_btn.setFont(QFont("Arial", 9))
        self.folder_btn.setStyleSheet("""
            QPushButton { 
                background-color: #fffde7; 
                color: #f57f17; 
                border: 1px solid #fff9c4; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #fff9c4; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.folder_btn.clicked.connect(self.run_folder_creation)
        bottom_btn_layout.addWidget(self.folder_btn)

        self.cancel_btn = QPushButton('取消任务')
        self.cancel_btn.setFont(QFont("Arial", 9))
        self.cancel_btn.setStyleSheet("""
            QPushButton { 
                background-color: #ffebee; 
                color: #c62828; 
                border: 1px solid #ffcdd2; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;
            }
            QPushButton:hover { 
                background-color: #ffcdd2; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.cancel_btn.clicked.connect(self.cancel_task)
        self.cancel_btn.setEnabled(False)
        bottom_btn_layout.addWidget(self.cancel_btn)
        button_layout.addLayout(bottom_btn_layout)

        button_group.setLayout(button_layout)
        layout.addWidget(button_group)

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

        # 日志显示组
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #fafafa;
                padding: 5px;
            }
        """)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 状态栏
        self.statusBar().showMessage('就绪')

    # -------------------------- 帮助菜单功能 --------------------------
    def create_help_menu(self):
        """创建问号按钮的下拉菜单"""
        help_menu = QMenu(self)

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
        QMessageBox.information(self, "版本信息", "Version: V3.0\n 更新内容：新增使用说明；优化搜索txt算法；加快爬虫速度；增加文件搜索；增加pdf转换", QMessageBox.Ok)

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
        self.excel_path, message = find_excel_file()
        self.excel_label.setText(message)
        excel_exists = self.excel_path is not None
        self.outlook_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)
        self.folder_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)

    def refresh_excel_data(self):
        """重新读取Excel文件，刷新数据"""
        self.update_log("正在刷新Excel数据...")
        self.excel_path, message = find_excel_file()
        self.excel_label.setText(message)

        excel_exists = self.excel_path is not None
        self.outlook_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)

        if excel_exists:
            self.update_log(" Excel数据已刷新（修改内容已生效）")
        else:
            self.update_log(" 未找到Excel文件，刷新失败")

    def _prepare_task(self):
        """准备任务：禁用按钮、启用取消按钮、显示进度条、清空日志"""
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
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setVisible(False)

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
        if os.path.exists(self.pdf_input_dir):
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

        self._prepare_task()
        self.update_log("开始执行PDF表格提取任务...")
        self.update_log(f"PDF输入路径：{self.pdf_input_dir}")
        self.update_log(f"TXT输出路径：{self.pdf_output_dir}")

        self.current_thread = PdfExtractThread(
            input_dir=self.pdf_input_dir,
            output_dir=self.pdf_output_dir
        )
        self.current_thread.log.connect(self.update_log)
        self.current_thread.progress.connect(self.update_progress)
        self.current_thread.finished.connect(self.on_pdf_finished)
        self.current_thread.start()

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
        self.current_thread = None

    # -------------------------- 其他功能方法 --------------------------
    def run_outlook(self):
        if not self.excel_path:
            QMessageBox.warning(self, "错误", "未找到Excel文件，请检查tool文件夹")
            return
        self._prepare_task()
        self.update_log("开始生成Outlook邮件...")
        self.current_thread = OutlookEmailThread(self.excel_path)
        self.current_thread.progress.connect(self.update_log)
        self.current_thread.finished.connect(self.on_outlook_finished)
        self.current_thread.start()

    def on_outlook_finished(self, success):
        self._reset_task_state()
        if success:
            self.update_log("Outlook邮件生成完成！")
            self.statusBar().showMessage("Outlook邮件生成完成")
        else:
            self.update_log("Outlook邮件生成失败！")
            self.statusBar().showMessage("Outlook邮件生成失败")
        self.current_thread = None

    def run_folder_creation(self):
        self._prepare_task()
        self.update_log("开始执行文件夹创建+文件检索流程...")
        self.current_thread = FolderThread()
        self.current_thread.progress.connect(self.update_log)
        self.current_thread.finished.connect(self.on_folder_finished)
        self.current_thread.start()

    def on_folder_finished(self, success):
        self._reset_task_state()
        if success:
            self.update_log("文件夹创建+文件检索流程完成！")
            self.statusBar().showMessage("文件夹流程完成")
        else:
            self.update_log("文件夹创建+文件检索流程失败！")
            self.statusBar().showMessage("文件夹流程失败")
        self.current_thread = None

    def run_memo(self):
        if not self.excel_path:
            QMessageBox.warning(self, "错误", "未找到Excel文件，请检查tool文件夹")
            return

        template_path = os.path.join(os.path.expanduser("~"), "Desktop", "tool", "MemoTemplate.docx")
        if not os.path.exists(template_path):
            QMessageBox.warning(
                self, "模板缺失",
                f"未找到备忘录模板：{template_path}\n请将MemoTemplate.docx放入tool文件夹后重试"
            )
            return

        self._prepare_task()
        self.update_log("开始生成备忘录...")
        self.current_thread = MemoThread(excel_path=self.excel_path)
        self.current_thread.progress.connect(self.update_log)
        self.current_thread.finished.connect(self.on_memo_finished)
        self.current_thread.start()

    def on_memo_finished(self, success, msg):
        self._reset_task_state()
        self.update_log(f"\n{msg}")
        self.statusBar().showMessage(msg)
        if success:
            QMessageBox.information(self, "生成成功", msg)
        self.current_thread = None

    def run_file_search(self):
        """运行文件搜索功能 - 弹出独立窗口"""
        # 创建独立的搜索窗口
        search_window = FileSearchDialog(self)
        search_window.exec_()

    def run_file_converter(self):
        """运行文件转换器功能 - 使用独立的UI界面"""
        self.update_log("🚀 启动文件转换器...")
        
        # 创建文件转换器UI窗口
        self.file_converter_ui = FileConverterUI()
        
        # 设置为模态对话框
        self.file_converter_ui.setWindowModality(Qt.ApplicationModal)
        
        # 显示窗口
        self.file_converter_ui.show()
        
        self.update_log(" 文件转换器UI已启动")

    def cancel_task(self):
        if not self.current_thread or not self.current_thread.isRunning():
            QMessageBox.information(self, "提示", "当前没有正在执行的任务")
            return

        reply = QMessageBox.question(
            self, "确认取消", "确定要取消当前任务吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if hasattr(self.current_thread, 'cancel'):
                self.current_thread.cancel()
            self.cancel_btn.setEnabled(False)


# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    # 为Windows系统添加任务栏图标支持
    import sys
    if sys.platform == 'win32':
        import ctypes
        # 设置应用程序用户模型ID，确保任务栏图标正确显示
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("automation.tool.correction.v2.0")
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # 设置应用程序图标（影响任务栏图标）
    # 直接从桌面tool文件夹查找图标文件
    desktop_path = os.path.expanduser("~/Desktop")
    tool_folder = os.path.join(desktop_path, "tool")
    icon_path = os.path.join(tool_folder, 'robot-solid-full.svg')
    
    print(f"正在查找图标文件: {icon_path}")
    if os.path.exists(icon_path):
        print(f"图标文件存在，正在加载: {icon_path}")
        app_icon = QIcon(icon_path)
        print(f"图标加载成功，尺寸: {app_icon.availableSizes()}")
        app.setWindowIcon(app_icon)
    else:
        print(f"应用程序图标文件不存在: {icon_path}")
        print("请确保图标文件位于桌面tool文件夹中:")
        print(f"路径: {icon_path}")
        # 如果图标不存在，使用默认图标
        app.setWindowIcon(QIcon())
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

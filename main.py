import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout,
                             QWidget, QLabel, QMessageBox, QTextEdit, QProgressBar, QGroupBox,
                             QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

# 导入模块
from modules.outlook_automation import OutlookEmailThread
from modules.wechat_automation_tool import auto_fill_wechat_report
from modules.folder_creation import FolderCreator
from modules.memo_generator import generate_memo
from utils.file_utils import find_excel_file
from modules.pdf_extractor import PdfTableExtractor


# -------------------------- 线程类 --------------------------
class WeChatThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self.is_canceled = False
        self.setTerminationEnabled(True)

    def run(self):
        try:
            result_dict = auto_fill_wechat_report(excel_path=self.excel_path, cancel_flag=self)
            if self.is_canceled:
                self.progress.emit("任务已被用户取消")
                self.finished.emit(False)
                return
            is_success = result_dict.get('status') == 'success'
            self.progress.emit(f"企业微信填写结果：{result_dict.get('message', '无详细信息')}")
            self.progress.emit(f"已填充数据：{result_dict.get('filled_data', {})}")
            self.finished.emit(is_success)
        except Exception as e:
            if not self.is_canceled:
                self.progress.emit(f"企业微信自动化执行出错：{str(e)}")
                self.finished.emit(False)

    def cancel(self):
        self.is_canceled = True
        self.progress.emit("正在取消任务...")


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
        self.pdf_input_dir = PdfTableExtractor.DEFAULT_INPUT_DIR  # 使用固定输入路径
        self.pdf_output_dir = PdfTableExtractor.DEFAULT_OUTPUT_DIR
        self.initUI()
        self.find_and_display_excel()

    def initUI(self):
        self.setWindowTitle('自动化工具集')
        self.setGeometry(300, 300, 900, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 标题    ooooooooooo
        title_label = QLabel('自动化小工具')
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50; margin: 15px;")
        layout.addWidget(title_label)

        # Excel文件信息组
        file_group = QGroupBox("Excel文件信息")
        file_layout = QVBoxLayout()
        #刷新按钮
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
        self.refresh_excel_btn.clicked.connect(self.refresh_excel_data)  # 绑定刷新方法
        file_layout.addWidget(self.refresh_excel_btn)

        self.excel_label = QLabel('正在查找Excel文件...')
        self.excel_label.setWordWrap(True)
        file_layout.addWidget(self.excel_label)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # PDF路径选择组
        pdf_path_group = QGroupBox("PDF提取路径设置")
        pdf_path_layout = QHBoxLayout()

        # PDF输入路径（固定路径，显示不可修改）
        self.pdf_input_btn = QPushButton('查看PDF输入文件夹')
        self.pdf_input_btn.setFont(QFont("Arial", 9))
        # 优化：浅色调按钮，添加倒角
        self.pdf_input_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                padding: 8px; 
                margin: 5px; 
                border-radius: 6px;  /* 倒角效果 */
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
        """)
        self.pdf_input_btn.clicked.connect(self.show_pdf_input_dir)
        self.pdf_input_label = QLabel(self.pdf_input_dir)
        self.pdf_input_label.setWordWrap(True)
        self.pdf_input_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")

        # PDF输出路径选择
        self.pdf_output_btn = QPushButton('选择TXT输出文件夹')
        self.pdf_output_btn.setFont(QFont("Arial", 9))
        # 优化：浅色调按钮，添加倒角
        self.pdf_output_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #bbdefb; 
                padding: 8px; 
                margin: 5px; 
                border-radius: 6px;  /* 倒角效果 */
            }
            QPushButton:hover { 
                background-color: #bbdefb; 
            }
        """)
        self.pdf_output_btn.clicked.connect(self.select_pdf_output_dir)
        self.pdf_output_label = QLabel(self.pdf_output_dir)
        self.pdf_output_label.setWordWrap(True)
        self.pdf_output_label.setStyleSheet("color: #7f8c8d; font-size: 15px;")

        # 布局路径选择组件
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

        # 第一行：Outlook + 企业微信 + 备忘录 + PDF提取
        top_btn_layout = QHBoxLayout()
        # 1. Outlook按钮
        self.outlook_btn = QPushButton('生成Outlook邮件')
        self.outlook_btn.setFont(QFont("Arial", 9))
        # 优化：浅蓝色系，倒角设计
        self.outlook_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e3f2fd; 
                color: #1565c0; 
                border: 1px solid #90caf9; 
                padding: 10px; 
                margin: 3px; 
                border-radius: 8px;  /* 倒角效果 */
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

        # 2. 企业微信按钮
        self.wechat_btn = QPushButton('测试-自动填写工具发运')
        self.wechat_btn.setFont(QFont("Arial", 9))
        # 优化：浅青色系，倒角设计，与其他按钮协调
        self.wechat_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e0f7fa; 
                color: #00695c; 
                border: 1px solid #b2ebf2; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;  /* 倒角效果 */
            }
            QPushButton:hover { 
                background-color: #b2ebf2; 
            }
            QPushButton:disabled { 
                background-color: #f5f5f5; 
                color: #bdbdbd;
                border: 1px solid #e0e0e0;
            }
        """)
        self.wechat_btn.clicked.connect(self.run_wechat)
        top_btn_layout.addWidget(self.wechat_btn)

        # 3. 备忘录按钮
        self.memo_btn = QPushButton('生成MEMO')
        self.memo_btn.setFont(QFont("Arial", 9))
        # 优化：浅蓝紫色系，倒角设计
        self.memo_btn.setStyleSheet("""
            QPushButton { 
                background-color: #f3e5f5; 
                color: #6a1b9a; 
                border: 1px solid #ce93d8; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;  /* 倒角效果 */
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

        # 4. PDF提取按钮
        self.pdf_btn = QPushButton('收集云盘步距规数据')
        self.pdf_btn.setFont(QFont("Arial", 9))
        # 优化：浅橙色系，倒角设计，保持协调
        self.pdf_btn.setStyleSheet("""
            QPushButton { 
                background-color: #fff3e0; 
                color: #e65100; 
                border: 1px solid #ffe0b2; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;  /* 倒角效果 */
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

        button_layout.addLayout(top_btn_layout)

        # 第二行：文件夹创建 + 取消按钮
        bottom_btn_layout = QHBoxLayout()
        # 5. 文件夹按钮
        self.folder_btn = QPushButton('创建DATA文件夹&检索tool文件')
        self.folder_btn.setFont(QFont("Arial", 9))
        # 优化：浅黄色系，倒角设计
        self.folder_btn.setStyleSheet("""
            QPushButton { 
                background-color: #fffde7; 
                color: #f57f17; 
                border: 1px solid #fff9c4; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;  /* 倒角效果 */
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

        # 取消按钮（通用所有线程）
        self.cancel_btn = QPushButton('取消任务')
        self.cancel_btn.setFont(QFont("Arial", 9))
        # 优化：浅红色系，倒角设计，但保持柔和
        self.cancel_btn.setStyleSheet("""
            QPushButton { 
                background-color: #ffebee; 
                color: #c62828; 
                border: 1px solid #ffcdd2; 
                padding: 10px; 
                margin: 2px; 
                border-radius: 8px;  /* 倒角效果 */
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

        # 进度条 - 优化样式
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
        # 优化日志框样式
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

    # -------------------------- 辅助方法 --------------------------
    def find_and_display_excel(self):
        self.excel_path, message = find_excel_file()
        self.excel_label.setText(message)
        # 启用/禁用依赖Excel的功能
        excel_exists = self.excel_path is not None
        self.outlook_btn.setEnabled(excel_exists)
        self.wechat_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)
        self.folder_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)  # PDF提取不依赖Excel

    def refresh_excel_data(self):
        """重新读取Excel文件，刷新数据"""
        self.update_log("正在刷新Excel数据...")
        # 调用原有的查找Excel方法，重新获取数据
        self.excel_path, message = find_excel_file()
        self.excel_label.setText(message)

        # 重新启用/禁用依赖Excel的功能按钮
        excel_exists = self.excel_path is not None
        self.outlook_btn.setEnabled(excel_exists)
        self.wechat_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)

        if excel_exists:
            self.update_log("✅ Excel数据已刷新（修改内容已生效）")
        else:
            self.update_log("⚠️ 未找到Excel文件，刷新失败")
    def _prepare_task(self):
        """准备任务：禁用按钮、启用取消按钮、显示进度条、清空日志"""
        self.outlook_btn.setEnabled(False)
        self.wechat_btn.setEnabled(False)
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
        self.wechat_btn.setEnabled(excel_exists)
        self.memo_btn.setEnabled(excel_exists)
        self.folder_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setVisible(False)

    def update_log(self, message):
        """更新日志显示"""
        self.log_text.append(message)
        self.statusBar().showMessage(message)
        # 自动滚动到底部
        self.log_text.moveCursor(self.log_text.textCursor().End)

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    # -------------------------- PDF相关方法 --------------------------
    def show_pdf_input_dir(self):
        """显示PDF输入文件夹（固定路径，不可修改）"""
        if os.path.exists(self.pdf_input_dir):
            # 打开文件夹
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
        # 检查输入路径是否存在
        if not os.path.exists(self.pdf_input_dir):
            QMessageBox.warning(self, "路径错误", f"PDF输入文件夹不存在：\n{self.pdf_input_dir}")
            return

        self._prepare_task()
        self.update_log("开始执行PDF表格提取任务...")
        self.update_log(f"PDF输入路径：{self.pdf_input_dir}")
        self.update_log(f"TXT输出路径：{self.pdf_output_dir}")

        # 创建并启动PDF提取线程
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
            # 询问是否打开输出文件夹
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

    def run_wechat(self):
        if not self.excel_path:
            QMessageBox.warning(self, "错误", "未找到Excel文件，请检查tool文件夹")
            return
        self._prepare_task()
        self.update_log("开始自动填写企业微信...")
        self.current_thread = WeChatThread(self.excel_path)
        self.current_thread.progress.connect(self.update_log)
        self.current_thread.finished.connect(self.on_wechat_finished)
        self.current_thread.start()

    def on_wechat_finished(self, success):
        self._reset_task_state()
        if success:
            self.update_log("企业微信自动填写完成！")
            self.statusBar().showMessage("企业微信自动填写完成")
        else:
            self.update_log("企业微信自动填写失败！")
            self.statusBar().showMessage("企业微信自动填写失败")
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
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格，跨平台一致性更好
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

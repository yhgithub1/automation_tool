#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文件转换器用户界面
提供友好的图形界面来使用文件转换功能
"""

import os
import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QTextEdit, 
                            QProgressBar, QGroupBox, QRadioButton, QButtonGroup,
                            QMessageBox, QListWidget, QListWidgetItem, QComboBox,
                            QCheckBox, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPixmap
import qtawesome as qta
from pathlib import Path

# 导入转换器模块
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from modules.file_converter import FileConverter


class ConversionThread(QThread):
    """转换线程"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, converter, input_files, output_dir=None, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.input_files = input_files
        self.output_dir = output_dir
        
    def run(self):
        """执行转换任务"""
        success_count, failed_count, results = self.converter.batch_convert(
            self.input_files, self.output_dir
        )
        self.finished_signal.emit(success_count > 0, f"成功: {success_count}, 失败: {failed_count}")


class FileConverterUI(QWidget):
    """文件转换器主界面"""
    
    def __init__(self):
        super().__init__()
        self.converter = FileConverter()
        self.conversion_thread = None
        self.init_ui()
        self.setup_connections()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("文件转换器 - Excel/Word/图片转PDF")
        self.setWindowIcon(qta.icon('fa5s.file-pdf', color='red'))
        self.setGeometry(100, 100, 800, 600)
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 标题区域
        title_layout = QHBoxLayout()
        title_label = QLabel("文件转换器")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(title_label)
        main_layout.addLayout(title_layout)
        
        # 文件选择区域
        file_group = QGroupBox(" 文件选择")
        file_layout = QVBoxLayout()
        
        # 单文件转换
        single_layout = QHBoxLayout()
        self.single_file_btn = QPushButton(" 选择单个文件")
        self.single_file_btn.setIcon(qta.icon('fa5s.file'))
        self.single_file_label = QLabel("未选择文件")
        self.single_file_label.setStyleSheet("color: gray;")
        single_layout.addWidget(self.single_file_btn)
        single_layout.addWidget(self.single_file_label)
        single_layout.addStretch()
        file_layout.addLayout(single_layout)
        
        # 批量文件转换
        batch_layout = QHBoxLayout()
        self.batch_files_btn = QPushButton(" 选择多个文件")
        self.batch_files_btn.setIcon(qta.icon('fa5s.folder-open'))
        self.batch_dir_btn = QPushButton("选择文件夹")
        self.batch_dir_btn.setIcon(qta.icon('fa5s.folder'))
        self.batch_files_label = QLabel("未选择文件")
        self.batch_files_label.setStyleSheet("color: gray;")
        batch_layout.addWidget(self.batch_files_btn)
        batch_layout.addWidget(self.batch_dir_btn)
        batch_layout.addWidget(self.batch_files_label)
        batch_layout.addStretch()
        file_layout.addLayout(batch_layout)
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)
        self.file_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        file_layout.addWidget(QLabel("选择的文件列表:"))
        file_layout.addWidget(self.file_list)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # 输出设置区域
        output_group = QGroupBox(" 输出设置")
        output_layout = QVBoxLayout()
        
        # 输出目录设置
        output_dir_layout = QHBoxLayout()
        self.output_dir_edit = QLabel(os.path.join(os.path.expanduser("~"), "Desktop", "converted_pdfs"))
        self.output_dir_btn = QPushButton(" 选择输出目录")
        self.output_dir_btn.setIcon(qta.icon('fa5s.folder'))
        output_dir_layout.addWidget(QLabel("输出目录:"))
        output_dir_layout.addWidget(self.output_dir_edit)
        output_dir_layout.addWidget(self.output_dir_btn)
        output_layout.addLayout(output_dir_layout)
        
        # 输出格式选项
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("输出格式:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF (推荐)", "保留原格式"])
        self.format_combo.setCurrentIndex(0)
        format_layout.addWidget(self.format_combo)
        format_layout.addStretch()
        output_layout.addLayout(format_layout)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # 转换控制区域
        control_group = QGroupBox(" 转换控制")
        control_layout = QHBoxLayout()
        
        # 转换按钮
        btn_layout = QVBoxLayout()
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setIcon(qta.icon('fa5s.play', color='green'))
        self.convert_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")
        self.convert_btn.setFont(QFont("Arial", 10, QFont.Bold))
        
        self.cancel_btn = QPushButton(" 取消转换")
        self.cancel_btn.setIcon(qta.icon('fa5s.stop', color='red'))
        self.cancel_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; padding: 10px; }")
        self.cancel_btn.setEnabled(False)
        
        self.clear_btn = QPushButton("清除列表")
        self.clear_btn.setIcon(qta.icon('fa5s.trash'))
        self.clear_btn.setStyleSheet("QPushButton { background-color: #9E9E9E; color: white; padding: 10px; }")
        
        btn_layout.addWidget(self.convert_btn)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        
        # 进度条和状态
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #ddd;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
            }
        """)
        
        progress_layout.addWidget(QLabel("转换进度:"))
        progress_layout.addWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("color: blue; font-weight: bold;")
        progress_layout.addWidget(self.status_label)
        
        control_layout.addLayout(btn_layout)
        control_layout.addLayout(progress_layout)
        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)
        
        # 日志显示区域
        log_group = QGroupBox(" 转换日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 10pt;
            }
        """)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        # 设置主布局
        self.setLayout(main_layout)
        
        # 添加一些样式
        self.setStyleSheet("""
            QWidget {
                font-family: 'Microsoft YaHei', Arial, sans-serif;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 6px;
                margin-top: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                subline-control: none;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QPushButton {
                border-radius: 4px;
                border: 1px solid #ddd;
                padding: 8px 16px;
                margin: 2px;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
            }
            QLabel {
                padding: 2px;
            }
        """)
        
    def setup_connections(self):
        """设置信号连接"""
        # 文件选择按钮
        self.single_file_btn.clicked.connect(self.select_single_file)
        self.batch_files_btn.clicked.connect(self.select_multiple_files)
        self.batch_dir_btn.clicked.connect(self.select_directory)
        self.output_dir_btn.clicked.connect(self.select_output_directory)
        
        # 控制按钮
        self.convert_btn.clicked.connect(self.start_conversion)
        self.cancel_btn.clicked.connect(self.cancel_conversion)
        self.clear_btn.clicked.connect(self.clear_file_list)
        
        # 转换器信号连接
        self.converter.log_signal.connect(self.update_log)
        self.converter.progress_signal.connect(self.update_progress)
        self.converter.finished_signal.connect(self.on_conversion_finished)
    
    def select_single_file(self):
        """选择单个文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", 
            "支持的文件类型 (*.xlsx *.xls *.docx *.doc *.jpg *.jpeg *.png *.bmp *.gif *.tiff);;Excel文件 (*.xlsx *.xls);;Word文件 (*.docx *.doc);;图片文件 (*.jpg *.jpeg *.png *.bmp *.gif *.tiff)"
        )
        
        if file_path:
            self.clear_file_list()
            self.add_file_to_list(file_path)
            self.single_file_label.setText(f"✓ {os.path.basename(file_path)}")
            self.status_label.setText(f"已选择单个文件: {os.path.basename(file_path)}")
    
    def select_multiple_files(self):
        """选择多个文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择多个文件", "", 
            "支持的文件类型 (*.xlsx *.xls *.docx *.doc *.jpg *.jpeg *.png *.bmp *.gif *.tiff)"
        )
        
        if file_paths:
            self.clear_file_list()
            for file_path in file_paths:
                self.add_file_to_list(file_path)
            self.batch_files_label.setText(f"✓ 已选择 {len(file_paths)} 个文件")
            self.status_label.setText(f"已选择 {len(file_paths)} 个文件")
    
    def select_directory(self):
        """选择文件夹"""
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹")
        
        if directory:
            # 获取文件夹中的所有支持的文件
            supported_extensions = {'.xlsx', '.xls', '.docx', '.doc', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}
            file_paths = []
            
            for file in os.listdir(directory):
                file_path = os.path.join(directory, file)
                if os.path.isfile(file_path):
                    ext = Path(file_path).suffix.lower()
                    if ext in supported_extensions:
                        file_paths.append(file_path)
            
            if file_paths:
                self.clear_file_list()
                for file_path in file_paths:
                    self.add_file_to_list(file_path)
                self.batch_files_label.setText(f"✓ 已选择 {len(file_paths)} 个文件")
                self.status_label.setText(f"已选择 {len(file_paths)} 个文件")
            else:
                QMessageBox.information(self, "提示", "该文件夹中没有找到支持的文件格式")
    
    def add_file_to_list(self, file_path):
        """添加文件到列表"""
        item = QListWidgetItem()
        item.setText(os.path.basename(file_path))
        item.setToolTip(file_path)
        
        # 根据文件类型设置图标
        ext = Path(file_path).suffix.lower()
        if ext in ['.xlsx', '.xls']:
            item.setIcon(qta.icon('fa5.file-excel', color='#217346'))
        elif ext in ['.docx', '.doc']:
            item.setIcon(qta.icon('fa5.file-word', color='#2B579A'))
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
            item.setIcon(qta.icon('fa5.file-image', color='#0078D4'))
        else:
            item.setIcon(qta.icon('fa5.file', color='gray'))
        
        self.file_list.addItem(item)
    
    def select_output_directory(self):
        """选择输出目录"""
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if directory:
            self.output_dir_edit.setText(directory)
    
    def start_conversion(self):
        """开始转换"""
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "请先选择要转换的文件！")
            return
        
        # 禁用按钮，启用取消按钮
        self.set_controls_enabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText(" 正在转换中...")
        self.progress_bar.setValue(0)
        
        # 清空日志
        self.log_text.clear()
        
        # 收集文件路径
        file_paths = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            # 从工具提示中获取完整路径
            file_paths.append(item.toolTip())
        
        # 创建转换线程
        self.conversion_thread = ConversionThread(
            self.converter, 
            file_paths, 
            self.output_dir_edit.text()
        )
        self.conversion_thread.log_signal.connect(self.update_log)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.finished_signal.connect(self.on_batch_conversion_finished)
        self.conversion_thread.start()
        
        self.update_log(" 开始批量转换任务...")
    
    def cancel_conversion(self):
        """取消转换"""
        if self.conversion_thread and self.conversion_thread.isRunning():
            self.converter.cancel_conversion()
            self.status_label.setText(" 正在取消转换...")
            self.update_log(" 用户取消了转换任务")
    
    def clear_file_list(self):
        """清除文件列表"""
        self.file_list.clear()
        self.single_file_label.setText("未选择文件")
        self.batch_files_label.setText("未选择文件")
        self.status_label.setText("准备就绪")
        self.log_text.clear()
        self.progress_bar.setValue(0)
    
    def set_controls_enabled(self, enabled):
        """设置控件启用状态"""
        self.single_file_btn.setEnabled(enabled)
        self.batch_files_btn.setEnabled(enabled)
        self.batch_dir_btn.setEnabled(enabled)
        self.output_dir_btn.setEnabled(enabled)
        self.format_combo.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)
        self.convert_btn.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)
    
    def update_log(self, message):
        """更新日志"""
        timestamp = self.get_current_time()
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.append(formatted_message)
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def on_conversion_finished(self, success, output_path):
        """单个文件转换完成"""
        if success:
            self.status_label.setText(f" 转换完成: {os.path.basename(output_path)}")
            self.set_controls_enabled(True)
            self.cancel_btn.setEnabled(False)
        else:
            self.status_label.setText(" 转换失败")
            self.set_controls_enabled(True)
            self.cancel_btn.setEnabled(False)
    
    def on_batch_conversion_finished(self, success, message):
        """批量转换完成"""
        self.status_label.setText(" 批量转换完成")
        self.set_controls_enabled(True)
        self.cancel_btn.setEnabled(False)
        self.update_log(f"🎉 {message}")
        
        # 如果有成功的转换，询问是否打开输出文件夹
        if success:
            reply = QMessageBox.question(
                self, "完成", 
                "转换完成！是否打开输出文件夹？", 
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                os.startfile(self.output_dir_edit.text())
    
    def get_current_time(self):
        """获取当前时间字符串"""
        from datetime import datetime
        return datetime.now().strftime("%H:%M:%S")
    
    def closeEvent(self, event):
        """关闭事件处理"""
        if self.conversion_thread and self.conversion_thread.isRunning():
            self.converter.cancel_conversion()
            event.ignore()  # 忽略关闭事件，等待转换完成
        else:
            event.accept()


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置应用程序样式
    app.setStyle('Fusion')
    
    # 创建并显示主窗口
    window = FileConverterUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

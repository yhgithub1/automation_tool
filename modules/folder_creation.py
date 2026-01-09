# modules/folder_creation.py
import os
import tkinter as tk
from tkinter import simpledialog, messagebox
import shutil
import re
from PyQt5.QtCore import QObject, pyqtSignal  # 用于传递日志（适配主窗口日志框）


class FolderCreator(QObject):
    """文件夹创建器（支持日志信号传递，适配主窗口日志显示）"""
    log_signal = pyqtSignal(str)  # 传递日志到主窗口的信号
    finished = pyqtSignal(bool)   # 任务完成信号（成功/失败）

    def create_folders(self):
        """核心：创建文件夹+文件检索复制逻辑"""
        try:
            # 1. 隐藏Tkinter主窗口（避免额外弹窗干扰）
            root = tk.Tk()
            root.withdraw()
            self.log_signal.emit("开始执行文件夹创建流程...")

            # 2. 让用户输入主文件夹名称
            folder_name = simpledialog.askstring("输入", "请输入文件夹名称:", parent=root)
            if not folder_name:
                self.log_signal.emit("用户未输入文件夹名称，操作取消")
                messagebox.showinfo("取消", "未输入文件夹名称，操作取消。")
                self.finished.emit(False)
                root.destroy()
                return

            # 3. 构建桌面路径和主文件夹路径
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            main_folder_path = os.path.join(desktop_path, folder_name)
            self.log_signal.emit(f"主文件夹目标路径：{main_folder_path}")

            # 4. 创建主文件夹（处理已存在情况）
            if os.path.exists(main_folder_path):
                self.log_signal.emit(f"警告：主文件夹已存在（{main_folder_path}）")
                messagebox.showwarning("警告", f"文件夹已存在: {main_folder_path}")
                self.finished.emit(False)
                root.destroy()
                return
            os.makedirs(main_folder_path)
            self.log_signal.emit(f"✅ 成功创建主文件夹：{main_folder_path}")

            # 5. 创建DATA子文件夹
            data_folder_path = os.path.join(main_folder_path, 'DATA')
            if not os.path.exists(data_folder_path):
                os.makedirs(data_folder_path)
                self.log_signal.emit(f"✅ 成功创建DATA文件夹：{data_folder_path}")
            else:
                self.log_signal.emit(f"警告：DATA文件夹已存在（{data_folder_path}）")
                messagebox.showwarning("警告", f"DATA文件夹已存在: {data_folder_path}")
                self.finished.emit(False)
                root.destroy()
                return

            # 6. 创建X/Y/Z/FR/FL/RL/RR子文件夹（批量创建）
            sub_folders = ["X", "Y", "Z", "FR", "FL", "RL", "RR"]
            for sub_folder in sub_folders:
                sub_folder_path = os.path.join(data_folder_path, sub_folder)
                os.makedirs(sub_folder_path)
                self.log_signal.emit(f"✅ 成功创建子文件夹：{sub_folder_path}")

            # 7. 检查tool文件夹是否存在（用于后续文件检索）
            tool_folder_path = os.path.join(desktop_path, 'tool')
            if not os.path.exists(tool_folder_path):
                self.log_signal.emit(f"❌ 错误：tool文件夹不存在（{tool_folder_path}）")
                messagebox.showwarning("警告", f"tool文件夹不存在: {tool_folder_path}")
                self.finished.emit(False)
                root.destroy()
                return
            self.log_signal.emit(f"✅ 找到tool文件夹：{tool_folder_path}（开始文件检索）")

            # 8. 循环检索并复制TXT文件（支持多次检索，点击Cancel退出）
            while True:
                customer_input = simpledialog.askstring(
                    "输入", "请输入检索值（点击'Cancel'退出检索）:", parent=root
                )
                if customer_input is None:  # 用户点击Cancel
                    self.log_signal.emit("用户退出文件检索流程")
                    break
                customer_input = customer_input.strip().lower()
                if customer_input == '退出':  # 支持输入“退出”关键词
                    self.log_signal.emit("用户输入'退出'，结束文件检索")
                    break
                if not customer_input:  # 空输入跳过
                    messagebox.showinfo("提示", "检索值不能为空，请重新输入")
                    continue

                # 检索匹配的TXT文件 - 使用通用数字_数字格式
                pattern_general = re.compile(r'(\d+)_(\d+)')

                files_with_date = []
                files_without_date = []

                for file in os.listdir(tool_folder_path):
                    if customer_input.lower() in file.lower() and file.endswith('.txt'):
                        file_path = os.path.join(tool_folder_path, file)
                        modify_time = os.path.getmtime(file_path)

                        # 尝试通用数字_数字格式
                        general_match = pattern_general.search(file)
                        if general_match:
                            # 文件名包含日期格式 (数字_数字)
                            date_str = general_match.group(2)
                            files_with_date.append((date_str, modify_time, file_path, file))
                            self.log_signal.emit(f"  {file} -> 分类为有日期格式: {date_str}")
                        else:
                            # 文件名不包含日期格式
                            files_without_date.append((modify_time, file_path, file))
                            self.log_signal.emit(f"  {file} -> 分类为无日期格式")
                
                # 详细日志输出
                self.log_signal.emit(f"找到 {len(files_with_date)} 个有日期格式的文件")
                self.log_signal.emit(f"找到 {len(files_without_date)} 个无日期格式的文件")
                
                # 选择逻辑
                if files_with_date and files_without_date:
                    # 如果既有有日期的文件，又有无日期的文件
                    # 比较所有文件的修改时间，选择修改时间最新的
                    all_files = []
                    # 添加有日期的文件
                    for date_str, modify_time, file_path, file_name in files_with_date:
                        all_files.append((modify_time, file_path, file_name))
                    # 添加无日期的文件
                    for modify_time, file_path, file_name in files_without_date:
                        all_files.append((modify_time, file_path, file_name))
                    
                    self.log_signal.emit(f"混合文件情况：合并 {len(all_files)} 个文件进行比较")
                    
                    # 从所有文件中选择修改时间最新的
                    latest_file_info = max(all_files, key=lambda x: x[0])
                    latest_file_name = latest_file_info[2]
                    latest_file_path = latest_file_info[1]
                    
                    self.log_signal.emit(f"混合文件选择结果：选择 {latest_file_name} (按修改时间)")
                    
                elif files_with_date:
                    # 只有有日期的文件，优先按日期选择
                    latest_file_info = max(files_with_date, key=lambda x: x[0])
                    latest_file_name = latest_file_info[3]
                    latest_file_path = latest_file_info[2]
                    
                    self.log_signal.emit(f"只有有日期文件：选择 {latest_file_name} (按日期)")
                    
                elif files_without_date:
                    # 只有无日期的文件，按修改时间选择
                    latest_file_info = max(files_without_date, key=lambda x: x[0])
                    latest_file_name = latest_file_info[2]
                    latest_file_path = latest_file_info[1]
                    
                    self.log_signal.emit(f"只有无日期文件：选择 {latest_file_name} (按修改时间)")
                    
                else:
                    self.log_signal.emit(f"未找到匹配的文件（检索值: {customer_input}）")
                    messagebox.showinfo("提示", "未找到匹配的文件")
                    continue  # 继续下一次检索
                
                # 复制最新文件到总文件夹
                shutil.copy(latest_file_path, main_folder_path)
                self.log_signal.emit(f"✅ 复制最新文件到总文件夹: {latest_file_name} 成功")
                messagebox.showinfo("成功", f"复制最新文件到总文件夹: {latest_file_name}")

            # 9. 流程结束
            self.log_signal.emit("文件夹创建+文件检索流程全部完成")
            messagebox.showinfo("完成", "所有文件夹创建完成！")
            self.finished.emit(True)
            root.destroy()

        except Exception as e:
            # 捕获异常并反馈
            error_msg = f"文件夹创建过程出错：{str(e)}"
            self.log_signal.emit(error_msg)
            messagebox.showerror("错误", error_msg)
            self.finished.emit(False)
            root.destroy()

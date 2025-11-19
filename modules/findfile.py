import os
import re
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, scrolledtext

def setup_gui(root):
    """设置GUI界面"""
    root.title("文件内容搜索工具 (优化版)")
    root.geometry("800x600")

    # 顶部框架：控制按钮
    control_frame = tk.Frame(root, pady=10)
    control_frame.pack(fill=tk.X, padx=10)

    tk.Button(control_frame, text="选择目录...", command=start_search).pack(side=tk.LEFT, padx=5)
    tk.Button(control_frame, text="清空结果", command=clear_results).pack(side=tk.LEFT, padx=5)

    # 中部框架：结果显示区
    result_frame = tk.Frame(root)
    result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    global result_text
    result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, state='disabled')
    result_text.pack(fill=tk.BOTH, expand=True)


def add_result(message):
    """向结果区添加文本"""
    result_text.config(state='normal')
    result_text.insert(tk.END, message + "\n")
    result_text.see(tk.END)  # 自动滚动到底部
    result_text.config(state='disabled')
    root.update_idletasks()

def clear_results():
    """清空结果区"""
    result_text.config(state='normal')
    result_text.delete('1.0', tk.END)
    result_text.config(state='disabled')

def find_files_with_progress(root_dir, search_content, file_names=None, case_sensitive=False):
    """
    搜索指定目录下的文件
    """
    found_files = []

    # 1. 如果没有指定文件名，默认搜索常见配置文件
    if file_names is None:
        file_names = ['config.kmg']

    # 2. 开始搜索
    add_result(f"\n开始搜索内容: '{search_content}' (文件名: {', '.join(file_names)})")

    for root, _, files in os.walk(root_dir):
        for file in files:
            file_lower = file.lower()
            # 检查文件名是否匹配（不区分大小写）
            if any(file_lower == name.lower() for name in file_names):
                file_path = os.path.join(root, file)
                
                try:
                    # 尝试多种编码读取文件
                    encodings = ['utf-8', 'gb18030', 'gbk', 'latin-1']
                    content_found = False

                    for encoding in encodings:
                        try:
                            with open(file_path, 'r', encoding=encoding) as f:
                                for line_num, line in enumerate(f, 1):
                                    if search_content.lower() in line.lower():
                                        # 找到匹配内容
                                        add_result(f"\n 找到匹配文件: {file_path}")
                                        add_result(f"   行号: {line_num}, 匹配行: {line.strip()}")
                                        found_files.append((file_path, line_num, line.strip()))
                                        content_found = True
                                        break  # 找到后跳出编码循环
                            if content_found:
                                break
                        except UnicodeDecodeError:
                            continue  # 尝试下一种编码
                        except Exception as e:
                            add_result(f"\n 读取文件 {file_path} 时出错: {e}")
                            break

                except Exception as e:
                    add_result(f"\n访问文件 {file_path} 时发生错误: {e}")

    # 3. 完成搜索
    add_result("\n" + "="*50)
    if found_files:
        add_result(f"搜索完成! 共找到 {len(found_files)} 个匹配项。")
    else:
        add_result(f"搜索完成! 未找到包含 '{search_content}' 的文件。")

def start_search():
    """开始搜索的触发函数"""
    # 1. 选择目录
    target_dir = filedialog.askdirectory(title="请选择要搜索的根目录")
    if not target_dir:
        return

    # 2. 获取搜索内容
    search_content = simpledialog.askstring("输入搜索内容", "请输入要搜索的内容:", initialvalue="Install_typ     = 101206")
    if not search_content:
        messagebox.showwarning("提示", "请输入搜索内容。")
        return

    # 3. 获取文件名列表
    file_names_input = simpledialog.askstring("输入文件名", "请输入要搜索的文件全名(用逗号分隔, 留空则搜索常见配置文件):", initialvalue="config.kmg")
    file_names = [name.strip() for name in file_names_input.split(',')] if file_names_input else None

    # 4. 清空之前的结果并开始搜索
    clear_results()
    find_files_with_progress(target_dir, search_content, file_names)

def main():
    """主函数"""
    global root
    root = tk.Tk()
    setup_gui(root)
    root.mainloop()

if __name__ == "__main__":
    main()

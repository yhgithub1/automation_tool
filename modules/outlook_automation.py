import win32com.client as win32
import pandas as pd
import re
import os
import time
import glob
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
import sys
import os

# Get the directory of the current script
current_dir = os.path.dirname(os.path.abspath(__file__))
# Get the project root (parent of modules directory)
project_root = os.path.dirname(current_dir)
# Add project root to Python path if not already there
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from utils.file_utils import find_excel_file


class OutlookEmailThread(QThread):
    """Outlook邮件生成线程（支持COM初始化，处理Excel并生成邮件）"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        # 初始化 COM 环境（必须在操作 Outlook 前调用）
        pythoncom.CoInitialize()
        try:
            result = self._generate_emails_from_excel()
            self.finished.emit(result)
        except Exception as e:
            # 处理 COM 注册问题
            if "CLSIDToClassMap" in str(e):
                self.progress.emit(f"Outlook COM 组件需要重新注册，尝试修复...")
                try:
                    # 尝试重新初始化
                    pythoncom.CoUninitialize()
                    pythoncom.CoInitialize()
                    result = self._generate_emails_from_excel()
                    self.finished.emit(result)
                    return
                except Exception as retry_e:
                    self.progress.emit(f"重新尝试失败: {str(retry_e)}")
            
            self.progress.emit(f"全局错误: {str(e)}")
            self.finished.emit(False)
        finally:
            # 释放 COM 资源，避免内存泄漏
            pythoncom.CoUninitialize()

    def _generate_emails_from_excel(self):
        """核心逻辑：从Excel读取数据，生成Outlook邮件"""
        try:
            # -------- 1. 读取Excel文件 --------
            self.progress.emit("正在读取Excel文件...")
            df = pd.read_excel(self.excel_path, header=None)
            self.progress.emit(f"成功读取Excel文件，共{len(df)}行数据")

            # -------- 2. 启动Outlook --------
            self.progress.emit("启动Outlook应用程序...")
            outlook = self._get_outlook_application()
            if not outlook:
                self.progress.emit("无法启动Outlook应用程序，请确保Outlook已安装且正常运行")
                return False
            self.progress.emit("Outlook应用程序已启动")

            # -------- 3. 捕获Outlook签名 --------
            self.progress.emit("正在获取Outlook签名...")
            signature = self._capture_outlook_signature(outlook)
            if not signature or len(signature) <= 50:
                self.progress.emit("警告: 未捕获到有效签名，邮件将不含签名")

            # -------- 4. 循环生成邮件 --------
            total_rows = len(df)
            for index, row in df.iterrows():
                self.progress.emit(f"正在处理第{index + 1}/{total_rows}行数据...")
                try:
                    # 提取公司名称（F列，索引2）
                    company_full = str(row.iloc[2])
                    company_name = company_full.split('/')[-1].strip() if '/' in company_full else company_full

                    # 提取设备信息（型号：H列索引4；SN：B列索引1）
                    model = str(row.iloc[4])
                    sn = str(row.iloc[1])

                    # 提取地址和联系人（N列，索引13）
                    contact_info = re.sub(r'\s+', ' ', str(row.iloc[13])).strip()

                    # 构建邮件主题和正文
                    subject = f"{company_name} {model} SN:{sn}包装箱回收"
                    body = f"""Dear Mr. Zhang:

如下客户请回收包装箱，麻烦尽快安排：
客户：{company_name}
CMM: {model} SN: {sn}
地址及联系人：{contact_info}
谢谢！

"""

                    # 创建HTML格式邮件
                    html_body = f"""
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
    body {{ font-family: 'Calibri', sans-serif; font-size: 11pt; }}
    pre {{ font-family: 'Calibri', sans-serif; font-size: 11pt; }}
</style>
</head>
<body>
<pre>{body}</pre>
<br>
{signature if signature else ''}
</body>
</html>
"""

                    # 创建并显示邮件
                    mail = outlook.CreateItem(0)
                    mail.Subject = subject
                    mail.HTMLBody = html_body
                    mail.To = "Zhang, Zicheng"
                    mail.CC = "Zhu, Zhiming"
                    mail.Display()

                    self.progress.emit(f"已创建邮件 #{index + 1}: {subject}")
                    time.sleep(1)  # 给Outlook留处理时间

                except Exception as e:
                    self.progress.emit(f"处理行{index + 1}时出错: {str(e)}")

            self.progress.emit("邮件创建完成！请检查Outlook窗口")
            return True

        except Exception as e:
            self.progress.emit(f"生成邮件时出错: {str(e)}")
            return False

    def _capture_outlook_signature(self, outlook):
        """多方法捕获Outlook签名"""
        try:
            # 方法1：从默认路径读取签名文件
            appdata = os.getenv('APPDATA')
            signature_path = os.path.join(appdata, 'Microsoft', 'Signatures')
            if os.path.exists(signature_path):
                html_files = glob.glob(os.path.join(signature_path, '*.htm')) or glob.glob(os.path.join(signature_path, '*.html'))
                if html_files:
                    latest_file = max(html_files, key=os.path.getmtime)
                    with open(latest_file, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    # 修正签名中的图片路径
                    base_name = os.path.splitext(os.path.basename(latest_file))[0]
                    image_folder = os.path.join(signature_path, base_name)
                    if os.path.exists(image_folder):
                        content = content.replace(f'"{base_name}/', f'"{image_folder}/')
                        content = content.replace(f"'{base_name}/", f"'{image_folder}/")
                    return content if len(content) > 50 else ""

            # 方法2：通过临时邮件获取签名
            temp_mail = outlook.CreateItem(0)
            temp_mail.Display()
            time.sleep(2)
            signature = ""
            for _ in range(5):
                signature = temp_mail.HTMLBody
                if len(signature) > 100 and ("<p>" in signature or "<div>" in signature):
                    break
                time.sleep(1)
            temp_mail.Close(0)  # 不保存临时邮件

            # 提取签名（通过<hr>标签分割）
            if "<hr" in signature:
                signature = signature.split("<hr", 1)[-1].split(">", 1)[-1]
            return signature if len(signature) > 50 else ""

        except Exception as e:
            self.progress.emit(f"捕获签名失败: {str(e)}")
            return ""

    def _get_outlook_application(self):
        """获取Outlook应用程序对象，包含多种COM初始化方法"""
        try:
            # 方法1：直接使用Dispatch
            try:
                outlook = win32.Dispatch('Outlook.Application')
                return outlook
            except Exception as e1:
                self.progress.emit(f"直接Dispatch失败: {str(e1)}")
            
            # 方法2：使用gencache确保分发
            try:
                import win32com.client.gencache
                outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
                self.progress.emit("使用gencache方法成功")
                return outlook
            except Exception as e2:
                self.progress.emit(f"gencache方法失败: {str(e2)}")
            
            # 方法3：清除并重建COM缓存
            try:
                self.progress.emit("尝试清除损坏的COM缓存...")
                self._clear_com_cache()
                import win32com.client.gencache
                outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
                self.progress.emit("清除缓存后成功")
                return outlook
            except Exception as e3:
                self.progress.emit(f"清除缓存后仍失败: {str(e3)}")
            
            # 方法4：检查Outlook是否正在运行并启动
            try:
                import psutil
                outlook_running = any(proc.name().lower() == 'outlook.exe' for proc in psutil.process_iter(['name']))
                if not outlook_running:
                    self.progress.emit("检测到Outlook未运行，尝试启动...")
                    import subprocess
                    subprocess.Popen(['outlook.exe'])
                    time.sleep(10)  # 增加等待时间
                    # 重试
                    outlook = win32.Dispatch('Outlook.Application')
                    return outlook
            except ImportError:
                self.progress.emit("无法检查Outlook进程状态（psutil未安装）")
            except Exception as e4:
                self.progress.emit(f"启动Outlook失败: {str(e4)}")
            
            # 方法5：使用CLSID直接访问
            try:
                outlook = win32.Dispatch('{0006F03A-0000-0000-C000-000000000046}')
                self.progress.emit("使用CLSID方法成功")
                return outlook
            except Exception as e5:
                self.progress.emit(f"CLSID方法也失败: {str(e5)}")
            
            # 如果所有方法都失败，提供详细指导
            self.progress.emit("Outlook连接失败，请按以下步骤手动修复：")
            self.progress.emit("1. 关闭所有Outlook窗口")
            self.progress.emit("2. 按 Win+R，输入 'cmd'，运行以下命令：")
            self.progress.emit("   cd %APPDATA%\\Python\\Pythonwin32\\gen_py")
            self.progress.emit("   rmdir /s /q *")
            self.progress.emit("3. 重新启动Outlook应用程序")
            self.progress.emit("4. 重新运行此工具")
            self.progress.emit("或者尝试：控制面板 → 程序和功能 → 修复Microsoft Office")
            return None

        except Exception as e:
            self.progress.emit(f"获取Outlook应用程序时发生未知错误: {str(e)}")
            return None

    def _clear_com_cache(self):
        """清除损坏的COM缓存文件"""
        try:
            import shutil
            # 获取Python COM缓存目录
            gen_py_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Python", "Pythonwin32", "gen_py")
            
            if os.path.exists(gen_py_dir):
                self.progress.emit(f"正在清除COM缓存目录: {gen_py_dir}")
                # 备份当前目录名并删除
                backup_dir = gen_py_dir + "_backup_" + str(int(time.time()))
                if os.path.exists(backup_dir):
                    shutil.rmtree(backup_dir)
                os.rename(gen_py_dir, backup_dir)
                self.progress.emit(f"已备份旧缓存到: {backup_dir}")
            else:
                self.progress.emit("COM缓存目录不存在，无需清除")
                
        except Exception as e:
            self.progress.emit(f"清除COM缓存时出错: {str(e)}")
            # 即使清除失败，继续尝试其他方法


# 测试代码（直接运行该脚本时执行）
if __name__ == "__main__":
    excel_path, msg = find_excel_file()
    if not excel_path:
        print(msg)
        exit(1)
    print(f"找到Excel文件: {excel_path}")
    # 创建并启动线程
    thread = OutlookEmailThread(excel_path)
    thread.progress.connect(print)
    thread.finished.connect(lambda success: print(f"执行完成: {success}"))
    thread.start()
    thread.wait()  # 等待线程结束（仅测试用）

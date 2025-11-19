# modules/pdf_extractor.py
import os
import pdfplumber
from PyQt5.QtCore import QObject, pyqtSignal
import sys  # 用于独立运行时的命令行交互


class PdfTableExtractor(QObject):
    """PDF表格提取器：提取第三页表格的“实测值”列，生成TXT文件"""
    log_signal = pyqtSignal(str)  # 传递日志到主窗口
    progress_signal = pyqtSignal(int)  # 传递进度（0-100）
    finished_signal = pyqtSignal(bool)  # 任务完成信号（成功/失败）

    # -------------------------- 核心修改：固定输入路径 --------------------------
    DEFAULT_INPUT_DIR = r"H:\Shanghai\IMT\Service\Management Tools\量具\标准器校准证书最新\02步距规"
    DEFAULT_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "tool")

    def __init__(self, input_dir=None, output_dir=None):
        super().__init__()
        # 优先使用传入路径，无传入则用默认路径（input_dir固定为DEFAULT_INPUT_DIR）
        self.input_dir = input_dir if input_dir else self.DEFAULT_INPUT_DIR
        self.output_dir = output_dir if output_dir else self.DEFAULT_OUTPUT_DIR
        self.is_canceled = False  # 取消标记

    def set_paths(self, input_dir=None, output_dir=None):
        """设置路径（input_dir默认固定，output_dir支持动态修改）"""
        # input_dir固定为默认路径，不允许外部修改（若需临时修改，可注释此行）
        self.input_dir = self.DEFAULT_INPUT_DIR
        if output_dir:  # 仅允许修改输出路径
            self.output_dir = output_dir
        self.log_signal.emit(f"📌 已设置路径：")
        self.log_signal.emit(f"   输入（固定）：{self.input_dir}")
        self.log_signal.emit(f"   输出：{self.output_dir}")

    def cancel_extract(self):
        """取消当前提取任务"""
        self.is_canceled = True
        self.log_signal.emit("⏹️  正在取消PDF提取任务...")

    def _extract_single_pdf(self, pdf_path):
        """提取单个PDF的“实测值”列数据（核心逻辑不变）"""
        try:
            if self.is_canceled:
                return None, "任务已取消"

            # 打开PDF并检查页数
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) < 3:
                    return None, "页数不足3页（需至少3页，从第3页提取表格）"

                # 提取第三页表格（索引2 = 第3页）
                page = pdf.pages[2]
                tables = page.extract_tables()
                if not tables:
                    return None, "未找到表格（第三页无表格数据）"

                # 处理第一个表格（默认目标表格）
                table = tables[0]
                if len(table) == 0:
                    return None, "表格为空（第三页表格无数据）"

                # 查找所有包含“实测值”的列索引
                header_row = table[0]
                target_col_indices = [
                    i for i, cell in enumerate(header_row)
                    if "实测值" in str(cell)  # 匹配“实测值”相关列
                ]
                if not target_col_indices:
                    return None, "未找到'实测值'列（表头无匹配字段）"

                # 按列提取数据（忽略第一行表头）
                merged_data = []
                for col_idx in target_col_indices:
                    for row_idx, row in enumerate(table):
                        if row_idx > 0 and len(row) > col_idx:  # 跳过表头，确保列存在
                            cell_data = str(row[col_idx]).strip()
                            if cell_data:  # 过滤空值
                                merged_data.append(cell_data)

                if not merged_data:
                    return None, "未提取到有效数据（'实测值'列无内容）"

                # 返回合并后的“实测值”数据（按行拼接）
                return "\n".join(merged_data), "提取成功"

        except Exception as e:
            return None, f"提取失败：{str(e)}"

    def batch_extract(self):
        """批量处理输入文件夹中的所有PDF（核心逻辑不变）"""
        try:
            # 1. 校验路径合法性
            if self.is_canceled:
                self.log_signal.emit("❌ PDF提取任务已取消")
                self.finished_signal.emit(False)
                return

            # 检查固定输入文件夹是否存在
            if not os.path.exists(self.input_dir):
                raise FileNotFoundError(f"PDF输入文件夹不存在（固定路径）：{self.input_dir}")

            # 2. 创建输出文件夹（若不存在）
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
                self.log_signal.emit(f"✅ 已创建TXT输出文件夹：{self.output_dir}")

            # 3. 获取所有PDF文件（过滤非PDF）
            pdf_files = [
                f for f in os.listdir(self.input_dir)
                if f.lower().endswith(".pdf")  # 忽略大小写，支持.PDF/.pdf
            ]
            if not pdf_files:
                self.log_signal.emit("ℹ️  未找到任何PDF文件（输入文件夹中无.pdf后缀文件）")
                self.finished_signal.emit(True)  # 无文件也算“任务完成”
                return

            total_files = len(pdf_files)
            self.log_signal.emit(f"📊 开始批量处理PDF：共{total_files}个文件")

            # 4. 批量处理每个PDF（带进度计算）
            success_count = 0
            for idx, filename in enumerate(pdf_files, 1):
                if self.is_canceled:
                    self.log_signal.emit(f"❌ 任务取消，已处理{idx - 1}/{total_files}个文件")
                    self.finished_signal.emit(False)
                    return

                # 计算当前进度（百分比）
                progress = int((idx / total_files) * 100)
                self.progress_signal.emit(progress)

                # 处理单个PDF
                pdf_path = os.path.join(self.input_dir, filename)
                self.log_signal.emit(f"\n🔄 正在处理（{idx}/{total_files}）：{filename}")

                data, msg = self._extract_single_pdf(pdf_path)
                if data:
                    # 提取成功：生成TXT文件
                    txt_filename = os.path.splitext(filename)[0] + ".txt"
                    txt_path = os.path.join(self.output_dir, txt_filename)
                    with open(txt_path, "w", encoding="utf-8") as f:
                        f.write(data)
                    success_count += 1
                    self.log_signal.emit(f"✅ 处理成功：{txt_filename}（已保存到输出文件夹）")
                else:
                    # 提取失败：记录错误原因
                    self.log_signal.emit(f"❌ 处理失败：{filename} - {msg}")

            # 5. 任务完成：汇总结果
            self.log_signal.emit(f"\n🎉 批量处理完成！")
            self.log_signal.emit(
                f"📈 处理统计：共{total_files}个文件，成功{success_count}个，失败{total_files - success_count}个")
            self.log_signal.emit(f"📁 TXT文件保存路径：{self.output_dir}")
            self.progress_signal.emit(100)  # 进度条拉满
            self.finished_signal.emit(True)

        except Exception as e:
            # 捕获全局异常
            err_msg = f"❌ PDF批量提取出错：{str(e)}"
            self.log_signal.emit(err_msg)
            self.finished_signal.emit(False)


# -------------------------- 新增：独立运行测试逻辑 --------------------------
def run_independent_test():
    """独立测试入口：无需依赖主程序，直接运行模块即可测试"""
    print("=" * 50)
    print("📝 PDF表格提取模块 - 独立测试")
    print("=" * 50)

    # 1. 初始化提取器（自动使用固定input_dir和默认output_dir）
    extractor = PdfTableExtractor()
    print(f"\n📌 固定输入路径：{extractor.input_dir}")
    print(f"📌 默认输出路径：{extractor.output_dir}")

    # 2. 路径预检查（提前提示用户问题）
    if not os.path.exists(extractor.input_dir):
        print(f"\n❌ 错误：固定输入文件夹不存在！")
        print(f"   路径：{extractor.input_dir}")
        print(f"   请检查路径是否正确，或修改代码中的DEFAULT_INPUT_DIR")
        return

    # 3. 询问用户是否修改输出路径（可选）
    print(f"\nℹ️ 当前输出路径：{extractor.output_dir}")
    change_output = input("是否需要修改输出路径？（y/n，默认n）：").strip().lower()
    if change_output == "y":
        new_output = input("请输入新的输出文件夹路径：").strip()
        if new_output:
            extractor.set_paths(output_dir=new_output)  # 仅修改输出路径
        else:
            print("⚠️  输入为空，使用默认输出路径")

    # 4. 绑定日志和进度回调（命令行显示）
    def log_callback(msg):
        print(f"[日志] {msg}")

    def progress_callback(progress):
        print(f"[进度] {progress}%", end="\r")  # 动态刷新进度

    extractor.log_signal.connect(log_callback)
    extractor.progress_signal.connect(progress_callback)

    # 5. 启动提取并等待完成
    print(f"\n🚀 开始PDF提取任务（按Ctrl+C可中断）...")
    try:
        # 手动触发批量提取（独立运行时无需线程）
        extractor.batch_extract()
    except KeyboardInterrupt:
        extractor.cancel_extract()
        print(f"\n\n⏹️  任务已被用户中断")
    except Exception as e:
        print(f"\n\n❌ 测试过程出错：{str(e)}")

    print(f"\n" + "=" * 50)
    print("📝 独立测试结束")
    print("=" * 50)


# -------------------------- 独立运行入口（直接执行模块时触发） --------------------------
if __name__ == "__main__":
    # 检查依赖（确保pdfplumber已安装）
    try:
        import pdfplumber
    except ImportError:
        print("❌ 缺少依赖库：pdfplumber")
        print("   请先安装：pip install pdfplumber")
        sys.exit(1)

    # 启动独立测试
    run_independent_test()
# modules/memo_generator.py
import openpyxl
from docx import Document
from datetime import datetime, timedelta
import os
import sys

# Get the directory of the current script
current_dir = os.path.dirname(os.path.abspath(__file__))
# Get the project root (parent of modules directory)
project_root = os.path.dirname(current_dir)
# Add project root to Python path if not already there
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from utils.file_utils import find_excel_file


def generate_memo(excel_path=None, template_path=None, output_path=None, progress_callback=None):
    """
    生成MEMO：从Excel读取数据，填充Word模板并保存
    :param excel_path: Excel文件路径（默认：tool/1.xlsx）
    :param template_path: Word模板路径（默认：tool/MemoTemplate.docx）
    :param output_path: 生成文件保存路径（默认：tool/Filled_Memo.docx）
    :param progress_callback: 日志回调函数（传递进度到主窗口）
    :return: tuple (success: bool, message: str, output_path: str)
    """
    # 日志发送辅助函数
    def send_log(msg):
        if progress_callback and callable(progress_callback):
            progress_callback(msg)
        print(msg)

    # 1. 初始化默认路径
    try:
        # 基础路径：桌面/tool
        tool_folder = os.path.join(os.path.expanduser("~"), "Desktop", "tool")
        if not os.path.exists(tool_folder):
            os.makedirs(tool_folder)
            send_log(f"✅ 已创建tool文件夹: {tool_folder}")

        # 默认Excel路径
        if not excel_path:
            excel_path, msg = find_excel_file()
            if not excel_path:
                send_log(f"❌ {msg}")
                return (False, msg, "")
        # 默认模板路径
        if not template_path:
            template_path = os.path.join(tool_folder, "MemoTemplate.docx")
        # 默认输出路径
        if not output_path:
            output_path = os.path.join(tool_folder, "Filled_Memo.docx")

        send_log(f"📋 开始执行MEMO生成流程")
        send_log(f"Excel路径：{excel_path}")
        send_log(f"模板路径：{template_path}")
        send_log(f"输出路径：{output_path}")

        # 2. 读取Excel数据
        send_log("\n🔍 正在读取Excel数据...")
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel文件不存在：{excel_path}")

        # 读取Excel（无表头，取Sheet1工作表）
        workbook = openpyxl.load_workbook(excel_path, read_only=True)
        sheet_names = workbook.sheetnames
        send_log(f"Excel包含工作表：{sheet_names}")
        sheet = workbook['Sheet1']

        # 读取所有数据
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(list(row))

        # Get the actual maximum column count from the sheet
        max_column = sheet.max_column

        if len(data) == 0:
            raise ValueError("Excel文件中无任何数据行")
        if max_column < 2:  # 至少需要2列（B列=1、C列=2）
            raise ValueError(f"Excel列数不足（当前{max_column}列，需至少2列）")

        # 提取第一行关键数据（原逻辑保持不变）
        row = data[0]
        send_log(f"第一行数据：{row}")

        # 解析公司名称、型号、序列号
        company_full = str(row[2]) if len(row) > 2 else ""
        company_name = company_full.split('/')[-1].strip() if '/' in company_full else company_full.strip()
        model = str(row[4]) if len(row) > 4 else ""
        sn = str(row[1]) if len(row) > 1 else ""

        # 校验关键数据
        if not company_name:
            raise ValueError("未从Excel C列（索引2）提取到公司名称")
        if not model:
            raise ValueError("未从Excel H列（索引4）提取到设备型号")
        if not sn:
            raise ValueError("未从Excel B列（索引1）提取到序列号")

        send_log(f"✅ 提取数据完成：")
        send_log(f"  公司名称：{company_name}")
        send_log(f"  设备型号：{model}")
        send_log(f"  序列号：{sn}")

        # 计算日期（结束日期=今天，开始日期=2天前）
        end_date = datetime.now()
        start_date = end_date - timedelta(days=2)
        excel_data = {
            "买方": company_name,
            "设备型号": model,
            "序列号": sn,
            "安装开始日期": start_date.strftime("%Y.%m.%d"),
            "安装结束日期": end_date.strftime("%Y.%m.%d")
        }
        send_log(f"✅ 日期计算完成：{excel_data['安装开始日期']} - {excel_data['安装结束日期']}")

        # 3. 填充Word模板
        send_log("\n📄 正在填充Word模板...")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Word模板不存在：{template_path}")

        doc = Document(template_path)
        keyword_mapping = {  # 关键词→数据字段的映射（原逻辑保持不变）
            "买方：": "买方",
            "已完成": "设备型号",
            "序列号：": "序列号",
            "日期从": "安装开始日期",
            "至": "安装结束日期"
        }
        placeholder_count = 0  # 成功替换的占位符数量

        # 处理段落中的下划线占位符
        send_log("  处理段落中的占位符...")
        for paragraph in doc.paragraphs:
            for keyword, data_key in keyword_mapping.items():
                if keyword in paragraph.text:
                    found_keyword = False
                    for run in paragraph.runs:
                        # 先找到关键词，再找后续的下划线
                        if not found_keyword and keyword in run.text:
                            found_keyword = True
                            continue
                        # 替换关键词后的第一个下划线
                        if found_keyword and run.underline:
                            run.text = excel_data[data_key]
                            placeholder_count += 1
                            send_log(f"    替换段落占位符：'{keyword}'→'{excel_data[data_key]}'")
                            break  # 只替换第一个匹配的下划线

        # 处理表格中的下划线占位符
        send_log("  处理表格中的占位符...")
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for keyword, data_key in keyword_mapping.items():
                            if keyword in paragraph.text:
                                found_keyword = False
                                for run in paragraph.runs:
                                    if not found_keyword and keyword in run.text:
                                        found_keyword = True
                                        continue
                                    if found_keyword and run.underline:
                                        run.text = excel_data[data_key]
                                        placeholder_count += 1
                                        send_log(f"    替换表格占位符：'{keyword}'→'{excel_data[data_key]}'")
                                        break

        # 校验替换结果
        if placeholder_count == 0:
            raise ValueError("❌ 未替换任何占位符！请检查模板中的关键词和下划线格式")
        send_log(f"✅ 模板填充完成，共替换{placeholder_count}个占位符")

        # 4. 保存生成的MEMO
        send_log(f"\n💾 正在保存生成的MEMO...")
        doc.save(output_path)
        if not os.path.exists(output_path):
            raise Exception(f"MEMO保存失败（文件未生成）：{output_path}")

        send_log(f"✅ MEMO生成成功！路径：{output_path}")
        return (True, f"MEMO生成成功（{output_path}）", output_path)

    except FileNotFoundError as e:
        err_msg = f"文件错误：{str(e)}"
        send_log(f"❌ {err_msg}")
        return (False, err_msg, "")
    except ValueError as e:
        err_msg = f"数据错误：{str(e)}"
        send_log(f"❌ {err_msg}")
        return (False, err_msg, "")
    except IndexError as e:
        err_msg = f"索引错误：{str(e)}（Excel数据格式可能异常）"
        send_log(f"❌ {err_msg}")
        return (False, err_msg, "")
    except Exception as e:
        err_msg = f"未知错误：{str(e)}"
        send_log(f"❌ {err_msg}")
        import traceback
        traceback.print_exc()  # 打印详细堆栈（调试用）
        return (False, err_msg, "")
# memo_generator.py 末尾添加测试代码
if __name__ == "__main__":
    # 1. 定义日志打印函数（模拟主程序的回调）
    def test_log_callback(msg):
        print(f"[测试日志] {msg}")  # 打印每一步执行日志

    # 2. 手动指定路径（避免路径问题）
    tool_folder = os.path.join(os.path.expanduser("~"), "Desktop", "tool")
    excel_path, msg = find_excel_file()
    if not excel_path:
        print(msg)
        exit(1)
    template_path = os.path.join(tool_folder, "MemoTemplate.docx")  # 模板路径
    output_path = os.path.join(tool_folder, "Filled_Memo_test.docx")  # 测试输出路径

    # 3. 调用MEMO生成函数（带日志回调）
    test_log_callback("开始测试MEMO生成...")
    success, message, result_path = generate_memo(
        excel_path=excel_path,
        template_path=template_path,
        output_path=output_path,
        progress_callback=test_log_callback  # 传递日志函数
    )

    # 4. 打印最终结果
    test_log_callback(f"\n测试结束：")
    test_log_callback(f"是否成功：{'是' if success else '否'}")
    test_log_callback(f"结果信息：{message}")
    test_log_callback(f"生成文件路径：{result_path}")

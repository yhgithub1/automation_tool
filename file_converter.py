#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文件转换器主程序
支持Excel、Word、图片一键转换为PDF格式
"""

import sys
import os
import argparse
from pathlib import Path

# 添加模块路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from modules.file_converter_always import FileConverter

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="文件转换器 - 支持Excel、Word、图片一键转换为PDF格式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python file_converter.py input.xlsx
  python file_converter.py input.docx output.pdf
  python file_converter.py input.jpg ./output/
        """
    )
    
    parser.add_argument("--gui", action="store_true", help="启动图形界面")
    parser.add_argument("--verbose", "-v", action="store_true", help="显示详细信息")
    parser.add_argument("input", nargs="?", help="输入文件路径或文件夹路径")
    parser.add_argument("output", nargs="?", help="输出文件路径或输出目录")
    
    args = parser.parse_args()
    
    if args.gui:
        # 启动图形界面
        try:
            from modules.file_converter_ui import main as gui_main
            gui_main()
        except ImportError as e:
            print(f"错误: 无法启动图形界面: {e}")
            print("请确保已安装必要的依赖: pip install PyQt5 qtawesome")
            sys.exit(1)
        return
    
    # 命令行模式
    converter = FileConverter(verbose=args.verbose)
    
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"错误: 输入文件或目录不存在: {args.input}")
        sys.exit(1)
    
    if input_path.is_dir():
        # 批量转换目录中的文件
        print(f"正在批量转换目录: {input_path}")
        output_dir = args.output
        if not output_dir:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "converted_pdfs")
        
        success_count, failed_count, results = converter.batch_convert(
            [str(f) for f in input_path.iterdir() if f.is_file()],
            output_dir
        )
        
        print(f"\n转换完成!")
        print(f"成功: {success_count} 个文件")
        print(f"失败: {failed_count} 个文件")
        
        if args.verbose and results:
            print("\n详细结果:")
            for result in results:
                status = "✅" if result["success"] else "❌"
                print(f"{status} {result['input']} -> {result['output']}")
                
    else:
        # 单文件转换
        output_path = args.output
        if not output_path:
            # 自动生成输出路径
            output_path = str(input_path.parent / f"{input_path.stem}.pdf")
        
        output_path = Path(output_path)
        
        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        print(f"正在转换: {input_path}")
        print(f"输出到: {output_path}")
        
        success, result_path = converter.convert_to_pdf(str(input_path), str(output_path))
        
        if success:
            print(f"✅转换成功: {result_path}")
        else:
            print(f"❌ 转换失败: {result_path}")

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文件转换器模块
支持一键将Excel、Word、图片转换为PDF文件（已修复中文乱码）
"""

import os
import sys
import logging
import tempfile
import subprocess
import comtypes.client
import time
from pathlib import Path
from datetime import datetime
from PyQt5.QtCore import QObject, pyqtSignal

try:
    from PIL import Image
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
except ImportError as e:
    logging.error(f"缺少依赖库: {e}")
    raise


class FileConverter(QObject):
    """文件转换器：支持Excel、Word、图片转PDF（中文乱码修复版）"""
    
    # 信号定义
    log_signal = pyqtSignal(str)  # 日志信号
    progress_signal = pyqtSignal(int)  # 进度信号 (0-100)
    finished_signal = pyqtSignal(bool, str)  # 完成信号 (成功/失败, 输出文件路径)
    
    def __init__(self, verbose=False):
        super().__init__()
        self.verbose = verbose
        self.logger = logging.getLogger(__name__)
        self.is_canceled = False
        
    def cancel_conversion(self):
        """取消转换任务"""
        self.is_canceled = True
        self.log_signal.emit("正在取消转换任务...")
    
    def _kill_processes(self, process_name):
        """通用进程清理函数"""
        if sys.platform != "win32":
            self.log_signal.emit("仅Windows系统支持进程清理")
            return
        
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", process_name],
                check=False,
                capture_output=True,
                text=True
            )
            if "成功" in result.stdout:
                self.log_signal.emit(f"已强制结束所有残留的{process_name}进程")
            elif "找不到进程" in result.stderr:
                self.log_signal.emit(f"没有发现残留的{process_name}进程")
            else:
                self.log_signal.emit(f"结束{process_name}进程时出错: {result.stderr}")
        except Exception as e:
            self.log_signal.emit(f"进程清理异常: {str(e)}")
    
    def convert_to_pdf(self, input_file, output_file=None):
        """
        一键转换文件到PDF
        
        Args:
            input_file (str): 输入文件路径
            output_file (str, optional): 输出PDF文件路径，如果为None则自动生成
            
        Returns:
            tuple: (success: bool, output_path: str)
        """
        try:
            if self.is_canceled:
                return False, ""
                
            # 检查输入文件
            if not os.path.exists(input_file):
                error_msg = f"输入文件不存在: {input_file}"
                self.log_signal.emit(f"{error_msg}")
                return False, ""
            
            # 获取文件扩展名
            file_ext = Path(input_file).suffix.lower()
            
            # 如果没有指定输出文件，自动生成
            if not output_file:
                output_file = self._generate_output_path(input_file)
            
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # 根据文件类型进行转换
            if file_ext in ['.xlsx', '.xls']:
                self._kill_processes("EXCEL.EXE")  # 转换前清理Excel进程
                return self._excel_to_pdf(input_file, output_file)
            elif file_ext in ['.docx', '.doc']:
                self._kill_processes("WINWORD.EXE")  # 转换前清理Word进程
                return self._word_to_pdf(input_file, output_file)
            elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
                return self._image_to_pdf(input_file, output_file)
            else:
                error_msg = f"不支持的文件格式: {file_ext}"
                self.log_signal.emit(f"{error_msg}")
                return False, ""
                
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return False, ""
    
    def _generate_output_path(self, input_file):
        """生成输出文件路径"""
        input_path = Path(input_file)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{input_path.stem}_{timestamp}.pdf"
        output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "converted_pdfs")
        return os.path.join(output_dir, output_filename)
    
    def _excel_to_pdf(self, excel_file, pdf_file):
        """Excel转PDF（优化后）"""
        try:
            if self.is_canceled:
                return False, ""
                
            self.log_signal.emit(f"开始转换Excel文件: {os.path.basename(excel_file)}")
            self.progress_signal.emit(10)
            
            # 首先尝试使用Office内存转换方法（主要方法）
            try:
                return self._excel_to_pdf_office_memory(excel_file, pdf_file)
            except Exception as memory_error:
                self.log_signal.emit(f"Office内存转换失败: {str(memory_error)}")
                self.log_signal.emit("尝试使用COM对象方法...")
                
                # 如果内存方法失败，使用COM对象方法作为备用
                try:
                    return self._excel_to_pdf_com(excel_file, pdf_file)
                except Exception as com_error:
                    self.log_signal.emit(f"COM对象方法也失败: {str(com_error)}")
                    self.finished_signal.emit(False, "")
                    return False, ""
                
        except Exception as e:
            error_msg = f"Excel转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            self.finished_signal.emit(False, "")
            return False, ""
    
    def _excel_to_pdf_com(self, excel_file, pdf_file):
        """使用COM对象转换Excel到PDF"""
        try:
            if self.is_canceled:
                return False, ""
            
            # 使用comtypes将Excel转换为PDF
            excel = comtypes.client.CreateObject('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.Interactive = False
            
            try:
                excel.ActivePrinter = "Microsoft Print to PDF"
                excel.PrintCommunication = False
                
                # 尝试多种方式打开工作簿
                try:
                    wb = excel.Workbooks.Open(excel_file)
                    self.log_signal.emit("工作簿打开成功")
                except Exception as open_error:
                    self.log_signal.emit(f"直接打开失败，尝试备用方法: {str(open_error)}")
                    try:
                        abs_path = os.path.abspath(excel_file)
                        wb = excel.Workbooks.Open(abs_path)
                        self.log_signal.emit("使用绝对路径打开成功")
                    except Exception as abs_error:
                        self.log_signal.emit(f"绝对路径也失败，尝试兼容模式: {str(abs_error)}")
                        wb = excel.Workbooks.Open(
                            excel_file,
                            UpdateLinks=0,
                            ReadOnly=True,
                            Format=2
                        )
                        self.log_signal.emit("兼容模式打开成功")
                
                self.progress_signal.emit(50)
                
                # 获取工作表信息
                sheet_count = wb.Worksheets.Count
                self.log_signal.emit(f"检测到 {sheet_count} 个工作表")
                
                # 检查工作表内容
                for i in range(1, min(sheet_count + 1, 10)):
                    try:
                        ws = wb.Worksheets(i)
                        if ws.Visible == -1:
                            used_range = ws.UsedRange
                            if used_range.Rows.Count > 0 and used_range.Columns.Count > 0:
                                self.log_signal.emit(f"工作表 '{ws.Name}' 包含数据")
                                break
                    except:
                        continue
                
                self.log_signal.emit(f"开始导出Excel为PDF...")
                
                # 执行PDF导出
                wb.ExportAsFixedFormat(0, pdf_file, 1, False, False, 1, 50000, False)
                self.log_signal.emit(f"Excel PDF导出完成")
                
                # 关闭工作簿
                wb.Close(False)
                import time
                time.sleep(0.5)
                
                excel.Quit()
                
                # 验证生成的PDF文件
                if os.path.exists(pdf_file) and os.path.getsize(pdf_file) > 0:
                    file_size = os.path.getsize(pdf_file)
                    self.log_signal.emit(f"生成PDF文件大小: {file_size} 字节")
                    self.progress_signal.emit(100)
                    self.log_signal.emit(f"Excel转换完成: {os.path.basename(pdf_file)}")
                    return True, pdf_file
                else:
                    error_msg = f"生成的PDF文件无效或为空: {pdf_file}"
                    self.log_signal.emit(f"{error_msg}")
                    return False, ""
                
            except Exception as e:
                self.log_signal.emit(f"Excel操作异常，尝试清理资源...")
                try:
                    excel.Quit()
                except:
                    pass
                raise e
                
        except Exception as e:
            raise e
    
    def _excel_to_pdf_backup(self, excel_file, pdf_file):
        """使用Office内存转换方法（备用）"""
        try:
            if self.is_canceled:
                return False, ""
            
            self.log_signal.emit("使用Office内存转换方法（备用）")
            self.progress_signal.emit(20)
            
            return self._excel_to_pdf_office_memory(excel_file, pdf_file)
                
        except Exception as e:
            error_msg = f"Office内存转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return False, ""
    
    def _excel_to_pdf_office_memory(self, excel_file, pdf_file):
        """使用Office内存方式转换Excel到PDF"""
        try:
            if self.is_canceled:
                return False, ""
            
            self.log_signal.emit("使用Office内存转换...")
            
            excel = None
            wb = None
            
            try:
                excel = comtypes.client.CreateObject('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                excel.Interactive = False
                
                # 尝试打开工作簿
                try:
                    wb = excel.Workbooks.Open(
                        excel_file,
                        UpdateLinks=0,
                        ReadOnly=True,
                        Format=2
                    )
                    self.log_signal.emit("工作簿在内存中打开成功")
                except Exception as open_error:
                    self.log_signal.emit(f"内存打开失败，尝试标准模式: {str(open_error)}")
                    wb = excel.Workbooks.Open(excel_file)
                    self.log_signal.emit("标准模式打开成功")
                
                self.progress_signal.emit(50)
                
                # 检查工作表
                sheet_count = wb.Worksheets.Count
                self.log_signal.emit(f"检测到 {sheet_count} 个工作表")
                
                # 检查是否有数据
                has_data = False
                for i in range(1, min(sheet_count + 1, 10)):
                    try:
                        ws = wb.Worksheets(i)
                        if ws.Visible == -1:
                            used_range = ws.UsedRange
                            if used_range.Rows.Count > 0 and used_range.Columns.Count > 0:
                                self.log_signal.emit(f"工作表 '{ws.Name}' 包含数据")
                                has_data = True
                                break
                    except:
                        continue
                
                if not has_data:
                    self.log_signal.emit("没有找到包含数据的工作表")
                    return False, ""
                
                self.log_signal.emit("使用Excel默认页面设置，避免权限错误")
                self.progress_signal.emit(70)
                
                # 执行PDF导出
                self.log_signal.emit("开始导出为PDF...")
                wb.ExportAsFixedFormat(
                    Type=0,
                    Filename=pdf_file,
                    Quality=1,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    From=1,
                    To=50000,
                    OpenAfterPublish=False
                )
                
                self.log_signal.emit("Office PDF导出完成")
                
                # 关闭工作簿
                wb.Close(SaveChanges=False)
                wb = None
                
                import time
                time.sleep(1)
                
                # 退出Excel
                excel.Quit()
                excel = None
                
                # 验证生成的PDF文件
                if os.path.exists(pdf_file) and os.path.getsize(pdf_file) > 0:
                    file_size = os.path.getsize(pdf_file)
                    self.log_signal.emit(f"生成PDF文件大小: {file_size} 字节")
                    self.progress_signal.emit(100)
                    self.log_signal.emit(f"Office内存转换完成: {os.path.basename(pdf_file)}")
                    return True, pdf_file
                else:
                    error_msg = f"生成的PDF文件无效或为空: {pdf_file}"
                    self.log_signal.emit(f"{error_msg}")
                    return False, ""
                
            except Exception as e:
                # 清理资源
                try:
                    if wb:
                        wb.Close(SaveChanges=False)
                    if excel:
                        excel.Quit()
                except:
                    pass
                
                raise e
                
        except Exception as e:
            error_msg = f"Office内存转换异常: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return False, ""
    
    def _word_to_pdf(self, word_file, pdf_file):
        """Word转PDF（仅使用Office COM：深度优化中文支持）"""
        try:
            if self.is_canceled:
                return False, ""
                
            self.log_signal.emit(f"开始转换Word文件: {os.path.basename(word_file)}")
            self.progress_signal.emit(10)
            
            # 直接使用Office COM方法（删除Python方法）
            return self._word_to_pdf_com(word_file, pdf_file)
                    
        except Exception as e:
            error_msg = f"Word转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return False, ""
    
    def _word_to_pdf_com(self, word_file, pdf_file):
        """Word转PDF（Office COM：深度优化中文支持）"""
        word = None
        doc = None
        try:
            if self.is_canceled:
                return False, ""
                
            self.log_signal.emit("使用Microsoft Office转换（深度优化中文支持）")
            self.progress_signal.emit(20)
            
            # 创建Word实例（关键：设置语言环境为中文）
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            
            # 核心优化1：强制设置Word语言环境为中文（简体）
            try:
                word.Language = 2052  # 2052 = 简体中文
                self.log_signal.emit("✅ 已设置Word语言环境为简体中文")
            except:
                self.log_signal.emit("⚠️ 无法设置Word语言环境，继续尝试...")
            
            # 核心优化2：尝试多种编码方式打开文档（优先GBK，再UTF-8）
            encodings = [
                (936, "GBK/GB2312（中文Windows默认）"),  # 中文Windows默认编码
                (65001, "UTF-8（通用编码）"),
                (0, "自动检测（系统默认）")
            ]
            
            doc = None
            for codepage, desc in encodings:
                if self.is_canceled:
                    return False, ""
                    
                try:
                    self.log_signal.emit(f"尝试用 {desc} 打开文档...")
                    word.Application.DefaultTextEncoding = codepage
                    
                    # 关键参数：NoEncodingDialog=True 防止弹出编码选择对话框
                    doc = word.Documents.Open(
                        FileName=word_file,
                        ReadOnly=True,
                        NoEncodingDialog=True,
                        ConfirmConversions=False,
                        Encoding=codepage if codepage != 0 else None
                    )
                    self.log_signal.emit(f"✅ 文档成功用 {desc} 打开")
                    break
                except Exception as e:
                    self.log_signal.emit(f"❌ {desc} 打开失败: {str(e)}")
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass
                    doc = None
            
            if not doc:
                raise Exception("所有编码尝试均失败，无法打开文档")
            
            self.progress_signal.emit(50)
            
            # 核心优化3：强制设置文档语言为中文（修复样式乱码）
            try:
                for story_range in doc.StoryRanges:
                    story_range.LanguageID = 2052  # 简体中文
                self.log_signal.emit("✅ 已强制设置文档语言为简体中文")
            except Exception as e:
                self.log_signal.emit(f"⚠️ 设置文档语言失败: {str(e)}")
            
            # 核心优化4：检查并修复中文字体（确保使用系统中文字体）
            try:
                chinese_fonts = ["宋体", "SimSun", "Microsoft YaHei", "微软雅黑", "黑体"]
                for font in word.Fonts:
                    if font.Name in chinese_fonts or "Sim" in font.Name or "YaHei" in font.Name:
                        font.NameAscii = font.Name
                        font.NameOther = font.Name
                self.log_signal.emit("✅ 已修复中文字体映射")
            except Exception as e:
                self.log_signal.emit(f"⚠️ 字体修复失败: {str(e)}")
            
            self.log_signal.emit("正在导出为PDF...")
            self.progress_signal.emit(70)
            
            # 核心优化5：使用更稳定的PDF导出方法（修复参数错误）
            try:
                # 先保存为PDF格式（使用正确的参数）
                doc.ExportAsFixedFormat(
                    pdf_file,  # 输出文件路径
                    17,        # wdExportFormatPDF = 17
                    False,     # OpenAfterExport
                    0,         # ExportOptimizeFor = wdExportOptimizeForPrint
                    0,         # Range = wdExportAllDocument
                    1,         # From = 1
                    50000,     # To = 50000
                    False,     # IncludeDocProperties
                    False,     # KeepIRM
                    1,         # CreateBookmarks = wdExportCreateNoBookmarks
                    0,         # DocStructureTags = True
                    False,     # BitmapMissingFonts
                    False,     # UseISO19005_1
                    False      # OptimizeForBestPrintQuality
                )
                self.log_signal.emit("✅ ExportAsFixedFormat成功")
            except Exception as export_error:
                self.log_signal.emit(f"ExportAsFixedFormat失败: {str(export_error)}")
                self.log_signal.emit("尝试使用PrintOut方法...")
                
                # 备用方案：使用打印到PDF（简化参数）
                try:
                    word.ActivePrinter = "Microsoft Print to PDF"
                    # 使用最少的必要参数
                    doc.PrintOut(
                        Background=False,
                        Range=0,  # wdPrintAllDocument
                        Item=0,   # wdPrintDocumentContent
                        Copies=1,
                        Pages="",
                        PageType=0,  # wdPrintAllPages
                        PrintToFile=True,
                        FileName=pdf_file
                    )
                    self.log_signal.emit("✅ PrintOut方法成功")
                except Exception as print_error:
                    self.log_signal.emit(f"PrintOut方法也失败: {str(print_error)}")
                    raise print_error
            
            # 等待PDF生成完成（关键：增加超时等待）
            self.log_signal.emit("等待PDF文件生成...")
            max_wait = 30  # 增加到30秒（大文档需要更长时间）
            wait_count = 0
            file_ready = False
            
            while wait_count < max_wait and not file_ready:
                time.sleep(1)
                wait_count += 1
                
                if os.path.exists(pdf_file):
                    # 检查文件是否仍在写入（大小变化）
                    size1 = os.path.getsize(pdf_file)
                    time.sleep(1)
                    if os.path.exists(pdf_file):  # 再次检查文件是否存在
                        size2 = os.path.getsize(pdf_file)
                        if size2 > 100 and size1 == size2:  # 文件大小稳定
                            file_ready = True
                            break
                
                self.log_signal.emit(f"已等待 {wait_count}/{max_wait} 秒...")
            
            if not file_ready:
                # 尝试强制关闭文档再检查
                try:
                    doc.Close(SaveChanges=False)
                    time.sleep(2)
                    if os.path.exists(pdf_file) and os.path.getsize(pdf_file) > 100:
                        file_ready = True
                        self.log_signal.emit("✅ PDF文件已生成（强制关闭后）")
                except:
                    pass
            
            if not file_ready:
                error_msg = f"PDF生成超时（超过{max_wait}秒）或文件无效"
                self.log_signal.emit(f"{error_msg}")
                return False, ""
            
            self.log_signal.emit("PDF导出完成，验证文件有效性...")
            self.progress_signal.emit(90)
            
            # 验证PDF（增加验证步骤）
            if not os.path.exists(pdf_file):
                raise FileNotFoundError(f"PDF文件未生成: {pdf_file}")
                
            file_size = os.path.getsize(pdf_file)
            if file_size < 100:
                raise ValueError(f"生成的PDF文件过小 ({file_size}字节)，可能损坏")
            

            
            self.progress_signal.emit(100)
            self.log_signal.emit(f"Word转换完成: {os.path.basename(pdf_file)} (大小: {file_size}字节)")
            return True, pdf_file
                
        except Exception as e:
            error_msg = f"Office转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return False, ""
        finally:
            # 增强的资源清理（关键：确保进程完全退出）
            try:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                        self.log_signal.emit("✅ Word文档已关闭")
                    except Exception as e:
                        self.log_signal.emit(f"⚠️ 关闭文档失败: {str(e)}")
                    doc = None
            except:
                pass
            
            try:
                if word:
                    try:
                        word.Quit()
                        self.log_signal.emit("✅ Word进程已退出")
                    except Exception as e:
                        self.log_signal.emit(f"⚠️ 退出Word失败: {str(e)}")
                    word = None
            except:
                pass
            
            # 强制结束残留WINWORD.EXE（双重保险）
            if sys.platform == "win32":
                try:
                    # 先尝试优雅退出
                    subprocess.run(
                        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                        check=False, 
                        capture_output=True, 
                        text=True,
                        timeout=5
                    )
                    time.sleep(1)  # 等待进程终止
                    
                    # 再次检查并强制结束
                    result = subprocess.run(
                        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
                        capture_output=True,
                        text=True
                    )
                    if "WINWORD.EXE" in result.stdout:
                        subprocess.run(
                            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                            check=False
                        )
                        self.log_signal.emit("✅ 已强制结束残留Word进程")
                except Exception as e:
                    self.log_signal.emit(f"⚠️ 进程清理异常: {str(e)}")

    
    def _image_to_pdf(self, image_file, pdf_file):
        """图片转PDF（保持原功能，支持透明背景处理）"""
        try:
            if self.is_canceled:
                return False, ""
                
            self.log_signal.emit(f"开始转换图片文件: {os.path.basename(image_file)}")
            self.progress_signal.emit(20)
            
            # 使用PIL处理图片
            with Image.open(image_file) as img:
                # 适配A4页面（保持图片比例）
                img_w, img_h = img.size
                if img_w > img_h:
                    pdf_w, pdf_h = A4[1], A4[0]  # 横向A4
                else:
                    pdf_w, pdf_h = A4  # 纵向A4
                
                # 计算居中显示尺寸（留边距）
                margin = 50
                max_display_w = pdf_w - 2 * margin
                max_display_h = pdf_h - 2 * margin
                scale = min(max_display_w / img_w, max_display_h / img_h, 1.0)
                display_w = img_w * scale
                display_h = img_h * scale
                pos_x = (pdf_w - display_w) / 2
                pos_y = (pdf_h - display_h) / 2
                
                # 处理PNG透明背景（转为白色背景）
                if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    # 创建白色背景的RGB图片
                    img_rgb = Image.new('RGB', img.size, (255, 255, 255))
                    img_rgb.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else img.info.get('transparency'))
                    img = img_rgb
                
                # 创建PDF并插入图片
                c = canvas.Canvas(pdf_file, pagesize=(pdf_w, pdf_h))
                self.progress_signal.emit(60)
                
                # 临时保存图片（避免PIL直接传参问题）
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                    temp_path = temp_img.name
                    try:
                        img.save(temp_path, 'JPEG', quality=95)
                        c.drawImage(temp_path, pos_x, pos_y, display_w, display_h)
                    finally:
                        if os.path.exists(temp_path):
                            os.unlink(temp_path)  # 删除临时文件
                
                c.showPage()
                c.save()
                
                self.progress_signal.emit(100)
                self.log_signal.emit(f"图片转换完成: {os.path.basename(pdf_file)}")
                self.finished_signal.emit(True, pdf_file)
                return True, pdf_file
                
        except Exception as e:
            error_msg = f"图片转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            self.finished_signal.emit(False, "")
            return False, ""
    
    def batch_convert(self, input_files, output_dir=None):
        """批量转换文件（保持原功能，支持进度反馈）"""
        try:
            if not input_files:
                self.log_signal.emit("没有文件需要转换")
                return 0, 0, []
            
            if self.is_canceled:
                return 0, 0, []
            
            # 设置输出目录
            if not output_dir:
                output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "converted_pdfs")
            os.makedirs(output_dir, exist_ok=True)
            
            self.log_signal.emit(f"输出目录: {output_dir}")
            self.log_signal.emit(f"开始批量转换 {len(input_files)} 个文件...")
            
            success_count = 0
            failed_count = 0
            results = []
            
            for i, input_file in enumerate(input_files):
                if self.is_canceled:
                    self.log_signal.emit("批量转换已取消")
                    break
                
                # 进度更新
                progress = int((i / len(input_files)) * 100)
                self.progress_signal.emit(progress)
                
                # 生成输出路径（带时间戳）
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                input_stem = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_stem}_{timestamp}.pdf")
                
                # 执行转换
                success, output_path = self.convert_to_pdf(input_file, output_file)
                
                if success:
                    success_count += 1
                    results.append((input_file, output_path, True))
                    self.log_signal.emit(f"{i+1}/{len(input_files)} 转换成功: {os.path.basename(output_path)}")
                else:
                    failed_count += 1
                    results.append((input_file, output_path, False))
                    self.log_signal.emit(f"{i+1}/{len(input_files)} 转换失败: {os.path.basename(input_file)}")
            
            # 批量完成
            self.progress_signal.emit(100)
            self.log_signal.emit(f"批量转换完成！成功: {success_count}, 失败: {failed_count}")
            self.finished_signal.emit(success_count > 0, output_dir)
            
            return success_count, failed_count, results
            
        except Exception as e:
            error_msg = f"批量转换失败: {str(e)}"
            self.log_signal.emit(f"{error_msg}")
            return 0, len(input_files), []


def main():
    """独立测试函数（支持命令行调用）"""
    import sys
    
    if len(sys.argv) < 2:
        print("使用方法: python file_converter.py <输入文件路径> [输出文件路径]")
        print("支持格式: Excel(.xlsx/.xls)、Word(.docx/.doc)、图片(.jpg/.png/.bmp/.gif/.tiff)")
        print("说明: Word转换使用Microsoft Office（深度优化中文支持）")
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # 创建转换器实例
    converter = FileConverter(verbose=True)
    
    # 日志回调（命令行打印）
    def log_print(msg):
        print(f"[日志] {msg}")
    
    converter.log_signal.connect(log_print)
    
    # 执行转换
    print(f"=== 开始转换文件: {input_file} ===")
    success, output_path = converter.convert_to_pdf(input_file, output_file)
    
    if success:
        print(f"=== 转换成功！输出文件: {output_path} ===")
    else:
        print("=== 转换失败！===")


if __name__ == "__main__":
    main()

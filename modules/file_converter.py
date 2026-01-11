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
import threading
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


class OfficeInstanceManager:
    """Office实例管理器：重用Office应用程序实例以提升性能"""

    _excel_instance = None
    _word_instance = None
    _instance_lock = threading.Lock()

    @classmethod
    def get_excel_instance(cls):
        """获取或创建Excel实例"""
        with cls._instance_lock:
            if cls._excel_instance is None:
                try:
                    cls._excel_instance = comtypes.client.CreateObject('Excel.Application')
                    cls._excel_instance.Visible = False
                    cls._excel_instance.DisplayAlerts = False
                    cls._excel_instance.ScreenUpdating = False
                    cls._excel_instance.Interactive = False
                    logging.info("Excel实例创建成功")
                except Exception as e:
                    logging.error(f"创建Excel实例失败: {e}")
                    raise
            return cls._excel_instance

    @classmethod
    def get_word_instance(cls):
        """获取或创建Word实例"""
        with cls._instance_lock:
            if cls._word_instance is None:
                try:
                    cls._word_instance = comtypes.client.CreateObject('Word.Application')
                    cls._word_instance.Visible = False
                    cls._word_instance.DisplayAlerts = 0
                    cls._word_instance.ScreenUpdating = False
                    logging.info("Word实例创建成功")
                except Exception as e:
                    logging.error(f"创建Word实例失败: {e}")
                    raise
            return cls._word_instance

    @classmethod
    def cleanup_instances(cls):
        """清理所有Office实例"""
        with cls._instance_lock:
            try:
                if cls._excel_instance:
                    cls._excel_instance.Quit()
                    cls._excel_instance = None
                    logging.info("Excel实例已清理")
            except:
                pass

            try:
                if cls._word_instance:
                    cls._word_instance.Quit()
                    cls._word_instance = None
                    logging.info("Word实例已清理")
            except:
                pass


class FileConverter(QObject):
    """文件转换器：支持Excel、Word、图片转PDF（高性能优化版）"""

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

    def cleanup_resources(self):
        """清理资源：在应用程序关闭时调用"""
        try:
            OfficeInstanceManager.cleanup_instances()
            self.log_signal.emit("Office实例已清理")
        except Exception as e:
            self.log_signal.emit(f"清理资源时出错: {str(e)}")
    
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
                return self._excel_to_pdf(input_file, output_file)
            elif file_ext in ['.docx', '.doc']:
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
            
            # 直接使用Office内存转换方法（成功率最高）
            return self._excel_to_pdf_office_memory(excel_file, pdf_file)
                
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
        """使用Office内存方式转换Excel到PDF（重用实例优化版）"""
        try:
            if self.is_canceled:
                return False, ""

            self.log_signal.emit("使用Office内存转换...")

            # 获取或创建Excel实例（重用实例）
            excel = OfficeInstanceManager.get_excel_instance()
            wb = None

            try:
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

                # 执行PDF导出（移除不必要的延迟）
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

                # 关闭工作簿（但保持Excel实例活跃）
                wb.Close(SaveChanges=False)
                wb = None

                # 立即验证PDF文件（移除1秒延迟）
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
                # 清理工作簿资源
                try:
                    if wb:
                        wb.Close(SaveChanges=False)
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
        """Word转PDF（Office COM：深度优化中文支持，重用实例版）"""
        doc = None
        try:
            if self.is_canceled:
                return False, ""

            self.log_signal.emit("使用Microsoft Office转换（深度优化中文支持）")
            self.progress_signal.emit(20)

            # 获取或创建Word实例（重用实例）
            word = OfficeInstanceManager.get_word_instance()

            # 直接用GBK编码打开文档（中文Windows默认，成功率最高）
            word.Application.DefaultTextEncoding = 936  # GBK编码
            doc = word.Documents.Open(
                FileName=word_file,
                ReadOnly=True,
                NoEncodingDialog=True,
                ConfirmConversions=False,
                Encoding=936
            )
            self.log_signal.emit("✅ 文档成功打开")

            self.progress_signal.emit(50)

            # 核心优化3：强制设置文档语言为中文（修复样式乱码）
            try:
                for story_range in doc.StoryRanges:
                    story_range.LanguageID = 2052  # 简体中文
                self.log_signal.emit("✅ 已强制设置文档语言为简体中文")
            except Exception as e:
                self.log_signal.emit(f"⚠️ 设置文档语言失败: {str(e)}")

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

            # 等待PDF生成完成（优化：减少等待时间，使用更智能的检测）
            self.log_signal.emit("等待PDF文件生成...")
            max_wait = 15  # 从30秒减少到15秒（优化后更快）
            wait_count = 0
            file_ready = False

            while wait_count < max_wait and not file_ready:
                # 减少检查间隔，从1秒改为0.5秒
                time.sleep(0.5)
                wait_count += 0.5

                if os.path.exists(pdf_file):
                    # 检查文件是否仍在写入（大小变化）
                    size1 = os.path.getsize(pdf_file)
                    # 减少检查间隔
                    time.sleep(0.5)
                    if os.path.exists(pdf_file):  # 再次检查文件是否存在
                        size2 = os.path.getsize(pdf_file)
                        if size2 > 100 and size1 == size2:  # 文件大小稳定
                            file_ready = True
                            break

                if wait_count % 2 == 0:  # 每2秒显示一次进度
                    self.log_signal.emit(f"已等待 {int(wait_count)}/{max_wait} 秒...")

            if not file_ready:
                # 尝试强制关闭文档再检查（移除额外的2秒延迟）
                try:
                    doc.Close(SaveChanges=False)
                    time.sleep(1)  # 从2秒减少到1秒
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
            # 优化的资源清理（保持Word实例活跃，只关闭文档）
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
                temp_path = None
                try:
                    with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                        temp_path = temp_img.name
                    img.save(temp_path, 'JPEG', quality=95)
                    c.drawImage(temp_path, pos_x, pos_y, display_w, display_h)
                finally:
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.unlink(temp_path)  # 删除临时文件
                        except:
                            pass  # 文件可能已被使用，忽略删除错误
                
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
    
    def batch_convert(self, input_files, output_dir=None, max_workers=3):
        """批量转换文件（并行处理优化版，支持进度反馈）"""
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
            self.log_signal.emit(f"开始批量转换 {len(input_files)} 个文件... (使用 {max_workers} 个并行线程)")

            import concurrent.futures
            import threading

            success_count = 0
            failed_count = 0
            results = []
            completed_count = 0
            progress_lock = threading.Lock()

            def convert_single_file(input_file):
                nonlocal completed_count
                try:
                    if self.is_canceled:
                        return (input_file, "", False)

                    # 生成输出路径（带时间戳）
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    input_stem = Path(input_file).stem
                    output_file = os.path.join(output_dir, f"{input_stem}_{timestamp}.pdf")

                    # 执行转换
                    success, output_path = self.convert_to_pdf(input_file, output_file)

                    # 更新进度
                    with progress_lock:
                        completed_count += 1
                        progress = int((completed_count / len(input_files)) * 100)
                        self.progress_signal.emit(progress)

                    if success:
                        self.log_signal.emit(f"{completed_count}/{len(input_files)} 转换成功: {os.path.basename(output_path)}")
                        return (input_file, output_path, True)
                    else:
                        self.log_signal.emit(f"{completed_count}/{len(input_files)} 转换失败: {os.path.basename(input_file)}")
                        return (input_file, output_path, False)

                except Exception as e:
                    with progress_lock:
                        completed_count += 1
                        progress = int((completed_count / len(input_files)) * 100)
                        self.progress_signal.emit(progress)
                    self.log_signal.emit(f"{completed_count}/{len(input_files)} 转换失败: {os.path.basename(input_file)} - {str(e)}")
                    return (input_file, "", False)

            # 使用线程池并行处理
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 提交所有任务
                future_to_file = {executor.submit(convert_single_file, file): file for file in input_files}

                # 等待所有任务完成
                for future in concurrent.futures.as_completed(future_to_file):
                    if self.is_canceled:
                        # 取消剩余任务
                        executor.shutdown(wait=False)
                        break

                    try:
                        result = future.result()
                        results.append(result)
                        if result[2]:  # success
                            success_count += 1
                        else:
                            failed_count += 1
                    except Exception as e:
                        self.log_signal.emit(f"任务执行异常: {str(e)}")
                        failed_count += 1

            if self.is_canceled:
                self.log_signal.emit("批量转换已取消")
            else:
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
        print("批量测试: python file_converter.py --batch <文件1> <文件2> ...")
        print("支持格式: Excel(.xlsx/.xls)、Word(.docx/.doc)、图片(.jpg/.png/.bmp/.gif/.tiff)")
        print("说明: Word转换使用Microsoft Office（深度优化中文支持）")
        return

    if sys.argv[1] == "--batch" and len(sys.argv) > 2:
        # 批量测试模式
        input_files = sys.argv[2:]
        print(f"=== 开始批量转换测试: {len(input_files)} 个文件 ===")

        import time
        start_time = time.time()

        converter = FileConverter(verbose=True)

        # 日志回调（命令行打印）
        def log_print(msg):
            print(f"[日志] {msg}")

        converter.log_signal.connect(log_print)

        success_count, failed_count, results = converter.batch_convert(input_files, max_workers=3)

        end_time = time.time()
        total_time = end_time - start_time

        print(f"=== 批量转换完成 ===")
        print(f"总耗时: {total_time:.2f} 秒")
        print(f"平均每个文件耗时: {total_time/len(input_files):.2f} 秒")
        print(f"成功: {success_count}, 失败: {failed_count}")

        # 清理资源
        converter.cleanup_resources()

    else:
        # 单文件测试模式
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

        # 清理资源
        converter.cleanup_resources()


if __name__ == "__main__":
    main()

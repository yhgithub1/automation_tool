#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
自动化工具配置文件
集中管理所有配置项，便于维护和调整
"""

import os
from dataclasses import dataclass
from typing import Dict, List, Optional

@dataclass
class WeChatConfig:
    """企业微信自动化配置"""
    
    # 窗口相关配置
    window_title_keywords: List[str] = None
    window_activate_retries: int = 4
    window_activate_interval: float = 0.6
    window_maximize_timeout: float = 0.8
    
    # 模板匹配配置
    workbench_template: str = None
    report_template: str = None
    new_report_template: str = None
    delivery_tool_template: str = None
    template_match_threshold: float = 0.7
    
    # OCR配置
    tesseract_config: str = r'--oem 3 --psm 6 -l chi_sim+eng'
    ocr_confidence_threshold: float = 20.0  # 降低置信度阈值，提高识别率
    ocr_preprocess_enabled: bool = False  # 关闭预处理，避免影响识别精度
    
    # 等待配置
    wait_for_text_timeout: float = 15.0
    wait_for_text_interval: float = 1.0
    form_load_timeout: float = 30.0
    scroll_wait_time: float = 2.0
    
    # 滚动配置
    scroll_amount: int = -800
    scroll_interval: float = 1.5
    
    # 点击配置
    click_duration: float = 0.12
    click_post_sleep: float = 0.35
    click_pre_sleep: float = 0.06
    
    # 输入配置
    type_select_all_wait: float = 0.6
    type_clear_wait: float = 0.4
    type_paste_wait: float = 0.8
    
    # 表单字段配置
    form_field_mapping: Dict[str, str] = None
    form_field_offsets: Dict[str, tuple] = None
    
    def __post_init__(self):
        if self.window_title_keywords is None:
            self.window_title_keywords = ["企业微信"]
        
        # 设置模板文件的绝对路径
        if self.workbench_template is None:
            self.workbench_template = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "workbench.png"
            )
        if self.report_template is None:
            self.report_template = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "report.png"
            )
        if self.new_report_template is None:
            self.new_report_template = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "new_report.png"
            )
        if self.delivery_tool_template is None:
            self.delivery_tool_template = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "delivery_tool.png"
            )
        
        if self.form_field_mapping is None:
            self.form_field_mapping = {
                "company": "客户公司名称",
                "address": "客户详细地址", 
                "name": "客户姓名",
                "phone": "客户手机号码"
            }
        
        if self.form_field_offsets is None:
            self.form_field_offsets = {
                "company": (0, 40),  # x_offset, y_offset
                "address": (500, 0),
                "name": (500, 0), 
                "phone": (500, 0)
            }

@dataclass
class ExcelConfig:
    """Excel文件配置"""
    
    # 默认Excel路径
    default_excel_path: str = None
    excel_sheet_name: str = None
    excel_data_mapping: Dict[str, str] = None
    
    def __post_init__(self):
        if self.default_excel_path is None:
            self.default_excel_path = os.path.join(
                os.path.expanduser("~"), "Desktop", "tool", "1.xlsx"
            )
        
        if self.excel_data_mapping is None:
            self.excel_data_mapping = {
                "company_name": "C1",
                "address_info": "N1"
            }

@dataclass
class DebugConfig:
    """调试配置"""
    
    # 调试目录
    debug_dir: str = None
    save_debug_screenshots: bool = False
    debug_log_level: str = "INFO"
    
    # OCR调试
    ocr_save_input: bool = False
    save_not_found_screenshots: bool = False
    
    def __post_init__(self):
        if self.debug_dir is None:
            self.debug_dir = os.path.join(os.path.expanduser("~"), "ocr_debug")

@dataclass
class PerformanceConfig:
    """性能优化配置"""
    
    # 缓存配置
    enable_ocr_cache: bool = True
    ocr_cache_ttl: int = 30  # 秒
    
    # 截图优化
    enable_high_resolution: bool = True
    high_res_scale: float = 2.0
    adaptive_scale_enabled: bool = True
    
    # 重试机制
    enable_retry_mechanism: bool = True
    max_retry_attempts: int = 3
    retry_delay: float = 1.0

# 全局配置实例
WECHAT_CONFIG = WeChatConfig()
EXCEL_CONFIG = ExcelConfig() 
DEBUG_CONFIG = DebugConfig()
PERFORMANCE_CONFIG = PerformanceConfig()

# Tesseract路径配置
TESSERACT_PATH = r'C:\Users\zchangyu\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

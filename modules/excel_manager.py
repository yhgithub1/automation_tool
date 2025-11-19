#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel数据管理工具类
负责读取Excel文件和解析数据
"""

import os
import re
import logging
from typing import Dict, Any, Optional

try:
    import openpyxl
except ImportError as e:
    logging.error(f"缺少依赖库: {e}")
    raise

from .config import EXCEL_CONFIG

class AddressParser:
    """地址解析器"""
    
    @staticmethod
    def parse_address_info(content: str) -> Dict[str, str]:
        """
        解析地址、姓名、电话信息
        
        Args:
            content: 包含地址、姓名、电话的文本
            
        Returns:
            Dict: 解析结果，包含address、name、phone字段
        """
        if not content:
            return {"address": "", "name": "", "phone": ""}
        
        # 清理文本
        content = content.replace('\n', ' ').strip()
        content = re.sub(r'\s+', ' ', content)
        
        # 定义解析模式（从精确到模糊）
        patterns = [
            # 最精确的模式：地址格式（路/街/巷/道 + 门牌号）
            (r'([\u4e00-\u9fa50-9a-zA-Z\s\-\.号楼路街巷区镇省市县]+?[路街巷道]\s*\d+[号号楼]?)\s+([^\d]+?)\s+(\d{10,11})$',
             ['address', 'name', 'phone']),
            
            # 次精确模式：包含数字的地址
            (r'([\u4e00-\u9fa50-9a-zA-Z\s\-\.号楼路街巷区镇省市县]+?\d+)\s+([^\d]+?)\s+(\d{10,11})$',
             ['address', 'name', 'phone']),
            
            # 宽松模式：任意格式，电话必须是10-11位数字
            (r'(.*?)\s+([^\d]+?)\s+(\d{10,11})$',
             ['address', 'name', 'phone'])
        ]
        
        for pattern, keys in patterns:
            match = re.match(pattern, content)
            if match:
                return {
                    keys[0]: match.group(1).strip(),
                    keys[1]: match.group(2).strip(),
                    keys[2]: match.group(3).strip()
                }
        
        # 如果没有匹配到任何模式，返回空值
        return {"address": "", "name": "", "phone": ""}

class ExcelManager:
    """Excel管理器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.parser = AddressParser()
    
    def read_excel_data(self, excel_path: Optional[str] = None) -> Dict[str, Any]:
        """
        读取Excel数据
        
        Args:
            excel_path: Excel文件路径，如果为None则使用默认路径
            
        Returns:
            Dict: 包含公司名称和地址信息的字典
            
        Raises:
            FileNotFoundError: 文件不存在
            Exception: 读取失败
        """
        if not excel_path:
            excel_path = EXCEL_CONFIG.default_excel_path
        
        self.logger.info(f"读取Excel文件: {excel_path}")
        
        # 检查文件是否存在
        if not os.path.exists(excel_path):
            error_msg = f"Excel文件不存在: {excel_path}"
            self.logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        
        try:
            # 打开工作簿
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet = wb.active
            
            # 读取数据
            company_name_cell = EXCEL_CONFIG.excel_data_mapping["company_name"]
            address_info_cell = EXCEL_CONFIG.excel_data_mapping["address_info"]
            
            company_name = sheet[company_name_cell].value or ""
            address_info = sheet[address_info_cell].value or ""
            
            # 处理公司名称（提取最后部分）
            content_to_paste = company_name.split("/")[-1] if company_name and "/" in company_name else company_name
            
            # 解析地址信息
            parsed_info = self.parser.parse_address_info(address_info)
            
            result = {
                "company_name": content_to_paste,
                "customer_address": parsed_info["address"],
                "customer_name": parsed_info["name"],
                "customer_phone": parsed_info["phone"]
            }
            
            self.logger.info("Excel数据读取成功")
            self.logger.debug(f"读取到的数据: {result}")
            
            return result
            
        except Exception as e:
            error_msg = f"读取Excel文件失败: {e}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
    
    def validate_excel_data(self, data: Dict[str, Any]) -> bool:
        """
        验证Excel数据的完整性
        
        Args:
            data: 要验证的数据字典
            
        Returns:
            bool: 数据是否有效
        """
        required_fields = ["company_name", "customer_address", "customer_name", "customer_phone"]
        
        for field in required_fields:
            if field not in data:
                self.logger.warning(f"缺少必需字段: {field}")
                return False
        
        # 检查电话号码格式
        phone = data.get("customer_phone", "")
        if phone and not re.match(r'^\d{10,11}$', phone):
            self.logger.warning(f"电话号码格式不正确: {phone}")
            return False
        
        return True
    
    def get_default_excel_path(self) -> str:
        """获取默认Excel路径"""
        return EXCEL_CONFIG.default_excel_path
    
    def set_excel_path(self, path: str) -> None:
        """设置Excel路径"""
        EXCEL_CONFIG.default_excel_path = path
        self.logger.info(f"设置Excel路径: {path}")

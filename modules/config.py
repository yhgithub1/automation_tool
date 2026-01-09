# modules/config.py
# 硬编码所有默认路径，避免在启动时导入其他模块
import os
# PDF相关路径
PDF_INPUT_DIR = r"H:\Shanghai\IMT\Service\Management Tools\量具\标准器校准证书最新\02步距规"
PDF_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "tool")

# 其他配置
DEFAULT_SEARCH_DIR = r"C:\Zeiss\CMM_Tools\FW_C99\backup"
DEFAULT_SEARCH_CONTENT = "Install_version = V47.04"
DEFAULT_FILE_NAMES = "config.kmg"

# Excel相关
EXCEL_SEARCH_PATHS = [
    os.path.join(os.path.expanduser("~/Desktop"), "tool"),
    os.path.join(os.path.expanduser("~/Desktop"), "tool", "*.xls"),
    os.path.join(os.path.expanduser("~/Desktop"), "tool", "*.xlsx"),
]

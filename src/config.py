"""应用配置常量"""
from typing import Dict, Final

# 文件格式
EXCEL_FILE_FILTERS: Final[str] = "Excel Files (*.xlsx *.xls *.xlsm)"

# UI相关
WINDOW_MIN_WIDTH: Final[int] = 800
WINDOW_MIN_HEIGHT: Final[int] = 600

# 应用信息
APP_NAME: Final[str] = "Excel数据比对工具"
APP_VERSION: Final[str] = "1.0.0"

# 时间格式
TIMESTAMP_FORMAT: Final[str] = "%Y%m%d_%H%M%S"

# 消息
MESSAGES: Final[Dict[str, str]] = {
    "no_file_selected": "未选择文件",
    "select_file_first": "请导入Excel文件开始操作",
    "select_columns": "请至少选择一列进行比对",
    "preview_first": "请先预览数据",
    "enter_filter_text": "请输入要比对的内容",
    "no_data_found": "没有找到匹配的数据",
    "export_success": "已成功导出Excel文件",
    "filter_success": "筛选成功",
    "import_success": "成功导入文件",
}

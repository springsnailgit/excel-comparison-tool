"""应用配置管理模块"""
import json
import logging
from typing import Dict, Final, Any
from pathlib import Path

# 应用基本信息
APP_NAME: Final[str] = "Excel数据比对工具"
APP_VERSION: Final[str] = "1.2.1"

# 默认配置
DEFAULT_CONFIG = {
    # 文件格式
    "excel_file_filters": "Excel Files (*.xlsx *.xls *.xlsm)",

    # UI相关
    "window_min_width": 800,
    "window_min_height": 600,
    "table_max_display_rows": 1000,  # 表格最大显示行数

    # 数据处理
    "max_file_size_mb": 100,  # 最大文件大小(MB)
    "chunk_size": 10000,  # 数据处理块大小
    "max_filter_conditions": 50,  # 最大筛选条件数

    # 时间格式
    "timestamp_format": "%Y%m%d_%H%M%S",
    "log_date_format": "%Y-%m-%d %H:%M:%S",

    # 日志配置
    "log_level": "INFO",
    "log_file_max_size": 10,  # MB
    "log_backup_count": 5,

    # 导出配置
    "export_sheet_name_max_length": 31,
    "export_filename_max_length": 200,
}

# 消息配置
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
    "file_too_large": "文件过大，请选择小于 {max_size}MB 的文件",
    "invalid_file_format": "不支持的文件格式",
    "data_validation_failed": "数据验证失败",
    "operation_cancelled": "操作已取消",
    "memory_warning": "数据量较大，处理可能需要较长时间",
}


class ConfigManager:
    """配置管理器"""

    def __init__(self):
        self._config = DEFAULT_CONFIG.copy()
        self._config_file = self._get_config_file_path()
        self._load_config()

    def _get_config_file_path(self) -> Path:
        """获取配置文件路径"""
        # 优先使用用户目录下的配置文件
        user_config_dir = Path.home() / ".excel_comparison_tool"
        user_config_dir.mkdir(exist_ok=True)
        return user_config_dir / "config.json"

    def _load_config(self) -> None:
        """加载配置文件"""
        try:
            if self._config_file.exists():
                with open(self._config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    self._config.update(user_config)
        except Exception as e:
            logging.warning(f"加载配置文件失败: {e}，使用默认配置")

    def save_config(self) -> bool:
        """保存配置到文件"""
        try:
            with open(self._config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            logging.error(f"保存配置文件失败: {e}")
            return False

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值"""
        return self._config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """设置配置值"""
        self._config[key] = value

    def reset_to_default(self) -> None:
        """重置为默认配置"""
        self._config = DEFAULT_CONFIG.copy()


# 全局配置实例
config = ConfigManager()

# 向后兼容的常量
EXCEL_FILE_FILTERS: Final[str] = config.get("excel_file_filters")
WINDOW_MIN_WIDTH: Final[int] = config.get("window_min_width")
WINDOW_MIN_HEIGHT: Final[int] = config.get("window_min_height")
TIMESTAMP_FORMAT: Final[str] = config.get("timestamp_format")

"""自定义异常类"""


class ExcelComparisonError(Exception):
    """Excel比对工具基础异常类"""
    
    def __init__(self, message: str, error_code: str = None):
        super().__init__(message)
        self.message = message
        self.error_code = error_code
    
    def __str__(self) -> str:
        if self.error_code:
            return f"[{self.error_code}] {self.message}"
        return self.message


class DataValidationError(ExcelComparisonError):
    """数据验证异常"""
    
    def __init__(self, message: str, field: str = None):
        super().__init__(message, "DATA_VALIDATION")
        self.field = field


class FileProcessingError(ExcelComparisonError):
    """文件处理异常"""
    
    def __init__(self, message: str, file_path: str = None):
        super().__init__(message, "FILE_PROCESSING")
        self.file_path = file_path


class FilterOperationError(ExcelComparisonError):
    """筛选操作异常"""
    
    def __init__(self, message: str, filter_condition: str = None):
        super().__init__(message, "FILTER_OPERATION")
        self.filter_condition = filter_condition


class ExportError(ExcelComparisonError):
    """导出异常"""
    
    def __init__(self, message: str, export_path: str = None):
        super().__init__(message, "EXPORT_ERROR")
        self.export_path = export_path


class ConfigurationError(ExcelComparisonError):
    """配置异常"""
    
    def __init__(self, message: str, config_key: str = None):
        super().__init__(message, "CONFIGURATION")
        self.config_key = config_key

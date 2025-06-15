"""数据验证工具"""
import os
from pathlib import Path
from typing import List, Optional, Tuple, Any
import pandas as pd
from .exceptions import DataValidationError, FileProcessingError
from ..config import config


class DataValidator:
    """数据验证器"""
    
    @staticmethod
    def validate_file_path(file_path: str) -> Tuple[bool, Optional[str]]:
        """验证文件路径
        
        Args:
            file_path: 文件路径
            
        Returns:
            tuple: (is_valid: bool, error_message: Optional[str])
        """
        if not file_path:
            return False, "文件路径不能为空"
        
        path = Path(file_path)
        
        if not path.exists():
            return False, f"文件不存在: {file_path}"
        
        if not path.is_file():
            return False, f"路径不是文件: {file_path}"
        
        # 检查文件大小
        max_size_mb = config.get("max_file_size_mb", 100)
        file_size_mb = path.stat().st_size / (1024 * 1024)
        if file_size_mb > max_size_mb:
            return False, f"文件过大 ({file_size_mb:.1f}MB)，最大支持 {max_size_mb}MB"
        
        # 检查文件扩展名
        allowed_extensions = {'.xlsx', '.xls', '.xlsm'}
        if path.suffix.lower() not in allowed_extensions:
            return False, f"不支持的文件格式: {path.suffix}"
        
        return True, None
    
    @staticmethod
    def validate_excel_data(df: pd.DataFrame) -> Tuple[bool, Optional[str]]:
        """验证Excel数据
        
        Args:
            df: pandas DataFrame
            
        Returns:
            tuple: (is_valid: bool, error_message: Optional[str])
        """
        if df is None:
            return False, "数据为空"
        
        if df.empty:
            return False, "Excel文件不包含数据"
        
        if len(df.columns) == 0:
            return False, "Excel文件不包含列"
        
        # 检查列名是否有效
        invalid_columns = []
        for col in df.columns:
            if pd.isna(col) or str(col).strip() == "":
                invalid_columns.append(col)
        
        if invalid_columns:
            return False, f"存在无效的列名: {invalid_columns}"
        
        return True, None
    
    @staticmethod
    def validate_column_selection(columns: List[str], available_columns: List[str]) -> Tuple[bool, Optional[str]]:
        """验证列选择
        
        Args:
            columns: 选择的列名列表
            available_columns: 可用的列名列表
            
        Returns:
            tuple: (is_valid: bool, error_message: Optional[str])
        """
        if not columns:
            return False, "必须至少选择一列"
        
        invalid_columns = [col for col in columns if col not in available_columns]
        if invalid_columns:
            return False, f"选择的列不存在: {invalid_columns}"
        
        return True, None
    
    @staticmethod
    def validate_filter_text(filter_text: str) -> Tuple[bool, Optional[str]]:
        """验证筛选文本
        
        Args:
            filter_text: 筛选文本
            
        Returns:
            tuple: (is_valid: bool, error_message: Optional[str])
        """
        if not filter_text or not filter_text.strip():
            return False, "筛选条件不能为空"
        
        # 检查长度限制
        if len(filter_text) > 1000:
            return False, "筛选条件过长，最多支持1000个字符"
        
        return True, None
    
    @staticmethod
    def validate_export_path(export_path: str) -> Tuple[bool, Optional[str]]:
        """验证导出路径
        
        Args:
            export_path: 导出路径
            
        Returns:
            tuple: (is_valid: bool, error_message: Optional[str])
        """
        if not export_path:
            return False, "导出路径不能为空"
        
        path = Path(export_path)
        
        # 检查父目录是否存在且可写
        parent_dir = path.parent
        if not parent_dir.exists():
            return False, f"导出目录不存在: {parent_dir}"
        
        if not os.access(parent_dir, os.W_OK):
            return False, f"没有写入权限: {parent_dir}"
        
        # 检查文件是否被占用
        if path.exists():
            try:
                # 尝试打开文件检查是否被占用
                with open(path, 'a'):
                    pass
            except PermissionError:
                return False, f"文件被占用或没有写入权限: {path}"
        
        return True, None
    
    @staticmethod
    def sanitize_sheet_name(name: str) -> str:
        """清理工作表名称
        
        Args:
            name: 原始名称
            
        Returns:
            str: 清理后的名称
        """
        if not name:
            return "Sheet1"
        
        # Excel工作表名称不能包含的字符
        invalid_chars = ['/', '\\', '?', '*', '[', ']', ':']
        sanitized = name
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        # 限制长度
        max_length = config.get("export_sheet_name_max_length", 31)
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length-3] + "..."
        
        return sanitized.strip()
    
    @staticmethod
    def sanitize_filename(name: str) -> str:
        """清理文件名
        
        Args:
            name: 原始文件名
            
        Returns:
            str: 清理后的文件名
        """
        if not name:
            return "export"
        
        # 文件名不能包含的字符
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        sanitized = name
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        # 限制长度
        max_length = config.get("export_filename_max_length", 200)
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length-3] + "..."
        
        return sanitized.strip()

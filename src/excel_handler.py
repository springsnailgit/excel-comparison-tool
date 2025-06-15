from typing import Dict, List, Tuple, Union, Optional, Protocol, Any
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from .utils.logger import get_logger
from .utils.validators import DataValidator
from .utils.performance import monitor_performance, check_memory_usage, optimize_dataframe_memory
from .config import config


class FilterStrategy(Protocol):
    """筛选策略接口"""

    def apply_filter(self, df: pd.DataFrame, columns: List[str], condition: str) -> pd.Series:
        """应用筛选条件"""
        ...


class ContainsFilter:
    """包含筛选策略"""

    def apply_filter(self, df: pd.DataFrame, columns: List[str], condition: str) -> pd.Series:
        """应用包含筛选"""
        mask = pd.Series(False, index=df.index)
        for col in columns:
            if col in df.columns:
                mask |= df[col].astype(str).str.contains(
                    condition, na=False, case=False, regex=False
                )
        return mask


class ExactMatchFilter:
    """精确匹配筛选策略"""

    def apply_filter(self, df: pd.DataFrame, columns: List[str], condition: str) -> pd.Series:
        """应用精确匹配筛选"""
        mask = pd.Series(False, index=df.index)
        for col in columns:
            if col in df.columns:
                mask |= df[col].astype(str).str.strip().str.lower() == condition.strip().lower()
        return mask


class RegexFilter:
    """正则表达式筛选策略"""

    def apply_filter(self, df: pd.DataFrame, columns: List[str], condition: str) -> pd.Series:
        """应用正则表达式筛选"""
        mask = pd.Series(False, index=df.index)
        for col in columns:
            if col in df.columns:
                try:
                    mask |= df[col].astype(str).str.contains(
                        condition, na=False, case=False, regex=True
                    )
                except Exception:
                    # 如果正则表达式无效，回退到普通包含匹配
                    mask |= df[col].astype(str).str.contains(
                        condition, na=False, case=False, regex=False
                    )
        return mask


class ExcelHandler:
    """Excel文件处理类，负责数据的加载、筛选和导出"""

    def __init__(self):
        self.logger = get_logger(self.__class__.__name__)
        self.excel_file_path: Optional[str] = None
        self.dataframe: Optional[pd.DataFrame] = None
        self.original_dataframe: Optional[pd.DataFrame] = None  # 保存原始数据副本
        self.filtered_sheets: Dict[str, pd.DataFrame] = {}  # 保存筛选结果 {sheet_name: dataframe}

        # 筛选策略
        self.filter_strategies = {
            'contains': ContainsFilter(),
            'exact': ExactMatchFilter(),
            'regex': RegexFilter(),
        }
        self.current_filter_strategy = 'contains'

    @monitor_performance("load_excel")
    def load_excel(self, file_path: str) -> Tuple[bool, Union[List[str], str]]:
        """加载Excel文件

        Args:
            file_path: Excel文件路径

        Returns:
            tuple: (success: bool, result: list|str) 成功时返回列名列表，失败时返回错误信息
        """
        self.logger.info(f"开始加载Excel文件: {file_path}")

        try:
            # 验证文件路径
            is_valid, error_msg = DataValidator.validate_file_path(file_path)
            if not is_valid:
                self.logger.error(f"文件验证失败: {error_msg}")
                return False, error_msg

            # 加载Excel文件
            self.excel_file_path = file_path

            # 根据文件大小选择加载策略
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            chunk_size = config.get("chunk_size", 10000)

            if file_size_mb > 50:  # 大文件使用分块读取
                self.logger.info(f"检测到大文件 ({file_size_mb:.1f}MB)，使用分块读取")
                self.dataframe = self._load_large_excel(file_path, chunk_size)
            else:
                self.dataframe = pd.read_excel(file_path)

            # 验证数据
            is_valid, error_msg = DataValidator.validate_excel_data(self.dataframe)
            if not is_valid:
                self.logger.error(f"数据验证失败: {error_msg}")
                return False, error_msg

            # 优化内存使用
            self.dataframe = optimize_dataframe_memory(self.dataframe)

            # 保存原始数据副本
            self.original_dataframe = self.dataframe.copy()

            # 检查内存使用
            check_memory_usage()

            self.logger.info(f"成功加载Excel文件，共 {len(self.dataframe)} 行，{len(self.dataframe.columns)} 列")
            return True, list(self.dataframe.columns)

        except pd.errors.EmptyDataError:
            error_msg = "Excel文件为空或不包含数据"
            self.logger.error(error_msg)
            return False, error_msg
        except pd.errors.ParserError as e:
            error_msg = f"Excel文件格式错误，无法解析: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
        except FileNotFoundError:
            error_msg = f"找不到文件: {file_path}"
            self.logger.error(error_msg)
            return False, error_msg
        except PermissionError:
            error_msg = f"没有权限访问文件: {file_path}"
            self.logger.error(error_msg)
            return False, error_msg
        except MemoryError:
            error_msg = "文件过大，内存不足，请尝试处理较小的文件"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"加载Excel文件失败: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            return False, error_msg

    def _load_large_excel(self, file_path: str, chunk_size: int) -> pd.DataFrame:
        """加载大型Excel文件

        Args:
            file_path: 文件路径
            chunk_size: 分块大小

        Returns:
            pd.DataFrame: 加载的数据
        """
        try:
            # 对于大文件，我们仍然需要一次性加载，但可以进行一些优化
            # 例如只读取需要的列，或者使用更高效的引擎
            return pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            self.logger.error(f"加载大文件失败: {str(e)}")
            raise

    def get_column_names(self) -> List[str]:
        """获取列名列表"""
        return list(self.dataframe.columns) if self.dataframe is not None else []

    def set_filter_strategy(self, strategy: str) -> bool:
        """设置筛选策略

        Args:
            strategy: 筛选策略名称 ('contains', 'exact', 'regex')

        Returns:
            bool: 设置是否成功
        """
        if strategy in self.filter_strategies:
            self.current_filter_strategy = strategy
            self.logger.info(f"筛选策略已设置为: {strategy}")
            return True
        else:
            self.logger.warning(f"未知的筛选策略: {strategy}")
            return False

    @monitor_performance("filter_data")
    def filter_data(self, selected_columns: List[str], filter_value: str,
                   strategy: Optional[str] = None) -> Tuple[bool, Union[pd.DataFrame, str]]:
        """根据条件筛选数据

        Args:
            selected_columns: 要搜索的列名列表
            filter_value: 筛选条件
            strategy: 筛选策略，为None时使用当前策略

        Returns:
            tuple: (success: bool, result: DataFrame|str)
        """
        self.logger.info(f"开始筛选数据，条件: '{filter_value}', 列: {selected_columns}")

        if self.dataframe is None:
            error_msg = "没有加载数据"
            self.logger.error(error_msg)
            return False, error_msg

        # 验证列选择
        is_valid, error_msg = DataValidator.validate_column_selection(selected_columns, self.get_column_names())
        if not is_valid:
            self.logger.error(f"列选择验证失败: {error_msg}")
            return False, error_msg

        # 验证筛选条件
        is_valid, error_msg = DataValidator.validate_filter_text(filter_value)
        if not is_valid:
            self.logger.error(f"筛选条件验证失败: {error_msg}")
            return False, error_msg

        try:
            # 选择筛选策略
            filter_strategy = strategy or self.current_filter_strategy
            if filter_strategy not in self.filter_strategies:
                filter_strategy = 'contains'  # 默认策略

            # 创建筛选条件
            mask = self.filter_strategies[filter_strategy].apply_filter(
                self.dataframe, selected_columns, filter_value
            )
            filtered_data = self.dataframe[mask].copy()

            if filtered_data.empty:
                error_msg = f"没有找到匹配 '{filter_value}' 的数据"
                self.logger.info(error_msg)
                return False, error_msg

            # 生成安全的工作表名称
            sheet_name = DataValidator.sanitize_sheet_name(filter_value)

            # 保存筛选结果并从原数据中移除
            self.filtered_sheets[sheet_name] = filtered_data
            self.dataframe = self.dataframe[~mask].reset_index(drop=True)

            self.logger.info(f"筛选成功，找到 {len(filtered_data)} 行数据，剩余 {len(self.dataframe)} 行")
            return True, filtered_data

        except KeyError as e:
            error_msg = f"列名错误: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"筛选数据时出错: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            return False, error_msg
    
    def filter_data_batch(self, selected_columns: List[str], filter_values: List[str],
                         logic_operator: str = 'AND') -> Tuple[bool, Union[Dict[str, pd.DataFrame], str]]:
        """批量筛选数据

        Args:
            selected_columns: 要搜索的列名列表
            filter_values: 筛选条件列表
            logic_operator: 逻辑操作符 ('AND' 或 'OR')

        Returns:
            tuple: (success: bool, result: dict|str)
        """
        self.logger.info(f"开始批量筛选，条件数: {len(filter_values)}, 逻辑: {logic_operator}")

        if self.dataframe is None:
            error_msg = "没有加载数据"
            self.logger.error(error_msg)
            return False, error_msg

        # 验证列选择
        is_valid, error_msg = DataValidator.validate_column_selection(selected_columns, self.get_column_names())
        if not is_valid:
            self.logger.error(f"列选择验证失败: {error_msg}")
            return False, error_msg

        if not filter_values:
            error_msg = "未提供筛选条件"
            self.logger.error(error_msg)
            return False, error_msg

        # 检查筛选条件数量限制
        max_conditions = config.get("max_filter_conditions", 50)
        if len(filter_values) > max_conditions:
            error_msg = f"筛选条件过多，最多支持 {max_conditions} 个条件"
            self.logger.error(error_msg)
            return False, error_msg

        try:
            # 创建组合筛选条件
            if logic_operator.upper() == 'AND':
                final_mask = pd.Series(True, index=self.dataframe.index)
                for filter_value in filter_values:
                    current_mask = self.filter_strategies[self.current_filter_strategy].apply_filter(
                        self.dataframe, selected_columns, filter_value
                    )
                    final_mask &= current_mask
                condition_name = " 与 ".join(filter_values)
            else:  # OR
                final_mask = pd.Series(False, index=self.dataframe.index)
                for filter_value in filter_values:
                    current_mask = self.filter_strategies[self.current_filter_strategy].apply_filter(
                        self.dataframe, selected_columns, filter_value
                    )
                    final_mask |= current_mask
                condition_name = " 或 ".join(filter_values)

            filtered_data = self.dataframe[final_mask].copy()

            if filtered_data.empty:
                error_msg = f"没有找到满足筛选条件的数据"
                self.logger.info(error_msg)
                return False, error_msg

            # 生成安全的工作表名称
            sheet_name = DataValidator.sanitize_sheet_name(condition_name)
            self.filtered_sheets[sheet_name] = filtered_data
            self.dataframe = self.dataframe[~final_mask].reset_index(drop=True)

            self.logger.info(f"批量筛选成功，找到 {len(filtered_data)} 行数据")
            return True, {sheet_name: filtered_data}

        except KeyError as e:
            error_msg = f"列名错误: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"批量筛选数据时出错: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            return False, error_msg
    
    def get_filtered_data(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """获取指定筛选条件的数据"""
        return self.filtered_sheets.get(sheet_name)

    def get_all_filtered_sheets(self) -> List[str]:
        """获取所有筛选结果的名称列表"""
        return list(self.filtered_sheets.keys())

    def get_data_summary(self) -> Dict[str, Any]:
        """获取数据摘要信息"""
        summary = {
            'original_rows': len(self.original_dataframe) if self.original_dataframe is not None else 0,
            'current_rows': len(self.dataframe) if self.dataframe is not None else 0,
            'filtered_sheets_count': len(self.filtered_sheets),
            'total_filtered_rows': sum(len(df) for df in self.filtered_sheets.values()),
            'columns': self.get_column_names(),
            'file_path': self.excel_file_path,
        }
        return summary

    def reset_data(self) -> bool:
        """重置数据到原始状态"""
        try:
            if self.original_dataframe is not None:
                self.dataframe = self.original_dataframe.copy()
                self.filtered_sheets.clear()
                self.logger.info("数据已重置到原始状态")
                return True
            else:
                self.logger.warning("没有原始数据可重置")
                return False
        except Exception as e:
            self.logger.error(f"重置数据失败: {str(e)}")
            return False

    def clear_filtered_data(self, sheet_name: Optional[str] = None) -> bool:
        """清除筛选数据

        Args:
            sheet_name: 要清除的工作表名称，为None时清除所有

        Returns:
            bool: 操作是否成功
        """
        try:
            if sheet_name is None:
                self.filtered_sheets.clear()
                self.logger.info("已清除所有筛选数据")
            elif sheet_name in self.filtered_sheets:
                del self.filtered_sheets[sheet_name]
                self.logger.info(f"已清除筛选数据: {sheet_name}")
            else:
                self.logger.warning(f"未找到筛选数据: {sheet_name}")
                return False
            return True
        except Exception as e:
            self.logger.error(f"清除筛选数据失败: {str(e)}")
            return False
    
    @monitor_performance("export_excel")
    def export_final_excel(self, save_directory: Optional[str] = None,
                          filename: Optional[str] = None) -> Tuple[bool, str]:
        """导出包含所有筛选结果的Excel文件

        Args:
            save_directory: 保存目录，为None时使用原文件目录
            filename: 自定义文件名，为None时自动生成

        Returns:
            tuple: (success: bool, result: str) 成功时返回文件路径，失败时返回错误信息
        """
        self.logger.info("开始导出Excel文件")

        if not self.filtered_sheets:
            error_msg = "没有筛选数据可导出"
            self.logger.error(error_msg)
            return False, error_msg

        if self.excel_file_path is None:
            error_msg = "未加载原始Excel文件"
            self.logger.error(error_msg)
            return False, error_msg

        try:
            # 生成文件名
            if filename is None:
                filename = self._generate_export_filename()
            else:
                filename = DataValidator.sanitize_filename(filename)
                if not filename.endswith('.xlsx'):
                    filename += '.xlsx'

            # 确定保存路径
            if save_directory and os.path.isdir(save_directory):
                file_path = os.path.join(save_directory, filename)
            else:
                file_path = os.path.join(os.path.dirname(self.excel_file_path), filename)

            # 验证导出路径
            is_valid, error_msg = DataValidator.validate_export_path(file_path)
            if not is_valid:
                self.logger.error(f"导出路径验证失败: {error_msg}")
                return False, error_msg

            # 创建新工作簿
            workbook = Workbook()
            workbook.remove(workbook.active)  # 删除默认工作表

            # 添加所有筛选结果
            for sheet_name, df in self.filtered_sheets.items():
                safe_sheet_name = DataValidator.sanitize_sheet_name(sheet_name)
                ws = workbook.create_sheet(safe_sheet_name)

                # 优化大数据写入
                if len(df) > config.get("table_max_display_rows", 1000):
                    self.logger.info(f"工作表 '{safe_sheet_name}' 包含大量数据 ({len(df)} 行)，正在写入...")

                for row in dataframe_to_rows(df, index=False, header=True):
                    ws.append(row)

            # 保存文件
            workbook.save(file_path)

            self.logger.info(f"Excel文件导出成功: {file_path}")
            return True, file_path

        except PermissionError:
            error_msg = "没有权限保存文件，请检查文件是否被其他程序占用"
            self.logger.error(error_msg)
            return False, error_msg
        except MemoryError:
            error_msg = "内存不足，无法导出大文件"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"导出Excel时出错: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            return False, error_msg
    
    def _generate_export_filename(self) -> str:
        """生成导出文件名的辅助方法

        Returns:
            str: 生成的文件名
        """
        timestamp = datetime.now().strftime(config.get("timestamp_format", "%Y%m%d_%H%M%S"))

        # 提取所有唯一的筛选条件
        all_conditions = []
        for sheet_name in self.filtered_sheets.keys():
            # 处理不同的连接符
            separators = [" 与 ", " 或 ", " AND ", " OR "]
            conditions = [sheet_name]

            for sep in separators:
                new_conditions = []
                for cond in conditions:
                    if sep in cond:
                        new_conditions.extend(cond.split(sep))
                    else:
                        new_conditions.append(cond)
                conditions = new_conditions

            all_conditions.extend(conditions)

        # 去重并限制数量
        unique_conditions = list(dict.fromkeys(all_conditions))  # 保持顺序的去重

        # 限制文件名长度和条件数量
        max_conditions = 3
        if len(unique_conditions) > max_conditions:
            conditions_str = "+".join(unique_conditions[:max_conditions]) + "等"
        else:
            conditions_str = "+".join(unique_conditions)

        # 使用验证器清理文件名
        conditions_str = DataValidator.sanitize_filename(conditions_str)

        # 限制文件名总长度
        max_length = config.get("export_filename_max_length", 200) - len(timestamp) - 7  # 7 for "_" and ".xlsx"
        if len(conditions_str) > max_length:
            conditions_str = conditions_str[:max_length-3] + "..."

        return f"{conditions_str}_{timestamp}.xlsx"

    def get_available_filter_strategies(self) -> List[str]:
        """获取可用的筛选策略列表"""
        return list(self.filter_strategies.keys())

    def get_current_filter_strategy(self) -> str:
        """获取当前筛选策略"""
        return self.current_filter_strategy

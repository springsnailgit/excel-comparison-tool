from typing import Dict, List, Tuple, Union, Optional
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelHandler:
    """Excel文件处理类，负责数据的加载、筛选和导出"""
    
    def __init__(self):
        self.excel_file_path: Optional[str] = None
        self.dataframe: Optional[pd.DataFrame] = None
        self.filtered_sheets: Dict[str, pd.DataFrame] = {}  # 保存筛选结果 {sheet_name: dataframe}

    def load_excel(self, file_path: str) -> Tuple[bool, Union[List[str], str]]:
        """加载Excel文件
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            tuple: (success: bool, result: list|str) 成功时返回列名列表，失败时返回错误信息
        """
        try:
            self.excel_file_path = file_path
            self.dataframe = pd.read_excel(file_path)
            return True, list(self.dataframe.columns)
        except pd.errors.EmptyDataError:
            return False, "Excel文件为空或不包含数据"
        except pd.errors.ParserError:
            return False, "Excel文件格式错误，无法解析"
        except FileNotFoundError:
            return False, f"找不到文件: {file_path}"
        except PermissionError:
            return False, f"没有权限访问文件: {file_path}"
        except Exception as e:
            return False, f"加载Excel文件失败: {str(e)}"

    def get_column_names(self) -> List[str]:
        """获取列名列表"""
        return list(self.dataframe.columns) if self.dataframe is not None else []

    def filter_data(self, selected_columns: List[str], filter_value: str) -> Tuple[bool, Union[pd.DataFrame, str]]:
        """根据条件筛选数据
        
        Args:
            selected_columns: 要搜索的列名列表
            filter_value: 筛选条件
            
        Returns:
            tuple: (success: bool, result: DataFrame|str)
        """
        if self.dataframe is None or not selected_columns:
            return False, "没有加载数据或未选择列"
        
        try:
            # 创建筛选条件：任一选定列包含筛选值
            mask = self._create_filter_mask(selected_columns, filter_value)
            filtered_data = self.dataframe[mask].copy()
            
            if filtered_data.empty:
                return False, f"没有找到匹配 '{filter_value}' 的数据"
            
            # 保存筛选结果并从原数据中移除
            self.filtered_sheets[filter_value] = filtered_data
            self.dataframe = self.dataframe[~mask].reset_index(drop=True)
            
            return True, filtered_data
        except KeyError as e:
            return False, f"列名错误: {str(e)}"
        except Exception as e:
            return False, f"筛选数据时出错: {str(e)}"
    
    def filter_data_batch(self, selected_columns: List[str], filter_values: List[str]) -> Tuple[bool, Union[Dict[str, pd.DataFrame], str]]:
        """批量筛选数据，多个条件之间是AND关系
        
        Args:
            selected_columns: 要搜索的列名列表
            filter_values: 筛选条件列表
            
        Returns:
            tuple: (success: bool, result: dict|str)
        """
        if self.dataframe is None or not selected_columns:
            return False, "没有加载数据或未选择列"
        
        if not filter_values:
            return False, "未提供筛选条件"
        
        try:
            # 创建组合筛选条件：所有条件都必须满足
            final_mask = pd.Series(True, index=self.dataframe.index)
            
            for filter_value in filter_values:
                current_mask = self._create_filter_mask(selected_columns, filter_value)
                final_mask &= current_mask
            
            filtered_data = self.dataframe[final_mask].copy()
            
            if filtered_data.empty:
                return False, "没有找到同时满足所有筛选条件的数据"
            
            # 使用"与"连接条件名称
            condition_name = " 与 ".join(filter_values)
            self.filtered_sheets[condition_name] = filtered_data
            self.dataframe = self.dataframe[~final_mask].reset_index(drop=True)
            
            return True, {condition_name: filtered_data}
        except KeyError as e:
            return False, f"列名错误: {str(e)}"
        except Exception as e:
            return False, f"批量筛选数据时出错: {str(e)}"
    
    def _create_filter_mask(self, selected_columns: List[str], filter_value: str) -> pd.Series:
        """创建筛选掩码的辅助方法
        
        Args:
            selected_columns: 要搜索的列名列表
            filter_value: 筛选条件
            
        Returns:
            pd.Series: 布尔掩码，表示每行是否匹配筛选条件
        """
        if self.dataframe is None:
            raise ValueError("没有加载数据")
            
        mask = pd.Series(False, index=self.dataframe.index)
        for col in selected_columns:
            if col in self.dataframe.columns:
                mask |= self.dataframe[col].astype(str).str.contains(
                    filter_value, na=False, case=False, regex=False
                )
        return mask
    
    def get_filtered_data(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """获取指定筛选条件的数据"""
        return self.filtered_sheets.get(sheet_name)
    
    def get_all_filtered_sheets(self) -> List[str]:
        """获取所有筛选结果的名称列表"""
        return list(self.filtered_sheets.keys())
    
    def export_final_excel(self, save_directory: Optional[str] = None) -> Tuple[bool, str]:
        """导出包含所有筛选结果的Excel文件
        
        Args:
            save_directory: 保存目录，为None时使用原文件目录
            
        Returns:
            tuple: (success: bool, result: str) 成功时返回文件路径，失败时返回错误信息
        """
        if not self.filtered_sheets:
            return False, "没有筛选数据可导出"
        
        if self.excel_file_path is None:
            return False, "未加载原始Excel文件"
        
        try:
            # 创建新工作簿
            workbook = Workbook()
            workbook.remove(workbook.active)  # 删除默认工作表
            
            # 添加所有筛选结果
            for sheet_name, df in self.filtered_sheets.items():
                # 工作表名称长度限制为31个字符
                safe_sheet_name = sheet_name[:31]
                ws = workbook.create_sheet(safe_sheet_name)
                for row in dataframe_to_rows(df, index=False, header=True):
                    ws.append(row)
            
            # 生成文件名
            file_name = self._generate_export_filename()
            
            # 确定保存路径
            if save_directory and os.path.isdir(save_directory):
                file_path = os.path.join(save_directory, file_name)
            else:
                file_path = os.path.join(os.path.dirname(self.excel_file_path), file_name)
            
            workbook.save(file_path)
            return True, file_path
        except PermissionError:
            return False, "没有权限保存文件，请检查文件是否被其他程序占用"
        except Exception as e:
            return False, f"导出Excel时出错: {str(e)}"
    
    def _generate_export_filename(self) -> str:
        """生成导出文件名的辅助方法
        
        Returns:
            str: 生成的文件名
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 提取所有唯一的筛选条件
        all_conditions = []
        for sheet_name in self.filtered_sheets.keys():
            if " 与 " in sheet_name:
                all_conditions.extend(sheet_name.split(" 与 "))
            else:
                all_conditions.append(sheet_name)
        
        unique_conditions = list(set(all_conditions))
        
        # 限制文件名长度
        if len(unique_conditions) > 3:
            conditions_str = "+".join(unique_conditions[:3]) + "等"
        else:
            conditions_str = "+".join(unique_conditions)
        
        # 确保文件名不包含非法字符
        conditions_str = conditions_str.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_")
        conditions_str = conditions_str.replace("?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")
        
        # 限制文件名总长度
        max_length = 200 - len(timestamp) - 7  # 7 for "_" and ".xlsx"
        if len(conditions_str) > max_length:
            conditions_str = conditions_str[:max_length-3] + "..."
        
        return f"{conditions_str}_{timestamp}.xlsx"

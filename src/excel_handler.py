import pandas as pd
import os
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelHandler:
    def __init__(self):
        self.excel_file_path = None
        self.workbook = None
        self.dataframe = None
        self.sheet_name = None
        self.filtered_sheets = {}  # 保存每次筛选后的数据，格式 {sheet_name: dataframe}

    def load_excel(self, file_path):
        """导入Excel文件"""
        try:
            self.excel_file_path = file_path
            self.dataframe = pd.read_excel(file_path)
            self.workbook = openpyxl.load_workbook(file_path)
            self.sheet_name = self.workbook.sheetnames[0]  # 默认使用第一个工作表
            return True, self.get_column_names()
        except Exception as e:
            return False, str(e)

    def get_column_names(self):
        """获取数据框的列名"""
        if self.dataframe is not None:
            return list(self.dataframe.columns)
        return []

    def filter_data(self, selected_columns, filter_value):
        """根据选择的列和筛选条件过滤数据"""
        if self.dataframe is None or not selected_columns:
            return False, "没有加载数据或未选择列"
        
        try:
            # 创建复合条件
            mask = pd.Series(False, index=self.dataframe.index)
            for col in selected_columns:
                # 将筛选内容与列内容进行比较，只要有一个列符合条件即可
                mask |= self.dataframe[col].astype(str).str.contains(filter_value, na=False)
            
            filtered_data = self.dataframe[mask].copy()
            
            if filtered_data.empty:
                return False, f"没有找到匹配 '{filter_value}' 的数据"
            
            # 保存筛选结果
            self.filtered_sheets[filter_value] = filtered_data
            
            # 从原数据中删除筛选出的行
            self.dataframe = self.dataframe[~mask].reset_index(drop=True)
            
            return True, filtered_data
        except Exception as e:
            return False, f"筛选数据时出错: {str(e)}"
    
    def filter_data_batch(self, selected_columns, filter_values):
        """根据选择的列和多个筛选条件批量过滤数据，条件之间是"逻辑与"关系
        
        Args:
            selected_columns: 选择的列名列表
            filter_values: 筛选条件列表
            
        Returns:
            (success, result_dict): 是否成功和筛选结果字典 {筛选条件组合名: dataframe}
        """
        if self.dataframe is None or not selected_columns:
            return False, "没有加载数据或未选择列"
        
        try:
            # 创建初始掩码，全部为True
            final_mask = pd.Series(True, index=self.dataframe.index)
            
            # 对每个筛选条件应用逻辑与操作
            for filter_value in filter_values:
                # 针对当前筛选条件创建掩码
                current_mask = pd.Series(False, index=self.dataframe.index)
                for col in selected_columns:
                    # 对于每个列，只要有一个匹配就算这个条件满足
                    current_mask |= self.dataframe[col].astype(str).str.contains(filter_value, na=False)
                
                # 将当前条件的掩码与最终掩码进行逻辑与操作
                final_mask &= current_mask
            
            # 获取满足所有条件的数据
            filtered_data = self.dataframe[final_mask].copy()
            
            if filtered_data.empty:
                return False, f"没有找到同时满足所有筛选条件的数据"
            
            # 条件名称使用"与"连接所有条件
            condition_name = " 与 ".join(filter_values)
            
            # 保存筛选结果
            self.filtered_sheets[condition_name] = filtered_data
            
            # 从原数据中删除筛选出的行
            self.dataframe = self.dataframe[~final_mask].reset_index(drop=True)
            
            # 返回结果字典
            results = {condition_name: filtered_data}
            return True, results
                
        except Exception as e:
            return False, f"批量筛选数据时出错: {str(e)}"
    
    def get_filtered_data(self, sheet_name):
        """获取指定筛选条件的数据"""
        return self.filtered_sheets.get(sheet_name, None)
    
    def get_all_filtered_sheets(self):
        """获取所有筛选后的工作表名称"""
        return list(self.filtered_sheets.keys())
    
    def export_final_excel(self, save_directory=None):
        """导出最终的Excel文件，包含所有筛选后的工作表
        
        Args:
            save_directory: 指定保存目录，如果为None则使用原始文件的目录
        """
        if not self.filtered_sheets:
            return False, "没有筛选数据可导出"
        
        try:
            # 创建新的工作簿
            new_workbook = Workbook()
            # 删除默认的工作表
            new_workbook.remove(new_workbook.active)
            
            # 获取所有筛选条件，用于生成文件名
            all_conditions = []
            
            # 将每个筛选结果添加到新的工作表
            for sheet_name, df in self.filtered_sheets.items():
                ws = new_workbook.create_sheet(sheet_name)
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # 收集所有筛选条件（工作表名称）
                if " 与 " in sheet_name:
                    # 如果工作表名称中包含"与"，说明是多条件筛选
                    all_conditions.extend(sheet_name.split(" 与 "))
                else:
                    all_conditions.append(sheet_name)
            
            # 去重筛选条件
            unique_conditions = list(set(all_conditions))
            
            # 生成新文件名
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            
            # 如果筛选条件太多，只取前3个
            if len(unique_conditions) > 3:
                filename_conditions = "+".join(unique_conditions[:3]) + "等"
            else:
                filename_conditions = "+".join(unique_conditions)
            
            file_name = f"{filename_conditions}_{timestamp}.xlsx"
            
            # 确定保存路径
            if save_directory and os.path.isdir(save_directory):
                new_file_path = os.path.join(save_directory, file_name)
            else:
                new_file_path = os.path.join(
                    os.path.dirname(self.excel_file_path),
                    file_name
                )
            
            # 保存工作簿
            new_workbook.save(new_file_path)
            return True, new_file_path
        except Exception as e:
            return False, f"导出Excel时出错: {str(e)}"
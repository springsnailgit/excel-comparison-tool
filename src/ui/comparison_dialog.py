from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QTableView, QMessageBox, QApplication, QCheckBox
)
from PyQt6.QtCore import Qt, QAbstractTableModel
import pandas as pd
import sys


class PandasTableModel(QAbstractTableModel):
    """用于在QTableView中显示pandas DataFrame的模型"""
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section])
        return None


class ComparisonDialog(QDialog):
    def __init__(self, parent, selected_columns, excel_handler):
        super().__init__(parent)
        
        self.selected_columns = selected_columns
        self.excel_handler = excel_handler
        self.all_columns = self.excel_handler.get_column_names()
        
        # 设置对话框属性
        self.setWindowTitle("数据比对")
        self.setMinimumSize(800, 600)
        
        # 创建布局
        layout = QVBoxLayout(self)
        
        # 添加选中列信息和编辑按钮
        columns_layout = QHBoxLayout()
        self.columns_info_label = QLabel(f"已选择列: {', '.join(selected_columns)}")
        self.edit_columns_button = QPushButton("更改列选择")
        self.edit_columns_button.clicked.connect(self.edit_selected_columns)
        
        columns_layout.addWidget(self.columns_info_label, 1)
        columns_layout.addWidget(self.edit_columns_button)
        layout.addLayout(columns_layout)
        
        # 添加比对内容输入区域
        input_layout = QHBoxLayout()
        input_label = QLabel("比对内容:")
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("请输入要比对的内容")
        
        input_layout.addWidget(input_label)
        input_layout.addWidget(self.input_edit, 1)
        
        # 添加预览和筛选按钮
        self.preview_button = QPushButton("预览数据")
        self.preview_button.clicked.connect(self.preview_data)
        
        filter_button = QPushButton("执行筛选")
        filter_button.clicked.connect(self.filter_data)
        
        input_layout.addWidget(self.preview_button)
        input_layout.addWidget(filter_button)
        
        layout.addLayout(input_layout)
        
        # 添加数据表格预览
        preview_label = QLabel("筛选结果预览:")
        layout.addWidget(preview_label)
        
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        layout.addWidget(self.table_view, 1)
        
        # 添加当前筛选状态提示
        self.status_label = QLabel("当前状态: 请输入第一个筛选条件")
        layout.addWidget(self.status_label)
        
        # 添加底部按钮区域
        button_layout = QHBoxLayout()
        
        self.continue_button = QPushButton("继续比对")
        self.continue_button.clicked.connect(self.continue_comparison)
        
        self.finish_button = QPushButton("完成")
        self.finish_button.clicked.connect(self.accept)
        
        button_layout.addWidget(self.continue_button)
        button_layout.addWidget(self.finish_button)
        
        layout.addLayout(button_layout)
        
        # 初始禁用继续按钮和完成按钮
        self.continue_button.setEnabled(False)
        self.finish_button.setEnabled(False)
        
        # 存储当前的预览数据和已应用的筛选条件
        self.preview_dataframe = None
        self.current_filter = None
        self.applied_filters = []
    
    def preview_data(self):
        """预览当前输入条件匹配的数据"""
        filter_value = self.input_edit.text().strip()
        
        if not filter_value:
            QMessageBox.warning(self, "警告", "请输入要比对的内容")
            return
        
        try:
            # 如果已经有预览数据（上一次筛选条件的结果），对其进行筛选
            if self.preview_dataframe is not None:
                # 对当前预览数据应用新的筛选条件
                mask = pd.Series(False, index=self.preview_dataframe.index)
                for col in self.selected_columns:
                    if col in self.preview_dataframe.columns:  # 确保列存在
                        mask |= self.preview_dataframe[col].astype(str).str.contains(filter_value, na=False)
                
                filtered_data = self.preview_dataframe[mask].copy()
                
                if filtered_data.empty:
                    QMessageBox.warning(self, "警告", f"在当前预览数据中没有找到匹配 '{filter_value}' 的数据")
                    return
                
                # 更新预览数据
                self.preview_dataframe = filtered_data
                self.current_filter = filter_value
                
                # 更新状态标签
                conditions = [f"'{f}'" for f in self.applied_filters]
                conditions.append(f"'{filter_value}'")
                self.status_label.setText(f"当前状态: 同时满足条件 {' 和 '.join(conditions)} 的数据")
                
            else:
                # 首次筛选，直接从原始数据中筛选
                mask = pd.Series(False, index=self.excel_handler.dataframe.index)
                for col in self.selected_columns:
                    mask |= self.excel_handler.dataframe[col].astype(str).str.contains(filter_value, na=False)
                
                filtered_data = self.excel_handler.dataframe[mask].copy()
                
                if filtered_data.empty:
                    QMessageBox.warning(self, "警告", f"没有找到匹配 '{filter_value}' 的数据")
                    return
                
                # 保存预览数据
                self.preview_dataframe = filtered_data
                self.current_filter = filter_value
                
                # 更新状态标签
                self.status_label.setText(f"当前状态: 满足条件 '{filter_value}' 的数据")
            
            # 显示预览数据
            self.display_filtered_data(self.preview_dataframe)
            
            # 清空输入框，准备输入新条件
            self.input_edit.clear()
            
            # 显示成功消息
            QMessageBox.information(self, "预览成功", f"找到匹配的数据 {len(self.preview_dataframe)} 行")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览数据时出错: {str(e)}")
    
    def filter_data(self):
        """执行筛选，将预览数据应用为最终筛选结果"""
        if self.preview_dataframe is None:
            QMessageBox.warning(self, "警告", "请先预览数据")
            return
        
        try:
            # 从原始数据中移除预览数据匹配的行
            # 创建索引的掩码
            matching_indices = []
            for idx, row in self.preview_dataframe.iterrows():
                for orig_idx, orig_row in self.excel_handler.dataframe.iterrows():
                    if all(row[col] == orig_row[col] for col in self.preview_dataframe.columns if col in self.excel_handler.dataframe.columns):
                        matching_indices.append(orig_idx)
                        break
            
            # 创建掩码
            mask = self.excel_handler.dataframe.index.isin(matching_indices)
            
            # 筛选数据
            filtered_data = self.excel_handler.dataframe[mask].copy()
            
            if not filtered_data.empty:
                # 给这次的结果一个有意义的名称（用所有已应用的筛选条件组合）
                if self.applied_filters:
                    all_filters = self.applied_filters.copy()
                    if self.current_filter:
                        all_filters.append(self.current_filter)
                    condition_name = " 与 ".join(all_filters)
                else:
                    condition_name = self.current_filter or "筛选结果"
                
                # 保存筛选结果
                self.excel_handler.filtered_sheets[condition_name] = filtered_data
                
                # 从原始数据中删除筛选出的行
                self.excel_handler.dataframe = self.excel_handler.dataframe[~mask].reset_index(drop=True)
                
                # 重置预览数据和状态
                self.preview_dataframe = None
                self.applied_filters = []
                self.current_filter = None
                self.status_label.setText("当前状态: 请输入第一个筛选条件")
                
                # 启用继续按钮和完成按钮
                self.continue_button.setEnabled(True)
                self.finish_button.setEnabled(True)
                
                # 清空表格预览
                self.table_view.setModel(None)
                
                # 显示成功消息
                QMessageBox.information(self, "筛选成功", f"已成功将筛选结果保存为 '{condition_name}'，共 {len(filtered_data)} 行")
            else:
                QMessageBox.warning(self, "警告", "没有找到符合条件的数据")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"筛选数据时出错: {str(e)}")
    
    def display_filtered_data(self, dataframe):
        """在表格中显示筛选结果"""
        model = PandasTableModel(dataframe)
        self.table_view.setModel(model)
        
        # 自动调整列宽
        for i in range(len(dataframe.columns)):
            self.table_view.resizeColumnToContents(i)
    
    def continue_comparison(self):
        """清空当前输入，准备下一次比对"""
        self.input_edit.clear()
        self.table_view.setModel(None)
        self.continue_button.setEnabled(False)
        self.finish_button.setEnabled(False)
        self.status_label.setText("当前状态: 请输入第一个筛选条件")
        self.preview_dataframe = None
        self.applied_filters = []
        self.current_filter = None
    
    def edit_selected_columns(self):
        """更改列选择，同时保存当前的筛选条件"""
        # 如果当前有过滤条件，保存它
        if self.current_filter and self.current_filter not in self.applied_filters:
            self.applied_filters.append(self.current_filter)
            self.current_filter = None
        
        # 打开列选择对话框
        class ColumnSelectionDialog(QDialog):
            def __init__(self, parent, all_columns, selected_columns):
                super().__init__(parent)
                self.setWindowTitle("选择列")
                self.setMinimumSize(400, 300)
                
                self.selected_columns = selected_columns
                self.checkboxes = []
                
                layout = QVBoxLayout(self)
                
                for column in all_columns:
                    checkbox = QCheckBox(column)
                    checkbox.setChecked(column in selected_columns)
                    self.checkboxes.append(checkbox)
                    layout.addWidget(checkbox)
                
                button_layout = QHBoxLayout()
                ok_button = QPushButton("确定")
                ok_button.clicked.connect(self.accept)
                cancel_button = QPushButton("取消")
                cancel_button.clicked.connect(self.reject)
                
                button_layout.addWidget(ok_button)
                button_layout.addWidget(cancel_button)
                layout.addLayout(button_layout)
            
            def get_selected_columns(self):
                return [cb.text() for cb in self.checkboxes if cb.isChecked()]
        
        dialog = ColumnSelectionDialog(self, self.all_columns, self.selected_columns)
        if dialog.exec():
            self.selected_columns = dialog.get_selected_columns()
            self.columns_info_label.setText(f"已选择列: {', '.join(self.selected_columns)}")
            
            # 清空输入框
            self.input_edit.clear()
            self.input_edit.setFocus()
            
            # 更新状态标签
            if self.applied_filters:
                conditions = [f"'{f}'" for f in self.applied_filters]
                self.status_label.setText(f"当前状态: 已应用条件 {' 和 '.join(conditions)}，请输入下一个条件")
            else:
                self.status_label.setText("当前状态: 请输入第一个筛选条件")


# 测试代码，如果直接运行此文件
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 这里仅作示例，实际使用时需要传入有效的参数
    dialog = ComparisonDialog(None, ["Column1", "Column2"], None)
    dialog.exec()
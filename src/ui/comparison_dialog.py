from typing import List, Dict, Tuple, Union, Optional, Any
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QTableView, QMessageBox, QCheckBox, QApplication
)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex
import pandas as pd
import sys
from ..excel_handler import ExcelHandler
from ..config import MESSAGES


class PandasTableModel(QAbstractTableModel):
    """用于在QTableView中显示pandas DataFrame的模型"""
    
    def __init__(self, data: pd.DataFrame):
        super().__init__()
        self._data = data

    def rowCount(self, parent: Optional[QModelIndex] = None) -> int:
        return len(self._data)

    def columnCount(self, parent: Optional[QModelIndex] = None) -> int:
        return len(self._data.columns)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section])
        return None


class ColumnSelectionDialog(QDialog):
    """列选择对话框"""
    
    def __init__(self, parent: QDialog, all_columns: List[str], selected_columns: List[str]):
        super().__init__(parent)
        self.setWindowTitle("选择列")
        self.setMinimumSize(400, 300)
        
        self.checkboxes: List[QCheckBox] = []
        layout = QVBoxLayout(self)
        
        # 创建复选框
        for column in all_columns:
            checkbox = QCheckBox(column)
            checkbox.setChecked(column in selected_columns)
            self.checkboxes.append(checkbox)
            layout.addWidget(checkbox)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        ok_button = QPushButton("确定")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def get_selected_columns(self) -> List[str]:
        """获取选中的列"""
        return [cb.text() for cb in self.checkboxes if cb.isChecked()]


class ComparisonDialog(QDialog):
    """数据比对对话框"""
    
    def __init__(self, parent: Any, selected_columns: List[str], excel_handler: ExcelHandler):
        super().__init__(parent)
        
        self.selected_columns: List[str] = selected_columns
        self.excel_handler: ExcelHandler = excel_handler
        self.all_columns: List[str] = excel_handler.get_column_names()
        self.preview_dataframe: Optional[pd.DataFrame] = None
        self.current_filter: Optional[str] = None
        self.applied_filters: List[str] = []
        
        # UI元素
        self.columns_info_label: QLabel = None
        self.edit_columns_button: QPushButton = None
        self.input_edit: QLineEdit = None
        self.preview_button: QPushButton = None
        self.filter_button: QPushButton = None
        self.table_view: QTableView = None
        self.status_label: QLabel = None
        self.continue_button: QPushButton = None
        self.finish_button: QPushButton = None
        
        self._setup_ui()
        self._connect_signals()
        
    def _setup_ui(self) -> None:
        """设置用户界面"""
        self.setWindowTitle("数据比对")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # 列选择区域
        layout.addLayout(self._create_column_section())
        
        # 输入区域
        layout.addLayout(self._create_input_section())
        
        # 预览区域
        layout.addWidget(QLabel("筛选结果预览:"))
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        layout.addWidget(self.table_view, 1)
        
        # 状态区域
        self.status_label = QLabel("当前状态: 请输入第一个筛选条件")
        layout.addWidget(self.status_label)
        
        # 按钮区域
        layout.addLayout(self._create_button_section())
    
    def _create_column_section(self) -> QHBoxLayout:
        """创建列选择区域
        
        Returns:
            QHBoxLayout: 包含列选择控件的布局
        """
        layout = QHBoxLayout()
        
        self.columns_info_label = QLabel(f"已选择列: {', '.join(self.selected_columns)}")
        self.edit_columns_button = QPushButton("更改列选择")
        
        layout.addWidget(self.columns_info_label, 1)
        layout.addWidget(self.edit_columns_button)
        return layout
    
    def _create_input_section(self) -> QHBoxLayout:
        """创建输入区域
        
        Returns:
            QHBoxLayout: 包含输入控件的布局
        """
        layout = QHBoxLayout()
        
        layout.addWidget(QLabel("比对内容:"))
        
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("请输入要比对的内容")
        layout.addWidget(self.input_edit, 1)
        
        self.preview_button = QPushButton("预览数据")
        self.filter_button = QPushButton("执行筛选")
        
        layout.addWidget(self.preview_button)
        layout.addWidget(self.filter_button)
        return layout
    
    def _create_button_section(self) -> QHBoxLayout:
        """创建底部按钮区域
        
        Returns:
            QHBoxLayout: 包含底部按钮的布局
        """
        layout = QHBoxLayout()
        
        self.continue_button = QPushButton("继续比对")
        self.continue_button.setEnabled(False)
        
        self.finish_button = QPushButton("完成")
        self.finish_button.setEnabled(False)
        
        layout.addWidget(self.continue_button)
        layout.addWidget(self.finish_button)
        return layout
    
    def _connect_signals(self) -> None:
        """连接信号和槽"""
        self.edit_columns_button.clicked.connect(self.edit_selected_columns)
        self.preview_button.clicked.connect(self.preview_data)
        self.filter_button.clicked.connect(self.filter_data)
        self.continue_button.clicked.connect(self.continue_comparison)
        self.finish_button.clicked.connect(self.accept)
    
    def preview_data(self) -> None:
        """预览当前输入条件匹配的数据"""
        filter_value = self.input_edit.text().strip()
        
        if not filter_value:
            QMessageBox.warning(self, "警告", MESSAGES["enter_filter_text"])
            return
        
        try:
            # 确定要筛选的数据源
            source_data = self.preview_dataframe if self.preview_dataframe is not None else self.excel_handler.dataframe
            
            if source_data is None:
                QMessageBox.warning(self, "警告", "没有可用的数据源")
                return
                
            # 创建筛选掩码
            mask = self._create_filter_mask(source_data, filter_value)
            filtered_data = source_data[mask].copy()
            
            if filtered_data.empty:
                scope = "当前预览数据中" if self.preview_dataframe is not None else ""
                QMessageBox.warning(self, "警告", f"在{scope}没有找到匹配 '{filter_value}' 的数据")
                return
            
            # 更新预览状态
            self.preview_dataframe = filtered_data
            self.current_filter = filter_value
            self._update_status_label()
            self._display_filtered_data(filtered_data)
            
            self.input_edit.clear()
            QMessageBox.information(self, "预览成功", f"找到匹配的数据 {len(filtered_data)} 行")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览数据时出错: {str(e)}")
    
    def filter_data(self) -> None:
        """执行筛选，将预览数据应用为最终筛选结果"""
        if self.preview_dataframe is None:
            QMessageBox.warning(self, "警告", MESSAGES["preview_first"])
            return
        
        try:
            # 在原始数据中找到匹配的行并移除
            indices_to_remove = self._find_matching_indices()
            
            if indices_to_remove:
                # 获取筛选数据
                mask = self.excel_handler.dataframe.index.isin(indices_to_remove)
                filtered_data = self.excel_handler.dataframe[mask].copy()
                
                # 生成条件名称
                condition_name = self._generate_condition_name()
                
                # 保存筛选结果并移除原始数据
                self.excel_handler.filtered_sheets[condition_name] = filtered_data
                self.excel_handler.dataframe = self.excel_handler.dataframe[~mask].reset_index(drop=True)
                
                # 重置状态
                self._reset_filter_state()
                
                QMessageBox.information(self, "筛选成功", 
                    f"已成功将筛选结果保存为 '{condition_name}'，共 {len(filtered_data)} 行")
            else:
                QMessageBox.warning(self, "警告", MESSAGES["no_data_found"])
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"筛选数据时出错: {str(e)}")
    
    def _create_filter_mask(self, data: pd.DataFrame, filter_value: str) -> pd.Series:
        """创建筛选掩码
        
        Args:
            data: 要筛选的数据
            filter_value: 筛选条件
            
        Returns:
            pd.Series: 布尔掩码，表示每行是否匹配筛选条件
        """
        mask = pd.Series(False, index=data.index)
        for col in self.selected_columns:
            if col in data.columns:
                mask |= data[col].astype(str).str.contains(filter_value, na=False, case=False, regex=False)
        return mask
    
    def _find_matching_indices(self) -> List[int]:
        """在原始数据中找到与预览数据匹配的行索引
        
        Returns:
            List[int]: 匹配行的索引列表
        """
        if self.preview_dataframe is None or self.excel_handler.dataframe is None:
            return []
            
        indices = []
        for _, preview_row in self.preview_dataframe.iterrows():
            for idx, orig_row in self.excel_handler.dataframe.iterrows():
                if self._rows_match(preview_row, orig_row):
                    indices.append(idx)
                    break
        return indices
    
    def _rows_match(self, row1: pd.Series, row2: pd.Series) -> bool:
        """检查两行数据是否匹配
        
        Args:
            row1: 第一行数据
            row2: 第二行数据
            
        Returns:
            bool: 如果两行匹配则为True，否则为False
        """
        if self.preview_dataframe is None or self.excel_handler.dataframe is None:
            return False
            
        for col in self.preview_dataframe.columns:
            if col in self.excel_handler.dataframe.columns:
                # 将值转换为字符串进行比较，避免类型不匹配问题
                if str(row1[col]) != str(row2[col]):
                    return False
        return True
    
    def _generate_condition_name(self) -> str:
        """生成条件名称
        
        Returns:
            str: 生成的条件名称
        """
        all_filters = self.applied_filters.copy()
        if self.current_filter:
            all_filters.append(self.current_filter)
        return " 与 ".join(all_filters) if len(all_filters) > 1 else all_filters[0]
    
    def _update_status_label(self) -> None:
        """更新状态标签"""
        conditions = [f"'{f}'" for f in self.applied_filters]
        if self.current_filter:
            conditions.append(f"'{self.current_filter}'")
        
        if len(conditions) == 1:
            self.status_label.setText(f"当前状态: 满足条件 {conditions[0]} 的数据")
        else:
            self.status_label.setText(f"当前状态: 同时满足条件 {' 和 '.join(conditions)} 的数据")
    
    def _reset_filter_state(self) -> None:
        """重置筛选状态"""
        self.preview_dataframe = None
        self.applied_filters = []
        self.current_filter = None
        self.status_label.setText("当前状态: 请输入第一个筛选条件")
        self.table_view.setModel(None)
        self.continue_button.setEnabled(True)
        self.finish_button.setEnabled(True)
    
    def _display_filtered_data(self, dataframe: pd.DataFrame) -> None:
        """在表格中显示筛选结果
        
        Args:
            dataframe: 要显示的数据
        """
        model = PandasTableModel(dataframe)
        self.table_view.setModel(model)
        
        # 自动调整列宽
        for i in range(len(dataframe.columns)):
            self.table_view.resizeColumnToContents(i)
    
    def continue_comparison(self) -> None:
        """清空当前输入，准备下一次比对"""
        self.input_edit.clear()
        self.table_view.setModel(None)
        self.continue_button.setEnabled(False)
        self.finish_button.setEnabled(False)
        self.status_label.setText("当前状态: 请输入第一个筛选条件")
        self.preview_dataframe = None
        self.applied_filters = []
        self.current_filter = None
    
    def edit_selected_columns(self) -> None:
        """更改列选择，同时保存当前的筛选条件"""
        # 如果当前有过滤条件，保存它
        if self.current_filter and self.current_filter not in self.applied_filters:
            self.applied_filters.append(self.current_filter)
            self.current_filter = None
        
        # 打开列选择对话框
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

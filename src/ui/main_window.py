import sys
import os
from typing import List, Tuple, Optional
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QFileDialog, QMessageBox, QListWidget, QAbstractItemView, QStatusBar
)

from ..excel_handler import ExcelHandler
from .comparison_dialog import ComparisonDialog
from ..config import EXCEL_FILE_FILTERS, WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT, APP_NAME, MESSAGES


class MainWindow(QMainWindow):
    """主窗口类，负责应用的主界面和用户交互"""
    
    def __init__(self):
        super().__init__()
        self.excel_handler = ExcelHandler()
        self.file_label: QLabel = None
        self.file_button: QPushButton = None
        self.columns_list: QListWidget = None
        self.compare_button: QPushButton = None
        self.export_button: QPushButton = None
        self.filtered_list: QListWidget = None
        self.status_bar: QStatusBar = None
        
        self._setup_ui()
        self._connect_signals()
        
    def _setup_ui(self) -> None:
        """设置用户界面"""
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        
        # 创建中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 文件导入区域
        layout.addWidget(self._create_file_section())
        
        # 列选择区域
        layout.addWidget(QLabel("可用列:"))
        self.columns_list = QListWidget()
        self.columns_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.columns_list)
        
        # 操作按钮区域
        layout.addWidget(self._create_button_section())
        
        # 筛选结果区域
        layout.addWidget(QLabel("已完成的筛选结果:"))
        self.filtered_list = QListWidget()
        layout.addWidget(self.filtered_list)
        
        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(MESSAGES["select_file_first"])
    
    def _create_file_section(self) -> QWidget:
        """创建文件导入区域
        
        Returns:
            QWidget: 包含文件导入控件的小部件
        """
        widget = QWidget()
        layout = QHBoxLayout(widget)
        
        self.file_label = QLabel(MESSAGES["no_file_selected"])
        self.file_button = QPushButton("导入Excel文件")
        
        layout.addWidget(self.file_label, 1)
        layout.addWidget(self.file_button)
        return widget
    
    def _create_button_section(self) -> QWidget:
        """创建操作按钮区域
        
        Returns:
            QWidget: 包含操作按钮的小部件
        """
        widget = QWidget()
        layout = QHBoxLayout(widget)
        
        self.compare_button = QPushButton("开始比对")
        self.compare_button.setEnabled(False)
        
        self.export_button = QPushButton("导出结果")
        self.export_button.setEnabled(False)
        
        layout.addWidget(self.compare_button)
        layout.addWidget(self.export_button)
        return widget
    
    def _connect_signals(self) -> None:
        """连接信号和槽"""
        self.file_button.clicked.connect(self.import_excel)
        self.compare_button.clicked.connect(self.start_comparison)
        self.export_button.clicked.connect(self.export_results)
    
    def import_excel(self) -> None:
        """导入Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", EXCEL_FILE_FILTERS
        )
        
        if not file_path:
            return
        
        success, result = self.excel_handler.load_excel(file_path)
        
        if success:
            # 更新UI
            self.file_label.setText(os.path.basename(file_path))
            self._update_columns_list(result)
            self.compare_button.setEnabled(True)
            self.status_bar.showMessage(f"{MESSAGES['import_success']}: {os.path.basename(file_path)}")
        else:
            QMessageBox.critical(self, "导入错误", result)
    
    def _update_columns_list(self, columns: List[str]) -> None:
        """更新列列表
        
        Args:
            columns: 列名列表
        """
        self.columns_list.clear()
        for column in columns:
            self.columns_list.addItem(column)
    
    def start_comparison(self) -> None:
        """开始比对操作"""
        selected_columns = [item.text() for item in self.columns_list.selectedItems()]
        
        if not selected_columns:
            QMessageBox.warning(self, "警告", MESSAGES["select_columns"])
            return
        
        dialog = ComparisonDialog(self, selected_columns, self.excel_handler)
        if dialog.exec():
            self._update_filtered_list()
            self.export_button.setEnabled(True)
    
    def _update_filtered_list(self) -> None:
        """更新筛选结果列表"""
        self.filtered_list.clear()
        for sheet_name in self.excel_handler.get_all_filtered_sheets():
            self.filtered_list.addItem(sheet_name)
    
    def export_results(self) -> None:
        """导出最终Excel文件"""
        default_dir = ""
        if self.excel_handler.excel_file_path:
            default_dir = os.path.dirname(self.excel_handler.excel_file_path)
            
        save_directory = QFileDialog.getExistingDirectory(
            self, "选择保存目录", default_dir
        )
        
        success, result = self.excel_handler.export_final_excel(
            save_directory if save_directory else None
        )
        
        if success:
            QMessageBox.information(self, "导出成功", f"{MESSAGES['export_success']}:\n{result}")
            self._reset_ui()
        else:
            QMessageBox.critical(self, "导出错误", result)
            
    def _reset_ui(self) -> None:
        """重置UI界面"""
        self.excel_handler = ExcelHandler()
        self.file_label.setText(MESSAGES["no_file_selected"])
        self.columns_list.clear()
        self.filtered_list.clear()
        self.compare_button.setEnabled(False)
        self.export_button.setEnabled(False)
        self.status_bar.showMessage(MESSAGES["select_file_first"])

import sys
import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QFileDialog, QMessageBox, QApplication, QListWidget,
    QListWidgetItem, QAbstractItemView, QStatusBar
)
from PyQt6.QtCore import Qt

from ..excel_handler import ExcelHandler
from .comparison_dialog import ComparisonDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 初始化Excel处理器
        self.excel_handler = ExcelHandler()
        
        # 设置窗口属性
        self.setWindowTitle("Excel数据比对工具")
        self.setMinimumSize(800, 600)
        
        # 创建中央控件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 添加文件导入部分
        file_section = QWidget()
        file_layout = QHBoxLayout(file_section)
        
        self.file_label = QLabel("未选择文件")
        self.file_button = QPushButton("导入Excel文件")
        self.file_button.clicked.connect(self.import_excel)
        
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.file_button)
        
        main_layout.addWidget(file_section)
        
        # 添加列选择部分标题
        columns_label = QLabel("可用列:")
        main_layout.addWidget(columns_label)
        
        # 添加列选择列表
        self.columns_list = QListWidget()
        self.columns_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        main_layout.addWidget(self.columns_list)
        
        # 添加比对操作按钮
        button_section = QWidget()
        button_layout = QHBoxLayout(button_section)
        
        self.compare_button = QPushButton("开始比对")
        self.compare_button.clicked.connect(self.start_comparison)
        self.compare_button.setEnabled(False)
        
        self.export_button = QPushButton("导出结果")
        self.export_button.clicked.connect(self.export_results)
        self.export_button.setEnabled(False)
        
        button_layout.addWidget(self.compare_button)
        button_layout.addWidget(self.export_button)
        
        main_layout.addWidget(button_section)
        
        # 添加筛选结果列表标题
        filtered_label = QLabel("已完成的筛选结果:")
        main_layout.addWidget(filtered_label)
        
        # 添加筛选结果列表
        self.filtered_list = QListWidget()
        main_layout.addWidget(self.filtered_list)
        
        # 添加状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # 显示初始消息
        self.status_bar.showMessage("请导入Excel文件开始操作")
    
    def import_excel(self):
        """导入Excel文件并显示列信息"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        
        if not file_path:
            return
        
        success, result = self.excel_handler.load_excel(file_path)
        
        if success:
            self.file_label.setText(os.path.basename(file_path))
            self.status_bar.showMessage(f"成功导入文件: {os.path.basename(file_path)}")
            
            # 更新列列表
            self.columns_list.clear()
            for column in result:
                item = QListWidgetItem(column)
                self.columns_list.addItem(item)
            
            # 启用比对按钮
            self.compare_button.setEnabled(True)
        else:
            QMessageBox.critical(self, "导入错误", f"导入Excel文件失败: {result}")
    
    def start_comparison(self):
        """开始比对操作"""
        # 获取选中的列
        selected_columns = [item.text() for item in self.columns_list.selectedItems()]
        
        if not selected_columns:
            QMessageBox.warning(self, "警告", "请至少选择一列进行比对")
            return
        
        # 打开比对对话框
        dialog = ComparisonDialog(self, selected_columns, self.excel_handler)
        if dialog.exec():
            # 如果有比对结果，更新筛选结果列表
            self.update_filtered_list()
            
            # 启用导出按钮
            self.export_button.setEnabled(True)
    
    def update_filtered_list(self):
        """更新筛选结果列表"""
        self.filtered_list.clear()
        for sheet_name in self.excel_handler.get_all_filtered_sheets():
            item = QListWidgetItem(sheet_name)
            self.filtered_list.addItem(item)
    
    def export_results(self):
        """导出最终Excel文件"""
        # 打开目录选择对话框
        save_directory = QFileDialog.getExistingDirectory(
            self, "选择保存目录", 
            os.path.dirname(self.excel_handler.excel_file_path) if self.excel_handler.excel_file_path else ""
        )
        
        # 如果用户取消了选择，则使用默认目录
        success, result = self.excel_handler.export_final_excel(save_directory if save_directory else None)
        
        if success:
            QMessageBox.information(
                self, "导出成功", 
                f"已成功导出Excel文件:\n{result}"
            )
            
            # 清空UI内容，为下一次对比做好准备
            self.reset_ui()
        else:
            QMessageBox.critical(self, "导出错误", f"导出Excel文件失败: {result}")
            
    def reset_ui(self):
        """重置UI界面，清空所有内容为下一次对比做准备"""
        # 重置Excel处理器
        self.excel_handler = ExcelHandler()
        
        # 重置文件标签
        self.file_label.setText("未选择文件")
        
        # 清空列表
        self.columns_list.clear()
        self.filtered_list.clear()
        
        # 禁用按钮
        self.compare_button.setEnabled(False)
        self.export_button.setEnabled(False)
        
        # 更新状态栏
        self.status_bar.showMessage("请导入Excel文件开始操作")


# 测试代码，如果直接运行此文件
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
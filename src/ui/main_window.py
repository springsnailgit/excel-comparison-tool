import os
from typing import List
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QMessageBox, QListWidget, QAbstractItemView,
    QStatusBar, QComboBox, QGroupBox, QTextEdit
)
from PyQt6.QtCore import QTimer

from ..excel_handler import ExcelHandler
from .comparison_dialog import ComparisonDialog
from ..config import config, MESSAGES, APP_NAME
from ..utils.logger import get_logger


class MainWindow(QMainWindow):
    """主窗口类，负责应用的主界面和用户交互"""

    def __init__(self):
        super().__init__()
        self.logger = get_logger(self.__class__.__name__)
        self.excel_handler = ExcelHandler()

        # UI组件
        self.file_label: QLabel = None
        self.file_button: QPushButton = None
        self.columns_list: QListWidget = None
        self.compare_button: QPushButton = None
        self.export_button: QPushButton = None
        self.reset_button: QPushButton = None
        self.filtered_list: QListWidget = None
        self.status_bar: QStatusBar = None
        self.filter_strategy_combo: QComboBox = None
        self.data_summary_text: QTextEdit = None

        self._setup_ui()
        self._connect_signals()
        self._update_ui_state()
        
    def _setup_ui(self) -> None:
        """设置用户界面"""
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(config.get("window_min_width", 800), config.get("window_min_height", 600))
        
        # 创建中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 文件导入区域
        layout.addWidget(self._create_file_section())

        # 筛选策略选择区域
        layout.addWidget(self._create_strategy_section())

        # 列选择区域
        columns_group = QGroupBox("可用列")
        columns_layout = QVBoxLayout(columns_group)
        self.columns_list = QListWidget()
        self.columns_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        columns_layout.addWidget(self.columns_list)
        layout.addWidget(columns_group)

        # 操作按钮区域
        layout.addWidget(self._create_button_section())

        # 数据摘要区域
        layout.addWidget(self._create_summary_section())

        # 筛选结果区域
        results_group = QGroupBox("已完成的筛选结果")
        results_layout = QVBoxLayout(results_group)
        self.filtered_list = QListWidget()
        results_layout.addWidget(self.filtered_list)
        layout.addWidget(results_group)
        
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
    
    def _create_strategy_section(self) -> QWidget:
        """创建筛选策略选择区域

        Returns:
            QWidget: 包含策略选择的小部件
        """
        group = QGroupBox("筛选策略")
        layout = QHBoxLayout(group)

        layout.addWidget(QLabel("筛选方式:"))
        self.filter_strategy_combo = QComboBox()
        self.filter_strategy_combo.addItems([
            ("contains", "包含匹配"),
            ("exact", "精确匹配"),
            ("regex", "正则表达式")
        ])
        self.filter_strategy_combo.setCurrentText("包含匹配")
        layout.addWidget(self.filter_strategy_combo)

        return group

    def _create_button_section(self) -> QWidget:
        """创建操作按钮区域

        Returns:
            QWidget: 包含操作按钮的小部件
        """
        widget = QWidget()
        layout = QHBoxLayout(widget)

        self.compare_button = QPushButton("开始比对")
        self.compare_button.setEnabled(False)

        self.reset_button = QPushButton("重置数据")
        self.reset_button.setEnabled(False)

        self.export_button = QPushButton("导出结果")
        self.export_button.setEnabled(False)

        layout.addWidget(self.compare_button)
        layout.addWidget(self.reset_button)
        layout.addWidget(self.export_button)
        return widget

    def _create_summary_section(self) -> QWidget:
        """创建数据摘要区域

        Returns:
            QWidget: 包含数据摘要的小部件
        """
        group = QGroupBox("数据摘要")
        layout = QVBoxLayout(group)

        self.data_summary_text = QTextEdit()
        self.data_summary_text.setMaximumHeight(100)
        self.data_summary_text.setReadOnly(True)
        layout.addWidget(self.data_summary_text)

        return group
    
    def _connect_signals(self) -> None:
        """连接信号和槽"""
        self.file_button.clicked.connect(self.import_excel)
        self.compare_button.clicked.connect(self.start_comparison)
        self.reset_button.clicked.connect(self.reset_data)
        self.export_button.clicked.connect(self.export_results)
        self.filter_strategy_combo.currentTextChanged.connect(self.on_strategy_changed)

        # 定时更新数据摘要
        self.summary_timer = QTimer()
        self.summary_timer.timeout.connect(self.update_data_summary)
        self.summary_timer.start(2000)  # 每2秒更新一次
    
    def import_excel(self) -> None:
        """导入Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", config.get("excel_file_filters", "Excel Files (*.xlsx *.xls *.xlsm)")
        )
        
        if not file_path:
            return
        
        success, result = self.excel_handler.load_excel(file_path)
        
        if success:
            # 更新UI
            self.file_label.setText(os.path.basename(file_path))
            self._update_columns_list(result)
            self._update_ui_state()
            self.update_data_summary()
            self.status_bar.showMessage(f"{MESSAGES['import_success']}: {os.path.basename(file_path)}")
            self.logger.info(f"成功导入文件: {file_path}")
        else:
            QMessageBox.critical(self, "导入错误", result)
            self.logger.error(f"导入文件失败: {result}")
    
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
            self._update_ui_state()
            self.update_data_summary()
    
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
            self.logger.info(f"成功导出文件: {result}")
        else:
            QMessageBox.critical(self, "导出错误", result)
            self.logger.error(f"导出失败: {result}")

    def reset_data(self) -> None:
        """重置数据到原始状态"""
        if self.excel_handler.original_dataframe is None:
            QMessageBox.warning(self, "警告", "没有原始数据可重置")
            return

        reply = QMessageBox.question(
            self, "确认重置",
            "确定要重置所有数据到原始状态吗？这将清除所有筛选结果。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.excel_handler.reset_data():
                self._update_filtered_list()
                self._update_ui_state()
                self.update_data_summary()
                self.status_bar.showMessage("数据已重置到原始状态")
                self.logger.info("用户重置了数据")
            else:
                QMessageBox.critical(self, "错误", "重置数据失败")

    def on_strategy_changed(self, strategy_text: str) -> None:
        """筛选策略改变时的处理"""
        strategy_map = {
            "包含匹配": "contains",
            "精确匹配": "exact",
            "正则表达式": "regex"
        }

        strategy = strategy_map.get(strategy_text, "contains")
        if self.excel_handler.set_filter_strategy(strategy):
            self.status_bar.showMessage(f"筛选策略已设置为: {strategy_text}")
            self.logger.info(f"筛选策略已更改为: {strategy}")

    def update_data_summary(self) -> None:
        """更新数据摘要"""
        if self.excel_handler.dataframe is None:
            self.data_summary_text.setText("未加载数据")
            return

        summary = self.excel_handler.get_data_summary()
        summary_text = f"""文件: {os.path.basename(summary['file_path']) if summary['file_path'] else '未知'}
原始行数: {summary['original_rows']:,}
当前行数: {summary['current_rows']:,}
已筛选行数: {summary['total_filtered_rows']:,}
筛选结果数: {summary['filtered_sheets_count']}
列数: {len(summary['columns'])}"""

        self.data_summary_text.setText(summary_text)

    def _update_ui_state(self) -> None:
        """更新UI状态"""
        has_data = self.excel_handler.dataframe is not None
        has_filtered_data = len(self.excel_handler.filtered_sheets) > 0
        has_original_data = self.excel_handler.original_dataframe is not None

        self.compare_button.setEnabled(has_data)
        self.reset_button.setEnabled(has_original_data and has_filtered_data)
        self.export_button.setEnabled(has_filtered_data)

    def _reset_ui(self) -> None:
        """重置UI界面"""
        self.excel_handler = ExcelHandler()
        self.file_label.setText(MESSAGES["no_file_selected"])
        self.columns_list.clear()
        self.filtered_list.clear()
        self.data_summary_text.clear()
        self._update_ui_state()
        self.status_bar.showMessage(MESSAGES["select_file_first"])

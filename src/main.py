#!/usr/bin/env python3
import sys
import os

# 将项目根目录添加到 Python 路径中，以便可以正确导入模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from src.ui.main_window import MainWindow


def main():
    """程序主入口函数"""
    # 创建QApplication实例
    app = QApplication(sys.argv)
    app.setApplicationName("Excel数据比对工具")
    
    # 设置应用样式
    app.setStyle("Fusion")
    
    # 创建并显示主窗口
    main_window = MainWindow()
    main_window.show()
    
    # 进入应用的主事件循环
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""Excel数据比对工具主程序入口"""

import sys
import os
from typing import NoReturn

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

print("正在导入PyQt6...")
from PyQt6.QtWidgets import QApplication
print("正在导入MainWindow...")
from src.ui.main_window import MainWindow
print("正在导入APP_NAME...")
from src.config import APP_NAME


def main() -> NoReturn:
    """程序主入口函数
    
    启动应用程序并显示主窗口
    """
    print("正在初始化应用程序...")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    
    print("正在创建主窗口...")
    # 创建并显示主窗口
    window = MainWindow()
    print("正在显示主窗口...")
    window.show()
    
    print("正在启动应用程序事件循环...")
    # 启动应用程序
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        print("程序开始执行...")
        main()
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        import traceback
        traceback.print_exc()

#!/usr/bin/env python3
"""Excel数据比对工具主程序入口"""

import sys
import os
from typing import NoReturn

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from PyQt6.QtWidgets import QApplication
from src.ui.main_window import MainWindow
from src.config import APP_NAME
from src.utils.logger import setup_logger, get_logger


def main() -> NoReturn:
    """程序主入口函数

    启动应用程序并显示主窗口
    """
    # 设置日志
    logger = setup_logger()
    logger.info("程序开始启动...")

    try:
        logger.info("正在初始化应用程序...")
        app = QApplication(sys.argv)
        app.setApplicationName(APP_NAME)
        app.setStyle("Fusion")

        logger.info("正在创建主窗口...")
        # 创建并显示主窗口
        window = MainWindow()
        logger.info("正在显示主窗口...")
        window.show()

        logger.info("正在启动应用程序事件循环...")
        # 启动应用程序
        sys.exit(app.exec())

    except Exception as e:
        logger.error(f"程序启动失败: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

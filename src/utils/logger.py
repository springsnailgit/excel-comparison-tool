"""日志管理模块"""
import logging
import logging.handlers
from pathlib import Path
from typing import Optional
from ..config import config


def setup_logger(name: str = "excel_comparison", log_file: Optional[str] = None) -> logging.Logger:
    """设置日志记录器
    
    Args:
        name: 日志记录器名称
        log_file: 日志文件路径，为None时使用默认路径
        
    Returns:
        logging.Logger: 配置好的日志记录器
    """
    logger = logging.getLogger(name)
    
    # 避免重复添加处理器
    if logger.handlers:
        return logger
    
    # 设置日志级别
    log_level = getattr(logging, config.get("log_level", "INFO").upper())
    logger.setLevel(log_level)
    
    # 创建格式化器
    formatter = logging.Formatter(
        fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt=config.get("log_date_format", "%Y-%m-%d %H:%M:%S")
    )
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 文件处理器
    if log_file is None:
        log_dir = Path.home() / ".excel_comparison_tool" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "app.log"
    
    # 使用RotatingFileHandler避免日志文件过大
    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=config.get("log_file_max_size", 10) * 1024 * 1024,  # MB to bytes
        backupCount=config.get("log_backup_count", 5),
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = "excel_comparison") -> logging.Logger:
    """获取日志记录器
    
    Args:
        name: 日志记录器名称
        
    Returns:
        logging.Logger: 日志记录器
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        return setup_logger(name)
    return logger


# 创建默认日志记录器
default_logger = setup_logger()

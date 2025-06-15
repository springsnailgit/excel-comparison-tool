"""性能监控和优化工具"""
import time
import functools
from typing import Callable, Any, Dict
from .logger import get_logger

# 可选依赖
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

logger = get_logger(__name__)


class PerformanceMonitor:
    """性能监控器"""
    
    def __init__(self):
        self.metrics = {}
    
    def measure_time(self, func_name: str = None):
        """装饰器：测量函数执行时间"""
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper(*args, **kwargs) -> Any:
                name = func_name or f"{func.__module__}.{func.__name__}"
                start_time = time.time()
                
                try:
                    result = func(*args, **kwargs)
                    execution_time = time.time() - start_time
                    
                    # 记录性能指标
                    if name not in self.metrics:
                        self.metrics[name] = []
                    self.metrics[name].append(execution_time)
                    
                    # 如果执行时间超过阈值，记录警告
                    if execution_time > 5.0:  # 5秒阈值
                        logger.warning(f"函数 {name} 执行时间较长: {execution_time:.2f}秒")
                    else:
                        logger.debug(f"函数 {name} 执行时间: {execution_time:.2f}秒")
                    
                    return result
                    
                except Exception as e:
                    execution_time = time.time() - start_time
                    logger.error(f"函数 {name} 执行失败 (耗时 {execution_time:.2f}秒): {str(e)}")
                    raise
                    
            return wrapper
        return decorator
    
    def get_memory_usage(self) -> Dict[str, float]:
        """获取内存使用情况"""
        if not HAS_PSUTIL:
            return {
                'rss_mb': 0.0,
                'vms_mb': 0.0,
                'percent': 0.0,
            }

        try:
            process = psutil.Process()
            memory_info = process.memory_info()

            return {
                'rss_mb': memory_info.rss / 1024 / 1024,  # 物理内存
                'vms_mb': memory_info.vms / 1024 / 1024,  # 虚拟内存
                'percent': process.memory_percent(),       # 内存使用百分比
            }
        except Exception as e:
            logger.warning(f"获取内存使用情况失败: {e}")
            return {
                'rss_mb': 0.0,
                'vms_mb': 0.0,
                'percent': 0.0,
            }
    
    def get_performance_summary(self) -> Dict[str, Any]:
        """获取性能摘要"""
        summary = {}
        
        for func_name, times in self.metrics.items():
            if times:
                summary[func_name] = {
                    'count': len(times),
                    'total_time': sum(times),
                    'avg_time': sum(times) / len(times),
                    'min_time': min(times),
                    'max_time': max(times),
                }
        
        # 添加内存信息
        summary['memory'] = self.get_memory_usage()
        
        return summary
    
    def log_performance_summary(self):
        """记录性能摘要到日志"""
        summary = self.get_performance_summary()
        
        logger.info("=== 性能摘要 ===")
        
        # 内存使用情况
        memory = summary.get('memory', {})
        logger.info(f"内存使用: {memory.get('rss_mb', 0):.1f}MB (物理), "
                   f"{memory.get('vms_mb', 0):.1f}MB (虚拟), "
                   f"{memory.get('percent', 0):.1f}%")
        
        # 函数执行时间
        for func_name, metrics in summary.items():
            if func_name != 'memory':
                logger.info(f"{func_name}: "
                           f"调用{metrics['count']}次, "
                           f"总耗时{metrics['total_time']:.2f}秒, "
                           f"平均{metrics['avg_time']:.2f}秒")


# 全局性能监控器实例
performance_monitor = PerformanceMonitor()


def monitor_performance(func_name: str = None):
    """性能监控装饰器的便捷函数"""
    return performance_monitor.measure_time(func_name)


def check_memory_usage(threshold_mb: float = 500.0) -> bool:
    """检查内存使用是否超过阈值

    Args:
        threshold_mb: 内存阈值(MB)

    Returns:
        bool: 是否超过阈值
    """
    if not HAS_PSUTIL:
        return False

    memory_info = performance_monitor.get_memory_usage()
    current_mb = memory_info['rss_mb']

    if current_mb > threshold_mb:
        logger.warning(f"内存使用过高: {current_mb:.1f}MB (阈值: {threshold_mb}MB)")
        return True

    return False


def optimize_dataframe_memory(df) -> Any:
    """优化DataFrame内存使用
    
    Args:
        df: pandas DataFrame
        
    Returns:
        优化后的DataFrame
    """
    try:
        import pandas as pd
        
        if not isinstance(df, pd.DataFrame):
            return df
        
        original_memory = df.memory_usage(deep=True).sum() / 1024 / 1024
        
        # 优化数值类型
        for col in df.select_dtypes(include=['int64']).columns:
            df[col] = pd.to_numeric(df[col], downcast='integer')
        
        for col in df.select_dtypes(include=['float64']).columns:
            df[col] = pd.to_numeric(df[col], downcast='float')
        
        # 优化字符串类型
        for col in df.select_dtypes(include=['object']).columns:
            if df[col].dtype == 'object':
                try:
                    df[col] = df[col].astype('category')
                except:
                    pass  # 如果转换失败，保持原样
        
        optimized_memory = df.memory_usage(deep=True).sum() / 1024 / 1024
        reduction = (original_memory - optimized_memory) / original_memory * 100
        
        if reduction > 5:  # 只有在显著减少内存时才记录
            logger.info(f"DataFrame内存优化: {original_memory:.1f}MB -> {optimized_memory:.1f}MB "
                       f"(减少 {reduction:.1f}%)")
        
        return df
        
    except Exception as e:
        logger.warning(f"DataFrame内存优化失败: {str(e)}")
        return df


class ProgressTracker:
    """进度跟踪器"""
    
    def __init__(self, total: int, description: str = "处理中"):
        self.total = total
        self.current = 0
        self.description = description
        self.start_time = time.time()
        self.last_log_time = self.start_time
    
    def update(self, increment: int = 1):
        """更新进度"""
        self.current += increment
        current_time = time.time()
        
        # 每5秒或完成时记录一次进度
        if (current_time - self.last_log_time > 5.0) or (self.current >= self.total):
            progress_percent = (self.current / self.total) * 100
            elapsed_time = current_time - self.start_time
            
            if self.current > 0:
                estimated_total_time = elapsed_time * self.total / self.current
                remaining_time = estimated_total_time - elapsed_time
                
                logger.info(f"{self.description}: {self.current}/{self.total} "
                           f"({progress_percent:.1f}%) - "
                           f"预计剩余时间: {remaining_time:.1f}秒")
            
            self.last_log_time = current_time
    
    def finish(self):
        """完成进度跟踪"""
        total_time = time.time() - self.start_time
        logger.info(f"{self.description}完成: {self.total}项，总耗时: {total_time:.1f}秒")

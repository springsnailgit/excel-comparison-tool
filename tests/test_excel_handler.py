"""Excel处理器测试"""
import unittest
import pandas as pd
import tempfile
import os
from pathlib import Path
import sys

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.excel_handler import ExcelHandler, ContainsFilter, ExactMatchFilter
from src.utils.validators import DataValidator


class TestExcelHandler(unittest.TestCase):
    """Excel处理器测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.handler = ExcelHandler()
        
        # 创建测试数据
        self.test_data = pd.DataFrame({
            'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
            'Age': [25, 30, 35, 28, 32],
            'City': ['New York', 'London', 'Paris', 'Tokyo', 'Sydney'],
            'Department': ['IT', 'HR', 'IT', 'Finance', 'IT']
        })
        
        # 创建临时Excel文件
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.test_data.to_excel(self.temp_file.name, index=False)
        self.temp_file.close()
    
    def tearDown(self):
        """测试后清理"""
        if os.path.exists(self.temp_file.name):
            os.unlink(self.temp_file.name)
    
    def test_load_excel_success(self):
        """测试成功加载Excel文件"""
        success, result = self.handler.load_excel(self.temp_file.name)
        self.assertTrue(success)
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 4)  # 4列
        self.assertIn('Name', result)
    
    def test_load_excel_file_not_found(self):
        """测试加载不存在的文件"""
        success, result = self.handler.load_excel('nonexistent.xlsx')
        self.assertFalse(success)
        self.assertIn('不存在', result)
    
    def test_filter_data_contains(self):
        """测试包含筛选"""
        self.handler.load_excel(self.temp_file.name)
        success, result = self.handler.filter_data(['Department'], 'IT')
        
        self.assertTrue(success)
        self.assertIsInstance(result, pd.DataFrame)
        self.assertEqual(len(result), 3)  # 3个IT部门的人
        
        # 检查原数据是否正确移除
        self.assertEqual(len(self.handler.dataframe), 2)  # 剩余2行
    
    def test_filter_data_no_match(self):
        """测试无匹配结果的筛选"""
        self.handler.load_excel(self.temp_file.name)
        success, result = self.handler.filter_data(['Department'], 'NonExistent')
        
        self.assertFalse(success)
        self.assertIn('没有找到匹配', result)
    
    def test_filter_strategies(self):
        """测试不同的筛选策略"""
        self.handler.load_excel(self.temp_file.name)
        
        # 测试精确匹配
        self.handler.set_filter_strategy('exact')
        success, result = self.handler.filter_data(['Name'], 'Alice')
        self.assertTrue(success)
        self.assertEqual(len(result), 1)
        
        # 重新加载数据
        self.handler.load_excel(self.temp_file.name)
        
        # 测试包含匹配
        self.handler.set_filter_strategy('contains')
        success, result = self.handler.filter_data(['Name'], 'li')  # 匹配Alice和Charlie
        self.assertTrue(success)
        self.assertEqual(len(result), 2)
    
    def test_batch_filter_and_logic(self):
        """测试批量筛选AND逻辑"""
        self.handler.load_excel(self.temp_file.name)
        success, result = self.handler.filter_data_batch(
            ['Department', 'City'], ['IT', 'New York'], 'AND'
        )
        
        self.assertTrue(success)
        self.assertIsInstance(result, dict)
        # 应该只有Alice匹配（IT部门且在New York）
        filtered_data = list(result.values())[0]
        self.assertEqual(len(filtered_data), 1)
        self.assertEqual(filtered_data.iloc[0]['Name'], 'Alice')
    
    def test_batch_filter_or_logic(self):
        """测试批量筛选OR逻辑"""
        self.handler.load_excel(self.temp_file.name)
        success, result = self.handler.filter_data_batch(
            ['Department'], ['IT', 'HR'], 'OR'
        )
        
        self.assertTrue(success)
        filtered_data = list(result.values())[0]
        self.assertEqual(len(filtered_data), 4)  # 3个IT + 1个HR
    
    def test_reset_data(self):
        """测试数据重置"""
        self.handler.load_excel(self.temp_file.name)
        original_count = len(self.handler.dataframe)
        
        # 执行筛选
        self.handler.filter_data(['Department'], 'IT')
        self.assertLess(len(self.handler.dataframe), original_count)
        
        # 重置数据
        success = self.handler.reset_data()
        self.assertTrue(success)
        self.assertEqual(len(self.handler.dataframe), original_count)
        self.assertEqual(len(self.handler.filtered_sheets), 0)
    
    def test_get_data_summary(self):
        """测试数据摘要"""
        self.handler.load_excel(self.temp_file.name)
        summary = self.handler.get_data_summary()
        
        self.assertEqual(summary['original_rows'], 5)
        self.assertEqual(summary['current_rows'], 5)
        self.assertEqual(summary['filtered_sheets_count'], 0)
        self.assertEqual(summary['total_filtered_rows'], 0)
        self.assertEqual(len(summary['columns']), 4)


class TestDataValidator(unittest.TestCase):
    """数据验证器测试类"""
    
    def test_validate_file_path(self):
        """测试文件路径验证"""
        # 测试空路径
        is_valid, error = DataValidator.validate_file_path("")
        self.assertFalse(is_valid)
        self.assertIn("不能为空", error)
        
        # 测试不存在的文件
        is_valid, error = DataValidator.validate_file_path("nonexistent.xlsx")
        self.assertFalse(is_valid)
        self.assertIn("不存在", error)
    
    def test_validate_filter_text(self):
        """测试筛选文本验证"""
        # 测试空文本
        is_valid, error = DataValidator.validate_filter_text("")
        self.assertFalse(is_valid)
        
        # 测试正常文本
        is_valid, error = DataValidator.validate_filter_text("test")
        self.assertTrue(is_valid)
        self.assertIsNone(error)
        
        # 测试过长文本
        long_text = "a" * 1001
        is_valid, error = DataValidator.validate_filter_text(long_text)
        self.assertFalse(is_valid)
        self.assertIn("过长", error)
    
    def test_sanitize_sheet_name(self):
        """测试工作表名称清理"""
        # 测试包含非法字符的名称
        sanitized = DataValidator.sanitize_sheet_name("test/name*with?chars")
        self.assertNotIn("/", sanitized)
        self.assertNotIn("*", sanitized)
        self.assertNotIn("?", sanitized)
        
        # 测试过长名称
        long_name = "a" * 50
        sanitized = DataValidator.sanitize_sheet_name(long_name)
        self.assertLessEqual(len(sanitized), 31)


class TestFilterStrategies(unittest.TestCase):
    """筛选策略测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.test_df = pd.DataFrame({
            'text': ['Hello World', 'hello world', 'HELLO', 'world', 'test']
        })
    
    def test_contains_filter(self):
        """测试包含筛选"""
        filter_strategy = ContainsFilter()
        mask = filter_strategy.apply_filter(self.test_df, ['text'], 'hello')
        
        # 应该匹配前3行（不区分大小写）
        self.assertEqual(mask.sum(), 3)
    
    def test_exact_match_filter(self):
        """测试精确匹配筛选"""
        filter_strategy = ExactMatchFilter()
        mask = filter_strategy.apply_filter(self.test_df, ['text'], 'hello world')
        
        # 应该匹配前2行（不区分大小写）
        self.assertEqual(mask.sum(), 2)


if __name__ == '__main__':
    unittest.main()

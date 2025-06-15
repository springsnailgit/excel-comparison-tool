#!/usr/bin/env python3
"""运行测试脚本"""
import sys
import os
import unittest
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def run_tests():
    """运行所有测试"""
    print("开始运行测试...")
    
    # 发现并运行测试
    loader = unittest.TestLoader()
    start_dir = project_root / 'tests'
    suite = loader.discover(start_dir, pattern='test_*.py')
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # 输出结果
    if result.wasSuccessful():
        print(f"\n✅ 所有测试通过！运行了 {result.testsRun} 个测试")
        return True
    else:
        print(f"\n❌ 测试失败！{len(result.failures)} 个失败，{len(result.errors)} 个错误")
        return False

def check_code_quality():
    """检查代码质量"""
    print("\n检查代码质量...")
    
    # 检查是否有语法错误
    try:
        import src.main
        import src.excel_handler
        import src.config
        import src.utils.logger
        import src.utils.validators
        import src.utils.exceptions
        print("✅ 所有模块导入成功")
        return True
    except Exception as e:
        print(f"❌ 模块导入失败: {e}")
        return False

def main():
    """主函数"""
    print("=" * 50)
    print("Excel数据比对工具 - 测试和质量检查")
    print("=" * 50)
    
    # 检查代码质量
    quality_ok = check_code_quality()
    
    # 运行测试
    tests_ok = run_tests()
    
    # 总结
    print("\n" + "=" * 50)
    if quality_ok and tests_ok:
        print("🎉 所有检查通过！代码质量良好。")
        sys.exit(0)
    else:
        print("⚠️  存在问题，请检查上述输出。")
        sys.exit(1)

if __name__ == "__main__":
    main()

#!/usr/bin/env python
"""
测试运行器 - 运行项目中的所有测试
"""
import os
import sys
import pytest

def main():
    """运行测试并生成报告"""
    # 设置命令行参数
    args = [
        # 测试目录
        'tests',
        # 显示详细输出
        '-v',
        # 显示测试覆盖率报告
        '--cov=utils',
        # 在控制台中显示测试覆盖率报告
        '--cov-report=term',
        # 生成HTML格式的测试覆盖率报告
        '--cov-report=html:reports/coverage',
    ]
    
    # 添加命令行参数 
    args.extend(sys.argv[1:])
    
    # 运行pytest
    return pytest.main(args)

if __name__ == "__main__":
    # 确保reports目录存在
    os.makedirs('reports/coverage', exist_ok=True)
    
    # 运行测试
    sys.exit(main()) 
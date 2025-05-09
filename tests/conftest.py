"""
测试配置文件 - 用于设置测试环境和提供共享测试夹具
"""

import os
import sys
import pytest

# 添加父目录到sys.path，以便测试可以导入项目模块
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# 定义常用的测试目录路径
TEST_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_DOCS_DIR = os.path.join(TEST_DIR, 'test_docs')
TEST_OUTPUTS_DIR = os.path.join(TEST_DIR, 'test_outputs')

# 确保测试输出目录存在
os.makedirs(TEST_OUTPUTS_DIR, exist_ok=True)

@pytest.fixture
def test_document_path():
    """提供测试文档路径的夹具"""
    return os.path.join(TEST_DOCS_DIR, 'test_document.docx')

@pytest.fixture
def complex_test_document_path():
    """提供复杂测试文档路径的夹具"""
    return os.path.join(TEST_DOCS_DIR, 'complex_test_document.docx')

@pytest.fixture
def test_lists_document_path():
    """提供测试列表文档路径的夹具"""
    return os.path.join(TEST_DOCS_DIR, 'test_lists_simple.docx')

@pytest.fixture
def test_output_path():
    """提供测试输出文件路径的夹具"""
    return os.path.join(TEST_OUTPUTS_DIR, 'test_output.md')

@pytest.fixture
def complex_test_output_path():
    """提供复杂测试输出文件路径的夹具"""
    return os.path.join(TEST_OUTPUTS_DIR, 'complex_test_output.md')

@pytest.fixture
def test_output_images_dir():
    """提供测试输出图片目录的夹具"""
    images_dir = os.path.join(TEST_OUTPUTS_DIR, 'test_output_images')
    os.makedirs(images_dir, exist_ok=True)
    return images_dir 
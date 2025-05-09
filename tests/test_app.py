import os
import sys
import unittest

# 确保能找到应用模块
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

class TestAppFunctionality(unittest.TestCase):
    """测试应用的基本功能"""
    
    def test_app_imports(self):
        """测试应用模块能否正确导入"""
        try:
            import app
            from utils import docx_to_md, pdf_to_md, gitee_uploader
            self.assertTrue(True)
        except ImportError as e:
            self.fail(f"导入模块失败: {str(e)}")
    
    def test_directories_exist(self):
        """测试必要的目录是否存在"""
        required_dirs = ['uploads', 'outputs', 'outputs/images', 'templates', 'static', 'static/css', 'static/js']
        for dir_path in required_dirs:
            self.assertTrue(os.path.isdir(dir_path), f"目录 {dir_path} 不存在")
    
    def test_static_files_exist(self):
        """测试静态文件是否存在"""
        static_files = ['static/css/style.css', 'static/js/script.js', 'templates/index.html']
        for file_path in static_files:
            self.assertTrue(os.path.isfile(file_path), f"文件 {file_path} 不存在")

if __name__ == '__main__':
    unittest.main() 
import os
import shutil
import pytest
from utils.docx_to_md import convert_docx_to_md

def test_docx_conversion(test_document_path, test_output_path, test_output_images_dir):
    """测试docx到markdown的转换功能"""
    # 清理旧的测试结果
    if os.path.exists(test_output_path):
        os.remove(test_output_path)
    
    # 清理旧的图片目录
    if os.path.exists(test_output_images_dir):
        shutil.rmtree(test_output_images_dir)
        os.makedirs(test_output_images_dir, exist_ok=True)
    
    # 转换文档
    image_paths = convert_docx_to_md(test_document_path, test_output_path, test_output_images_dir)
    
    # 检查结果
    assert os.path.exists(test_output_path), "输出文件应该存在"
    
    # 检查图片是否被提取
    assert len(image_paths) > 0, "应该提取至少一张图片"
    for img_path in image_paths:
        assert os.path.exists(img_path), f"图片文件应该存在: {img_path}"
    
    # 检查输出内容是否正确
    with open(test_output_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
        # 检查基本格式元素存在
        assert '# ' in content, "应该包含一级标题"
        assert content.count('#') >= 1, "应该至少有一个标题"
        
        # 检查图片引用
        image_refs = content.count('![图片')
        assert image_refs > 0, "应该包含图片引用"
        assert image_refs == len(image_paths), "图片引用数应与提取的图片数匹配"

def test_complex_docx_conversion(complex_test_document_path, complex_test_output_path, test_output_images_dir):
    """测试复杂docx文档的转换功能"""
    # 清理旧的测试结果
    if os.path.exists(complex_test_output_path):
        os.remove(complex_test_output_path)
    
    # 清理旧的图片目录
    image_dir = os.path.join(os.path.dirname(complex_test_output_path), 'complex_test_output_images')
    if os.path.exists(image_dir):
        shutil.rmtree(image_dir)
        os.makedirs(image_dir, exist_ok=True)
    
    # 转换文档
    image_paths = convert_docx_to_md(complex_test_document_path, complex_test_output_path, image_dir)
    
    # 检查结果
    assert os.path.exists(complex_test_output_path), "输出文件应该存在"
    
    # 检查图片是否被提取
    assert len(image_paths) > 0, "应该提取至少一张图片"
    
    # 检查输出内容是否正确
    with open(complex_test_output_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
        # 检查是否包含标题层级
        assert '# ' in content, "应该包含一级标题"
        assert '## ' in content, "应该包含二级标题"
        assert content.count('#') >= 3, "应该至少有三个标题"
        
        # 检查是否包含列表
        assert '- ' in content, "应该包含无序列表"
        
        # 检查是否包含图片引用
        assert '![图片' in content, "应该包含图片引用"
        
        # 检查是否包含链接
        assert '[' in content and '](' in content, "应该包含链接"

if __name__ == "__main__":
    # 允许直接运行此测试文件
    pytest.main(['-v', __file__]) 
"""
使用本地docx2markdown库进行DOCX到Markdown的转换
这是一个可选的转换方式，作为主转换方法的替代
"""

import os
import logging
import sys
import importlib.util

logger = logging.getLogger(__name__)

def is_local_docx2md_available():
    """检查本地docx2markdown库是否可用"""
    try:
        # 检查项目根目录下的utils/docx2markdown目录
        docx2md_path = os.path.join('utils', 'docx2markdown')
        if os.path.exists(docx2md_path) and os.path.isdir(docx2md_path):
            converter_path = os.path.join(docx2md_path, 'docx_to_markdown_converter.py')
            if os.path.exists(converter_path):
                return True
    except Exception as e:
        logger.error(f"检查本地docx2markdown时出错: {str(e)}")
    return False

def convert_docx_to_md_cyrus(docx_path, output_path, image_dir=None):
    """
    使用本地docx2markdown库将DOCX文件转换为Markdown
    
    Args:
        docx_path (str): DOCX文件路径
        output_path (str): 输出的Markdown文件路径
        image_dir (str, optional): 图片保存目录，此参数在此方法中会被忽略，
                                  因为docx2markdown自动处理图片
    
    Returns:
        list: 包含图片路径的列表
    """
    try:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        if not is_local_docx2md_available():
            logger.error("未找到本地docx2markdown库，转换失败")
            return []
        
        # 使用本地docx2markdown库
        sys.path.append(os.path.abspath('utils'))
        from docx2markdown.docx_to_markdown_converter import docx_to_markdown
        
        # 执行转换
        logger.info(f"使用本地CYRUS方法转换DOCX文件: {docx_path}")
        docx_to_markdown(docx_path, output_path)
        
        # 获取生成的图片路径
        # 图片通常保存在output_path所在目录的images文件夹中
        image_output_dir = os.path.join(os.path.dirname(output_path), 'images')
        image_paths = []
        
        if os.path.exists(image_output_dir):
            logger.info(f"找到图片目录: {image_output_dir}")
            for img_file in os.listdir(image_output_dir):
                if img_file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    img_path = os.path.join(image_output_dir, img_file)
                    image_paths.append(img_path)
        
        logger.info(f"本地CYRUS方法转换完成，共提取了 {len(image_paths)} 张图片")
        return image_paths
            
    except Exception as e:
        logger.error(f"使用本地docx2markdown转换出错: {str(e)}", exc_info=True)
        return [] 
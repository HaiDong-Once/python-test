"""
DOCX到Markdown转换接口
支持选择不同的转换方式
"""

import os
import logging
from enum import Enum

# 导入两种转换方法
from utils.docx_to_md import convert_docx_to_md as convert_default
from utils.cyrus_docx_converter import convert_docx_to_md_cyrus as convert_cyrus

logger = logging.getLogger(__name__)

class ConversionMethod(Enum):
    """转换方法枚举"""
    DEFAULT = "default"  # 默认方法
    CYRUS = "cyrus"      # CYRUS-STUDIO方法

def convert_docx_to_markdown(
    docx_path, 
    output_path, 
    method=ConversionMethod.DEFAULT,
    image_dir=None
):
    """
    DOCX到Markdown转换主接口
    
    Args:
        docx_path (str): DOCX文件路径
        output_path (str): 输出的Markdown文件路径
        method (ConversionMethod): 转换方法，默认使用自带方法
        image_dir (str, optional): 图片保存目录
    
    Returns:
        list: 包含图片路径的列表
    """
    logger.info(f"使用 {method.value} 方法转换 {docx_path} 到 {output_path}")
    
    if not os.path.exists(docx_path):
        logger.error(f"DOCX文件不存在: {docx_path}")
        return []
    
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 根据选择的方法执行转换
    if method == ConversionMethod.CYRUS:
        logger.info("使用本地CYRUS方法转换")
        return convert_cyrus(docx_path, output_path, image_dir)
    else:
        logger.info("使用默认方法进行转换")
        return convert_default(docx_path, output_path, image_dir)

def get_available_methods():
    """
    获取可用的转换方法
    
    Returns:
        list: 可用转换方法列表
    """
    methods = [ConversionMethod.DEFAULT]
    
    # 检查本地CYRUS实现是否可用
    try:
        from utils.cyrus_docx_converter import is_local_docx2md_available
        if is_local_docx2md_available():
            logger.info("发现本地CYRUS方法")
            methods.append(ConversionMethod.CYRUS)
        else:
            logger.info("本地CYRUS方法不可用")
    except ImportError:
        logger.warning("无法导入CYRUS模块")
    
    return methods 
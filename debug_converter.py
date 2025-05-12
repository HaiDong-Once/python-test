import os
import sys
import logging
import traceback
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_cyrus_converter():
    """测试Cyrus方法的转换"""
    try:
        logger.info("开始测试Cyrus转换方法")
        
        # 确保脚本路径在sys.path中
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
        
        # 检查上传目录中是否有docx文件
        uploads_dir = "uploads"
        if not os.path.exists(uploads_dir):
            logger.error(f"上传目录 {uploads_dir} 不存在")
            return
        
        # 查找第一个docx文件
        docx_file = None
        for file in os.listdir(uploads_dir):
            if file.lower().endswith(".docx"):
                docx_file = os.path.join(uploads_dir, file)
                break
        
        if not docx_file:
            logger.error("未找到任何docx文件用于测试")
            return
        
        logger.info(f"找到测试文件: {docx_file}")
        
        # 设置输出路径
        output_dir = "debug_outputs"
        os.makedirs(output_dir, exist_ok=True)
        
        filename = os.path.basename(docx_file)
        filename_without_ext = os.path.splitext(filename)[0]
        output_path = os.path.join(output_dir, f"{filename_without_ext}.md")
        
        # 执行转换
        logger.info(f"开始转换 {docx_file} 到 {output_path}")
        
        try:
            # 直接使用相对路径访问模块文件
            sys.path.insert(0, os.path.join(current_dir, 'utils', 'docx2markdown'))
            
            # 直接从文件导入
            from docx_parser import DocxParser
            from docx_to_markdown_converter import DocxToMarkdownConverter, docx_to_markdown
            
            # 检查DocxParser是否有extract_image方法
            parser = DocxParser(docx_file)
            logger.info(f"DocxParser方法: {dir(parser)}")
            
            # 检查DocxToMarkdownConverter的初始化方法
            logger.info(f"DocxToMarkdownConverter初始化参数: {DocxToMarkdownConverter.__init__.__code__.co_varnames}")
            
            # 直接使用已有的docx_to_markdown函数
            logger.info("使用docx_to_markdown函数")
            markdown_content = docx_to_markdown(docx_file, output_path)
            logger.info(f"已生成Markdown文件: {output_path}")
            
        except Exception as direct_error:
            logger.error(f"直接转换失败: {str(direct_error)}")
            traceback.print_exc()
            
            # 尝试作为绝对路径导入
            logger.info("尝试使用绝对路径导入")
            try:
                from utils.docx2markdown.docx_parser import DocxParser
                from utils.docx2markdown.docx_to_markdown_converter import DocxToMarkdownConverter, docx_to_markdown
                
                # 检查DocxParser是否有extract_image方法
                parser = DocxParser(docx_file)
                logger.info(f"DocxParser方法: {dir(parser)}")
                
                # 检查DocxToMarkdownConverter的初始化方法
                logger.info(f"DocxToMarkdownConverter初始化参数: {DocxToMarkdownConverter.__init__.__code__.co_varnames}")
                
                # 直接使用docx_to_markdown函数
                logger.info("使用docx_to_markdown函数")
                markdown_content = docx_to_markdown(docx_file, output_path)
                logger.info(f"已生成Markdown文件: {output_path}")
                
            except Exception as absolute_error:
                logger.error(f"绝对路径导入也失败: {str(absolute_error)}")
                traceback.print_exc()
                
                # 尝试修改文件
                logger.info("尝试手动修改文件路径并导入")
                try:
                    # 确保在Python路径中
                    if 'utils' not in sys.path:
                        sys.path.insert(0, os.path.join(current_dir, 'utils'))
                    
                    # 导入Cyrus转换方法作为最后的尝试
                    from utils.cyrus_docx_converter import is_local_docx2md_available, convert_docx_to_md_cyrus
                    
                    if is_local_docx2md_available():
                        logger.info("本地docx2markdown库可用，尝试使用Cyrus方法")
                        image_paths = convert_docx_to_md_cyrus(docx_file, output_path)
                        logger.info(f"Cyrus方法转换成功，提取到 {len(image_paths)} 张图片")
                        for img in image_paths:
                            logger.info(f"  - {img}")
                    else:
                        logger.error("本地docx2markdown库不可用")
                except Exception as cyrus_error:
                    logger.error(f"Cyrus方法也失败: {str(cyrus_error)}")
                    traceback.print_exc()
    
    except Exception as e:
        logger.error(f"测试过程出错: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    logger.info("开始调试转换器")
    test_cyrus_converter()
    logger.info("调试完成") 
"""
修复版的Cyrus转换器
"""

import os
import sys
import logging
import traceback
import re
import zipfile

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 导入修复模块
import fix_docx_converter

# 确保修补已应用
fix_docx_converter.fix_all()

# Gitee基础URL
GITEE_BASE_URL = "https://gitee.com/comma-dong/image-projects/raw/master/"

def fix_cyrus_converter():
    """
    执行Cyrus转换器的修复操作
    
    Returns:
        bool: 是否成功修复
    """
    try:
        # 导入修复所需的模块
        sys.path.append(os.path.abspath('utils'))
        
        # 应用DocxParser的修复
        fix_docx_converter.patch_docx_parser()
        
        # 应用DocxToMarkdownConverter的修复
        fix_docx_converter.patch_docx_to_markdown_converter()
        
        logger.info("Cyrus转换器修复完成")
        return True
    except Exception as e:
        logger.error(f"Cyrus转换器修复失败: {str(e)}")
        traceback.print_exc()
        return False

def convert_docx_to_md(docx_path, output_path, image_dir=None):
    """
    使用修复版的docx2markdown库将DOCX文件转换为Markdown
    
    Args:
        docx_path (str): DOCX文件路径
        output_path (str): 输出的Markdown文件路径
        image_dir (str, optional): 图片保存目录
    
    Returns:
        list: 包含图片路径的列表
    """
    try:
        logger.info(f"使用修复版的Cyrus方法转换文件: {docx_path}")
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 导入需要的类
        from utils.docx2markdown.docx_parser import DocxParser
        
        # 确保DocxParser类有extract_image方法
        if not hasattr(DocxParser, 'extract_image'):
            fix_docx_converter.patch_docx_parser()
        
        # 获取文档名，用于创建图片目录
        doc_name = os.path.basename(output_path).split('.')[0]
        
        # 配置图片目录
        if image_dir is None:
            image_output_dir = os.path.join(os.path.dirname(output_path), doc_name + "_outputs")
        else:
            image_output_dir = image_dir
            
        # 确保图片目录存在
        os.makedirs(image_output_dir, exist_ok=True)
        
        # 创建解析器并提取所有媒体文件
        parser = DocxParser(docx_path)
        parser.extract_media(image_output_dir)
        
        # 解析文档
        document = parser.parse()
        
        # 收集提取的图片路径
        image_paths = []
        if os.path.exists(image_output_dir):
            for img_file in os.listdir(image_output_dir):
                if img_file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    img_path = os.path.join(image_output_dir, img_file)
                    image_paths.append(img_path)
        
        logger.info(f"从文档中提取到 {len(image_paths)} 张图片")
        
        # 生成Markdown内容
        markdown_content = []
        
        # 处理每个元素
        for element in document['elements']:
            if hasattr(element, 'text'):  # 如果是段落
                # 处理文本格式
                text = element.text
                style = element.style
                
                # 处理加粗、斜体、下划线
                if style.bold:
                    text = f"**{text}**"
                if style.italic:
                    text = f"*{text}*"
                if style.underline:
                    text = f"_{text}_"
                
                # 检查是否是标题
                if style.fonts.get("default", None):
                    try:
                        heading_level = int(style.fonts["default"])
                        if 1 <= heading_level <= 6:  # 1-6 级标题有效
                            text = f"{'#' * heading_level} {text}"
                    except (ValueError, TypeError):
                        pass
                
                # 添加段落文本
                markdown_content.append(text)
                
                # 处理图片
                if element.image:
                    image_name = os.path.basename(element.image['file'])
                    folder_name = doc_name + "_outputs"
                    remote_url = f"{GITEE_BASE_URL}{folder_name}/{image_name}"
                    markdown_content.append(f"\n![{image_name}]({remote_url})\n")
            
            elif hasattr(element, 'rows'):  # 如果是表格
                # 简单处理表格
                for row in element.rows:
                    markdown_content.append("| " + " | ".join(row) + " |")
                markdown_content.append("")
        
        # 合并所有内容
        final_content = "\n".join(markdown_content)
        
        # 写入输出文件
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_content)
        
        logger.info(f"已生成Markdown文件: {output_path}")
        
        # 更新图片链接
        _update_image_links(output_path, doc_name + "_outputs")
        
        return image_paths
    
    except Exception as e:
        logger.error(f"转换失败: {str(e)}")
        traceback.print_exc()
        return []

def _update_image_links(md_file_path, image_folder_name):
    """
    更新Markdown文件中的图片链接为Gitee远程链接
    
    Args:
        md_file_path (str): Markdown文件路径
        image_folder_name (str): 图片文件夹名称
    """
    try:
        # 读取Markdown文件内容
        with open(md_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 计数转换的图片链接数量
        replaced_count = 0
        
        # 替换所有本地图片链接为远程链接
        # 匹配Markdown图片语法 ![alt](path)，但排除已经是远程链接或base64图片的情况
        pattern = r'!\[(.*?)\]\(((?!data:image|http|https).*?)\)'
        
        def replace_link(match):
            nonlocal replaced_count
            alt_text = match.group(1)
            img_path = match.group(2)
            
            # 获取图片文件名
            img_filename = os.path.basename(img_path)
            
            # 创建新的远程链接
            remote_url = f"{GITEE_BASE_URL}{image_folder_name}/{img_filename}"
            
            # 增加计数
            replaced_count += 1
            
            # 返回替换后的链接
            return f"![{alt_text}]({remote_url})"
        
        # 替换所有图片链接
        new_content = re.sub(pattern, replace_link, content)
        
        # 写回文件
        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(new_content)
            
        logger.info(f"已更新Markdown文件中的{replaced_count}个图片链接为远程链接")
        
    except Exception as e:
        logger.error(f"更新图片链接时出错: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    # 执行修复
    if fix_cyrus_converter():
        logger.info("Cyrus转换器修复成功")
    else:
        logger.error("Cyrus转换器修复失败")
    
    # 测试转换功能
    try:
        # 查找上传目录中的第一个docx文件
        uploads_dir = "uploads"
        if not os.path.exists(uploads_dir):
            logger.error(f"上传目录 {uploads_dir} 不存在")
            sys.exit(1)
        
        # 查找第一个docx文件
        docx_file = None
        for file in os.listdir(uploads_dir):
            if file.lower().endswith(".docx"):
                docx_file = os.path.join(uploads_dir, file)
                break
        
        if not docx_file:
            logger.error("未找到任何docx文件用于测试")
            sys.exit(1)
        
        logger.info(f"找到测试文件: {docx_file}")
        
        # 设置输出路径
        output_dir = "fix_outputs"
        os.makedirs(output_dir, exist_ok=True)
        
        filename = os.path.basename(docx_file)
        filename_without_ext = os.path.splitext(filename)[0]
        output_path = os.path.join(output_dir, f"{filename_without_ext}.md")
        
        # 执行转换
        logger.info(f"开始转换 {docx_file} 到 {output_path}")
        image_paths = convert_docx_to_md(docx_file, output_path)
        
        if os.path.exists(output_path):
            logger.info(f"转换成功! 输出文件: {output_path}")
            logger.info(f"共提取了 {len(image_paths)} 张图片")
        else:
            logger.error("转换失败，未生成输出文件")
    
    except Exception as e:
        logger.error(f"测试过程出错: {str(e)}")
        traceback.print_exc() 
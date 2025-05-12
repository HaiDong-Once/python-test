"""
使用本地docx2markdown库进行DOCX到Markdown的转换
这是一个可选的转换方式，作为主转换方法的替代
"""

import os
import logging
import sys
import importlib.util
import re
import shutil
import zipfile

logger = logging.getLogger(__name__)

# Gitee基础URL
GITEE_BASE_URL = "https://gitee.com/comma-dong/image-projects/raw/master/"

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
        image_dir (str, optional): 图片保存目录，此参数将覆盖默认的图片目录
    
    Returns:
        list: 包含图片路径的列表
    """
    try:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        if not is_local_docx2md_available():
            logger.error("未找到本地docx2markdown库，转换失败")
            return []
        
        # 获取文档名，用于创建图片目录
        doc_name = os.path.basename(output_path).split('.')[0]
        
        # 配置图片目录
        if image_dir is None:
            # 创建与markdown文件同级的图片目录
            image_output_dir = os.path.join(os.path.dirname(output_path), doc_name + "_outputs")
        else:
            image_output_dir = image_dir
            
        # 确保图片目录存在
        os.makedirs(image_output_dir, exist_ok=True)
        
        # 执行转换
        logger.info(f"使用本地CYRUS方法转换DOCX文件: {docx_path}")
        logger.info(f"图片将保存在: {image_output_dir}")

        # 导入所需模块
        try:
            # 使用docx2markdown库
            sys.path.append(os.path.abspath('utils'))
            from docx2markdown.docx_parser import DocxParser
            from docx2markdown.docx_to_markdown_converter import DocxToMarkdownConverter, docx_to_markdown
            
            # 创建解析器并提取所有媒体文件
            parser = DocxParser(docx_path)
            
            # 确保DocxParser有extract_image方法
            if not hasattr(parser, 'extract_image'):
                logger.warning("DocxParser缺少extract_image方法，尝试动态添加...")
                
                def extract_image(self, image_path, output_path):
                    """
                    从.docx文件中提取指定的图片并保存到指定路径
                    """
                    try:
                        logger.info(f"尝试提取图片: {image_path} 到 {output_path}")
                        with zipfile.ZipFile(self.file_path, 'r') as docx_zip:
                            # 获取所有媒体文件列表
                            media_files = [name for name in docx_zip.namelist() if name.startswith('word/media/')]
                            logger.info(f"文档中的媒体文件数量: {len(media_files)}")
                            
                            # 检查文件是否存在
                            if image_path in docx_zip.namelist():
                                # 获取图片数据
                                image_data = docx_zip.read(image_path)
                                
                                # 确保输出目录存在
                                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                
                                # 保存图片
                                with open(output_path, 'wb') as img_file:
                                    img_file.write(image_data)
                                logger.info(f"成功提取图片: {image_path} 到 {output_path}")
                                return True
                            else:
                                # 尝试通过基本名称匹配
                                base_name = os.path.basename(image_path)
                                for name in media_files:
                                    if name.endswith(base_name) or base_name in name:
                                        logger.info(f"找到匹配文件: {name}")
                                        # 获取图片数据
                                        image_data = docx_zip.read(name)
                                        
                                        # 确保输出目录存在
                                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                        
                                        # 保存图片
                                        with open(output_path, 'wb') as img_file:
                                            img_file.write(image_data)
                                        logger.info(f"通过匹配文件名成功提取图片: {name} 到 {output_path}")
                                        return True
                                
                                # 如果找不到匹配的图片，尝试从media文件夹中提取第一个图片
                                if media_files:
                                    logger.info(f"未找到匹配的图片，尝试提取第一个媒体文件: {media_files[0]}")
                                    image_data = docx_zip.read(media_files[0])
                                    
                                    # 确保输出目录存在
                                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                    
                                    # 保存图片
                                    with open(output_path, 'wb') as img_file:
                                        img_file.write(image_data)
                                    logger.info(f"已提取替代图片: {media_files[0]} 到 {output_path}")
                                    return True
                                
                                logger.info(f"无法在文档中找到任何可用图片")
                                return False
                    except Exception as e:
                        logger.error(f"提取图片时出错: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        return False
                
                # 动态添加方法
                DocxParser.extract_image = extract_image
                logger.info("已动态添加extract_image方法")
                
            # 提取所有媒体文件到图片目录
            parser.extract_media(image_output_dir)
            
            # 尝试使用DocxToMarkdownConverter进行转换
            try:
                # 检查DocxToMarkdownConverter是否接受output_path参数
                if 'output_path' in importlib.import_module('utils.docx2markdown.docx_to_markdown_converter').__dict__['DocxToMarkdownConverter'].__init__.__code__.co_varnames:
                    logger.info("使用DocxToMarkdownConverter进行转换...")
                    converter = DocxToMarkdownConverter(docx_path, output_path)
                    markdown_content = converter.convert()
                    
                    # 将Markdown内容写入文件
                    with open(output_path, "w", encoding="utf-8") as f:
                        f.write(markdown_content)
                else:
                    # 使用docx_to_markdown函数
                    logger.info("使用docx_to_markdown函数进行转换...")
                    docx_to_markdown(docx_path, output_path)
                    
                logger.info(f"转换完成: {output_path}")
            except Exception as e:
                logger.error(f"使用DocxToMarkdownConverter转换失败: {str(e)}")
                logger.info("尝试手动解析并生成Markdown...")
                
                # 解析文档
                document = parser.parse()
                
                # 手动生成Markdown
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
                            # 尝试直接从文档中提取图片
                            new_image_name = f"image_{len(markdown_content)}{os.path.splitext(image_name)[1]}"
                            output_img_path = os.path.join(image_output_dir, new_image_name)
                            
                            if parser.extract_image(element.image['file'], output_img_path):
                                # 构建Gitee远程URL
                                folder_name = doc_name + "_outputs"
                                remote_url = f"{GITEE_BASE_URL}{folder_name}/{new_image_name}"
                                markdown_content.append(f"\n![{new_image_name}]({remote_url})\n")
                    
                    elif hasattr(element, 'rows'):  # 如果是表格
                        # 简单处理表格
                        for row in element.rows:
                            markdown_content.append("| " + " | ".join(row) + " |")
                        markdown_content.append("")
                
                # 合并所有内容并写入文件
                final_content = "\n".join(markdown_content)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(final_content)
                
                logger.info(f"手动生成Markdown完成: {output_path}")
        
        except ImportError as e:
            logger.error(f"导入模块失败: {str(e)}")
            return []
        
        # 获取图片路径列表
        image_paths = []
        if os.path.exists(image_output_dir):
            for img_file in os.listdir(image_output_dir):
                if img_file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    img_path = os.path.join(image_output_dir, img_file)
                    image_paths.append(img_path)
        
        # 打印提取的图片数量
        logger.info(f"共提取了 {len(image_paths)} 张图片")
        
        # 确保更新Markdown文件中的图片链接为Gitee远程链接
        if image_paths:
            _update_image_links(output_path, doc_name + "_outputs")
        
        return image_paths
            
    except Exception as e:
        logger.error(f"CYRUS转换出错: {str(e)}", exc_info=True)
        return []

def _update_image_links(md_file_path, image_folder_name):
    """
    更新Markdown文件中的图片链接为Gitee远程链接
    
    Args:
        md_file_path (str): Markdown文件路径
        image_folder_name (str): 图片文件夹名称
    """
    try:
        # 尝试导入base64图片提取函数
        try:
            from fix_docx_converter import extract_base64_images
            # 图片输出目录
            img_dir = os.path.join(os.path.dirname(md_file_path), image_folder_name)
            # 提取base64编码的图片
            base64_count = extract_base64_images(md_file_path, img_dir)
            logger.info(f"已提取{base64_count}张base64图片到{img_dir}")
        except ImportError:
            logger.warning("无法导入extract_base64_images函数，将跳过base64图片提取")
            base64_count = 0
        
        # 读取Markdown文件内容
        with open(md_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 计数转换的图片链接数量和base64图片数量
        replaced_count = 0
        
        # 替换所有base64编码的图片链接
        base64_pattern = r'!\[(.*?)\]\(data:image\/[^;]+;base64,[^\)]+\)'
        
        def replace_base64(match):
            nonlocal replaced_count
            alt_text = match.group(1)
            
            # 生成新的图片文件名
            new_image_name = f"base64_image_{replaced_count + 1}.png"
            
            # 创建远程URL
            remote_url = f"{GITEE_BASE_URL}{image_folder_name}/{new_image_name}"
            
            # 增加计数
            replaced_count += 1
            
            # 返回替换后的链接
            return f"![{alt_text}]({remote_url})"
        
        # 替换所有base64图片
        content = re.sub(base64_pattern, replace_base64, content)
        
        # 然后替换所有本地图片链接为远程链接
        # 匹配Markdown图片语法 ![alt](path)，但排除已经是远程链接的情况
        local_pattern = r'!\[(.*?)\]\(((?!data:image|http|https).*?)\)'
        
        local_replaced_count = 0
        def replace_local_link(match):
            nonlocal local_replaced_count
            alt_text = match.group(1)
            img_path = match.group(2)
            
            # 获取图片文件名
            img_filename = os.path.basename(img_path)
            
            # 创建新的远程链接
            remote_url = f"{GITEE_BASE_URL}{image_folder_name}/{img_filename}"
            
            # 增加计数
            local_replaced_count += 1
            
            # 返回替换后的链接
            return f"![{alt_text}]({remote_url})"
        
        # 替换所有本地图片链接
        content = re.sub(local_pattern, replace_local_link, content)
        
        # 写回文件
        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(content)
            
        logger.info(f"已更新Markdown文件中的图片链接，替换了{replaced_count}个base64图片和{local_replaced_count}个本地链接")
        
    except Exception as e:
        logger.error(f"更新图片链接时出错: {str(e)}", exc_info=True) 
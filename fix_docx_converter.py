"""
动态修复docx2markdown模块
"""

import os
import sys
import inspect
import logging
import traceback

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def patch_docx_parser():
    """动态给DocxParser类添加extract_image方法"""
    try:
        logger.info("准备给DocxParser类添加extract_image方法")
        
        # 导入需要修补的类
        from utils.docx2markdown.docx_parser import DocxParser
        import zipfile
        import os
        
        # 定义要添加的方法
        def extract_image(self, image_path, output_path):
            """
            从.docx文件中提取指定的图片并保存到指定路径
            
            Args:
                image_path (str): docx中的图片路径，例如: 'word/media/image1.png'
                output_path (str): 图片保存的目标路径
            
            Returns:
                bool: 是否成功提取图片
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
                        logger.info(f"图片 {image_path} 不存在于 .docx 文件中，尝试匹配文件名...")
                        
                        # 尝试通过基本名称匹配
                        base_name = os.path.basename(image_path)
                        for name in docx_zip.namelist():
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
                traceback.print_exc()
                return False
        
        # 将方法添加到类中
        DocxParser.extract_image = extract_image
        
        # 验证添加成功
        if hasattr(DocxParser, 'extract_image'):
            logger.info("成功添加extract_image方法到DocxParser类")
            return True
        else:
            logger.error("添加方法失败")
            return False
    
    except Exception as e:
        logger.error(f"修补DocxParser时出错: {str(e)}")
        traceback.print_exc()
        return False

def patch_docx_to_markdown_converter():
    """修补DocxToMarkdownConverter类，确保正确处理output_path参数"""
    try:
        logger.info("准备修补DocxToMarkdownConverter类")
        
        # 导入需要修补的类
        from utils.docx2markdown.docx_to_markdown_converter import DocxToMarkdownConverter
        
        # 检查现有的__init__方法
        if 'output_path' in inspect.signature(DocxToMarkdownConverter.__init__).parameters:
            logger.info("DocxToMarkdownConverter已经接受output_path参数，无需修补")
            return True
        
        # 保存原始方法
        original_init = DocxToMarkdownConverter.__init__
        
        # 定义新的初始化方法
        def new_init(self, docx_file, output_path=None):
            original_init(self, docx_file)
            self.output_path = output_path
            self.image_count = 0  # 用于计数和生成图片文件名
            self.extracted_images = []  # 保存提取的图片信息
            
            # 获取文档名，用于创建图片目录
            if output_path:
                self.doc_name = os.path.basename(output_path).split('.')[0]
                # 创建图片目录
                self.img_dir = os.path.join(os.path.dirname(output_path), self.doc_name + "_outputs")
                os.makedirs(self.img_dir, exist_ok=True)
        
        # 替换初始化方法
        DocxToMarkdownConverter.__init__ = new_init
        
        # 验证修补成功
        if 'output_path' in inspect.signature(DocxToMarkdownConverter.__init__).parameters:
            logger.info("成功修补DocxToMarkdownConverter.__init__方法")
            return True
        else:
            logger.error("修补DocxToMarkdownConverter.__init__方法失败")
            return False
    
    except Exception as e:
        logger.error(f"修补DocxToMarkdownConverter时出错: {str(e)}")
        traceback.print_exc()
        return False

def extract_base64_images(md_file_path, output_folder):
    """
    从Markdown文件中提取base64编码的图片并保存为文件
    
    Args:
        md_file_path (str): Markdown文件路径
        output_folder (str): 图片保存目录
    
    Returns:
        int: 提取的图片数量
    """
    try:
        logger.info(f"从Markdown中提取base64图片: {md_file_path}")
        
        import re
        import base64
        import os
        
        # 确保输出目录存在
        os.makedirs(output_folder, exist_ok=True)
        
        # 读取Markdown文件内容
        with open(md_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 查找所有base64编码的图片
        pattern = r'!\[(.*?)\]\(data:image\/([^;]+);base64,([^\)]+)\)'
        matches = re.findall(pattern, content)
        
        if not matches:
            logger.info("未找到base64编码的图片")
            return 0
        
        # 提取图片并保存到文件
        count = 0
        for alt_text, img_type, base64_data in matches:
            try:
                # 清理base64数据
                base64_data = base64_data.strip()
                
                # 生成图片文件名
                img_filename = f"base64_image_{count + 1}.{img_type}"
                img_path = os.path.join(output_folder, img_filename)
                
                # 解码base64数据并保存
                img_data = base64.b64decode(base64_data)
                with open(img_path, 'wb') as img_file:
                    img_file.write(img_data)
                
                logger.info(f"已保存base64图片: {img_filename}")
                count += 1
            except Exception as e:
                logger.error(f"处理base64图片时出错: {str(e)}")
                logger.error(f"Alt文本: {alt_text}, 图片类型: {img_type}")
        
        logger.info(f"共提取了{count}张base64图片到 {output_folder}")
        return count
    
    except Exception as e:
        logger.error(f"提取base64图片时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return 0

def fix_all():
    """应用所有修补"""
    # 清除缓存
    try:
        for module in list(sys.modules.keys()):
            if module.startswith('utils.docx2markdown'):
                del sys.modules[module]
        logger.info("已清除相关模块缓存")
    except:
        pass
    
    # 应用修补
    parser_patched = patch_docx_parser()
    converter_patched = patch_docx_to_markdown_converter()
    
    return parser_patched and converter_patched

if __name__ == "__main__":
    logger.info("开始修补docx2markdown模块")
    if fix_all():
        logger.info("修补成功完成")
    else:
        logger.error("修补失败") 
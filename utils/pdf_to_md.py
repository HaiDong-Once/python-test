import os
import re
import PyPDF2
# 修改导入方式，确保兼容性
try:
    import fitz  # PyMuPDF
except ImportError:
    try:
        import pymupdf as fitz
    except ImportError:
        raise ImportError("请安装PyMuPDF: pip install PyMuPDF")
from PIL import Image
from io import BytesIO

def extract_images_from_pdf(pdf_path, output_dir):
    """从PDF文件中提取图片并保存到指定目录"""
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    image_paths = []
    image_count = 0
    
    # 打开PDF文件
    doc = fitz.open(pdf_path)
    
    # 遍历每一页
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        image_list = page.get_images(full=True)
        
        # 提取每一页的图片
        for img_index, img in enumerate(image_list):
            image_count += 1
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            
            # 尝试确定图片格式
            try:
                img_format = base_image["ext"]
            except:
                img_format = "png"  # 默认使用png
            
            image_filename = f"image_{image_count}.{img_format}"
            image_path = os.path.join(output_dir, image_filename)
            
            # 保存图片
            with open(image_path, "wb") as f:
                f.write(image_bytes)
            
            # 保存图片路径，使用页码和图片索引作为引用
            ref = f"page_{page_num+1}_img_{img_index+1}"
            image_paths.append((ref, image_path))
    
    return image_paths

def identify_headings(text, font_sizes):
    """根据字体大小尝试识别标题"""
    # 按字体大小排序
    sorted_sizes = sorted(set(font_sizes), reverse=True)
    
    # 如果只有一种字体大小，则无法区分标题
    if len(sorted_sizes) <= 1:
        return text, 0
    
    # 假设最大的字体是标题，根据字体大小相对于最大字体的比例判断标题级别
    max_font = sorted_sizes[0]
    
    for i, size in enumerate(sorted_sizes):
        if size == font_sizes:
            # 根据字体大小确定标题级别
            if i == 0:  # 最大字体
                return text, 1
            elif i == 1:  # 第二大字体
                return text, 2
            elif i == 2:  # 第三大字体
                return text, 3
            else:
                return text, min(i+1, 6)  # 最多到h6
    
    # 如果没有匹配，则不是标题
    return text, 0

def detect_code_blocks(text):
    """尝试检测可能的代码块"""
    # 简单启发式方法：如果文本包含特定的编程关键字和符号，可能是代码块
    code_indicators = [
        '{', '}', '()', '[]', ';', 'function', 'class', 'def ', 'import ', 
        'var ', 'let ', 'const ', 'return ', 'if (', 'for (', 'while ('
    ]
    
    if sum(1 for ind in code_indicators if ind in text) >= 2:
        return f"```\n{text}\n```"
    return text

def convert_pdf_to_md(pdf_path, output_path, image_dir=None):
    """将PDF文件转换为markdown格式"""
    if image_dir is None:
        image_dir = os.path.join(os.path.dirname(output_path), 'images')
    
    # 提取图片
    image_paths = extract_images_from_pdf(pdf_path, image_dir)
    image_refs = {ref: path for ref, path in image_paths}
    
    # 打开PDF
    reader = PyPDF2.PdfReader(pdf_path)
    md_content = []
    
    # 使用PyMuPDF获取更多格式信息
    doc = fitz.open(pdf_path)
    
    # 处理每一页
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text = page.extract_text()
        
        # 使用PyMuPDF获取格式信息
        fitz_page = doc.load_page(page_num)
        blocks = fitz_page.get_text("dict")["blocks"]
        
        # 处理当前页的文本块
        for block in blocks:
            if "lines" not in block:
                continue
                
            for line in block["lines"]:
                line_text = ""
                font_sizes = []
                
                for span in line["spans"]:
                    span_text = span["text"]
                    font_size = span["size"]
                    font_flags = span["flags"]
                    
                    # 收集字体大小
                    font_sizes.append(font_size)
                    
                    # 处理粗体和斜体
                    if font_flags & 2:  # 粗体
                        span_text = f"**{span_text}**"
                    if font_flags & 1:  # 斜体
                        span_text = f"*{span_text}*"
                        
                    line_text += span_text
                
                # 处理可能的标题
                text, heading_level = identify_headings(line_text, max(font_sizes) if font_sizes else 0)
                
                if heading_level > 0:
                    md_content.append(f"{'#' * heading_level} {text}")
                else:
                    # 检测可能的代码块
                    text = detect_code_blocks(text)
                    
                    # 检测URL链接
                    url_pattern = r'(https?://\S+)'
                    text = re.sub(url_pattern, r'[\1](\1)', text)
                    
                    md_content.append(text)
    
    # 将提取的图片引用插入到Markdown中
    # 这里简单地将所有图片附加到文档末尾
    for ref, path in image_paths:
        filename = os.path.basename(path)
        md_content.append(f"\n![{ref}](images/{filename})\n")
    
    # 写入markdown文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(md_content))
    
    # 返回图片路径列表，以便后续处理
    return [path for _, path in image_paths] 
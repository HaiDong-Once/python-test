import os
import time
import logging
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort
from werkzeug.utils import secure_filename
import subprocess
import sys
import importlib.util
import socket
import inspect
import traceback

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 在应用启动前检查CYRUS本地库
try:
    logger.info("检查本地docx2markdown库...")
    
    # 检查是否已存在本地实现
    docx2md_path = os.path.join('utils', 'docx2markdown')
    if os.path.exists(docx2md_path) and os.path.isdir(docx2md_path):
        converter_path = os.path.join(docx2md_path, 'docx_to_markdown_converter.py')
        if os.path.exists(converter_path):
            logger.info("找到本地docx2markdown库")
            
            # 导入并检查是否需要修复
            try:
                # 清除相关模块缓存
                for module_name in list(sys.modules.keys()):
                    if module_name.startswith('utils.docx2markdown'):
                        del sys.modules[module_name]
                
                # 导入DocxParser类
                from utils.docx2markdown.docx_parser import DocxParser
                
                # 检查是否需要添加extract_image方法
                if not hasattr(DocxParser, 'extract_image'):
                    logger.info("DocxParser缺少extract_image方法，正在添加...")
                    import zipfile
                    
                    def extract_image(self, image_path, output_path):
                        """从.docx文件中提取指定的图片并保存到指定路径"""
                        try:
                            logger.info(f"尝试提取图片: {image_path} 到 {output_path}")
                            with zipfile.ZipFile(self.file_path, 'r') as docx_zip:
                                # 获取媒体文件列表
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
                                    # 尝试通过文件名匹配
                                    base_name = os.path.basename(image_path)
                                    for name in docx_zip.namelist():
                                        if name.endswith(base_name) or base_name in name:
                                            # 获取图片数据
                                            image_data = docx_zip.read(name)
                                            
                                            # 确保输出目录存在
                                            os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                            
                                            # 保存图片
                                            with open(output_path, 'wb') as img_file:
                                                img_file.write(image_data)
                                            logger.info(f"通过文件名匹配提取图片: {name} 到 {output_path}")
                                            return True
                                    
                                    # 如果仍找不到，尝试提取任意媒体文件
                                    if media_files:
                                        image_data = docx_zip.read(media_files[0])
                                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                        with open(output_path, 'wb') as img_file:
                                            img_file.write(image_data)
                                        logger.info(f"提取替代图片: {media_files[0]} 到 {output_path}")
                                        return True
                                    
                                    logger.warning(f"无法找到匹配的图片: {image_path}")
                                    return False
                        except Exception as e:
                            logger.error(f"提取图片时出错: {str(e)}")
                            traceback.print_exc()
                            return False
                    
                    # 将方法添加到DocxParser类
                    DocxParser.extract_image = extract_image
                    logger.info("成功添加extract_image方法到DocxParser类")
                
                # 检查DocxToMarkdownConverter类
                from utils.docx2markdown.docx_to_markdown_converter import DocxToMarkdownConverter
                
                # 检查构造函数是否接受output_path参数
                if 'output_path' not in inspect.signature(DocxToMarkdownConverter.__init__).parameters:
                    logger.info("DocxToMarkdownConverter缺少output_path参数，正在修复...")
                    
                    # 保存原始方法
                    original_init = DocxToMarkdownConverter.__init__
                    
                    # 定义新的初始化方法
                    def new_init(self, docx_file, output_path=None):
                        original_init(self, docx_file)
                        self.output_path = output_path
                        self.image_count = 0
                        self.extracted_images = []
                        
                        # 创建图片目录
                        if output_path:
                            self.doc_name = os.path.basename(output_path).split('.')[0]
                            self.img_dir = os.path.join(os.path.dirname(output_path), self.doc_name + "_outputs")
                            os.makedirs(self.img_dir, exist_ok=True)
                    
                    # 替换初始化方法
                    DocxToMarkdownConverter.__init__ = new_init
                    logger.info("成功修复DocxToMarkdownConverter类")
                
                logger.info("docx2markdown库检查和修复完成")
            
            except Exception as e:
                logger.error(f"修复docx2markdown库时出错: {str(e)}")
                traceback.print_exc()
        else:
            logger.warning("utils/docx2markdown目录存在，但缺少docx_to_markdown_converter.py文件")
    else:
        logger.warning("未找到utils/docx2markdown目录")
    
except Exception as e:
    logger.error(f"检查本地docx2markdown库时出错: {str(e)}")
    logger.info("应用将继续运行，使用可用的转换方法")

# 开始初始化应用
app = Flask(__name__)
app.secret_key = 'youdonote2markdown'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['ALLOWED_EXTENSIONS'] = {'docx', 'pdf'}
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 限制上传文件大小为50MB

# 修复docx转换模块
try:
    import fix_docx_converter
    import fix_cyrus_converter
    
    logger.info("正在应用docx转换模块修复...")
    # 应用基础修复
    docx_fix_result = fix_docx_converter.fix_all()
    logger.info(f"docx转换模块修复{'成功' if docx_fix_result else '失败'}")
    
    # 应用Cyrus转换器修复
    cyrus_fix_result = fix_cyrus_converter.fix_cyrus_converter()
    logger.info(f"Cyrus转换器修复{'成功' if cyrus_fix_result else '失败'}")
    
    if docx_fix_result and cyrus_fix_result:
        logger.info("所有修复已成功应用")
    else:
        logger.warning("部分修复未能成功应用，可能影响转换功能")
except Exception as e:
    logger.error(f"应用转换模块修复时出错: {str(e)}")
    logger.error("将以默认配置继续运行应用")

# 确保上传和输出目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(app.config['OUTPUT_FOLDER'], 'images'), exist_ok=True)

# 延迟导入转换工具，以避免循环导入
def get_conversion_modules():
    try:
        from utils.docx_to_md import convert_docx_to_md
        from utils.pdf_to_md import convert_pdf_to_md
        from utils.gitee_uploader import upload_images_to_gitee
        # 新增: 导入转换选择器
        from utils.docx_converter_selector import convert_docx_to_markdown, ConversionMethod, get_available_methods
        return convert_docx_to_md, convert_pdf_to_md, upload_images_to_gitee, convert_docx_to_markdown, ConversionMethod, get_available_methods
    except ImportError as e:
        logger.error(f"导入模块失败: {str(e)}")
        raise

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    logger.info("访问首页")
    uploaded_files = os.listdir(app.config['UPLOAD_FOLDER'])
    
    # 修改列出已转换文件的逻辑
    converted_files = []
    output_dir = app.config['OUTPUT_FOLDER']
    
    # 查找所有以_outputs结尾的文件夹
    for item in os.listdir(output_dir):
        item_path = os.path.join(output_dir, item)
        if os.path.isdir(item_path) and item.endswith('_outputs'):
            # 查找每个文件夹中的md文件
            for file in os.listdir(item_path):
                if file.endswith('.md'):
                    # 保存路径格式为 folder_name/file.md
                    converted_files.append(f"{item}/{file}")
    
    # 获取可用的转换方法
    try:
        _, _, _, _, _, get_available_methods = get_conversion_modules()
        conversion_methods = get_available_methods()
        method_names = [method.value for method in conversion_methods]
    except:
        method_names = ["default"]
    
    return render_template('index.html', 
                           uploaded_files=uploaded_files, 
                           converted_files=converted_files,
                           conversion_methods=method_names)

@app.route('/upload', methods=['POST'])
def upload_file():
    logger.info("上传文件请求")
    if 'file' not in request.files:
        logger.warning("没有选择文件")
        flash('没有选择文件', 'error')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        logger.warning("文件名为空")
        flash('没有选择文件', 'error')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        try:
            # 使用时间戳+原文件名作为文件名
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S_')
            filename = timestamp + secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            logger.info(f"保存文件: {filename}")
            file.save(file_path)
            flash(f'文件 {filename} 上传成功!', 'success')
        except Exception as e:
            logger.error(f"文件上传失败: {str(e)}", exc_info=True)
            flash(f'文件上传失败: {str(e)}', 'error')
        
        return redirect(url_for('index'))
    else:
        logger.warning(f"不支持的文件类型: {file.filename if file else 'None'}")
        flash('不支持的文件类型，请上传docx或pdf文件', 'error')
        return redirect(request.url)

@app.route('/convert/<filename>')
def convert_file(filename):
    logger.info(f"转换文件: {filename}")
    try:
        convert_docx_to_md, convert_pdf_to_md, upload_images_to_gitee, convert_docx_to_markdown, ConversionMethod, _ = get_conversion_modules()
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # 验证文件是否存在
        if not os.path.exists(input_path):
            logger.error(f"文件不存在: {input_path}")
            flash(f'文件 {filename} 不存在', 'error')
            return redirect(url_for('index'))
        
        filename_without_ext = os.path.splitext(filename)[0]
        
        # 创建专用的输出目录，同时存放markdown和图片
        output_folder_name = f"{filename_without_ext}_outputs"
        output_dir = os.path.join(app.config['OUTPUT_FOLDER'], output_folder_name)
        os.makedirs(output_dir, exist_ok=True)
        
        # 设置markdown文件路径
        output_filename = f"{filename_without_ext}.md"
        output_path = os.path.join(output_dir, output_filename)
        
        # 获取选择的转换方法
        conversion_method = request.args.get('method', 'default')
        logger.info(f"使用转换方法: {conversion_method}")
        
        # 根据文件类型选择转换方法
        if filename.lower().endswith('.docx'):
            logger.info(f"处理DOCX文件: {filename}")
            
            # 根据选择的方法进行转换
            if conversion_method == 'cyrus':
                # 使用CYRUS方法
                try:
                    logger.info(f"尝试使用CYRUS方法转换文件: {filename}")
                    image_paths = convert_docx_to_markdown(
                        input_path, 
                        output_path, 
                        method=ConversionMethod.CYRUS,
                        image_dir=output_dir
                    )
                    logger.info(f"CYRUS方法转换成功，提取到 {len(image_paths)} 张图片")
                    # 成功时给用户消息
                    flash(f'使用CYRUS方法成功转换文件 {filename}', 'success')
                except Exception as e:
                    logger.error(f"使用CYRUS方法转换失败: {str(e)}", exc_info=True)
                    # 告知用户转换失败，但系统会自动使用默认方法
                    flash(f'CYRUS方法转换失败，已自动降级到默认方法: {str(e)}', 'warning')
                    logger.info("尝试使用默认方法")
                    image_paths = convert_docx_to_md(input_path, output_path, output_dir)
                    logger.info(f"默认方法转换成功，提取到 {len(image_paths)} 张图片")
            else:
                # 使用默认方法
                logger.info(f"使用默认方法转换文件: {filename}")
                image_paths = convert_docx_to_md(input_path, output_path, output_dir)
                logger.info(f"默认方法转换成功，提取到 {len(image_paths) if image_paths else 0} 张图片")
                
        elif filename.lower().endswith('.pdf'):
            logger.info(f"处理PDF文件: {filename}")
            # 使用统一的图片和输出目录
            convert_pdf_to_md(input_path, output_path, output_dir)
        else:
            logger.error(f"不支持的文件类型: {filename}")
            flash(f'不支持的文件类型: {filename}', 'error')
            return redirect(url_for('index'))
        
        # 验证转换后的文件是否存在
        if not os.path.exists(output_path):
            logger.error(f"转换失败，输出文件不存在: {output_path}")
            flash(f'转换失败，无法生成输出文件', 'error')
            return redirect(url_for('index'))
        
        # 上传图片到Gitee并更新Markdown文件中的图片链接
        try:
            logger.info("上传图片到Gitee")
            # 使用统一的目录
            upload_images_to_gitee(output_path, output_dir)
        except Exception as e:
            logger.warning(f"上传图片到Gitee失败: {str(e)}")
            flash(f'文件已转换，但上传图片到Gitee失败: {str(e)}', 'warning')
        
        logger.info(f"文件 {filename} 转换成功")
        flash(f'文件 {filename} 转换成功!', 'success')
    except Exception as e:
        logger.error(f"转换失败: {str(e)}", exc_info=True)
        flash(f'转换失败: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/download/<path:filepath>')
def download_file(filepath):
    logger.info(f"下载文件: {filepath}")
    try:
        # 拆分文件夹和文件名
        folder, filename = os.path.split(filepath)
        folder_path = os.path.join(app.config['OUTPUT_FOLDER'], folder)
        file_path = os.path.join(folder_path, filename)
        
        # 验证文件是否存在
        if not os.path.exists(file_path):
            logger.error(f"下载文件不存在: {file_path}")
            flash(f'文件 {filename} 不存在', 'error')
            return redirect(url_for('index'))
        
        return send_from_directory(folder_path, filename, as_attachment=True)
    except Exception as e:
        logger.error(f"下载文件失败: {str(e)}", exc_info=True)
        flash(f'下载文件失败: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/delete/upload/<filename>')
def delete_upload(filename):
    logger.info(f"删除上传文件: {filename}")
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            flash(f'上传文件 {filename} 已删除', 'success')
        else:
            logger.warning(f"删除的文件不存在: {file_path}")
            flash(f'文件 {filename} 不存在', 'warning')
    except Exception as e:
        logger.error(f"删除文件失败: {str(e)}", exc_info=True)
        flash(f'删除文件失败: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/delete/converted/<path:filepath>')
def delete_converted(filepath):
    logger.info(f"删除转换文件: {filepath}")
    try:
        # 拆分文件夹和文件名
        folder, filename = os.path.split(filepath)
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], folder, filename)
        
        if os.path.exists(file_path):
            os.remove(file_path)
            flash(f'转换文件 {filename} 已删除', 'success')
            
            # 检查文件夹是否为空，如果为空则删除文件夹
            folder_path = os.path.join(app.config['OUTPUT_FOLDER'], folder)
            if os.path.exists(folder_path) and len(os.listdir(folder_path)) == 0:
                os.rmdir(folder_path)
                logger.info(f"删除空文件夹: {folder}")
        else:
            logger.warning(f"删除的文件不存在: {file_path}")
            flash(f'文件 {filename} 不存在', 'warning')
    except Exception as e:
        logger.error(f"删除文件失败: {str(e)}", exc_info=True)
        flash(f'删除文件失败: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.errorhandler(404)
def page_not_found(e):
    logger.warning(f"404错误: {request.path}")
    flash('请求的页面不存在', 'error')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_server_error(e):
    logger.error(f"500错误: {str(e)}", exc_info=True)
    flash('服务器内部错误，请查看日志', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    logger.info("启动应用...")
    print("有道云笔记转Markdown工具已启动，请访问: http://127.0.0.1:5000")
    app.run(debug=True) 
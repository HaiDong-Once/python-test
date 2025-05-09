#!/usr/bin/env python
"""
测试并比较不同的DOCX到Markdown转换方法
"""

import os
import sys
import logging
import time
import difflib
from pathlib import Path

# 添加父目录到sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# 导入转换模块
from utils.docx_converter_selector import convert_docx_to_markdown, ConversionMethod, get_available_methods

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def count_images_in_dir(directory):
    """统计目录中的图片数量"""
    if not os.path.exists(directory):
        return 0
        
    count = 0
    for file in os.listdir(directory):
        if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
            count += 1
    return count

def get_code_blocks_count(md_file):
    """统计Markdown文件中的代码块数量"""
    if not os.path.exists(md_file):
        return 0
        
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 简单计算```的数量并除以2（开始和结束标记）
    code_block_markers = content.count('```')
    return code_block_markers // 2

def get_images_count(md_file):
    """统计Markdown文件中的图片引用数量"""
    if not os.path.exists(md_file):
        return 0
        
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 计算![的数量
    image_markers = content.count('![')
    return image_markers

def compare_files(file1, file2):
    """比较两个文件的差异"""
    if not os.path.exists(file1) or not os.path.exists(file2):
        return "无法比较文件，至少有一个文件不存在"
        
    with open(file1, 'r', encoding='utf-8') as f1:
        content1 = f1.readlines()
    
    with open(file2, 'r', encoding='utf-8') as f2:
        content2 = f2.readlines()
    
    # 计算差异
    diff = difflib.unified_diff(
        content1, content2, 
        fromfile=os.path.basename(file1),
        tofile=os.path.basename(file2),
        lineterm=''
    )
    
    # 返回差异
    return '\n'.join(list(diff))

def test_conversion_methods(docx_file):
    """测试不同的转换方法并对比结果"""
    if not os.path.exists(docx_file):
        logger.error(f"测试文件不存在: {docx_file}")
        return
    
    # 创建测试输出目录
    test_dir = Path("tests/test_outputs/converter_comparison")
    test_dir.mkdir(parents=True, exist_ok=True)
    
    # 获取文件名（不带扩展名）
    filename = os.path.basename(docx_file)
    filename_without_ext = os.path.splitext(filename)[0]
    
    results = {}
    
    # 测试可用的转换方法
    available_methods = get_available_methods()
    logger.info(f"可用的转换方法: {[method.value for method in available_methods]}")
    
    for method in available_methods:
        logger.info(f"测试转换方法: {method.value}")
        
        # 设置输出路径
        output_dir = test_dir / method.value
        output_dir.mkdir(exist_ok=True)
        output_path = output_dir / f"{filename_without_ext}.md"
        
        # 记录开始时间
        start_time = time.time()
        
        # 执行转换
        try:
            image_paths = convert_docx_to_markdown(
                docx_file,
                str(output_path),
                method=method,
                image_dir=str(output_dir)
            )
            
            # 记录结束时间
            end_time = time.time()
            conversion_time = end_time - start_time
            
            # 收集结果
            results[method.value] = {
                'output_path': str(output_path),
                'success': True,
                'conversion_time': conversion_time,
                'image_count': len(image_paths),
                'image_dir_count': count_images_in_dir(output_dir),
                'code_blocks': get_code_blocks_count(output_path),
                'md_image_refs': get_images_count(output_path)
            }
            
            logger.info(f"转换完成: {method.value}, 耗时: {conversion_time:.2f}秒")
            
        except Exception as e:
            logger.error(f"转换方法 {method.value} 失败: {str(e)}", exc_info=True)
            results[method.value] = {
                'output_path': str(output_path),
                'success': False,
                'error': str(e)
            }
    
    # 比较结果
    compare_results(results)
    
    return results

def compare_results(results):
    """比较不同转换方法的结果"""
    logger.info("\n" + "="*50)
    logger.info("转换方法比较结果:")
    logger.info("="*50)
    
    # 打印结果表格
    print("\n{:<10} | {:<10} | {:<15} | {:<15} | {:<15} | {:<15}".format(
        "方法", "成功", "转换时间(秒)", "图片数", "代码块数", "图片引用数"
    ))
    print("-" * 90)
    
    for method, result in results.items():
        if result['success']:
            print("{:<10} | {:<10} | {:<15.2f} | {:<15} | {:<15} | {:<15}".format(
                method,
                "是" if result['success'] else "否",
                result['conversion_time'],
                result['image_count'],
                result['code_blocks'],
                result['md_image_refs']
            ))
        else:
            print("{:<10} | {:<10} | {:<15} | {:<15} | {:<15} | {:<15}".format(
                method,
                "是" if result['success'] else "否",
                "N/A",
                "N/A",
                "N/A",
                "N/A"
            ))
    
    # 如果有多个成功的方法，比较它们的差异
    successful_methods = [m for m, r in results.items() if r['success']]
    if len(successful_methods) > 1:
        logger.info("\n文件差异比较:")
        for i in range(len(successful_methods)-1):
            for j in range(i+1, len(successful_methods)):
                method1 = successful_methods[i]
                method2 = successful_methods[j]
                
                logger.info(f"\n比较 {method1} 与 {method2}:")
                diff = compare_files(
                    results[method1]['output_path'],
                    results[method2]['output_path']
                )
                
                # 计算差异的行数（粗略估计）
                diff_lines = diff.count('\n')
                logger.info(f"差异行数: 约 {diff_lines} 行")
                
                # 只显示差异摘要
                if diff_lines > 10:
                    diff_preview = '\n'.join(diff.split('\n')[:10]) + "\n... (更多差异已省略)"
                    logger.info(f"差异摘要:\n{diff_preview}")
                else:
                    logger.info(f"完整差异:\n{diff}")

if __name__ == "__main__":
    # 检查命令行参数
    if len(sys.argv) > 1:
        docx_file = sys.argv[1]
    else:
        # 使用默认测试文件
        docx_file = "tests/test_docs/complex_test_document.docx"
    
    # 确保测试文件存在
    if not os.path.exists(docx_file):
        logger.error(f"测试文件不存在: {docx_file}")
        sys.exit(1)
    
    # 运行测试
    test_conversion_methods(docx_file) 
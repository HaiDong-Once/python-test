#!/usr/bin/env python

import re
import os

def check_code_blocks(filepath):
    """检查markdown文件中代码块的格式"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查所有代码块格式
    code_block_pattern = r'```([a-zA-Z]*)\n(.*?)\n```'
    code_blocks = re.findall(code_block_pattern, content, re.DOTALL)
    
    print(f"在 {filepath} 中找到 {len(code_blocks)} 个代码块")
    print("-" * 50)
    
    for i, (lang, code) in enumerate(code_blocks):
        print(f"代码块 {i+1}:")
        print(f"语言: '{lang}'")
        print(f"代码长度: {len(code)} 字符")
        print(f"代码前10个字符: {code[:10]}")
        print("-" * 30)
    
    return len(code_blocks)

def check_images_position(filepath):
    """检查Markdown文件中图片的位置"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.readlines()
    
    image_positions = []
    
    for i, line in enumerate(content):
        if re.search(r'!\[图片\d+\]', line):
            image_match = re.search(r'!\[图片(\d+)\]', line)
            if image_match:
                image_num = image_match.group(1)
                # 获取上下文
                context_before = content[max(0, i-3):i]
                context_after = content[i+1:min(len(content), i+3)]
                
                image_positions.append({
                    'line_number': i+1,
                    'image_number': image_num,
                    'context_before': ''.join(context_before),
                    'image_line': line,
                    'context_after': ''.join(context_after)
                })
    
    print(f"在 {filepath} 中找到 {len(image_positions)} 张图片引用")
    print("-" * 50)
    
    for pos in image_positions:
        print(f"图片 {pos['image_number']} 在第 {pos['line_number']} 行:")
        print("上下文:")
        print(pos['context_before'], end='')
        print(f">>> {pos['image_line']}", end='')
        print(pos['context_after'])
        print("-" * 50)
    
    return len(image_positions)

def main():
    original_file = "outputs/20250509152655_RAGFlow_outputs/20250509152655_RAGFlow.md"
    improved_file = "outputs/20250509152655_RAGFlow_outputs/20250509152655_RAGFlow_improved.md"
    
    if not os.path.exists(improved_file):
        print(f"错误: 无法找到文件 {improved_file}")
        return
    
    # 检查代码块格式
    print("\n===== 代码块格式检查 =====")
    code_block_count = check_code_blocks(improved_file)
    print(f"文件中共有 {code_block_count} 个代码块")
    
    # 检查图片位置
    print("\n===== 图片位置检查 =====")
    image_count = check_images_position(improved_file)
    print(f"文件中共有 {image_count} 张图片")

if __name__ == "__main__":
    main() 
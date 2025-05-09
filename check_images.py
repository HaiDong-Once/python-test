import re

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

if __name__ == "__main__":
    filepath = "outputs/20250509152655_RAGFlow_outputs/20250509152655_RAGFlow_fixed.md"
    check_images_position(filepath) 
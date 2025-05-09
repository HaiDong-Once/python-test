import os
from utils.docx_to_md import convert_docx_to_md

# 测试文件
test_docx = "test_document.docx"
output_md = "outputs/test_result.md"

# 确保输出目录存在
os.makedirs("outputs", exist_ok=True)

# 执行转换
try:
    convert_docx_to_md(test_docx, output_md)
    with open("outputs/test_log.txt", "w", encoding="utf-8") as log:
        log.write(f"转换成功: {test_docx} -> {output_md}\n")
        
        # 读取结果文件并写入日志
        with open(output_md, "r", encoding="utf-8") as f:
            content = f.read()
            log.write("\n转换后的内容:\n")
            log.write("-" * 40 + "\n")
            log.write(content)
except Exception as e:
    with open("outputs/test_log.txt", "w", encoding="utf-8") as log:
        log.write(f"转换失败: {str(e)}") 
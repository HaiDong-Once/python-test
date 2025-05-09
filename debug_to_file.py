import os
import sys
import logging

# 创建日志文件
log_file = 'debug_log.txt'
with open(log_file, 'w') as f:
    f.write("开始日志记录\n")
    f.write(f"Python 版本: {sys.version}\n")
    f.write(f"当前目录: {os.getcwd()}\n")
    
    try:
        f.write(f"目录内容: {os.listdir('.')}\n")
    except Exception as e:
        f.write(f"无法列出目录内容: {str(e)}\n")
    
    # 检查依赖项
    f.write("\n检查依赖项:\n")
    try:
        import flask
        f.write(f"Flask 已安装\n")
    except ImportError as e:
        f.write(f"Flask导入失败: {e}\n")
    
    try:
        import docx
        f.write(f"python-docx 已安装\n")
    except ImportError as e:
        f.write(f"python-docx导入失败: {e}\n")
    
    # 检查项目结构
    f.write("\n检查项目结构:\n")
    try:
        for root, dirs, files in os.walk('.'):
            if '__pycache__' in root:
                continue
            f.write(f"目录: {root}\n")
            for file in files:
                f.write(f"  文件: {file}\n")
    except Exception as e:
        f.write(f"无法遍历目录: {str(e)}\n")
    
    # 尝试导入应用模块
    f.write("\n尝试导入应用和工具模块:\n")
    try:
        import app
        f.write("成功导入app.py\n")
    except Exception as e:
        f.write(f"导入app.py失败: {e}\n")
    
    try:
        from utils import docx_to_md
        f.write("成功导入docx_to_md\n")
    except Exception as e:
        f.write(f"导入docx_to_md失败: {e}\n")
    
    f.write("\n调试完成\n")

# 创建一个简单的HTML测试文件
with open('test.html', 'w') as f:
    f.write("""
<!DOCTYPE html>
<html>
<head>
    <title>测试页面</title>
</head>
<body>
    <h1>测试页面</h1>
    <p>这是一个测试页面，用于确认服务器是否正常工作。</p>
</body>
</html>
""") 
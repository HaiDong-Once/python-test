#!/usr/bin/env python
"""
专门用于安装CYRUS-STUDIO/docx2markdown库的脚本
"""

import subprocess
import sys
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """尝试安装CYRUS-STUDIO/docx2markdown库"""
    try:
        logger.info("正在安装CYRUS-STUDIO/docx2markdown库...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", 
            "git+https://github.com/CYRUS-STUDIO/docx2markdown.git"
        ])
        logger.info("CYRUS-STUDIO/docx2markdown库安装成功")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"安装CYRUS-STUDIO/docx2markdown库失败: {str(e)}")
        return False

if __name__ == "__main__":
    result = main()
    
    if result:
        print("✅ CYRUS-STUDIO/docx2markdown库安装成功！")
    else:
        print("❌ CYRUS-STUDIO/docx2markdown库安装失败，请查看日志获取详细信息。")
    
    # 等待用户按键退出
    input("按回车键退出...") 
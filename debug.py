import os
import sys
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

logger.info("检查Python版本和路径")
logger.info(f"Python 版本: {sys.version}")
logger.info(f"当前目录: {os.getcwd()}")
logger.info(f"目录内容: {os.listdir('.')}")

logger.info("\n检查依赖项:")
try:
    import flask
    logger.info(f"Flask 版本: {flask.__version__}")
except ImportError as e:
    logger.error(f"Flask导入失败: {e}")

try:
    import docx
    logger.info(f"python-docx 已安装")
except ImportError as e:
    logger.error(f"python-docx导入失败: {e}")

try:
    import PyPDF2
    logger.info(f"PyPDF2 版本: {PyPDF2.__version__}")
except ImportError as e:
    logger.error(f"PyPDF2导入失败: {e}")

try:
    import fitz
    logger.info(f"PyMuPDF (fitz) 已安装")
except ImportError as e:
    logger.error(f"PyMuPDF导入失败: {e}")

logger.info("\n检查项目结构:")
for root, dirs, files in os.walk('.'):
    if '__pycache__' in root:
        continue
    logger.info(f"目录: {root}")
    for file in files:
        logger.info(f"  文件: {file}")

logger.info("\n尝试导入应用和工具模块:")
try:
    import app
    logger.info("成功导入app.py")
except Exception as e:
    logger.error(f"导入app.py失败: {e}")

try:
    from utils import docx_to_md
    logger.info("成功导入docx_to_md")
except Exception as e:
    logger.error(f"导入docx_to_md失败: {e}")

try:
    from utils import pdf_to_md
    logger.info("成功导入pdf_to_md")
except Exception as e:
    logger.error(f"导入pdf_to_md失败: {e}")

try:
    from utils import gitee_uploader
    logger.info("成功导入gitee_uploader")
except Exception as e:
    logger.error(f"导入gitee_uploader失败: {e}")

print("\n调试完成，请检查上述日志信息以诊断问题") 
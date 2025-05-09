import sys

def test_import(module_name, package_name=None):
    if package_name is None:
        package_name = module_name
    try:
        module = __import__(module_name)
        print(f"✓ {package_name} 已成功导入")
        return True
    except ImportError as e:
        print(f"✗ {package_name} 导入失败: {e}")
        return False

print("测试项目依赖...")

# 核心依赖
deps = [
    ("flask", "Flask"),
    ("werkzeug", "Werkzeug"),
    ("docx", "python-docx"),
    ("PyPDF2", "PyPDF2"),
    ("fitz", "PyMuPDF"),
    ("PIL", "Pillow"),
    ("bs4", "beautifulsoup4"),
    ("markdown", "markdown"),
    ("flask_uploads", "Flask-Reuploaded"),
    ("dotenv", "python-dotenv"),
    ("git", "gitpython"),
    ("requests", "requests"),
    ("tqdm", "tqdm")
]

success = 0
failed = 0

for module_name, package_name in deps:
    if test_import(module_name, package_name):
        success += 1
    else:
        failed += 1

print(f"\n总结: {success}个依赖成功, {failed}个依赖失败")

if failed > 0:
    print("\n缺少的依赖，请运行以下命令安装:")
    print("pip install -r requirements.txt") 
print("开始测试PyMuPDF...")

try:
    import fitz
    print("PyMuPDF已成功导入!")
    print(f"PyMuPDF版本: {fitz.version}")
except ImportError as e:
    print(f"PyMuPDF导入失败: {e}")
    print("尝试安装PyMuPDF...")
    print("请运行: pip install PyMuPDF")

print("测试完成") 
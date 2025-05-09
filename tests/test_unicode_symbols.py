import unittest
from utils.docx_to_md import format_paragraph
from docx import Document
from docx.text.paragraph import Paragraph

class TestUnicodeSymbols(unittest.TestCase):
    """测试特殊Unicode符号处理"""
    
    def test_unicode_symbol_handling(self):
        """测试特殊Unicode符号处理功能"""
        # 创建一个简单的docx文档用于测试
        doc = Document()
        
        # 添加包含特殊符号的段落
        symbols = [
            '❓', '❗', '✅', '✓', '✔️', '✗', '✘', '★', '☆', '➔', '➤', '➡️', '⬅️', '⬆️', '⬇️',
            '📌', '📝', '📊', '📈', '📉', '📋', '⚠️', '⚡', '🔍', '🔎', '🔑', '🔒', '🔓', '💡'
        ]
        
        # 创建测试文本
        test_text = "这是一个包含特殊符号的文本: " + " ".join(symbols)
        p = doc.add_paragraph(test_text)
        
        # 使用format_paragraph函数处理
        formatted_text = format_paragraph(p)
        
        # 验证所有符号都被正确保留
        for symbol in symbols:
            self.assertIn(symbol, formatted_text, f"符号 {symbol} 应该被保留在格式化的文本中")
        
        print("格式化后的文本:", formatted_text)

if __name__ == '__main__':
    unittest.main() 
import os

from docx2markdown.docx_parser import DocxParser, Paragraph, Table

# 添加Gitee基础URL常量
GITEE_BASE_URL = "https://gitee.com/comma-dong/image-projects/raw/master/"

class DocxToMarkdownConverter:
    def __init__(self, docx_file, output_path=None):
        self.docx_file = docx_file
        self.output_path = output_path
        self.in_code_block = False  # 用于追踪是否在代码块中
        self.code_block_content = ""  # 存储代码块的内容
        self.image_count = 0  # 用于计数和生成图片文件名
        self.extracted_images = []  # 保存已提取的图片信息
        
        # 获取文档名，用于创建图片目录
        if output_path:
            self.doc_name = os.path.basename(output_path).split('.')[0]
            # 创建图片目录 - 与MD文件同级
            self.img_dir = os.path.join(os.path.dirname(output_path), self.doc_name + "_outputs")
            os.makedirs(self.img_dir, exist_ok=True)

    def _parse_text_with_hyperlink(self, paragraph):
        """
        如果文本中有超链接，转换为 Markdown 超链接格式
        """
        if paragraph.hyperlink:
            # 转换为 Markdown 超链接格式
            return paragraph.text.replace(paragraph.hyperlink.text, f"[{paragraph.hyperlink.text}]({paragraph.hyperlink.url})")
        return paragraph.text

    def _escaping_text(self, text):
        # 非代码块
        if not self.in_code_block:
            # < 转义
            text = text.replace('<', '\\<', )
        return text

    def _generate_markdown_from_paragraph(self, parser, paragraph):
        """
        根据段落信息生成相应的 Markdown 格式。
        """
        text = self._parse_text_with_hyperlink(paragraph)  # 获取段落文本并处理文本中的超链接
        style = paragraph.style
        image = paragraph.image
        numbering = paragraph.numbering  # 假设已经在解析时获取了编号信息
        background = style.background  # 获取背景填充信息

        markdown_text = ""

        # 处理加粗、斜体、下划线
        if style.bold:
            text = f"**{text}**"
        if style.italic:
            text = f"*{text}*"
        if style.underline:
            text = f"_{text}_"

        # 检查是否是代码块（背景填充不为空）
        if background and background.get('fill') == 'DBDBDB':  # 可以根据需要修改背景色
            if not self.in_code_block:
                # 如果之前没有在代码块中，开始新的代码块
                self.in_code_block = True
                markdown_text += "```\n"  # 开始代码块

            # 将当前段落的文本添加到代码块内容中
            markdown_text += text
        else:
            if self.in_code_block:
                # 如果之前处于代码块中，结束代码块
                self.in_code_block = False
                markdown_text += "```\n"  # 结束代码块

            # 判断是否是列表项
            if numbering is not None:

                # {'bullet': None, 'ilvl': '0', 'lvl_text': '%1.', 'numId': '1', 'num_format': 'decimal'}

                num_format = numbering.get('num_format')
                ilvl = int(numbering.get('ilvl'))
                lvl_text = numbering.get('lvl_text')

                if num_format == 'bullet':
                    # 无序列表
                    markdown_text += f"- {text}\n"
                elif num_format in ['decimal', 'lowerRoman', 'lowerLetter']:
                    # 有序列表，使用 numId 和 ilvl 生成编号
                    if num_format == 'decimal':
                        # 有序列表的格式：1. 2. 3.
                        markdown_text += f"{lvl_text.replace('%1', str(ilvl + 1))} {text}\n"
                    elif num_format == 'lowerRoman':
                        roman_numerals = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x']
                        roman = roman_numerals[ilvl] if ilvl < len(roman_numerals) else str(ilvl + 1)
                        markdown_text += f"{roman}) {text}\n"
                    elif num_format == 'lowerLetter':
                        letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']
                        letter = letters[ilvl] if ilvl < len(letters) else chr(ord('a') + ilvl)
                        markdown_text += f"{letter}) {text}\n"

            else:
                # 根据标题级别生成 Markdown
                if style.fonts.get("default", None):
                    try:
                        heading_level = int(style.fonts["default"])
                        heading_text = text.replace('**', '').strip()
                        if 1 <= heading_level <= 6:  # 1-6 级标题有效
                            markdown_text += f"{'#' * heading_level} {heading_text}\n"
                        else:
                            markdown_text += f"{heading_text}\n"
                    except ValueError:
                        markdown_text += f"{text}\n"  # 如果无法解析为数字，默认处理为普通文本
                else:
                    # 普通文本
                    text = self._escaping_text(text) # 转义
                    markdown_text += f"{text}\n"

        # 处理图片
        if image and self.output_path:
            image_filename = image['file']
            self.image_count += 1
            
            # 获取图片扩展名
            extension = os.path.splitext(image_filename)[1].lower()
            if not extension:
                extension = ".png"  # 默认png格式
            
            # 新的图片文件名
            new_image_name = f"image_{self.image_count}{extension}"
            output_path = os.path.join(self.img_dir, new_image_name)
            
            # 从docx中提取图片并保存到文件
            if parser.extract_image(image_filename, output_path):
                # 记录已提取的图片
                self.extracted_images.append(new_image_name)
                
                # 构建Gitee远程URL
                folder_name = self.doc_name + "_outputs"
                remote_url = f"{GITEE_BASE_URL}{folder_name}/{new_image_name}"
                
                # 使用Gitee链接格式添加图片
                markdown_text += f"\n![{new_image_name}]({remote_url})\n"
            else:
                print(f"无法提取图片: {image_filename}")

        return markdown_text

    def _generate_markdown_from_table(self, table):
        """
        将 Table 对象转换为 Markdown 格式的表格。
        :param table: Table 对象，包含表格的数据
        :return: Markdown 格式的表格字符串
        """
        markdown_table = []

        # 如果表格没有行，直接返回空字符串
        if not table.rows:
            return ""

        # 表头行
        header_row = table.rows[0]
        # 转义
        header_row = [self._escaping_text(s) for s in header_row]
        markdown_table.append("| " + " | ".join(header_row) + " |")

        # 分隔符行（Markdown 表头和表体的分隔线）
        markdown_table.append("|" + " | ".join(["---"] * len(header_row)) + "|")

        # 表格内容行
        for row in table.rows[1:]:
            # 转义
            row = [self._escaping_text(s) for s in row]
            markdown_table.append("| " + " | ".join(row) + " |")

        # 返回转换后的 Markdown 表格内容
        return "\n".join(markdown_table)

    def convert(self):
        """
        转换整个 docx 文件的内容为 markdown 格式。
        """
        parser = DocxParser(self.docx_file)
        document = parser.parse()
        markdown_content = ""

        for element in document['elements']:
            if isinstance(element, Paragraph):
                markdown_content += self._generate_markdown_from_paragraph(parser, element) + "\n"
            elif isinstance(element, Table):
                markdown_content += self._generate_markdown_from_table(element) + "\n"

        # 如果文件结尾处仍然有未关闭的代码块，关闭它
        if self.in_code_block:
            markdown_content += "```\n"

        return markdown_content


def docx_to_markdown(docx_file, output=None):
    converter = DocxToMarkdownConverter(docx_file, output)
    markdown_content = converter.convert()

    # 输出生成的 Markdown 内容
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        print(f"Markdown 文件已生成：{output}")
        
        # 输出图片相关信息
        if hasattr(converter, 'img_dir'):
            print(f"图片保存在：{converter.img_dir}")
            print(f"共提取了{len(converter.extracted_images)}张图片")

    return markdown_content

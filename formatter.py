""""排版模块"""
import re
from docx import Document
from docx.shared import Pt,Inches
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
class DocumentFormatter:
    DEFAULT_FONT = '宋体'
    DEFAULT_FONT_SIZE = 12
    DEFAULT_LINE_SPACING = 18
    DEFAULT_INDENT = 0.5
    def __init__(self):
        self.custom_settings={
            'font_name':self.DEFAULT_FONT,
            'font_size':self.DEFAULT_FONT_SIZE,
            'line_spacing':self.DEFAULT_LINE_SPACING,
            'indent':self.DEFAULT_INDENT,
            'first_line_indent': self.DEFAULT_INDENT
        }
    def clean_whitespace(self, content):
        """清理冗余格式：制表符替换、多空格压缩、连续空行合并"""
        # 1. 把制表符替换成空格
        content = content.replace('\t', ' ')

        # 2. 把多个连续空格压缩成单个空格
        content = re.sub(r' {2,}', ' ', content)

        # 3. 按行分割，只去掉每行末尾空格，保留开头缩进
        lines = content.split('\n')
        lines = [line.rstrip() for line in lines]

        # 4. 合并连续空行（最多保留2个）
        cleaned_lines = []
        empty_count = 0
        for line in lines:
            if line == '':
                empty_count += 1
                if empty_count <= 2:
                    cleaned_lines.append(line)
            else:
                empty_count = 0
                cleaned_lines.append(line)

        # 5. 去掉首尾空行
        while cleaned_lines and cleaned_lines[0] == '':
            cleaned_lines.pop(0)
        while cleaned_lines and cleaned_lines[-1] == '':
            cleaned_lines.pop(-1)

        return '\n'.join(cleaned_lines)

    def get_current_settings(self):
        """返回当前自定义格式设置"""
        return self.custom_settings.copy()

    def apply_basic_format(self, content):
        """
        应用基础格式（考核要求2：统一格式设置）
        默认：宋体小四、固定行距18磅
        """
        content = self.clean_whitespace(content)
        formatted_content = f"""/* 
    文档排版信息：
    - 字体：{self.custom_settings['font_name']}
    - 字号：{self.custom_settings['font_size']}pt（小四）
    - 行距：{self.custom_settings['line_spacing']}磅（固定值）
    - 首行缩进：{self.custom_settings['first_line_indent']}英寸
    =============================== */
    {content}
    """
        return formatted_content
    def _strip_format_header(self, content):
        """去掉排版信息头（/* ... */），只保留文档正文"""
        # 只在以 /* 开头时才处理，避免误删正文中的 */
        if content.strip().startswith('/*'):
            # 按行过滤：跳过 /* ... */ 之间的所有行
            lines = content.split('\n')
            in_header = True
            result_lines = []
            for line in lines:
                if in_header:
                    if '*/' in line:
                        in_header = False
                    # 跳过头部所有行
                else:
                    result_lines.append(line)
            content = '\n'.join(result_lines)
        return content.rstrip()

    def save_as_word(self, content, out_path):
        try:
            # 过滤掉排版信息头
            content = self._strip_format_header(content)

            doc = Document()
            style = doc.styles['Normal']
            font_name = self.custom_settings['font_name']
            style.font.name = font_name
            style.font.size = Pt(self.custom_settings['font_size'])
            style.paragraph_format.line_spacing = Pt(self.custom_settings['line_spacing'])
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

            # 同时也设置文档默认样式的东亚字体
            style_xml = style.element
            rPr = style_xml.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                style_xml.append(rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            # 同时设置 ascii、hAnsi、eastAsia 三种字体
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:eastAsia'), font_name)

            paragraphs = content.split('\n')
            for para_text in paragraphs:
                if para_text.strip():
                    paragraph = doc.add_paragraph(para_text)
                    paragraph.style = style
            doc.save(out_path)
            return True
        except Exception as e:
            raise Exception(f"保存失败: {out_path}") from e

    def get_statistics(self, content):
        """
        获取文本统计信息（进阶功能：字数统计）

        Args:
            content: 文本内容

        Returns:
            dict: 包含字符数、单词数、行数、段落数
        """
        if not content:
            return {
                'char_count': 0,
                'word_count': 0,
                'line_count': 0,
                'paragraph_count': 0
            }

        # 1. 字符数（不含空格、换行、回车）
        char_count = len(content.replace(' ', '').replace('\n', '').replace('\r', ''))

        # 2. 单词数（匹配中文汉字和英文单词）
        import re
        # 匹配中文：[\u4e00-\u9fff] 匹配所有常用汉字
        # 匹配英文：[a-zA-Z]+ 匹配连续英文字母
        words = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', content)
        word_count = len(words)

        # 3. 行数（按换行符分割）
        lines = content.split('\n')
        line_count = len(lines)

        # 4. 段落数（按空行分隔，过滤掉空段落）
        paragraphs = [p for p in content.split('\n\n') if p.strip()]
        paragraph_count = len(paragraphs)

        return {
            'char_count': char_count,
            'word_count': word_count,
            'line_count': line_count,
            'paragraph_count': paragraph_count
        }
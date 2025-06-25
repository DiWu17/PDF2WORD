import pypandoc
from docx import Document
from docx.shared import Inches
from loguru import logger
import os
import re

def _html_table_to_markdown(html_table_match):
    """
    A helper function to convert an HTML table string into a Markdown table.
    This is designed to be used with re.sub.
    """
    html_table = html_table_match.group(0)
    
    # Extract rows from the table
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html_table, re.DOTALL)
    if not rows:
        return html_table # Return original if no rows found

    md_rows = []
    header_generated = False
    for row_html in rows:
        # Extract cells (both th and td)
        cells = re.findall(r'<(t[hd])[^>]*>(.*?)</\1>', row_html, re.DOTALL)
        cell_contents = [cell[1].strip().replace('\n', ' ') for cell in cells]
        
        if not cell_contents:
            continue
            
        md_rows.append("| " + " | ".join(cell_contents) + " |")

        # Generate header separator after the first row
        if not header_generated:
            separator = "| " + " | ".join(["---"] * len(cell_contents)) + " |"
            md_rows.append(separator)
            header_generated = True
            
    return "\n".join(md_rows)

def convert_markdown_to_word_with_pandoc(markdown_file, output_docx_path):
    """
    Converts a Markdown file to a Word document using pypandoc,
    handles images, and adjusts their size in the final document.
    """
    try:
        # The resource path should be the directory containing the markdown file,
        # where associated images are expected to be.
        resource_path = os.path.dirname(markdown_file)
        logger.info(f"使用 pandoc 将 {markdown_file} 转换为 {output_docx_path}...")
        logger.info(f"Pandoc 资源路径 (图片等): {resource_path}")

        # 读取 Markdown 文件内容
        with open(markdown_file, 'r', encoding='utf-8') as f:
            markdown_content = f.read()

        # 预处理内容，替换 \textcircled，解决 "Could not convert TeX math" 警告
        processed_content = re.sub(r'\\textcircled\{(\d+)\}', r'(\1)', markdown_content)
        
        # 将HTML表格转换为Markdown表格
        processed_content = re.sub(
            r'<table[^>]*>.*?</table>',
            _html_table_to_markdown,
            processed_content,
            flags=re.DOTALL | re.IGNORECASE
        )

        extra_args = [
            '--resource-path', resource_path,
        ]
        
        # 定义输入格式，告诉 pandoc 输入的是 markdown，并启用解析 html 块的扩展
        input_format = 'markdown+markdown_in_html_blocks'

        pypandoc.convert_text(
            processed_content,
            'docx',
            format=input_format,
            outputfile=output_docx_path,
            extra_args=extra_args
        )

        logger.info(f"Markdown 转换成功，开始调整图片尺寸: {output_docx_path}")

        # 打开刚刚创建的文档以调整图片大小
        doc = Document(output_docx_path)
        
        # 获取页面内容区域的大致宽度（以英寸为单位）
        # 假设A4纸，左右边距各1英寸
        # 8.5 (页面宽) - 2 * 1 (边距) = 6.5英寸
        # 我们可以设置一个安全的内容宽度
        max_width_inches = 6.0 

        for shape in doc.inline_shapes:
            # 检查对象是否为图片
            # 'picture' in shape.type' 检查更为通用
            if shape.width and shape.height:
                # 将图片宽度从 EMU (English Metric Units) 转换为英寸
                current_width_inches = shape.width.inches
                
                if current_width_inches > max_width_inches:
                    # 计算缩放比例
                    ratio = max_width_inches / current_width_inches
                    # 调整尺寸
                    shape.width = Inches(max_width_inches)
                    shape.height = int(shape.height.emu * ratio)
                    logger.info(f"图片尺寸已从 {current_width_inches:.2f} 英寸宽调整为 {max_width_inches:.2f} 英寸宽。")

        # 4. 保存修改后的文档
        doc.save(output_docx_path)
        logger.info(f"图片尺寸调整完成，文档已保存到 {output_docx_path}")

    except Exception as e:
        logger.error(f"使用 pandoc 转换或处理文档时出错: {e}")
        raise

# 保留旧函数，但可以标记为已弃用或之后移除
# def convert_markdown_to_word_with_spire(input_file, output_file):
#     ... 
import fitz  # PyMuPDF
import json
import argparse
import os
import base64
import re
import math
from loguru import logger

def analysis_pdf(pdf_path, json_path):
    """
    Parses a PDF file and extracts text blocks, saving the output to a JSON file.

    :param pdf_path: Path to the input PDF file.
    :param json_path: Path to the output JSON file.
    """
    if not os.path.exists(pdf_path):
        logger.error(f"Error: PDF file not found at {pdf_path}")
        return

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        logger.error(f"Error opening PDF file: {e}")
        return

    pdf_data = {
        "file_path": pdf_path,
        "num_pages": doc.page_count,
        "pages": []
    }

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        page_dict = page.get_text("dict")
        
        for block in page_dict.get("blocks", []):
            if block.get("type") == 1 and "image" in block and isinstance(block["image"], bytes):
                block["image"] = base64.b64encode(block["image"]).decode('utf-8')

        page_data = {
            "page_number": page_num + 1,
            "width": page.rect.width,
            "height": page.rect.height,
            "blocks": page_dict["blocks"]
        }
        pdf_data["pages"].append(page_data)

    doc.close()

    with open(json_path, 'w', encoding='utf-8') as jf:
        json.dump(pdf_data, jf, ensure_ascii=False, indent=4)
    logger.info(f"Successfully converted {pdf_path} to {json_path}")


def estimate_line_count(text: str, bbox: list) -> int:
    """
    Estimates how many lines a given text block occupies based on its content and bounding box.

    :param text: The text content of the block.
    :param bbox: The bounding box [x0, y0, x1, y1] for the entire text block.
    :return: An estimated integer number of lines.
    """
    if not text or not bbox or len(bbox) != 4:
        return 1 if text else 0

    bbox_width = bbox[2] - bbox[0]
    bbox_height = bbox[3] - bbox[1]
    text_length = len(text)
    
    if bbox_width <= 0 or bbox_height <= 0 or text_length == 0:
        return 1

    # Step 1: Estimate font size from bbox area and character count
    avg_char_area = (bbox_width * bbox_height) / text_length
    
    is_chinese = re.search("[\u4e00-\u9fff]", text)
    
    # For CJK chars, area is roughly S*S. For Latin, roughly S*(0.6*S)
    char_width_to_height_ratio = 1.0 if is_chinese else 0.6
    
    # S^2 ≈ avg_char_area / ratio
    estimated_font_size_sq = avg_char_area / char_width_to_height_ratio
    if estimated_font_size_sq <= 0:
        return 1
    estimated_font_size = math.sqrt(estimated_font_size_sq)

    # Step 2: Estimate single line height from font size
    # Line height is typically font size + leading. A common value is 1.2 * font size.
    line_height_factor = 1.1 if is_chinese else 1.2
    estimated_line_height = estimated_font_size * line_height_factor
    
    if estimated_line_height <= 0:
        return 1

    # Step 3: Calculate number of lines
    num_lines = bbox_height / estimated_line_height
    
    # Round to the nearest integer. If result is 0, it must be at least 1 line.
    estimated_lines = round(num_lines)
    return estimated_lines if estimated_lines > 0 else 1


def estimate_font_size_advanced(text: str, bbox: list) -> float:
    """
    Estimates the font size for text that can be English or Chinese.

    :param text: The text content.
    :param bbox: A list representing the bounding box [x0, y0, x1, y1].
    :return: The estimated font size in points.
    """
    if not text or not bbox or len(bbox) != 4:
        return 0.0

    bbox_height = bbox[3] - bbox[1]
    
    # 使用正则表达式判断文本是否包含中文字符
    # CJK Unified Ideographs 的 Unicode 范围是 \u4e00-\u9fff
    if re.search("[\u4e00-\u9fff]", text):
        # 对于中文，bbox 高度非常接近字号
        # 我们可以使用一个微小的修正系数（如1.0 to 1.05）来补偿字体内部的微小边距
        # 但在大多数情况下，直接使用 bbox_height 已经足够准确
        correction_factor = 1.0 
        estimated_size = bbox_height * correction_factor
    else:
        # 对于英文，检查是否包含降部字符
        if any(c in 'gjpqy' for c in text):
            # 包含降部，使用经验系数
            correction_factor = 0.75 
            estimated_size = bbox_height * correction_factor
        else:
            # 不含降部（如大写字母）
            correction_factor = 0.9
            estimated_size = bbox_height * correction_factor
            
    return estimated_size

# if __name__ == '__main__':
#     pdf_path = "D:/python/PDF2WORD/input/sample_2.pdf"
#     json_path = "D:/python/PDF2WORD/output/sample_2_analysis.json"
#
#
#     parse_pdf(pdf_path, json_path)

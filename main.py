import os
import argparse
from pathlib import Path
from loguru import logger

from utils.pdf_utils import parse_pdf_to_files
from utils.docx_utils import convert_markdown_to_word_with_pandoc
from utils.pdf_converter import process_and_convert_pdf
import Config


os.environ["MINERU_MODEL_SOURCE"] = "modelscope"

def main():
    """
    主函数，根据用户选择的模式将PDF转换为Word文档。
    """
    parser = argparse.ArgumentParser(description="PDF to Word Converter")
    parser.add_argument("-i", "--input", help="输入PDF文件的路径", default="input/sample_2")
    parser.add_argument("-o", "--output", help="输出文件的目录", default="output")
    parser.add_argument(
        "-m", "--mode", 
        choices=["content", "format", "debug"], 
        default="format", 
        help="转换模式: 'content' (只保留内容，格式流畅), 'format' (尽量保留原始格式), 'debug' (仅解析PDF，不进行格式转换)"
    )

    args = parser.parse_args()

    # 确保输入文件存在
    if not os.path.exists(args.input):
        logger.error(f"输入文件不存在: {args.input}")
        return

    # 确保输出目录存在
    os.makedirs(args.output, exist_ok=True)

    file_base_name = Path(args.input).stem

    logger.info(f"开始处理文件: {args.input}, 使用模式: {args.mode}")

    if args.mode == 'content':
        # --- 内容模式 ---
        # 1. 解析PDF到Markdown
        md_file_path = parse_pdf_to_files(args.input, args.output)

        if md_file_path and os.path.exists(md_file_path):
            logger.info(f"PDF解析完成，生成的Markdown文件: {md_file_path}")
            
            # 2. 将Markdown转换为Word
            docx_file_name = f"{file_base_name}_content.docx"
            docx_file_path = os.path.join(args.output, docx_file_name)
            
            logger.info(f"正在将Markdown转换为Word: {docx_file_path}")
            try:
                convert_markdown_to_word_with_pandoc(md_file_path, docx_file_path)
                logger.info(f"内容模式转换成功! 输出文件: {docx_file_path}")
            except Exception as e:
                logger.error(f"从Markdown转换为Word失败: {e}")
        else:
            logger.error("PDF解析失败，未生成Markdown文件。")

    elif args.mode == 'format':
        # --- 格式模式 ---
        docx_file_name = f"{file_base_name}_format.docx"
        docx_file_path = os.path.join(args.output, docx_file_name)

        logger.info(f"正在转换PDF为Word（格式模式）: {docx_file_path}")
        try:
            process_and_convert_pdf(args.input, docx_file_path)
            logger.info(f"格式模式转换成功! 输出文件: {docx_file_path}")
        except Exception as e:
            logger.error(f"PDF到Word转换失败: {e}", exc_info=True)


if __name__ == "__main__":
    main() 
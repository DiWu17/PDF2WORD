from pdf2docx import Converter
import fitz
import os
import tempfile

def process_and_convert_pdf(input_pdf_path: str, output_docx_path: str):
    """
    将PDF文件转换为DOCX文件，并在此过程中处理图像以确保兼容性。

    此函数会打开一个PDF文件，遍历其中的图像，将任何具有非标准色彩空间（如CMYK）
    的图像转换为RGB，将修改后的PDF保存到临时文件中，然后将该临时PDF转换为DOCX文件。

    Args:
        input_pdf_path (str): 输入的PDF文件路径。
        output_docx_path (str): 输出的DOCX文件的保存路径。
    """
    # 创建一个临时文件用于保存修改后的PDF
    temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    modified_pdf_path = temp_pdf.name
    temp_pdf.close()

    doc = fitz.open(input_pdf_path)
    try:
        for page in doc:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                
                pix = fitz.Pixmap(doc, xref)

                # 如果色彩空间不是RGB或灰度，则将其转换为RGB
                if pix.colorspace not in (fitz.csRGB, fitz.csGRAY):
                    new_pix = fitz.Pixmap(fitz.csRGB, pix)
                    # 在PDF中更新图像
                    page.replace_image(xref, pixmap=new_pix)
                    new_pix = None # 释放内存
                
                pix = None # 释放内存
        
        # 保存修改后的PDF
        doc.save(modified_pdf_path, garbage=3, deflate=True)
    finally:
        doc.close()

    # 将修改后的PDF转换为DOCX，并确保临时文件被清理
    try:
        cv = Converter(modified_pdf_path)
        cv.convert(output_docx_path)
        cv.close()
    finally:
        os.remove(modified_pdf_path) 
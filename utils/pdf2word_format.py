import json
import os
import win32com.client as win32
from win32com.client import constants

class TextBoxWordReconstructor:
    def __init__(self, json_path, output_path="reconstructed_textbox_smart_font.docx"):
        """
        使用文本框重建 Word 文档，正确处理多页
        
        Args:
            json_path: JSON 文件路径
            output_path: 输出的 Word 文档路径
        """
        self.json_path = json_path
        self.output_path = output_path
        self.json_dir = os.path.dirname(json_path)
        
        # 加载 JSON 数据
        with open(json_path, 'r', encoding='utf-8') as f:
            self.data = json.load(f)
    
    def create_document_with_textboxes(self):
        """创建带文本框的 Word 文档，使用分节符处理多页"""
        # 创建 Word 应用实例
        word = win32.Dispatch('Word.Application')
        word.Visible = True
        
        # 创建新文档
        doc = word.Documents.Add()
        
        # 获取 PDF 信息
        pdf_info = self.data.get('pdf_info', [])
        
        print(f"总页数: {len(pdf_info)}")
        
        for page_idx, page_data in enumerate(pdf_info):
            print(f"处理第 {page_idx + 1} 页...")
            
            # 设定当前操作的节
            # 对于新文档，第一节就是 doc.Sections(1)
            # 对于后续页面，我们需要确保添加了新的页面/节
            if page_idx > 0:
                # 移动到文档末尾并插入分页符，这通常会创建新页面
                word.Selection.EndKey(Unit=6) # wdStory
                word.Selection.InsertBreak(Type=7) # wdPageBreak
            
            # 在新页面上操作，需要重新定位节
            # 使用 ActiveWindow.Selection 来获取当前光标所在节
            current_section = word.Selection.Sections(1)

            # 设置页面大小和边距
            page_size = page_data.get('page_size', [612, 792])
            try:
                current_section.PageSetup.PageWidth = page_size[0]
                current_section.PageSetup.PageHeight = page_size[1]
                current_section.PageSetup.LeftMargin = 0  # 设置为0
                current_section.PageSetup.RightMargin = 0
                current_section.PageSetup.TopMargin = 0
                current_section.PageSetup.BottomMargin = 0
            except Exception as e:
                print(f"设置页面大小失败: {e}")
            
            # 获取当前节的范围
            section_range = current_section.Range
            
            # 处理段落块
            para_blocks = page_data.get('para_blocks', [])
            sorted_blocks = sorted(para_blocks, key=lambda x: x.get('bbox', [0, 0, 0, 0])[1])
            
            print(f"  该页有 {len(sorted_blocks)} 个块")
            
            for block in sorted_blocks:
                self.add_block_to_section(doc, current_section, block, page_idx)
        
        # 保存文档
        doc.SaveAs(os.path.abspath(self.output_path))
        print(f"文档已保存到: {os.path.abspath(self.output_path)}")
        
        # 关闭文档和 Word 应用
        doc.Close(SaveChanges=False)
        word.Quit()
    
    def add_block_to_section(self, doc, section, block, page_idx):
        """将块添加到指定的节"""
        block_type = block.get('type', '')
        bbox = block.get('bbox')
        
        # 如果块类型缺失，尝试从子块中推断
        if not block_type:
            sub_blocks = block.get('blocks')
            for sub_block in sub_blocks:
                sub_block_type = sub_block.get('type')
                if sub_block_type == 'image_body':
                    block_type = 'image'
                    print(f"  推断块类型为 'image'")
                    break
                elif sub_block_type == 'table_body':
                    block_type = 'table'
                    print(f"  推断块类型为 'table'")
                    break

        # 获取节的范围
        section_range = section.Range
        
        # 将绝对坐标转换为相对于当前页面的坐标
        # PDF坐标系：原点在左上角，但需要转换为Word页面相对坐标
        # relative_bbox = self.convert_to_page_relative_coords(bbox, page_idx)
        relative_bbox = bbox
        
        if block_type in ['text', 'title']:
            # 创建文本框，锚定到当前节
            try:
                # 使用相对于页面的坐标
                left = relative_bbox[0]
                top = relative_bbox[1]
                width = relative_bbox[2] - relative_bbox[0]  # 确保最小宽度
                height = relative_bbox[3] - relative_bbox[1]  # 确保最小高度
                
                # 创建锚点
                anchor_para = section_range.Paragraphs.Add()
                anchor_range = anchor_para.Range
                
                # 创建文本框，锚定到特定段落
                shape = doc.Shapes.AddTextbox(
                    Orientation=1,  # msoTextOrientationHorizontal
                    Left=float(left),
                    Top=float(top),
                    Width=float(width),
                    Height=float(height),
                    Anchor=anchor_range
                )
                
                # 设置文本框的相对位置
                try:
                    shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
                    shape.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionPage
                except:
                    pass
                
                # 获取文本内容
                text_content = self.extract_text_from_block(block)
                
                # 设置文本框内容
                text_frame = shape.TextFrame
                text_frame.TextRange.Text = text_content
                
                # 优先使用 avg_size
                if 'avg_size' in block and isinstance(block['avg_size'], (int, float)):
                    calculated_font_size = block['avg_size']
                    print(f"  使用 avg_size: {calculated_font_size} for block index {block.get('index')}")
                else:
                    # 根据 bbox 和字符数量计算字号
                    print(f"  使用 calculate_font_size: {relative_bbox}, {text_content}, {block_type}")
                    calculated_font_size = self.calculate_font_size(relative_bbox, text_content, block_type)
                
                # 设置字体样式
                if block_type == 'title':
                    text_frame.TextRange.Font.Size = calculated_font_size
                    text_frame.TextRange.Font.Bold = True
                else:
                    text_frame.TextRange.Font.Size = calculated_font_size
                
                # Set paragraph line spacing and disable snap to grid
                try:
                    para_format = text_frame.TextRange.ParagraphFormat
                    para_format.LineSpacingRule = 5  # wdLineSpaceMultiple
                    para_format.LineSpacing = 12
                    
                    # Disable "Snap to grid if document grid is defined"
                    para_format.DisableLineHeightGrid = True
                except Exception as e:
                    print(f"  设置段落格式失败: {e}")
                
                # 设置文本框样式
                shape.Line.Visible = False  # 隐藏边框
                shape.Fill.Visible = False   # 透明背景
                text_frame.MarginLeft = 0
                text_frame.MarginRight = 0
                text_frame.MarginTop = 0
                text_frame.MarginBottom = 0
                
            except Exception as e:
                print(f"创建文本框失败: {e}")
                # 降级到普通段落
                try:
                    para = section_range.Paragraphs.Add()
                    para.Range.Text = self.extract_text_from_block(block)
                except Exception as e2:
                    print(f"添加段落也失败: {e2}")
        
        elif block_type == 'image':
            print(f"  检测到图片块，调用 add_image_to_section for block index {block.get('index')}")
            self.add_image_to_section(doc, section, block, relative_bbox)
        
        elif block_type == 'table':
            self.add_table_to_section(doc, section, block, relative_bbox)
    
    def add_image_to_section(self, doc, section, block, bbox):
        """添加图片到节"""
        image_info = self.find_image_info(block)
        
        if image_info and 'image_path' in image_info:
            image_filename = image_info['image_path']
            print(f"  在块 {block.get('index')} 中找到图片: {image_filename}, 准备插入...")
            image_path = os.path.join(self.json_dir, 'images', image_filename)
            
            if os.path.exists(image_path):
                try:
                    # 创建锚点段落
                    anchor_para = section.Range.Paragraphs.Add()
                    anchor_range = anchor_para.Range
                    
                    # 添加浮动图片
                    shape = doc.Shapes.AddPicture(
                        FileName=os.path.abspath(image_path),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=float(bbox[0]),
                        Top=float(bbox[1]),
                        Width=float(bbox[2] - bbox[0]),
                        Height=float(bbox[3] - bbox[1]),
                        Anchor=anchor_range
                    )
                    
                    print(f"  成功插入图片: {image_filename}")

                    # 设置相对位置
                    try:
                        shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
                        shape.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionPage
                    except:
                        pass
                    
                except Exception as e:
                    print(f"插入浮动图片失败: {e}")
                    # 降级到内联图片
                    try:
                        para = section.Range.Paragraphs.Add()
                        inline_shape = doc.InlineShapes.AddPicture(
                            os.path.abspath(image_path),
                            False,
                            True,
                            para.Range
                        )
                        # 设置大小
                        width = bbox[2] - bbox[0]
                        if width > 400:
                            scale = 400.0 / width
                            inline_shape.Width = 400
                            inline_shape.Height = (bbox[3] - bbox[1]) * scale
                        else:
                            inline_shape.Width = width
                            inline_shape.Height = bbox[3] - bbox[1]
                    except Exception as e2:
                        print(f"插入内联图片也失败: {e2}")
            else:
                print(f"图片文件未在磁盘上找到: {image_path}")
        else:
            print(f"  在块 {block.get('index')} 中未找到 image_info 或 image_path")
    
    def add_table_to_section(self, doc, section, block, bbox):
        """添加表格到节"""
        # 检查表格是否有图片
        table_image_info = self.find_table_image_info(block)
        
        if table_image_info and 'image_path' in table_image_info:
            # 表格被保存为图片，直接插入图片
            image_filename = table_image_info['image_path']
            image_path = os.path.join(self.json_dir, 'images', image_filename)
            
            if os.path.exists(image_path):
                try:
                    # 创建锚点段落
                    anchor_para = section.Range.Paragraphs.Add()
                    anchor_range = anchor_para.Range
                    
                    # 添加浮动图片
                    shape = doc.Shapes.AddPicture(
                        FileName=os.path.abspath(image_path),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=float(bbox[0]),
                        Top=float(bbox[1]),
                        Width=float(bbox[2] - bbox[0]),
                        Height=float(bbox[3] - bbox[1]),
                        Anchor=anchor_range
                    )
                    
                    # 设置相对位置
                    try:
                        shape.RelativeHorizontalPosition = 1
                        shape.RelativeVerticalPosition = 1
                    except:
                        pass
                    
                    print(f"插入表格图片: {image_filename}")
                    
                except Exception as e:
                    print(f"插入表格图片失败: {e}")
                    # 降级处理
                    self.add_table_as_text(doc, section, block, bbox)
            else:
                print(f"表格图片文件未在磁盘上找到: {image_path}")
                self.add_table_as_text(doc, section, block, bbox)
        else:
            # 没有图片，使用文本方式
            self.add_table_as_text(doc, section, block, bbox)
    
    def add_table_as_text(self, doc, section, block, bbox):
        """将表格作为文本添加"""
        try:
            anchor_para = section.Range.Paragraphs.Add()
            anchor_range = anchor_para.Range
            
            shape = doc.Shapes.AddTextbox(
                Orientation=1,
                Left=float(bbox[0]),
                Top=float(bbox[1]),
                Width=float(bbox[2] - bbox[0]),
                Height=float(bbox[3] - bbox[1]),
                Anchor=anchor_range
            )
            
            try:
                shape.RelativeHorizontalPosition = 1
                shape.RelativeVerticalPosition = 1
            except:
                pass
            
            # 获取表格内容
            table_content = self.extract_text_from_block(block)
            
            # 设置文本框内容
            text_frame = shape.TextFrame
            text_frame.TextRange.Text = f"[表格]\n{table_content}"
            text_frame.TextRange.Font.Size = 10
            
            # 设置边框
            shape.Line.Visible = True
            shape.Line.Weight = 0.5
            
        except Exception as e:
            print(f"创建表格文本框失败: {e}")
            try:
                para = section.Range.Paragraphs.Add()
                para.Range.Text = f"[表格]\n{self.extract_text_from_block(block)}"
            except Exception as e2:
                print(f"添加表格段落也失败: {e2}")
    
    def find_image_info(self, block):
        """递归深度优先搜索来查找图片信息"""
        if isinstance(block, dict):
            # 检查当前块是否是包含图片路径的图片span
            if block.get('type') == 'image' and 'image_path' in block:
                return block
            
            # 递归搜索 'blocks'
            if 'blocks' in block:
                for sub_block in block['blocks']:
                    found = self.find_image_info(sub_block)
                    if found:
                        return found
            
            # 递归搜索 'lines'
            if 'lines' in block:
                for line in block['lines']:
                    found = self.find_image_info(line)
                    if found:
                        return found
            
            # 递归搜索 'spans'
            if 'spans' in block:
                for span in block['spans']:
                    found = self.find_image_info(span)
                    if found:
                        return found

        # 如果传入的是列表，则遍历列表
        elif isinstance(block, list):
            for item in block:
                found = self.find_image_info(item)
                if found:
                    return found
        
        return None
    
    def find_table_image_info(self, block):
        """查找表格的图片信息"""
        # 不再检查顶层类型，因为可能是推断出来的
        blocks = block.get('blocks', [])
        for sub_block in blocks:
            if sub_block.get('type') == 'table_body':
                lines = sub_block.get('lines', [])
                for line in lines:
                    spans = line.get('spans', [])
                    for span in spans:
                        if span.get('type') == 'table' and 'image_path' in span:
                            return span
        return None
    
    def convert_to_page_relative_coords(self, bbox, page_idx):
        """
        将绝对坐标转换为相对于当前页面的坐标
        
        Args:
            bbox: 原始的绝对坐标 [x1, y1, x2, y2]
            page_idx: 当前页面索引（从0开始）
            
        Returns:
            转换后的相对坐标 [x1, y1, x2, y2]
        """
        if not self.data or 'layout_dets' not in self.data:
            # print(f"警告：无法获取页面信息，使用原始坐标")
            return bbox
        # 获取页面信息
        layout_dets = self.data['layout_dets']
        if page_idx >= len(layout_dets):
            # print(f"警告：页面索引 {page_idx} 超出范围，使用原始坐标")
            return bbox
            
        page_data = layout_dets[page_idx]
        page_size = page_data.get('page_size')  # 默认Letter大小
        page_width, page_height = page_size
        
        # 计算当前页面在整个文档中的起始Y坐标
        # 假设所有页面高度相同，第一页从0开始
        page_start_y = page_idx * page_height
        
        # 转换坐标：从绝对坐标转为页面相对坐标
        # x坐标不变，y坐标需要减去页面起始位置
        relative_x1 = bbox[0]
        relative_y1 = bbox[1] - page_start_y
        relative_x2 = bbox[2] 
        relative_y2 = bbox[3] - page_start_y
        
        # 确保坐标在页面范围内
        relative_x1 = max(0, min(relative_x1, page_width))
        relative_y1 = max(0, min(relative_y1, page_height))
        relative_x2 = max(relative_x1, min(relative_x2, page_width))
        relative_y2 = max(relative_y1, min(relative_y2, page_height))
        
        print(f"  坐标转换：页面{page_idx}, 原始{bbox} -> 相对{[relative_x1, relative_y1, relative_x2, relative_y2]}")
        
        return [relative_x1, relative_y1, relative_x2, relative_y2]

    def calculate_font_size(self, bbox, text, block_type='text'):
        """
        根据 bbox 和字符数量计算合适的字号
        
        Args:
            bbox: [x1, y1, x2, y2] 边界框坐标
            text: 文本内容
            block_type: 块类型 ('title' 或 'text')
        
        Returns:
            float: 计算出的字号
        """
        if block_type == 'title':
            return 12
        else:
            return 10
        
        if not bbox or len(bbox) < 4 or not text.strip():
            # 如果没有有效数据，返回默认字号
            print(f"没有有效数据，返回默认字号")
            return 11
        
        # 计算 bbox 尺寸
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        
        if width <= 0 or height <= 0:
            print(f"bbox 尺寸无效，返回默认字号")
            return 11
        
        # 清理文本并计算字符数（排除空格和换行）
        clean_text = text.replace(' ', '').replace('\n', '').replace('\t', '')
        char_count = len(clean_text)
        
        if char_count == 0:
            print(f"字符数为0，返回默认字号")
            return 11
        
        line_count = 1
        max_iterations = 20  # 防止无穷循环
        
        for i in range(max_iterations):
            estimated_font_size_from_height = (height / line_count) / 1.15
            
            # 如果字号太小，停止计算
            if estimated_font_size_from_height < 3:
                break
                
            # 计算这个字号下需要的宽度
            estimated_width_needed = estimated_font_size_from_height * (char_count / line_count) * 1.15
            
            # 如果估算的宽度小于等于实际宽度，说明可以容纳
            if estimated_width_needed <= width and (estimated_font_size_from_height+1)*line_count <= height:
                break
                
            line_count += 1

        # 方法1：基于高度估算字号
        # 假设每行的高度约等于字号的1.2倍（行距系数）
        estimated_font_size_from_height = (height / line_count) / 1.15
        font_size = estimated_font_size_from_height
        
        # 根据块类型调整
        # if block_type == 'title':
        #     # 标题通常比正文大一些
        #     # 设置合理的范围：12-20
        #     font_size = max(12, min(20, font_size))
        # else:
        #     # 正文字号范围：8-16
        #     font_size = max(8, min(16, font_size))
        
        # 调试信息
        print(f"  块类型: {block_type}, bbox: {bbox}, 字符数: {char_count}")
        print(f"  字号: {font_size:.1f}")
        
        return round(font_size, 1)

    def extract_text_from_block(self, block):
        """从块中提取文本内容"""
        text_parts = []
        
        lines = block.get('lines', [])
        for line in lines:
            spans = line.get('spans', [])
            for span in spans:
                if span.get('type') == 'text':
                    text_parts.append(span.get('content', ''))
                elif span.get('type') == 'inline_equation':
                    text_parts.append(f"[{span.get('content', '')}]")
        
        # 如果是表格，也尝试从 blocks 中提取文本
        if block.get('type') == 'table':
            blocks = block.get('blocks', [])
            for sub_block in blocks:
                if sub_block.get('type') == 'table_caption':
                    caption_text = self.extract_text_from_block(sub_block)
                    if caption_text:
                        text_parts.append(f"\n说明: {caption_text}")
        
        return ' '.join(text_parts)

def main():
    # 设置路径
    json_path = os.path.join('output', 'layout_enriched.json')  # 使用丰富后的JSON
    output_path = os.path.join('output', 'reconstructed_with_avg_font.docx')

    # 检查文件是否存在
    if not os.path.exists(json_path):
        print(f"错误: JSON 文件未找到 at {json_path}")
        return

    # 创建并运行重建器
    reconstructor = TextBoxWordReconstructor(json_path, output_path)
    reconstructor.create_document_with_textboxes()

if __name__ == "__main__":
    main() 
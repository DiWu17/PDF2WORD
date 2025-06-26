import os
import html
import cv2
import numpy as np
from loguru import logger
# from rapid_table import RapidTable, RapidTableInput #不再需要
from openai import OpenAI
import base64
import io

# 导入配置
import sys

# 将项目根目录添加到sys.path，以便能够找到Config.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..')))
from Config import DASHSCOPE_API_KEY, DASHSCOPE_BASE_URL, DASHSCOPE_STREAM_REQUEST


# from mineru.utils.enum_class import ModelPath #不再需要
# from mineru.utils.models_download_utils import auto_download_and_get_model_root_path #不再需要

def encode_image(image):
    buffered = io.BytesIO()
    image.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode("utf-8")

def escape_html(input_string):
    """Escape HTML Entities."""
    return html.escape(input_string)


class RapidTableModel(object):
    def __init__(self):
        """
        初始化RapidTableModel，使用Qwen作为表格识别引擎。
        """
        self.client = OpenAI(
            api_key=DASHSCOPE_API_KEY,
            base_url=DASHSCOPE_BASE_URL,
        )
        logger.info("RapidTableModel (VLM-based) initialized.")
        # slanet_plus_model_path = os.path.join(auto_download_and_get_model_root_path(ModelPath.slanet_plus), ModelPath.slanet_plus)
        # input_args = RapidTableInput(model_type='slanet_plus', model_path=slanet_plus_model_path)
        # self.table_model = RapidTable(input_args)
        # self.ocr_engine = ocr_engine


    # def predict(self, image):
    #     bgr_image = cv2.cvtColor(np.asarray(image), cv2.COLOR_RGB2BGR)
    #
    #     # First check the overall image aspect ratio (height/width)
    #     img_height, img_width = bgr_image.shape[:2]
    #     img_aspect_ratio = img_height / img_width if img_width > 0 else 1.0
    #     img_is_portrait = img_aspect_ratio > 1.2
    #
    #     if img_is_portrait:
    #
    #         det_res = self.ocr_engine.ocr(bgr_image, rec=False)[0]
    #         # Check if table is rotated by analyzing text box aspect ratios
    #         is_rotated = False
    #         if det_res:
    #             vertical_count = 0
    #
    #             for box_ocr_res in det_res:
    #                 p1, p2, p3, p4 = box_ocr_res
    #
    #                 # Calculate width and height
    #                 width = p3[0] - p1[0]
    #                 height = p3[1] - p1[1]
    #
    #                 aspect_ratio = width / height if height > 0 else 1.0
    #
    #                 # Count vertical vs horizontal text boxes
    #                 if aspect_ratio < 0.8:  # Taller than wide - vertical text
    #                     vertical_count += 1
    #                 # elif aspect_ratio > 1.2:  # Wider than tall - horizontal text
    #                 #     horizontal_count += 1
    #
    #             # If we have more vertical text boxes than horizontal ones,
    #             # and vertical ones are significant, table might be rotated
    #             if vertical_count >= len(det_res) * 0.3:
    #                 is_rotated = True
    #
    #             # logger.debug(f"Text orientation analysis: vertical={vertical_count}, det_res={len(det_res)}, rotated={is_rotated}")
    #
    #         # Rotate image if necessary
    #         if is_rotated:
    #             # logger.debug("Table appears to be in portrait orientation, rotating 90 degrees clockwise")
    #             image = cv2.rotate(np.asarray(image), cv2.ROTATE_90_CLOCKWISE)
    #             bgr_image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    #
    #     # Continue with OCR on potentially rotated image
    #     ocr_result = self.ocr_engine.ocr(bgr_image)[0]
    #     if ocr_result:
    #         ocr_result = [[item[0], escape_html(item[1][0]), item[1][1]] for item in ocr_result if
    #                   len(item) == 2 and isinstance(item[1], tuple)]
    #     else:
    #         ocr_result = None
    #
    #
    #     if ocr_result:
    #         table_results = self.table_model(np.asarray(image), ocr_result)
    #         html_code = table_results.pred_html
    #         table_cell_bboxes = table_results.cell_bboxes
    #         logic_points = table_results.logic_points
    #         elapse = table_results.elapse
    #         return html_code, table_cell_bboxes, logic_points, elapse
    #     else:
    #         return None, None, None, None

    def predict(self, image):
        """
        使用大语言模型（VLM）从图片中识别表格并返回HTML。
        根据配置，此方法可以以流式或非流式方式工作。

        Args:
            image: PIL.Image.Image, 输入的表格图片。

        Returns:
            A tuple containing:
            - html_code (str): The recognized table as an HTML string.
            - table_cell_bboxes: None (not provided by this model).
            - logic_points: None (not provided by this model).
            - elapse: A float representing the processing time.
        """
        logger.info(f"使用大模型解析表格 (流式: {DASHSCOPE_STREAM_REQUEST})...")
        start_time = cv2.getTickCount()

        base64_image = encode_image(image)

        prompt = """
        ## 角色
        你是一位精通网页设计的科研内容专家，擅长将学术图像中的表格内容准确提取、翻译，并生成语义清晰、结构合理、视觉美观的 HTML 页面代码。

        ## 任务
        根据提供的学术研究类表格图片，完成以下任务：
        1. 提取图片中的所有表格内容；
        2. 特别注意识别因换行或竖排造成的术语断裂；
        3. 使用语义化的 HTML 标签构建结构清晰的网页表格；
        4. 输出完整 HTML 代码。
        5. 特别注意不要把两个靠的比较近的列合并，比如发表年份材料制备方法，不要合并，分成"年份、材料、制备方法"三列

        ## 输出注意点
        - 表格结构应使用标准 HTML 表格语义标签，包括：
        - <table> 表格容器；
        - <thead> 表头区域；
        - <tbody> 表格主体；
        - <tr> 表格行；
        - <th> 表头单元格（加粗、居中）；
        - <td> 数据单元格（居中）；
        - 所有内容应忠实还原图片中的行列顺序与数据；
        - 不得添加任何解释性文字，仅输出 HTML 源代码。

        ## 输出示例

        ### 正确的输出示例
        <table>
        <thead>
            <tr>
            <th>Material</th>
            <th>Property</th>
            <th>Performance</th>
            </tr>
        </thead>
        <tbody>
            <tr>
            <td>Cellulose</td>
            <td>Thermal Stability</td>
            <td>High</td>
            </tr>
            <tr>
            <td>Graphene</td>
            <td>Conductivity</td>
            <td>Excellent</td>
            </tr>
        </tbody>
        </table>
        ### 错误的输出示例
        <h1>表格内容</h1>
        <p>纤维</p >
        <p>素</p >
        <!-- 错误1：未合并换行词语，未翻译 -->
        <!-- 错误2：使用了不语义化的标签结构 -->
        """
        
        try:
            if DASHSCOPE_STREAM_REQUEST:
                # --- 流式请求 ---
                stream = self.client.chat.completions.create(
                    model="qwen-vl-max-latest",
                    messages=[{
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                }
                            },
                            {"type": "text", "text": prompt}
                        ]
                    }],
                    stream=True,
                    max_tokens=4000,
                )

                html_code_pieces = []
                for chunk in stream:
                    content = chunk.choices[0].delta.content
                    if content is not None:
                        html_code_pieces.append(content)
                
                html_code = "".join(html_code_pieces)
            
            else:
                # --- 非流式请求 ---
                response = self.client.chat.completions.create(
                    model="qwen-vl-max-latest",
                    messages=[{
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                }
                            },
                            {"type": "text", "text": prompt}
                        ]
                    }],
                    stream=False,
                    max_tokens=4000,
                )
                html_code = response.choices[0].message.content

            # 清理返回结果，去除可能的markdown标记
            if html_code.strip().startswith("```html"):
                html_code = html_code.strip()[7:].strip()
                if html_code.endswith("```"):
                    html_code = html_code[:-3].strip()
            
            elapse = (cv2.getTickCount() - start_time) / cv2.getTickFrequency()
            logger.info(f"大模型解析表格成功，耗时: {elapse:.2f}s")
            return html_code, None, None, elapse

        except Exception as e:
            logger.error(f"调用大模型解析表格失败: {e}")
            # 在非流式模式下，我们可能已经有完整的 html_code
            error_context = ""
            try:
                if 'html_code' in locals() and html_code:
                    error_context = f"模型返回结果: {html_code}"
            except NameError:
                pass
            logger.error(f"调用大模型解析表格失败: {e}. {error_context}")
            return None, None, None, None

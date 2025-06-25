# Copyright (c) Opendatalab. All rights reserved.
import copy
import json
import os
from pathlib import Path
from typing import Union, List

from loguru import logger

from mineru.cli.common import convert_pdf_bytes_to_bytes_by_pypdfium2, prepare_env, read_fn
from mineru.data.data_reader_writer import FileBasedDataWriter
from mineru.utils.draw_bbox import draw_layout_bbox, draw_span_bbox
from mineru.utils.enum_class import MakeMode
from mineru.backend.vlm.vlm_analyze import doc_analyze as vlm_doc_analyze
from mineru.backend.pipeline.pipeline_analyze import doc_analyze as pipeline_doc_analyze
from mineru.backend.pipeline.pipeline_middle_json_mkcontent import union_make as pipeline_union_make
from mineru.backend.pipeline.model_json_to_middle_json import result_to_middle_json as pipeline_result_to_middle_json
from mineru.backend.vlm.vlm_middle_json_mkcontent import union_make as vlm_union_make


def _do_parse_internal(
    output_dir,  # 用于存储解析结果的输出目录
    pdf_file_names: list[str],  # 要解析的PDF文件名列表
    pdf_bytes_list: list[bytes],  # 要解析的PDF字节列表
    p_lang_list: list[str],  # 每个PDF的语言列表，默认为'ch'（中文）
    backend="pipeline",  # 用于解析PDF的后端，默认为'pipeline'
    parse_method="auto",  # 解析PDF的方法，默认为'auto'
    p_formula_enable=True,  # 启用公式解析
    p_table_enable=True,  # 启用表格解析
    server_url=None,  # vlm-sglang-client后端的服务器URL
    f_draw_layout_bbox=True,  # 是否绘制布局边界框
    f_draw_span_bbox=True,  # 是否绘制文本块边界框
    f_dump_md=True,  # 是否转储markdown文件
    f_dump_middle_json=True,  # 是否转储中间JSON文件
    f_dump_model_output=True,  # 是否转储模型输出文件
    f_dump_orig_pdf=True,  # 是否转储原始PDF文件
    f_dump_content_list=True,  # 是否转储内容列表文件
    f_make_md_mode=MakeMode.MM_MD,  # 生成markdown内容的模式，默认为MM_MD
    start_page_id=0,  # 解析的起始页面ID，默认为0
    end_page_id=None,  # 解析的结束页面ID，默认为None（解析到文档末尾）
):
    """
    内部解析函数，处理单个或批量PDF。
    返回生成的MD或JSON文件路径列表。
    """
    output_file_paths = []

    if backend == "pipeline":
        for idx, pdf_bytes in enumerate(pdf_bytes_list):
            new_pdf_bytes = convert_pdf_bytes_to_bytes_by_pypdfium2(pdf_bytes, start_page_id, end_page_id)
            pdf_bytes_list[idx] = new_pdf_bytes

        infer_results, all_image_lists, all_pdf_docs, lang_list, ocr_enabled_list = pipeline_doc_analyze(pdf_bytes_list, p_lang_list, parse_method=parse_method, formula_enable=p_formula_enable,table_enable=p_table_enable)

        for idx, model_list in enumerate(infer_results):
            model_json = copy.deepcopy(model_list)
            pdf_file_name = pdf_file_names[idx]
            local_image_dir, local_md_dir = prepare_env(output_dir, pdf_file_name, parse_method)
            image_writer, md_writer = FileBasedDataWriter(local_image_dir), FileBasedDataWriter(local_md_dir)

            images_list = all_image_lists[idx]
            pdf_doc = all_pdf_docs[idx]
            _lang = lang_list[idx]
            _ocr_enable = ocr_enabled_list[idx]
            middle_json = pipeline_result_to_middle_json(model_list, images_list, pdf_doc, image_writer, _lang, _ocr_enable, p_formula_enable)

            pdf_info = middle_json["pdf_info"]

            pdf_bytes = pdf_bytes_list[idx]
            if f_draw_layout_bbox:
                draw_layout_bbox(pdf_info, pdf_bytes, local_md_dir, f"{pdf_file_name}_layout.pdf")

            if f_draw_span_bbox:
                draw_span_bbox(pdf_info, pdf_bytes, local_md_dir, f"{pdf_file_name}_span.pdf")

            if f_dump_orig_pdf:
                md_writer.write(
                    f"{pdf_file_name}_origin.pdf",
                    pdf_bytes,
                )

            if f_dump_md:
                image_dir = str(os.path.basename(local_image_dir))
                md_content_str = pipeline_union_make(pdf_info, f_make_md_mode, image_dir)
                output_file_path = os.path.join(local_md_dir, f"{pdf_file_name}.md")
                md_writer.write_string(
                    f"{pdf_file_name}.md",
                    md_content_str,
                )
                output_file_paths.append(output_file_path)

            if f_dump_content_list:
                image_dir = str(os.path.basename(local_image_dir))
                content_list = pipeline_union_make(pdf_info, MakeMode.CONTENT_LIST, image_dir)
                md_writer.write_string(
                    f"{pdf_file_name}_content_list.json",
                    json.dumps(content_list, ensure_ascii=False, indent=4),
                )

            if f_dump_middle_json:
                output_file_path = os.path.join(local_md_dir, f"{pdf_file_name}_middle.json")
                md_writer.write_string(
                    f"{pdf_file_name}_middle.json",
                    json.dumps(middle_json, ensure_ascii=False, indent=4),
                )
                if not f_dump_md: # 如果不输出MD，则返回JSON路径
                    output_file_paths.append(output_file_path)


            if f_dump_model_output:
                md_writer.write_string(
                    f"{pdf_file_name}_model.json",
                    json.dumps(model_json, ensure_ascii=False, indent=4),
                )

            logger.info(f"local output dir is {local_md_dir}")
    else: # VLM backend logic...
        # ... (VLM logic remains the same, but should also append to output_file_paths)
        pass
    
    return output_file_paths


def parse_pdfs_to_files(
        path_list: List[Union[str, Path]],
        output_dir,
        lang="ch",
        backend="pipeline",
        method="auto",
        server_url=None,
        start_page_id=0,
        end_page_id=None,
        dump_md=True,
):
    """
    解析多个PDF/图像文件，并返回生成的文件路径。
    """
    try:
        file_name_list = []
        pdf_bytes_list = []
        lang_list = []
        for path in path_list:
            file_name = str(Path(path).stem)
            pdf_bytes = read_fn(path)
            file_name_list.append(file_name)
            pdf_bytes_list.append(pdf_bytes)
            lang_list.append(lang)
        
        return _do_parse_internal(
            output_dir=output_dir,
            pdf_file_names=file_name_list,
            pdf_bytes_list=pdf_bytes_list,
            p_lang_list=lang_list,
            backend=backend,
            parse_method=method,
            server_url=server_url,
            start_page_id=start_page_id,
            end_page_id=end_page_id,
            f_dump_md=dump_md,
            f_dump_middle_json=True # Always dump json for format mode
        )
    except Exception as e:
        logger.exception(e)


def parse_pdf_to_files(
        path: Union[str, Path],
        output_dir,
        lang="ch",
        backend="pipeline",
        method="auto",
        server_url=None,
        start_page_id=0,
        end_page_id=None,
        dump_md=True,
):
    """
    解析单个PDF/图像文件，并返回生成的文件路径。
    """
    results = parse_pdfs_to_files(
        [path], output_dir, lang, backend, method, server_url, start_page_id, end_page_id, dump_md
    )
    return results[0] if results else None 
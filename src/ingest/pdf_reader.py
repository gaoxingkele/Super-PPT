# -*- coding: utf-8 -*-
"""PDF 文件读取：提取文本、表格、图片。"""
from pathlib import Path


def read_pdf(file_path: Path, image_dir: Path = None) -> dict:
    """
    读取 PDF 文件，返回结构化内容。

    Args:
        file_path: PDF 文件路径
        image_dir: 图片保存目录（None 则不提取图片）

    Returns:
        {
            "text": str,
            "meta": dict,          # {title, author, ...}
            "tables": list[dict],  # [{headers, rows}, ...]
            "images": list[str],   # 保存的图片路径列表
        }
    """
    # 文本提取 (pypdf)
    from pypdf import PdfReader
    reader = PdfReader(str(file_path))
    pages_text = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages_text.append(text.strip())
    full_text = "\n\n".join(pages_text)

    # 后处理：合并单字符行（PDF 竖排提取常见问题）
    full_text = _fix_single_char_lines(full_text)

    # 元数据
    meta = {}
    if reader.metadata:
        if reader.metadata.title:
            meta["title"] = reader.metadata.title
        if reader.metadata.author:
            meta["author"] = reader.metadata.author

    # 表格提取 (pdfplumber)
    tables = []
    try:
        import pdfplumber
        with pdfplumber.open(str(file_path)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    if table and len(table) >= 2:
                        headers = [str(c or "") for c in table[0]]
                        rows = [[str(c or "") for c in row] for row in table[1:]]
                        tables.append({"headers": headers, "rows": rows})
    except ImportError:
        pass  # pdfplumber 可选

    # 图片提取
    images = []
    if image_dir:
        image_dir = Path(image_dir)
        image_dir.mkdir(parents=True, exist_ok=True)
        img_counter = 0
        for page in reader.pages:
            if hasattr(page, "images"):
                for img in page.images:
                    img_counter += 1
                    ext = ".png"
                    out_name = f"img_{img_counter:03d}{ext}"
                    out_path = image_dir / out_name
                    out_path.write_bytes(img.data)
                    images.append(str(out_path))

    return {"text": full_text, "meta": meta, "tables": tables, "images": images}


def _fix_single_char_lines(text: str) -> str:
    """
    修复 PDF 提取的单字符行问题。
    如果连续多行每行仅有 1 个字符（中日韩字符），则合并为一行。
    同时处理混合情况：正常行与单字行交替出现。
    """
    import re
    lines = text.split("\n")
    result = []
    buffer = []

    for line in lines:
        stripped = line.strip()
        # 单个字符行（中日韩字符、标点、或少数英文字符）
        if len(stripped) <= 1 and stripped:
            buffer.append(stripped)
        else:
            if buffer:
                result.append("".join(buffer))
                buffer = []
            if stripped:
                result.append(stripped)
            elif result and result[-1] != "":
                result.append("")  # 保留段落空行

    if buffer:
        result.append("".join(buffer))

    # 合并过短的连续行（可能是被错误分行的句子片段）
    merged = []
    for line in result:
        if (merged and len(merged[-1]) < 20 and len(line) < 20
                and merged[-1] and line and not line.startswith(("第", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"))
                and not re.match(r'^\d', line)):
            merged[-1] += line
        else:
            merged.append(line)

    return "\n".join(merged)

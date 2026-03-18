# -*- coding: utf-8 -*-
"""Markdown 文件读取：提取正文、前置 YAML 元数据、表格。"""
import re
from pathlib import Path


def read_markdown(file_path: Path) -> dict:
    """
    读取 Markdown 文件，返回结构化内容。

    Returns:
        {
            "text": str,           # 正文全文
            "meta": dict,          # YAML front matter (如有)
            "tables": list[dict],  # 提取的表格 [{headers, rows}, ...]
            "images": list[str],   # 内嵌图片路径
        }
    """
    text = file_path.read_text(encoding="utf-8", errors="replace")

    # 解析 YAML front matter
    meta = {}
    if text.startswith("---"):
        end = text.find("---", 3)
        if end > 0:
            front_matter = text[3:end].strip()
            try:
                import yaml
                meta = yaml.safe_load(front_matter) or {}
            except Exception:
                pass
            text = text[end + 3:].strip()

    # 提取表格
    tables = _extract_tables(text)

    # 提取图片引用
    images = re.findall(r"!\[.*?\]\((.+?)\)", text)

    return {"text": text, "meta": meta, "tables": tables, "images": images}


def _extract_tables(text: str) -> list:
    """从 Markdown 文本中提取表格为结构化数据。"""
    tables = []
    lines = text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # 检测表格头（含 |）
        if "|" in line and i + 1 < len(lines) and re.match(r"^\s*\|?[\s:]*-+[\s:]*\|", lines[i + 1]):
            headers = [c.strip() for c in line.strip("|").split("|")]
            rows = []
            j = i + 2  # 跳过分隔行
            while j < len(lines) and "|" in lines[j]:
                row = [c.strip() for c in lines[j].strip().strip("|").split("|")]
                rows.append(row)
                j += 1
            tables.append({"headers": headers, "rows": rows})
            i = j
        else:
            i += 1
    return tables

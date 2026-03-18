# -*- coding: utf-8 -*-
"""
Step0: 内容获取与统一化。
将各种来源（URL、PDF、DOCX、MD、文件夹）统一为结构化内容。
"""
import json
from pathlib import Path

import src  # noqa: F401
from config import OUTPUT_DIR


def run_ingest(source: str, base: str = None) -> dict:
    """
    统一入口：根据来源类型自动分发到对应读取器。

    Args:
        source: URL、文件路径或目录路径
        base: 输出名称前缀

    Returns:
        {
            "base": str,
            "output_dir": Path,          # output/{base}/
            "raw_content_path": Path,    # output/{base}/raw_content.md
            "raw_meta_path": Path,       # output/{base}/raw_meta.json
            "raw_tables_path": Path,     # output/{base}/raw_tables.json
        }
    """
    source = source.strip()
    source_path = Path(source)

    # 判断来源类型
    if source.startswith(("http://", "https://")):
        result = _ingest_url(source)
    elif source_path.is_dir():
        result = _ingest_directory(source_path)
    elif source_path.is_file():
        result = _ingest_file(source_path)
    else:
        raise FileNotFoundError(f"输入不存在: {source}")

    # 确定 base 名称
    if not base:
        if source.startswith(("http://", "https://")):
            from urllib.parse import urlparse
            base = urlparse(source).path.strip("/").split("/")[-1] or "web_content"
        elif source_path.is_dir():
            base = source_path.name
        else:
            base = source_path.stem
    base = base.replace(" ", "_")

    # 创建输出目录并保存
    out_dir = OUTPUT_DIR / base
    out_dir.mkdir(parents=True, exist_ok=True)
    image_dir = out_dir / "raw_images"

    # 保存原始内容
    raw_content_path = out_dir / "raw_content.md"
    raw_content_path.write_text(result.get("text", ""), encoding="utf-8")

    # 保存元数据
    raw_meta_path = out_dir / "raw_meta.json"
    meta = result.get("meta", {})
    meta["source"] = source
    meta["source_type"] = _detect_source_type(source)
    raw_meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    # 保存表格数据
    raw_tables_path = out_dir / "raw_tables.json"
    tables = result.get("tables", [])
    raw_tables_path.write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[Step0] 内容获取完成: {len(result.get('text', ''))} 字符, {len(tables)} 个表格")

    return {
        "base": base,
        "output_dir": out_dir,
        "raw_content_path": raw_content_path,
        "raw_meta_path": raw_meta_path,
        "raw_tables_path": raw_tables_path,
    }


def _detect_source_type(source: str) -> str:
    if source.startswith(("http://", "https://")):
        return "url"
    p = Path(source)
    if p.is_dir():
        return "directory"
    return p.suffix.lower().lstrip(".")


def _ingest_url(url: str) -> dict:
    """从 URL 爬取内容。"""
    from src.ingest.crawlers import crawl_url
    result = crawl_url(url)
    return {
        "text": result.text,
        "meta": {"title": result.title},
        "tables": [],
        "images": result.images,
    }


def _ingest_file(file_path: Path) -> dict:
    """从单个文件读取内容。"""
    suffix = file_path.suffix.lower()

    if suffix == ".md":
        from src.ingest.md_reader import read_markdown
        return read_markdown(file_path)

    elif suffix == ".docx":
        from src.ingest.docx_reader import read_docx
        return read_docx(file_path)

    elif suffix == ".pdf":
        from src.ingest.pdf_reader import read_pdf
        return read_pdf(file_path)

    elif suffix in (".txt", ".json", ".html"):
        text = file_path.read_text(encoding="utf-8", errors="replace")
        return {"text": text, "meta": {"title": file_path.stem}, "tables": [], "images": []}

    else:
        raise ValueError(f"不支持的文件类型: {suffix}")


def _ingest_directory(dir_path: Path) -> dict:
    """从目录递归读取所有文件。"""
    from src.ingest.dir_scanner import scan_directory
    return scan_directory(dir_path, recursive=True)

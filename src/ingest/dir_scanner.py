# -*- coding: utf-8 -*-
"""文件夹递归扫描：合并目录下所有支持的文件为统一内容。"""
from pathlib import Path

import src  # noqa: F401
from config import INPUT_EXTENSIONS


def scan_directory(dir_path: Path, recursive: bool = True) -> dict:
    """
    扫描目录下所有支持的文件，合并内容。

    Args:
        dir_path: 目录路径
        recursive: 是否递归子目录

    Returns:
        {
            "text": str,           # 合并后的文本
            "meta": dict,          # {title: 目录名, file_count: N}
            "tables": list[dict],  # 所有提取的表格
            "images": list[str],   # 所有提取的图片路径
            "files": list[str],    # 处理的文件列表
        }
    """
    pattern = "**/*" if recursive else "*"
    files = sorted(
        f for f in dir_path.glob(pattern)
        if f.is_file() and f.suffix.lower() in INPUT_EXTENSIONS
    )

    all_text = []
    all_tables = []
    all_images = []
    processed_files = []

    for file_path in files:
        suffix = file_path.suffix.lower()
        try:
            if suffix in (".txt", ".md"):
                if suffix == ".md":
                    from src.ingest.md_reader import read_markdown
                    result = read_markdown(file_path)
                else:
                    result = {"text": file_path.read_text(encoding="utf-8", errors="replace")}

            elif suffix == ".docx":
                from src.ingest.docx_reader import read_docx
                result = read_docx(file_path)

            elif suffix == ".pdf":
                from src.ingest.pdf_reader import read_pdf
                result = read_pdf(file_path)

            elif suffix in (".json", ".html"):
                result = {"text": file_path.read_text(encoding="utf-8", errors="replace")}

            else:
                continue

            text = result.get("text", "")
            if text.strip():
                all_text.append(f"--- {file_path.name} ---\n{text}")
                processed_files.append(str(file_path))
            all_tables.extend(result.get("tables", []))
            all_images.extend(result.get("images", []))

        except Exception as e:
            print(f"[警告] 跳过文件 {file_path.name}: {e}")
            continue

    return {
        "text": "\n\n".join(all_text),
        "meta": {"title": dir_path.name, "file_count": len(processed_files)},
        "tables": all_tables,
        "images": all_images,
        "files": processed_files,
    }

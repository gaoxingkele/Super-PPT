# -*- coding: utf-8 -*-
"""带重试的安全文件写入，应对网盘同步锁等临时文件锁定。"""
import time
from pathlib import Path


def safe_write_text(path: Path, text: str, encoding: str = "utf-8",
                    retries: int = 5, delay: float = 2.0):
    """写入文本文件，遇到 PermissionError 时自动重试。"""
    path = Path(path)
    for attempt in range(1, retries + 1):
        try:
            path.write_text(text, encoding=encoding)
            return
        except PermissionError:
            if attempt == retries:
                raise
            print(f"  [写入重试] {path.name} 被锁定，{delay}s 后第 {attempt+1} 次尝试",
                  flush=True)
            time.sleep(delay)

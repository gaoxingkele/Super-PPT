# -*- coding: utf-8 -*-
"""LaTeX 公式渲染为 PNG 图片。"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


def render_latex_to_png(latex: str, output_path: Path, fontsize: int = 24, dpi: int = 300) -> Path:
    """
    使用 matplotlib mathtext 将 LaTeX 公式渲染为 PNG。

    Args:
        latex: LaTeX 公式字符串
        output_path: 输出路径
        fontsize: 字号
        dpi: 分辨率

    Returns:
        输出文件路径
    """
    fig, ax = plt.subplots(figsize=(0.01, 0.01))
    ax.text(0, 0, f"${latex}$", fontsize=fontsize,
            transform=ax.transAxes, va="center", ha="center")
    ax.axis("off")
    fig.savefig(str(output_path), dpi=dpi, transparent=True, bbox_inches="tight")
    plt.close(fig)
    return output_path

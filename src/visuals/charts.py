# -*- coding: utf-8 -*-
"""
matplotlib/seaborn 统计图表渲染器。
将 slide_plan 中的 visual 规格渲染为 PNG 图片。
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # 非交互式后端
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

import src  # noqa: F401
from config import CHART_DPI, CHART_FIGSIZE


# ============ 字体配置 ============
def _setup_chinese_font():
    """配置中文字体支持。"""
    zh_fonts = ["SimHei", "Microsoft YaHei", "WenQuanYi Micro Hei", "Noto Sans CJK SC"]
    for font_name in zh_fonts:
        try:
            fm.findfont(font_name, fallback_to_default=False)
            plt.rcParams["font.sans-serif"] = [font_name, "Arial"]
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception:
            continue
    # 回退
    plt.rcParams["font.sans-serif"] = ["Arial"]


_setup_chinese_font()


# ============ 颜色工具 ============
def _get_colors(color_scheme: dict, n: int = 6) -> list:
    """从配色方案生成 N 个颜色。"""
    base_colors = [
        color_scheme.get("primary", "#1B365D"),
        color_scheme.get("accent", "#E8612D"),
        color_scheme.get("secondary", "#4A90D9"),
        "#2ECC71", "#9B59B6", "#F39C12", "#E74C3C", "#1ABC9C",
    ]
    return (base_colors * ((n // len(base_colors)) + 1))[:n]


# ============ 统一入口 ============
def render_chart(visual: dict, color_scheme: dict, output_path: Path):
    """
    根据 visual 规格渲染图表并保存为 PNG。

    Args:
        visual: slide_plan 中的 visual 字段
        color_scheme: 配色方案
        output_path: 输出 PNG 路径
    """
    chart_type = visual.get("chart", "bar")
    data = visual.get("data", {})

    renderer = CHART_RENDERERS.get(chart_type)
    if not renderer:
        raise ValueError(f"不支持的图表类型: {chart_type}")

    fig = renderer(data, color_scheme, visual)
    fig.savefig(str(output_path), dpi=CHART_DPI, transparent=True, bbox_inches="tight")
    plt.close(fig)


# ============ 各类图表渲染器 ============

def render_bar(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """柱状图。如果有多系列 series，自动转为分组柱状图。"""
    # 如果有 series 但没有 values，自动用 grouped_bar
    if data.get("series") and not data.get("values"):
        return render_grouped_bar(data, color_scheme, options)

    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    labels = data.get("labels", [])
    values = data.get("values", [])
    colors = _get_colors(color_scheme, len(labels))

    bars = ax.bar(labels, values, color=colors, edgecolor="white", linewidth=0.5)

    # 数据标签
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.02,
                str(val), ha="center", va="bottom", fontsize=14, fontweight="bold")

    # 高亮
    highlight = (options or {}).get("highlight", [])
    for idx in highlight:
        if 0 <= idx < len(bars):
            bars[idx].set_edgecolor(color_scheme.get("accent", "#E8612D"))
            bars[idx].set_linewidth(3)

    ax.set_ylabel(data.get("ylabel", options.get("ylabel", "")), fontsize=14)
    ax.set_title(data.get("title", ""), fontsize=18, fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.3)

    # 注释
    annotation = (options or {}).get("annotation")
    if annotation and "index" in annotation:
        idx = annotation["index"]
        if 0 <= idx < len(bars):
            ax.annotate(
                annotation.get("text", ""),
                xy=(bars[idx].get_x() + bars[idx].get_width() / 2, bars[idx].get_height()),
                xytext=(0, 20), textcoords="offset points",
                ha="center", fontsize=14, fontweight="bold",
                color=color_scheme.get("accent", "#E8612D"),
                arrowprops=dict(arrowstyle="->", color=color_scheme.get("accent", "#E8612D")),
            )

    fig.tight_layout()
    return fig


def render_line(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """折线图。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    labels = data.get("labels", [])
    colors = _get_colors(color_scheme)

    # 支持多系列
    series = data.get("series", {})
    if not series:
        series = {"数据": data.get("values", [])}

    for i, (name, values) in enumerate(series.items()):
        ax.plot(labels, values, marker="o", linewidth=2.5, markersize=8,
                label=name, color=colors[i % len(colors)])
        for j, v in enumerate(values):
            ax.text(j, v + max(values) * 0.02, str(v), ha="center", fontsize=11)

    ax.set_ylabel(data.get("ylabel", (options or {}).get("ylabel", "")), fontsize=14)
    ax.legend(fontsize=12)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(alpha=0.3)
    fig.tight_layout()
    return fig


def render_pie(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """饼图 / 环形图。"""
    fig, ax = plt.subplots(figsize=(12, 10))
    labels = data.get("labels", [])
    values = data.get("values", [])
    colors = _get_colors(color_scheme, len(labels))

    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors, autopct="%1.1f%%",
        startangle=90, pctdistance=0.8, textprops={"fontsize": 13},
    )
    for autotext in autotexts:
        autotext.set_fontweight("bold")

    ax.set_title(data.get("title", ""), fontsize=18, fontweight="bold", pad=20)
    fig.tight_layout()
    return fig


def render_radar(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """雷达图。"""
    categories = data.get("categories", [])
    series = data.get("series", {})
    colors = _get_colors(color_scheme, len(series))

    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(polar=True))

    for i, (name, values) in enumerate(series.items()):
        vals = values + values[:1]
        ax.plot(angles, vals, "o-", linewidth=2, label=name, color=colors[i])
        ax.fill(angles, vals, alpha=0.15, color=colors[i])

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=13)
    ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1), fontsize=12)
    ax.set_title(data.get("title", ""), fontsize=18, fontweight="bold", pad=30)
    fig.tight_layout()
    return fig


def render_heatmap(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """热力图。"""
    import seaborn as sns

    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    matrix = np.array(data.get("matrix", [[]]))
    x_labels = data.get("x_labels", [])
    y_labels = data.get("y_labels", [])

    sns.heatmap(matrix, annot=True, fmt=".1f", cmap="YlOrRd",
                xticklabels=x_labels, yticklabels=y_labels, ax=ax,
                linewidths=0.5, linecolor="white")
    ax.set_title(data.get("title", ""), fontsize=18, fontweight="bold")
    fig.tight_layout()
    return fig


def render_scatter(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """散点图。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    x = data.get("x", [])
    y = data.get("y", [])
    labels = data.get("labels", [])

    ax.scatter(x, y, s=100, c=color_scheme.get("primary", "#1B365D"),
               alpha=0.7, edgecolors="white")

    for i, label in enumerate(labels):
        if i < len(x):
            ax.annotate(label, (x[i], y[i]), fontsize=10, ha="center", va="bottom")

    ax.set_xlabel(data.get("xlabel", ""), fontsize=14)
    ax.set_ylabel(data.get("ylabel", ""), fontsize=14)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(alpha=0.3)
    fig.tight_layout()
    return fig


def render_donut(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """环形图。"""
    fig, ax = plt.subplots(figsize=(12, 10))
    labels = data.get("labels", [])
    values = data.get("values", [])
    colors = _get_colors(color_scheme, len(labels))

    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors, autopct="%1.1f%%",
        startangle=90, pctdistance=0.8, textprops={"fontsize": 13},
        wedgeprops={"width": 0.4},
    )
    centre_circle = plt.Circle((0, 0), 0.35, fc="white")
    ax.add_patch(centre_circle)

    fig.tight_layout()
    return fig


def render_grouped_bar(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """分组柱状图。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    labels = data.get("labels", [])
    series = data.get("series", {})
    colors = _get_colors(color_scheme, len(series))

    x = np.arange(len(labels))
    width = 0.8 / max(len(series), 1)

    for i, (name, values) in enumerate(series.items()):
        offset = (i - len(series) / 2 + 0.5) * width
        bars = ax.bar(x + offset, values, width, label=name, color=colors[i])
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                    str(val), ha="center", va="bottom", fontsize=10)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=12)
    ax.legend(fontsize=12)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def render_waterfall(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """瀑布图。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    labels = data.get("labels", [])
    values = data.get("values", [])

    cumulative = [0]
    for v in values[:-1]:
        cumulative.append(cumulative[-1] + v)

    positive_color = color_scheme.get("primary", "#1B365D")
    negative_color = color_scheme.get("accent", "#E8612D")
    total_color = color_scheme.get("secondary", "#4A90D9")

    bar_colors = []
    for i, v in enumerate(values):
        if i == len(values) - 1:
            bar_colors.append(total_color)
        elif v >= 0:
            bar_colors.append(positive_color)
        else:
            bar_colors.append(negative_color)

    bottoms = cumulative[:len(values)]
    bottoms[-1] = 0  # 最后一个（总计）从 0 开始

    ax.bar(labels, values, bottom=bottoms, color=bar_colors, edgecolor="white")

    for i, (v, b) in enumerate(zip(values, bottoms)):
        ax.text(i, b + v + max(abs(v) for v in values) * 0.02,
                str(v), ha="center", fontsize=12, fontweight="bold")

    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def render_funnel(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """漏斗图（横向条形图模拟）。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    labels = data.get("labels", [])
    values = data.get("values", [])
    colors = _get_colors(color_scheme, len(labels))

    max_val = max(values) if values else 1
    for i, (label, val) in enumerate(zip(labels, values)):
        width = val / max_val
        left = (1 - width) / 2
        ax.barh(len(labels) - 1 - i, width, left=left, height=0.7,
                color=colors[i], edgecolor="white")
        ax.text(0.5, len(labels) - 1 - i, f"{label}: {val}",
                ha="center", va="center", fontsize=14, fontweight="bold")

    ax.set_xlim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 渲染器注册表 ============
CHART_RENDERERS = {
    "bar": render_bar,
    "line": render_line,
    "pie": render_pie,
    "donut": render_donut,
    "radar": render_radar,
    "heatmap": render_heatmap,
    "scatter": render_scatter,
    "grouped_bar": render_grouped_bar,
    "waterfall": render_waterfall,
    "funnel": render_funnel,
}

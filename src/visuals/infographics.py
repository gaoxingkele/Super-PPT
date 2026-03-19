# -*- coding: utf-8 -*-
"""
信息图生成器。
优先使用 Google Gemini API 生成专业信息图 PNG，失败时回退到 matplotlib 本地渲染。
支持类型：process_flow, stat_display, timeline, hierarchy, comparison,
          matrix, network, pyramid, cycle。
"""
import base64
import json
import logging
import time
from datetime import datetime
from pathlib import Path

import httpx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

import src  # noqa: F401
from config import GEMINI_API_KEY, CHART_DPI, CLOUBIC_ENABLED, CLOUBIC_API_KEY, CLOUBIC_BASE_URL, CLOUBIC_IMAGE_MODEL

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 中文字体配置（matplotlib 回退渲染用）
# ---------------------------------------------------------------------------
import matplotlib.font_manager as fm

_zh_fonts = ["SimHei", "Microsoft YaHei", "WenQuanYi Micro Hei", "Noto Sans CJK SC"]
for _fn in _zh_fonts:
    try:
        fm.findfont(_fn, fallback_to_default=False)
        plt.rcParams["font.sans-serif"] = [_fn, "Arial"]
        plt.rcParams["axes.unicode_minus"] = False
        break
    except Exception:
        continue

# ---------------------------------------------------------------------------
# Gemini API 配置
# ---------------------------------------------------------------------------
_GEMINI_IMAGE_MODEL = "gemini-3.1-flash-image-preview"
_GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/"
    f"models/{_GEMINI_IMAGE_MODEL}:generateContent"
)
_GEMINI_TIMEOUT = 120  # seconds
_GEMINI_MAX_RETRIES = 2


def _ts() -> str:
    """返回当前时间戳字符串，用于日志。"""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Gemini 图片生成
# ---------------------------------------------------------------------------

_TYPE_GUIDELINES = {
    "process_flow": {
        "name": "Process Flow / Step-by-Step",
        "guidelines": (
            "PROCESS FLOW INFOGRAPHIC REQUIREMENTS:\n"
            "- Numbered steps (1, 2, 3...) clearly visible with large numbers\n"
            "- Directional arrows or connecting lines between steps\n"
            "- Action-oriented icons or illustrations for each step\n"
            "- Brief action text for each step (bold title + short description)\n"
            "- Clear start and end indicators\n"
            "- Logical flow direction (left-to-right or top-to-bottom)\n"
            "- Each step box has rounded corners with distinct background color\n"
            "- Modern flat design with subtle shadows"
        ),
    },
    "stat_display": {
        "name": "KPI / Big Number Statistics",
        "guidelines": (
            "STATISTICAL DISPLAY REQUIREMENTS:\n"
            "- Large, bold numbers that are immediately readable (biggest visual element)\n"
            "- Clear data visualization (bar charts, donut charts, or gauge meters)\n"
            "- Data callouts with context (e.g., '15% increase ↑')\n"
            "- Trend indicators (arrows, growth symbols) in green/red\n"
            "- Each KPI in its own card with subtle background\n"
            "- Clean grid alignment for data elements\n"
            "- Numbers should be at least 3x larger than label text\n"
            "- Use accent color for the most important metric"
        ),
    },
    "timeline": {
        "name": "Horizontal Timeline",
        "guidelines": (
            "TIMELINE INFOGRAPHIC REQUIREMENTS:\n"
            "- Clear chronological flow with a prominent connecting line/path\n"
            "- Prominent date/year markers on the line\n"
            "- Event nodes (circles or icons) at each date point\n"
            "- Brief event descriptions above or below (alternating)\n"
            "- Consistent spacing between events\n"
            "- Visual progression indicating time direction\n"
            "- Start and end points clearly marked with distinct shapes\n"
            "- Color gradient or distinct colors for different time periods"
        ),
    },
    "hierarchy": {
        "name": "Hierarchy / Org-Chart Tree",
        "guidelines": (
            "HIERARCHY INFOGRAPHIC REQUIREMENTS:\n"
            "- Clear tree structure from top to bottom\n"
            "- Distinct levels with visual separation and different colors\n"
            "- Top node is largest and most prominent (primary color)\n"
            "- Connecting lines between parent and child nodes\n"
            "- Labels inside each node box\n"
            "- Balanced, centered composition\n"
            "- Rounded rectangle nodes with subtle shadows\n"
            "- Child nodes slightly smaller than parent nodes"
        ),
    },
    "comparison": {
        "name": "Side-by-Side Comparison",
        "guidelines": (
            "COMPARISON INFOGRAPHIC REQUIREMENTS:\n"
            "- Symmetrical side-by-side layout with clear division\n"
            "- Clear headers for each option being compared (different colors)\n"
            "- Matching rows/categories for fair comparison\n"
            "- Visual indicators (checkmarks ✓, X marks ✗, star ratings)\n"
            "- Balanced visual weight on both sides\n"
            "- 'VS' badge or divider in the center\n"
            "- Summary or verdict section at the bottom if applicable\n"
            "- Use primary color for left, accent color for right"
        ),
    },
    "matrix": {
        "name": "Grid / Matrix Cards",
        "guidelines": (
            "MATRIX INFOGRAPHIC REQUIREMENTS:\n"
            "- Clean card-based grid layout (2x2, 2x3, or 3x3)\n"
            "- Each card has a colored top accent bar\n"
            "- Card title in bold, description in smaller text below\n"
            "- Icons or small illustrations for each card\n"
            "- Consistent card sizing and spacing\n"
            "- Light background colors with darker text\n"
            "- Subtle shadows for depth\n"
            "- Clear visual hierarchy within each card"
        ),
    },
    "network": {
        "name": "Network / Relationship Diagram",
        "guidelines": (
            "NETWORK INFOGRAPHIC REQUIREMENTS:\n"
            "- Central node (largest, primary color) connected to surrounding nodes\n"
            "- Circular or radial arrangement of outer nodes\n"
            "- Connecting lines showing relationships\n"
            "- Each node has a label inside or beside it\n"
            "- Different colors for different categories of nodes\n"
            "- Line thickness indicates relationship strength\n"
            "- Clean, uncluttered layout with adequate spacing\n"
            "- Optional: small icons inside nodes"
        ),
    },
    "pyramid": {
        "name": "Pyramid Layers",
        "guidelines": (
            "PYRAMID INFOGRAPHIC REQUIREMENTS:\n"
            "- Clear pyramid shape with distinct horizontal layers\n"
            "- Narrowest at top, widest at bottom\n"
            "- Each layer has a different color (gradient or distinct)\n"
            "- Layer labels centered within each section\n"
            "- White text on colored backgrounds for readability\n"
            "- Optional: brief description beside each layer\n"
            "- Balanced, centered composition\n"
            "- 3D effect or clean flat design"
        ),
    },
    "cycle": {
        "name": "Circular Cycle",
        "guidelines": (
            "CYCLE INFOGRAPHIC REQUIREMENTS:\n"
            "- Nodes arranged in a circular pattern\n"
            "- Curved arrows connecting each node to the next\n"
            "- Each node is a circle or rounded rectangle with distinct color\n"
            "- Node labels centered inside each shape\n"
            "- Arrow direction shows the cycle flow (clockwise)\n"
            "- Optional center text describing the overall cycle\n"
            "- Balanced spacing between all nodes\n"
            "- Professional, clean design"
        ),
    },
}


def _build_gemini_prompt(visual: dict, color_scheme: dict) -> str:
    """根据 visual 和 color_scheme 构建增强版 Gemini 信息图生成 prompt（Nano Banana 风格）。"""
    infographic_type = visual.get("infographic_type", "process_flow")
    data = visual.get("data", {})
    description = visual.get("description", "")

    primary = color_scheme.get("primary", "#1B365D")
    accent = color_scheme.get("accent", "#E8612D")
    secondary = color_scheme.get("secondary", "#4A90D9")

    type_info = _TYPE_GUIDELINES.get(infographic_type, {})
    type_name = type_info.get("name", infographic_type)
    type_guidelines = type_info.get("guidelines", "")

    # 将 data 序列化为可读文本
    if isinstance(data, dict) and data:
        data_text = json.dumps(data, ensure_ascii=False, indent=2, default=str)
    else:
        data_text = str(data) if data else ""

    parts = []
    parts.append(f"Create a PROFESSIONAL, PUBLICATION-QUALITY infographic in {type_name} style.")
    parts.append("")
    parts.append(f"COLOR SCHEME:")
    parts.append(f"- Primary: {primary} (main headings, key elements)")
    parts.append(f"- Secondary: {secondary} (supporting elements, backgrounds)")
    parts.append(f"- Accent: {accent} (highlights, call-to-action, important numbers)")
    parts.append(f"- Background: clean white (#FFFFFF)")
    parts.append("")

    if type_guidelines:
        parts.append(type_guidelines)
        parts.append("")

    if data_text:
        parts.append(f"DATA TO VISUALIZE:")
        parts.append(data_text)
        parts.append("")
    if description:
        parts.append(f"CONTEXT: {description}")
        parts.append("")

    parts.append("DESIGN REQUIREMENTS:")
    parts.append("- Image size: 1920x1080 pixels (16:9 ratio for presentation slides)")
    parts.append("- Modern flat design with clean typography")
    parts.append("- Clear visual hierarchy: most important elements should be most prominent")
    parts.append("- Professional corporate style suitable for business presentations")
    parts.append("- Bold headlines, readable body text (no tiny or overlapping text)")
    parts.append("- Sufficient contrast for readability")
    parts.append("- Harmonious colors that work together")
    parts.append("- Subtle shadows and rounded corners for modern look")
    parts.append("- CRITICAL: ALL text, labels, numbers, titles, annotations MUST be in Simplified Chinese (简体中文). Do NOT use any English text.")
    parts.append("- Text must be large enough to read clearly (minimum 24pt equivalent for labels, 36pt for headlines)")
    parts.append("- Maximum 6-8 key elements per infographic, do not overcrowd")
    parts.append("")
    parts.append("STRICTLY FORBIDDEN:")
    parts.append("- No English text anywhere in the image")
    parts.append("- No watermarks, borders, or meta-text")
    parts.append("- No instructions or annotations")
    parts.append("- No blurry or low-resolution elements")
    parts.append("- No overlapping text or cramped layouts")
    parts.append("- No tiny unreadable text")

    return "\n".join(parts)


import re as _re


def _extract_image_from_cloubic(data: dict) -> bytes:
    """从 Cloubic 响应中提取 base64 图片（支持 markdown/data URI/list 格式）。"""
    for choice in data.get("choices", []):
        content = choice.get("message", {}).get("content", "")
        if isinstance(content, str) and "base64," in content:
            match = _re.search(r"data:image/[a-z]+;base64,([A-Za-z0-9+/=\s]+)", content)
            if match:
                b64_data = match.group(1).replace("\n", "").replace(" ", "")
                try:
                    return base64.b64decode(b64_data)
                except Exception:
                    pass
        if isinstance(content, list):
            for part in content:
                if isinstance(part, dict):
                    img_url = ""
                    if part.get("type") == "image_url":
                        img_url = part.get("image_url", {}).get("url", "")
                    elif part.get("type") == "image":
                        img_url = part.get("url", "") or part.get("data", "")
                    if img_url and "base64," in img_url:
                        b64_data = img_url.split(",", 1)[1]
                        try:
                            return base64.b64decode(b64_data)
                        except Exception:
                            pass
    return b""


def _try_cloubic_image_generation(prompt: str, output_path: Path) -> bool:
    """通过 Cloubic OpenAI 兼容接口生成信息图。"""
    import config as _cfg
    if not _cfg.CLOUBIC_ENABLED or not _cfg.CLOUBIC_API_KEY:
        return False

    image_model = _cfg.CLOUBIC_IMAGE_MODEL or _GEMINI_IMAGE_MODEL
    payload = {
        "model": image_model,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 4096,
    }
    url = f"{_cfg.CLOUBIC_BASE_URL}/chat/completions"
    headers = {"Authorization": f"Bearer {_cfg.CLOUBIC_API_KEY}", "Content-Type": "application/json"}

    for attempt in range(1, _GEMINI_MAX_RETRIES + 1):
        try:
            logger.info("[%s] Cloubic 信息图生成 attempt %d/%d | model: %s",
                        _ts(), attempt, _GEMINI_MAX_RETRIES, image_model)
            with httpx.Client(timeout=_GEMINI_TIMEOUT, proxy=None) as client:
                resp = client.post(url, json=payload, headers=headers)

            if resp.status_code != 200:
                logger.warning("[%s] Cloubic API 返回 HTTP %d: %s", _ts(), resp.status_code, resp.text[:300])
                continue

            data = resp.json()
            img_bytes = _extract_image_from_cloubic(data)
            if img_bytes:
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(img_bytes)
                logger.info("[%s] Cloubic 信息图已保存: %s (%d bytes)", _ts(), output_path, len(img_bytes))
                return True

            logger.warning("[%s] Cloubic 响应中未找到图片数据 (attempt %d)", _ts(), attempt)

        except httpx.TimeoutException:
            logger.warning("[%s] Cloubic API 超时 (attempt %d)", _ts(), attempt)
        except Exception as exc:
            logger.warning("[%s] Cloubic API 异常 (attempt %d): %s", _ts(), attempt, exc)

    return False


def _try_gemini_generation(prompt: str, output_path: Path) -> bool:
    """
    调用 Gemini 图片生成 API，成功时将 PNG 写入 output_path 并返回 True。
    失败（无 API key、网络错误、超时、API 报错）时返回 False。
    最多重试 _GEMINI_MAX_RETRIES 次。
    """
    if not GEMINI_API_KEY:
        logger.info("[%s] Gemini API key 未配置，跳过 Gemini 生成", _ts())
        return False

    url = f"{_GEMINI_URL}?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseModalities": ["TEXT", "IMAGE"]},
    }

    for attempt in range(1, _GEMINI_MAX_RETRIES + 1):
        try:
            logger.info(
                "[%s] Gemini 信息图生成 attempt %d/%d ...",
                _ts(), attempt, _GEMINI_MAX_RETRIES,
            )
            with httpx.Client(timeout=_GEMINI_TIMEOUT) as client:
                resp = client.post(url, json=payload)

            if resp.status_code != 200:
                logger.warning(
                    "[%s] Gemini API 返回 HTTP %d: %s",
                    _ts(), resp.status_code, resp.text[:500],
                )
                continue

            result = resp.json()

            # 从响应中提取 inlineData 图片
            candidates = result.get("candidates", [])
            for candidate in candidates:
                content = candidate.get("content", {})
                parts = content.get("parts", [])
                for part in parts:
                    inline_data = part.get("inlineData")
                    if inline_data and inline_data.get("mimeType", "").startswith("image/"):
                        img_bytes = base64.b64decode(inline_data["data"])
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                        output_path.write_bytes(img_bytes)
                        logger.info(
                            "[%s] Gemini 信息图已保存: %s (%d bytes)",
                            _ts(), output_path, len(img_bytes),
                        )
                        return True

            logger.warning("[%s] Gemini 响应中未找到图片数据 (attempt %d)", _ts(), attempt)

        except httpx.TimeoutException:
            logger.warning("[%s] Gemini API 超时 (attempt %d)", _ts(), attempt)
        except Exception as exc:
            logger.warning("[%s] Gemini API 异常 (attempt %d): %s", _ts(), attempt, exc)

    logger.info("[%s] Gemini 生成失败，将回退到 matplotlib", _ts())
    return False


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def render_infographic(visual: dict, color_scheme: dict, output_path: Path):
    """生成信息图并保存为 PNG。优先 Gemini API，失败回退 matplotlib。"""
    output_path = Path(output_path)

    # 1) 尝试 Cloubic 路由
    prompt = _build_gemini_prompt(visual, color_scheme)
    if _try_cloubic_image_generation(prompt, output_path):
        return

    # 2) 尝试 Gemini 直连
    if _try_gemini_generation(prompt, output_path):
        return

    # 3) 回退到 matplotlib
    logger.info("[%s] 使用 matplotlib 回退渲染信息图", _ts())
    infographic_type = visual.get("infographic_type", "process_flow")
    data = visual.get("data", {})
    description = visual.get("description", "")

    renderer = _INFOGRAPHIC_RENDERERS.get(infographic_type, _render_placeholder)
    fig = renderer(data, description, color_scheme)
    fig.savefig(str(output_path), dpi=CHART_DPI, transparent=True, bbox_inches="tight")
    plt.close(fig)


# ===================================================================
#  以下为 matplotlib 回退渲染器（完整保留原有实现）
# ===================================================================

def _hex_to_rgb(hex_color: str):
    """Convert hex to (r, g, b) float tuple."""
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))


def _lighten(hex_color: str, factor: float = 0.3) -> tuple:
    """Lighten a hex color."""
    r, g, b = _hex_to_rgb(hex_color)
    return (r + (1 - r) * factor, g + (1 - g) * factor, b + (1 - b) * factor)


def _get_palette(color_scheme: dict, n: int) -> list:
    """Generate n colors from scheme."""
    base = [
        color_scheme.get("primary", "#1B365D"),
        color_scheme.get("accent", "#E8612D"),
        color_scheme.get("secondary", "#4A90D9"),
        "#2ECC71", "#9B59B6", "#F39C12", "#E74C3C", "#1ABC9C",
    ]
    colors = (base * ((n // len(base)) + 1))[:n]
    return colors


# ============ 流程图 ============
def _render_process_flow(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """流程图：横向圆角矩形 + 渐变色 + 箭头。"""
    stages = data.get("stages", [])
    if not stages:
        parts = description.replace("\u2192", "|").replace("->", "|").split("|")
        stages = [{"name": p.strip(), "detail": ""} for p in parts if p.strip()]
    if not stages:
        stages = [{"name": "步骤", "detail": ""}]

    n = len(stages)
    fig, ax = plt.subplots(figsize=(max(n * 4.5, 14), 5))

    colors = _get_palette(color_scheme, n)
    accent = color_scheme.get("accent", "#E8612D")

    box_w = 0.65 / n
    box_h = 0.45
    y_center = 0.5

    for i, stage in enumerate(stages):
        x = (i + 0.5) / n
        c = colors[i]
        # 圆角矩形 + 阴影
        shadow = FancyBboxPatch(
            (x - box_w / 2 + 0.005, y_center - box_h / 2 - 0.01), box_w, box_h,
            boxstyle="round,pad=0.02", facecolor="#00000010", edgecolor="none",
        )
        ax.add_patch(shadow)
        rect = FancyBboxPatch(
            (x - box_w / 2, y_center - box_h / 2), box_w, box_h,
            boxstyle="round,pad=0.02", facecolor=c, edgecolor="white", linewidth=2,
        )
        ax.add_patch(rect)

        # 编号圆圈
        circle = plt.Circle((x - box_w / 2 + 0.02, y_center + box_h / 2 - 0.02),
                            0.025, color="white", alpha=0.9, zorder=5)
        ax.add_patch(circle)
        ax.text(x - box_w / 2 + 0.02, y_center + box_h / 2 - 0.02,
                str(i + 1), ha="center", va="center", fontsize=10,
                fontweight="bold", color=c, zorder=6)

        # 文字
        ax.text(x, y_center + 0.06, stage.get("name", ""), ha="center", va="center",
                fontsize=13, fontweight="bold", color="white", zorder=5)
        if stage.get("detail"):
            ax.text(x, y_center - 0.08, stage["detail"], ha="center", va="center",
                    fontsize=9, color=(1, 1, 1, 0.8), zorder=5)

        # 箭头
        if i < n - 1:
            ax.annotate("",
                        xy=((i + 1.5) / n - box_w / 2 - 0.01, y_center),
                        xytext=((i + 0.5) / n + box_w / 2 + 0.01, y_center),
                        arrowprops=dict(arrowstyle="-|>", color=accent, lw=3,
                                        mutation_scale=20))

    ax.set_xlim(-0.02, 1.02)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ KPI 大数字 ============
def _render_stat_display(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """KPI 大数字展示，带装饰性底部色条。"""
    kpis = data.get("kpis", data.get("items", []))
    if not kpis:
        kpis = [{"label": "数据", "value": "N/A"}]

    n = len(kpis)
    colors = _get_palette(color_scheme, n)
    fig, axes = plt.subplots(1, n, figsize=(n * 5, 4.5))
    if n == 1:
        axes = [axes]

    for i, (ax, kpi) in enumerate(zip(axes, kpis)):
        c = colors[i]
        # 底部色条
        bar = FancyBboxPatch((0.1, 0.02), 0.8, 0.08,
                             boxstyle="round,pad=0.01", facecolor=c,
                             transform=ax.transAxes, clip_on=False)
        ax.add_patch(bar)

        # 大数字
        ax.text(0.5, 0.6, str(kpi.get("value", "")), ha="center", va="center",
                fontsize=52, fontweight="bold", color=c, transform=ax.transAxes)
        # 标签
        ax.text(0.5, 0.22, str(kpi.get("label", "")), ha="center", va="center",
                fontsize=15, color="#555555", transform=ax.transAxes)

        # 趋势箭头
        trend = kpi.get("trend", "")
        if trend == "up":
            ax.text(0.88, 0.65, "\u25b2", ha="center", fontsize=22, color="#2ECC71",
                    transform=ax.transAxes)
        elif trend == "down":
            ax.text(0.88, 0.65, "\u25bc", ha="center", fontsize=22, color="#E74C3C",
                    transform=ax.transAxes)

        ax.axis("off")

    fig.tight_layout(pad=2)
    return fig


# ============ 时间线 ============
def _render_timeline(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """水平时间线：交替上下布局。"""
    events = data.get("events", data.get("stages", []))
    if not events:
        parts = description.replace("\u2192", "|").replace("->", "|").replace("\u3001", "|").split("|")
        events = [{"date": "", "title": p.strip()} for p in parts if p.strip()]
    if not events:
        events = [{"date": "", "title": "事件"}]

    n = len(events)
    colors = _get_palette(color_scheme, n)
    primary = color_scheme.get("primary", "#1B365D")

    fig, ax = plt.subplots(figsize=(max(n * 3.5, 14), 6))

    # 主轴线
    ax.plot([0.05, 0.95], [0.5, 0.5], color=primary, linewidth=3, solid_capstyle="round")

    for i, event in enumerate(events):
        x = 0.05 + (i / max(n - 1, 1)) * 0.9
        above = i % 2 == 0

        # 节点圆
        circle = plt.Circle((x, 0.5), 0.025, color=colors[i], zorder=5)
        ax.add_patch(circle)
        ax.plot([x, x], [0.5, 0.68 if above else 0.32],
                color=colors[i], linewidth=2, zorder=4)

        # 文字
        y_title = 0.75 if above else 0.22
        y_date = 0.85 if above else 0.12
        va = "bottom" if above else "top"

        title = event.get("title", event.get("name", ""))
        date = event.get("date", event.get("detail", ""))

        ax.text(x, y_title, title[:15], ha="center", va=va,
                fontsize=11, fontweight="bold", color=colors[i])
        if date:
            ax.text(x, y_date, str(date)[:15], ha="center", va=va,
                    fontsize=9, color="#888888")

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 层级/树状图 ============
def _render_hierarchy(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """层级结构：从上到下的树状图。"""
    nodes = data.get("nodes", [])
    if not nodes:
        return _render_placeholder(data, description, color_scheme)

    colors = _get_palette(color_scheme, 8)
    primary = color_scheme.get("primary", "#1B365D")

    fig, ax = plt.subplots(figsize=(14, 8))

    def draw_node(x, y, text, color, width=0.15, height=0.08):
        rect = FancyBboxPatch((x - width / 2, y - height / 2), width, height,
                              boxstyle="round,pad=0.01", facecolor=color,
                              edgecolor="white", linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x, y, text[:12], ha="center", va="center",
                fontsize=10, fontweight="bold", color="white")

    # Layer 0: root
    root = nodes[0] if nodes else {"name": "Root"}
    draw_node(0.5, 0.85, root.get("name", "Root"), primary, 0.25, 0.1)

    # Layer 1: children
    children = root.get("children", nodes[1:] if len(nodes) > 1 else [])
    if isinstance(children, list) and children:
        nc = len(children)
        for i, child in enumerate(children):
            name = child if isinstance(child, str) else child.get("name", "")
            x = 0.15 + (i / max(nc - 1, 1)) * 0.7 if nc > 1 else 0.5
            draw_node(x, 0.55, name, colors[i % len(colors)], 0.18, 0.08)
            ax.plot([0.5, x], [0.8, 0.59], color="#CCCCCC", linewidth=1.5)

            # Layer 2: grandchildren
            grandchildren = child.get("children", []) if isinstance(child, dict) else []
            if grandchildren:
                ngc = len(grandchildren)
                x_start = max(0.05, x - 0.15)
                x_end = min(0.95, x + 0.15)
                for j, gc in enumerate(grandchildren[:4]):
                    gc_name = gc if isinstance(gc, str) else gc.get("name", "")
                    gx = x_start + (j / max(ngc - 1, 1)) * (x_end - x_start) if ngc > 1 else x
                    draw_node(gx, 0.25, gc_name, _lighten(colors[i % len(colors)], 0.2), 0.12, 0.06)
                    ax.plot([x, gx], [0.51, 0.28], color="#DDDDDD", linewidth=1)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 网络/关系图 ============
def _render_network(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """关系网络图：圆形布局 + 中心节点。"""
    nodes = data.get("nodes", data.get("items", []))
    if not nodes:
        parts = description.replace("\u3001", "|").replace("\uff0c", "|").replace(",", "|").split("|")
        nodes = [p.strip() for p in parts if p.strip()][:8]
    if not nodes:
        nodes = ["节点"]

    colors = _get_palette(color_scheme, len(nodes) + 1)
    primary = color_scheme.get("primary", "#1B365D")

    fig, ax = plt.subplots(figsize=(10, 10))

    n = len(nodes)
    # 中心节点
    center = data.get("center", data.get("title", "核心"))
    if isinstance(center, dict):
        center = center.get("name", "核心")
    circle = plt.Circle((0.5, 0.5), 0.08, color=primary, zorder=5)
    ax.add_patch(circle)
    ax.text(0.5, 0.5, str(center)[:8], ha="center", va="center",
            fontsize=14, fontweight="bold", color="white", zorder=6)

    # 外围节点
    for i, node in enumerate(nodes):
        name = node if isinstance(node, str) else node.get("name", str(node))
        angle = 2 * np.pi * i / n - np.pi / 2
        r = 0.32
        x = 0.5 + r * np.cos(angle)
        y = 0.5 + r * np.sin(angle)

        c = colors[i + 1]
        circle = plt.Circle((x, y), 0.06, color=c, zorder=5)
        ax.add_patch(circle)
        ax.text(x, y, str(name)[:8], ha="center", va="center",
                fontsize=10, fontweight="bold", color="white", zorder=6)

        # 连线
        ax.plot([0.5, x], [0.5, y], color="#CCCCCC", linewidth=1.5, zorder=3)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 矩阵图 ============
def _render_matrix(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """2xN 或 NxM 矩阵图，卡片式布局。"""
    items = data.get("items", data.get("cells", data.get("nodes", [])))
    if not items:
        parts = description.replace("\u3001", "|").replace("\uff1b", "|").replace(";", "|").split("|")
        items = [{"name": p.strip()} for p in parts if p.strip()]
    if not items:
        items = [{"name": "项目"}]

    n = len(items)
    cols = min(3, n)
    rows = (n + cols - 1) // cols

    colors = _get_palette(color_scheme, n)
    fig, ax = plt.subplots(figsize=(cols * 5, rows * 3.5))

    card_w = 0.85 / cols
    card_h = 0.8 / rows
    pad = 0.03

    for i, item in enumerate(items):
        row = i // cols
        col = i % cols
        x = 0.08 + col * (card_w + pad)
        y = 0.88 - (row + 1) * (card_h + pad)

        name = item if isinstance(item, str) else item.get("name", str(item))
        detail = item.get("detail", item.get("description", "")) if isinstance(item, dict) else ""

        rect = FancyBboxPatch((x, y), card_w, card_h,
                              boxstyle="round,pad=0.015", facecolor=_lighten(colors[i], 0.6),
                              edgecolor=colors[i], linewidth=2)
        ax.add_patch(rect)

        # 顶部色条
        bar = FancyBboxPatch((x, y + card_h - 0.03), card_w, 0.03,
                             boxstyle="round,pad=0.005", facecolor=colors[i],
                             edgecolor="none")
        ax.add_patch(bar)

        ax.text(x + card_w / 2, y + card_h * 0.6, str(name)[:15],
                ha="center", va="center", fontsize=12, fontweight="bold",
                color=colors[i])
        if detail:
            ax.text(x + card_w / 2, y + card_h * 0.3, str(detail)[:25],
                    ha="center", va="center", fontsize=9, color="#666666")

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 对比图 ============
def _render_comparison(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """左右对比图，带指标条形。"""
    items = data.get("items", [])
    if not items or len(items) < 2:
        return _render_placeholder(data, description, color_scheme)

    primary = color_scheme.get("primary", "#1B365D")
    accent = color_scheme.get("accent", "#E8612D")

    fig, ax = plt.subplots(figsize=(14, 8))

    left = items[0]
    right = items[1]
    left_name = left.get("name", "A")
    right_name = right.get("name", "B")

    # 标题
    ax.text(0.25, 0.92, left_name, ha="center", va="center",
            fontsize=22, fontweight="bold", color=primary)
    ax.text(0.75, 0.92, right_name, ha="center", va="center",
            fontsize=22, fontweight="bold", color=accent)

    # VS
    circle = plt.Circle((0.5, 0.92), 0.035, color="#F0F0F0", zorder=5)
    ax.add_patch(circle)
    ax.text(0.5, 0.92, "VS", ha="center", va="center",
            fontsize=12, fontweight="bold", color="#999999", zorder=6)

    # 对比指标
    left_metrics = left.get("metrics", left.get("features", left.get("bullets", [])))
    right_metrics = right.get("metrics", right.get("features", right.get("bullets", [])))

    if isinstance(left_metrics, dict):
        keys = list(left_metrics.keys())
        left_vals = list(left_metrics.values())
        right_vals = [right_metrics.get(k, 0) for k in keys] if isinstance(right_metrics, dict) else []
    elif isinstance(left_metrics, list):
        keys = [str(m) for m in left_metrics]
        left_vals = list(range(len(keys), 0, -1))
        right_vals = list(range(1, len(keys) + 1))
    else:
        keys = ["指标"]
        left_vals = [1]
        right_vals = [1]

    n_metrics = min(len(keys), 6)
    for i in range(n_metrics):
        y = 0.78 - i * 0.12
        ax.text(0.5, y, keys[i][:12], ha="center", va="center", fontsize=12, color="#888888")

        # 左侧条
        if i < len(left_vals):
            try:
                lv = float(left_vals[i])
                max_v = max(float(left_vals[i]), float(right_vals[i]) if i < len(right_vals) else 1) or 1
                lw = 0.2 * (lv / max_v)
            except (ValueError, TypeError):
                lw = 0.15
            bar_l = FancyBboxPatch((0.42 - lw, y - 0.02), lw, 0.04,
                                   boxstyle="round,pad=0.005", facecolor=primary, alpha=0.8)
            ax.add_patch(bar_l)

        # 右侧条
        if i < len(right_vals):
            try:
                rv = float(right_vals[i])
                rw = 0.2 * (rv / max_v)
            except (ValueError, TypeError):
                rw = 0.15
            bar_r = FancyBboxPatch((0.58, y - 0.02), rw, 0.04,
                                   boxstyle="round,pad=0.005", facecolor=accent, alpha=0.8)
            ax.add_patch(bar_r)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 金字塔图 ============
def _render_pyramid(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """金字塔层级图。"""
    levels = data.get("levels", data.get("items", data.get("stages", [])))
    if not levels:
        parts = description.replace("\u2192", "|").replace("\u3001", "|").split("|")
        levels = [p.strip() for p in parts if p.strip()]
    if not levels:
        levels = ["顶层"]

    n = len(levels)
    colors = _get_palette(color_scheme, n)
    fig, ax = plt.subplots(figsize=(12, 8))

    for i, level in enumerate(levels):
        name = level if isinstance(level, str) else level.get("name", str(level))
        # 梯形
        x_left_bottom = 0.5 - (0.45 * (n - i) / n)
        x_right_bottom = 0.5 + (0.45 * (n - i) / n)
        x_left_top = 0.5 - (0.45 * (n - i - 1) / n)
        x_right_top = 0.5 + (0.45 * (n - i - 1) / n)

        # 翻转：第 0 层在底部
        y_b = 0.1 + (n - 1 - i) * 0.8 / n
        y_t = 0.1 + (n - i) * 0.8 / n

        polygon = plt.Polygon(
            [(x_left_bottom, y_b), (x_right_bottom, y_b),
             (x_right_top, y_t), (x_left_top, y_t)],
            facecolor=colors[i], edgecolor="white", linewidth=2
        )
        ax.add_patch(polygon)
        ax.text(0.5, (y_b + y_t) / 2, str(name)[:20], ha="center", va="center",
                fontsize=13, fontweight="bold", color="white")

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 循环图 ============
def _render_cycle(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """环形循环图。"""
    stages = data.get("stages", data.get("items", []))
    if not stages:
        parts = description.replace("\u2192", "|").replace("->", "|").split("|")
        stages = [{"name": p.strip()} for p in parts if p.strip()]
    if not stages:
        stages = [{"name": "阶段"}]

    n = len(stages)
    colors = _get_palette(color_scheme, n)
    fig, ax = plt.subplots(figsize=(10, 10))

    r = 0.3
    for i, stage in enumerate(stages):
        name = stage if isinstance(stage, str) else stage.get("name", "")
        angle = 2 * np.pi * i / n - np.pi / 2
        x = 0.5 + r * np.cos(angle)
        y = 0.5 + r * np.sin(angle)

        circle = plt.Circle((x, y), 0.08, color=colors[i], zorder=5)
        ax.add_patch(circle)
        ax.text(x, y, str(name)[:8], ha="center", va="center",
                fontsize=11, fontweight="bold", color="white", zorder=6)

        # 弧形箭头到下一个节点
        if n > 1:
            next_angle = 2 * np.pi * ((i + 1) % n) / n - np.pi / 2
            ax.annotate("",
                        xy=(0.5 + (r - 0.09) * np.cos(next_angle),
                            0.5 + (r - 0.09) * np.sin(next_angle)),
                        xytext=(0.5 + (r + 0.09) * np.cos(angle) * 0.95,
                                0.5 + (r + 0.09) * np.sin(angle) * 0.95),
                        arrowprops=dict(arrowstyle="-|>", color="#CCCCCC", lw=2,
                                        connectionstyle="arc3,rad=0.3"),
                        zorder=3)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 占位符 ============
def _render_placeholder(data: dict, description: str, color_scheme: dict) -> plt.Figure:
    """通用占位信息图：卡片式展示描述文字。"""
    primary = color_scheme.get("primary", "#1B365D")
    fig, ax = plt.subplots(figsize=(14, 8))

    # 背景卡片
    rect = FancyBboxPatch(
        (0.03, 0.03), 0.94, 0.94,
        boxstyle="round,pad=0.03", facecolor=_lighten(primary, 0.85),
        edgecolor=primary, linewidth=2,
    )
    ax.add_patch(rect)

    # 顶部装饰条
    bar = FancyBboxPatch((0.03, 0.9), 0.94, 0.07,
                         boxstyle="round,pad=0.01", facecolor=primary, edgecolor="none")
    ax.add_patch(bar)

    # 描述文字（分行显示）
    text = description or "信息图"
    lines = []
    while text:
        lines.append(text[:35])
        text = text[35:]

    for i, line in enumerate(lines[:8]):
        ax.text(0.5, 0.75 - i * 0.08, line, ha="center", va="center",
                fontsize=14, color=primary, transform=ax.transAxes)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.tight_layout()
    return fig


# ============ 渲染器注册表 ============
_INFOGRAPHIC_RENDERERS = {
    "process_flow": _render_process_flow,
    "stat_display": _render_stat_display,
    "timeline": _render_timeline,
    "hierarchy": _render_hierarchy,
    "comparison": _render_comparison,
    "cycle": _render_cycle,
    "matrix": _render_matrix,
    "network": _render_network,
    "pyramid": _render_pyramid,
    "venn": _render_placeholder,
}

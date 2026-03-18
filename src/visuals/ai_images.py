# -*- coding: utf-8 -*-
"""
AI 图片生成器 — Google Gemini 图片生成版本。
使用 Gemini API 生成真实 AI 图片，失败时回退到 matplotlib 渐变背景。
"""
import base64
import logging
import time
from pathlib import Path

import httpx

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

import src  # noqa: F401
from config import GEMINI_API_KEY

# --------------- 日志配置 ---------------
logger = logging.getLogger(__name__)
if not logger.handlers:
    _handler = logging.StreamHandler()
    _handler.setFormatter(
        logging.Formatter("[%(asctime)s] %(levelname)s %(name)s: %(message)s",
                          datefmt="%Y-%m-%d %H:%M:%S")
    )
    logger.addHandler(_handler)
    logger.setLevel(logging.INFO)

# --------------- 常量 ---------------
GEMINI_IMAGE_MODEL = "gemini-3-pro-image-preview"
GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_IMAGE_MODEL}:generateContent"
)
MAX_RETRIES = 2
REQUEST_TIMEOUT = 120  # 秒


# --------------- 提示词增强 ---------------
def _enhance_prompt(prompt: str) -> str:
    """根据 prompt 内容追加专业风格描述，提升生成质量。"""
    prompt_lower = prompt.lower()

    if any(k in prompt_lower for k in ["military", "defense", "weapon", "army", "navy", "战", "军", "国防"]):
        suffix = (
            "Style: dramatic cinematic military photograph, deep navy blue and steel gray tones, "
            "high contrast lighting, sharp details, photorealistic, 8K resolution, widescreen 16:9 aspect ratio."
        )
    elif any(k in prompt_lower for k in ["tech", "digital", "cyber", "data", "network", "科技", "数据"]):
        suffix = (
            "Style: futuristic technology visualization, dark background with glowing blue and cyan accents, "
            "clean modern aesthetic, abstract digital elements, photorealistic, 8K resolution, widescreen 16:9."
        )
    elif any(k in prompt_lower for k in ["thank", "end", "谢", "结束", "结语"]):
        suffix = (
            "Style: elegant minimalist closing slide background, soft warm gradient, "
            "subtle abstract shapes, calming atmosphere, professional, 8K resolution, widescreen 16:9."
        )
    elif any(k in prompt_lower for k in ["medical", "health", "hospital", "医", "健康"]):
        suffix = (
            "Style: clean professional medical illustration, soft blue and white tones, "
            "modern healthcare aesthetic, photorealistic, 8K resolution, widescreen 16:9."
        )
    elif any(k in prompt_lower for k in ["finance", "business", "market", "金融", "商业", "商务"]):
        suffix = (
            "Style: professional corporate business visual, sophisticated blue and gold palette, "
            "modern clean design, photorealistic, 8K resolution, widescreen 16:9."
        )
    else:
        suffix = (
            "Style: professional presentation background, clean modern design, subtle gradient, "
            "sophisticated color palette, photorealistic, 8K resolution, widescreen 16:9 aspect ratio."
        )

    return f"{prompt}. {suffix}"


# --------------- Gemini API 调用 ---------------
def _call_gemini(prompt: str, output_path: Path) -> bool:
    """调用 Gemini 图片生成 API，成功返回 True 并保存 PNG 到 output_path。"""
    if not GEMINI_API_KEY:
        logger.warning("GEMINI_API_KEY 未配置，跳过 AI 图片生成")
        return False

    enhanced_prompt = _enhance_prompt(prompt)
    payload = {
        "contents": [{"parts": [{"text": enhanced_prompt}]}],
        "generationConfig": {"responseModalities": ["TEXT", "IMAGE"]},
    }
    url = f"{GEMINI_API_URL}?key={GEMINI_API_KEY}"

    for attempt in range(1, MAX_RETRIES + 1):
        logger.info("Gemini 图片生成 第 %d/%d 次尝试 | prompt: %.80s...", attempt, MAX_RETRIES, prompt)
        try:
            with httpx.Client(timeout=REQUEST_TIMEOUT) as client:
                resp = client.post(url, json=payload)

            if resp.status_code != 200:
                logger.warning(
                    "Gemini API 返回 HTTP %d: %s", resp.status_code, resp.text[:300]
                )
                if attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)
                continue

            data = resp.json()
            # 从 candidates 中提取 inlineData 图片
            candidates = data.get("candidates", [])
            for candidate in candidates:
                parts = candidate.get("content", {}).get("parts", [])
                for part in parts:
                    inline_data = part.get("inlineData")
                    if inline_data and "data" in inline_data:
                        image_bytes = base64.b64decode(inline_data["data"])
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                        output_path.write_bytes(image_bytes)
                        logger.info(
                            "Gemini 图片已保存: %s (%.1f KB)",
                            output_path,
                            len(image_bytes) / 1024,
                        )
                        return True

            logger.warning("Gemini 响应中未找到 inlineData 图片数据 (第 %d 次)", attempt)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

        except httpx.TimeoutException:
            logger.warning("Gemini API 请求超时 (第 %d 次, timeout=%ds)", attempt, REQUEST_TIMEOUT)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)
        except Exception as exc:
            logger.error("Gemini API 调用异常 (第 %d 次): %s", attempt, exc, exc_info=True)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

    return False


# --------------- matplotlib 回退 ---------------
def _fallback_gradient(prompt: str, output_path: Path):
    """生成 matplotlib 渐变背景作为兜底方案。"""
    logger.info("使用 matplotlib 渐变背景回退: %s", output_path)
    prompt_lower = prompt.lower()

    fig, ax = plt.subplots(figsize=(19.2, 10.8))

    # 根据主题选择配色
    if any(k in prompt_lower for k in ["military", "defense", "army", "战", "军"]):
        cmap, bg_color, accent = "Blues_r", "#0A1628", "#4A90D9"
    elif any(k in prompt_lower for k in ["tech", "digital", "cyber", "科技"]):
        cmap, bg_color, accent = "cool", "#0D1117", "#58A6FF"
    elif any(k in prompt_lower for k in ["thank", "end", "谢", "结束"]):
        cmap, bg_color, accent = "RdYlBu_r", "#1A1A2E", "#E8612D"
    else:
        cmap, bg_color, accent = "Blues", "#F0F4F8", "#1B365D"

    ax.set_facecolor(bg_color)

    # 径向渐变
    y = np.linspace(0, 1, 500)
    x = np.linspace(0, 1, 500)
    X, Y = np.meshgrid(x, y)
    Z = 1.0 - 0.7 * np.sqrt((X - 0.5) ** 2 + (Y - 0.5) ** 2)
    ax.imshow(Z, cmap=cmap, aspect="auto", extent=[0, 1, 0, 1], alpha=0.35)

    # 装饰圆
    for cx, cy, r, a in [(0.8, 0.3, 0.2, 0.07), (0.2, 0.7, 0.15, 0.06), (0.5, 0.5, 0.3, 0.04)]:
        ax.add_patch(plt.Circle((cx, cy), r, color=accent, alpha=a))

    # 细网格
    for i in range(25):
        ax.axhline(y=i / 25, color=accent, alpha=0.04, linewidth=0.5)
        ax.axvline(x=i / 25, color=accent, alpha=0.04, linewidth=0.5)

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    fig.savefig(str(output_path), dpi=150, bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    logger.info("matplotlib 回退图片已保存: %s", output_path)


# --------------- 主入口 ---------------
def generate_image(visual: dict, output_path: Path):
    """生成 AI 图片。优先使用 Gemini API，失败后回退 matplotlib。"""
    prompt = visual.get("prompt", "professional presentation background")
    logger.info("开始生成图片 | prompt: %.100s | output: %s", prompt, output_path)

    start = time.time()

    # 尝试 Gemini API
    success = _call_gemini(prompt, output_path)

    if not success:
        logger.info("Gemini 生成失败，回退至 matplotlib")
        _fallback_gradient(prompt, output_path)

    elapsed = time.time() - start
    logger.info("图片生成完成 | 耗时 %.1f 秒 | 方式: %s | %s",
                elapsed, "Gemini AI" if success else "matplotlib 回退", output_path)

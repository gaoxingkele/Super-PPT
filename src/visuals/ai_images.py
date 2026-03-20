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
from config import (GEMINI_API_KEY, CLOUBIC_ENABLED, CLOUBIC_API_KEY, CLOUBIC_BASE_URL, CLOUBIC_IMAGE_MODEL,
                    DOUBAO_API_KEY, DOUBAO_BASE_URL, DOUBAO_IMAGE_MODEL)

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
GEMINI_IMAGE_MODEL = "gemini-3.1-flash-image-preview"
GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_IMAGE_MODEL}:generateContent"
)
MAX_RETRIES = 3
REQUEST_TIMEOUT = 180  # 秒（Cloubic 图片生成可能较慢）


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

    # 全局要求：图中所有文字必须使用中文
    chinese_req = "IMPORTANT: All text, labels, captions, and annotations in the image MUST be in Simplified Chinese (简体中文). Do NOT use any English text."
    return f"{prompt}. {chinese_req} {suffix}"


# --------------- Cloubic 响应图片提取 ---------------
import re as _re

def _extract_image_from_cloubic_response(data: dict) -> bytes:
    """从 Cloubic chat completions 响应中提取 base64 图片。
    支持多种返回格式：
    1. 字符串含 markdown: ![image](data:image/png;base64,...)
    2. 字符串含裸 data URI: data:image/png;base64,...
    3. list 结构含 image_url 类型
    """
    for choice in data.get("choices", []):
        content = choice.get("message", {}).get("content", "")

        # 格式1/2: 字符串中包含 base64 图片
        if isinstance(content, str) and "base64," in content:
            match = _re.search(r"data:image/[a-z]+;base64,([A-Za-z0-9+/=\s]+)", content)
            if match:
                b64_data = match.group(1).replace("\n", "").replace(" ", "")
                try:
                    return base64.b64decode(b64_data)
                except Exception:
                    pass

        # 格式3: list 结构
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


# --------------- 豆包 Seedream 图片生成 ---------------
def _call_doubao_image(prompt: str, output_path: Path) -> bool:
    """调用豆包 Seedream 图片生成 API（火山引擎直连）。"""
    if not DOUBAO_API_KEY or not DOUBAO_IMAGE_MODEL:
        return False

    enhanced_prompt = _enhance_prompt(prompt)
    url = f"{DOUBAO_BASE_URL}/images/generations"
    headers = {"Authorization": f"Bearer {DOUBAO_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": DOUBAO_IMAGE_MODEL,
        "prompt": enhanced_prompt,
        "response_format": "b64_json",
        "size": "2560x1440",  # 16:9 适合 PPT
        "seed": -1,
        "watermark": False,
    }

    for attempt in range(1, MAX_RETRIES + 1):
        logger.info("豆包 Seedream 图片生成 第 %d/%d 次 | model: %s", attempt, MAX_RETRIES, DOUBAO_IMAGE_MODEL)
        try:
            with httpx.Client(timeout=REQUEST_TIMEOUT) as client:
                resp = client.post(url, json=payload, headers=headers)

            if resp.status_code != 200:
                logger.warning("豆包 API 返回 HTTP %d: %s", resp.status_code, resp.text[:300])
                if attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)
                continue

            data = resp.json()
            img_list = data.get("data", [])
            if img_list:
                b64_data = img_list[0].get("b64_json", "")
                if b64_data:
                    image_bytes = base64.b64decode(b64_data)
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    output_path.write_bytes(image_bytes)
                    logger.info("豆包图片已保存: %s (%.1f KB)", output_path, len(image_bytes) / 1024)
                    return True

            logger.warning("豆包响应中未找到图片数据 (第 %d 次)", attempt)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

        except httpx.TimeoutException:
            logger.warning("豆包 API 超时 (第 %d 次)", attempt)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)
        except Exception as exc:
            logger.error("豆包 API 异常 (第 %d 次): %s", attempt, exc)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

    return False


# --------------- Cloubic OpenAI 兼容图片生成 ---------------
def _call_cloubic_image(prompt: str, output_path: Path) -> bool:
    """通过 Cloubic OpenAI 兼容接口调用 Gemini 图片生成。"""
    import config as _cfg
    if not _cfg.CLOUBIC_ENABLED or not _cfg.CLOUBIC_API_KEY:
        return False

    enhanced_prompt = _enhance_prompt(prompt)
    # Cloubic 支持 OpenAI chat completions 格式调用图片模型
    image_model = _cfg.CLOUBIC_IMAGE_MODEL or GEMINI_IMAGE_MODEL
    payload = {
        "model": image_model,
        "messages": [{"role": "user", "content": enhanced_prompt}],
        "max_tokens": 4096,
    }
    url = f"{_cfg.CLOUBIC_BASE_URL}/chat/completions"
    headers = {"Authorization": f"Bearer {_cfg.CLOUBIC_API_KEY}", "Content-Type": "application/json"}

    for attempt in range(1, MAX_RETRIES + 1):
        logger.info("Cloubic 图片生成 第 %d/%d 次尝试 | model: %s", attempt, MAX_RETRIES, image_model)
        try:
            with httpx.Client(timeout=REQUEST_TIMEOUT, proxy=None) as client:
                resp = client.post(url, json=payload, headers=headers)

            if resp.status_code != 200:
                logger.warning("Cloubic API 返回 HTTP %d: %s", resp.status_code, resp.text[:300])
                if attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)
                continue

            data = resp.json()
            # 从 choices 中提取 base64 图片
            img_bytes = _extract_image_from_cloubic_response(data)
            if img_bytes:
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(img_bytes)
                logger.info("Cloubic 图片已保存: %s (%.1f KB)", output_path, len(img_bytes) / 1024)
                return True

            logger.warning("Cloubic 响应中未找到图片数据 (第 %d 次)", attempt)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

        except httpx.TimeoutException:
            logger.warning("Cloubic API 请求超时 (第 %d 次)", attempt)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)
        except Exception as exc:
            logger.error("Cloubic API 调用异常 (第 %d 次): %s", attempt, exc)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

    return False


# --------------- Cloubic /images/generations 调用（wan2.6-t2i 等） ---------------
def _call_cloubic_image_gen(prompt: str, output_path: Path, model: str) -> bool:
    """通过 Cloubic /images/generations 端点调用按次计费的图片模型。"""
    import config as _cfg
    if not _cfg.CLOUBIC_ENABLED or not _cfg.CLOUBIC_API_KEY:
        return False

    enhanced_prompt = _enhance_prompt(prompt)
    url = f"{_cfg.CLOUBIC_BASE_URL}/images/generations"
    headers = {"Authorization": f"Bearer {_cfg.CLOUBIC_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": model,
        "prompt": enhanced_prompt,
        "response_format": "b64_json",
        "size": "1024x1024",
    }

    for attempt in range(1, MAX_RETRIES + 1):
        logger.info("Cloubic 图片生成 第 %d/%d 次 | model: %s", attempt, MAX_RETRIES, model)
        try:
            with httpx.Client(timeout=REQUEST_TIMEOUT, proxy=None) as client:
                resp = client.post(url, json=payload, headers=headers)

            if resp.status_code != 200:
                logger.warning("Cloubic 图片API 返回 HTTP %d: %s", resp.status_code, resp.text[:300])
                if attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)
                continue

            data = resp.json()
            img_list = data.get("data", [])
            if img_list:
                # 优先 b64_json，其次 url
                b64 = img_list[0].get("b64_json", "")
                if b64:
                    image_bytes = base64.b64decode(b64)
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    output_path.write_bytes(image_bytes)
                    logger.info("Cloubic 图片已保存: %s (%.1f KB) | model: %s", output_path, len(image_bytes) / 1024, model)
                    return True
                img_url = img_list[0].get("url", "")
                if img_url:
                    img_resp = httpx.get(img_url, timeout=60)
                    if img_resp.status_code == 200:
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                        output_path.write_bytes(img_resp.content)
                        logger.info("Cloubic 图片已保存(url): %s (%.1f KB)", output_path, len(img_resp.content) / 1024)
                        return True

            logger.warning("Cloubic 图片响应无数据 (第 %d 次) | model: %s", attempt, model)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

        except Exception as exc:
            logger.warning("Cloubic 图片异常 (第 %d 次): %s", attempt, exc)
            if attempt < MAX_RETRIES:
                time.sleep(2 * attempt)

    return False


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
    """生成 AI 图片。优先级：Cloubic seedream → Cloubic wan2.6 → Cloubic qwen → 豆包直连 → matplotlib。"""
    import config as _cfg
    prompt = visual.get("prompt", "professional presentation background")
    logger.info("开始生成图片 | prompt: %.100s | output: %s", prompt, output_path)

    start = time.time()
    success = False
    method = "matplotlib 回退"

    # 1. Cloubic doubao-seedream-5-0（chat completions 格式）
    if not success and _cfg.CLOUBIC_ENABLED and _cfg.CLOUBIC_API_KEY:
        success = _call_cloubic_image(prompt, output_path)
        if success:
            method = f"Cloubic ({_cfg.CLOUBIC_IMAGE_MODEL})"

    # 2. 备选: Cloubic wan2.6-t2i（¥0.16/次）
    if not success and _cfg.CLOUBIC_ENABLED and _cfg.CLOUBIC_API_KEY:
        success = _call_cloubic_image_gen(prompt, output_path, "wan2.6-t2i")
        if success:
            method = "Cloubic (wan2.6-t2i)"

    # 3. 备选: Cloubic qwen-image-edit-plus（¥0.16/次）
    if not success and _cfg.CLOUBIC_ENABLED and _cfg.CLOUBIC_API_KEY:
        success = _call_cloubic_image_gen(prompt, output_path, "qwen-image-edit-plus")
        if success:
            method = "Cloubic (qwen-image-edit-plus)"

    # 4. 豆包 Seedream 直连
    if not success and DOUBAO_API_KEY and DOUBAO_IMAGE_MODEL:
        success = _call_doubao_image(prompt, output_path)
        if success:
            method = f"豆包直连 ({DOUBAO_IMAGE_MODEL})"

    # 5. matplotlib 回退
    if not success:
        logger.info("图片生成全部失败，回退至 matplotlib")
        _fallback_gradient(prompt, output_path)

    elapsed = time.time() - start
    logger.info("图片生成完成 | 耗时 %.1f 秒 | 方式: %s | %s", elapsed, method, output_path)

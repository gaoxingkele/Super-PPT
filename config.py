# -*- coding: utf-8 -*-
"""Super-PPT 项目配置：多模型 API Key、路径、内容限制。"""
import os
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()  # 加载 .env
load_dotenv(Path(__file__).resolve().parent / ".env.cloubic", override=False)  # 加载 .env.cloubic（不覆盖 .env）


def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default).strip()


# ============ 默认 Provider ============
LLM_PROVIDER = _env("LLM_PROVIDER", "grok")

# ============ Kimi（月之暗面） ============
KIMI_API_KEY = _env("KIMI_API_KEY")
KIMI_BASE_URL = "https://api.moonshot.cn/v1"
KIMI_MODEL = _env("KIMI_MODEL", "kimi-k2.5")
KIMI_VISION_MODEL = _env("KIMI_VISION_MODEL", "moonshot-v1-32k-vision-preview")

# ============ Google（Gemini） ============
GEMINI_API_KEY = _env("GEMINI_API_KEY")
GEMINI_MODEL = _env("GEMINI_MODEL", "gemini-3.1-pro-preview")

# ============ xAI（Grok） ============
GROK_API_KEY = _env("GROK_API_KEY")
GROK_BASE_URL = "https://api.x.ai/v1"
GROK_MODEL = _env("GROK_MODEL", "grok-4.2-0718")

# ============ MiniMax ============
MINIMAX_API_KEY = _env("MINIMAX_API_KEY")
MINIMAX_BASE_URL = _env("MINIMAX_BASE_URL") or "https://api.minimax.chat/v1"
MINIMAX_MODEL = _env("MINIMAX_MODEL", "MiniMax-M2.5")

# ============ GLM（智谱清言） ============
GLM_API_KEY = _env("GLM_API_KEY")
_raw_glm_url = _env("GLM_BASE_URL") or "https://open.bigmodel.cn/api/paas/v4"
GLM_BASE_URL = _raw_glm_url.removesuffix("/chat/completions")
GLM_MODEL = _env("GLM_MODEL", "glm-4.7")
GLM_VISION_MODEL = _env("GLM_VISION_MODEL", "glm-4v-plus")

# ============ Qwen（通义千问） ============
QWEN_API_KEY = _env("QWEN_API_KEY") or _env("DASHSCOPE_API_KEY")
QWEN_BASE_URL = _env("QWEN_BASE_URL") or "https://dashscope.aliyuncs.com/compatible-mode/v1"
QWEN_MODEL = _env("QWEN_MODEL", "qwen3.5-plus")

# ============ DeepSeek ============
DEEPSEEK_API_KEY = _env("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL = _env("DEEPSEEK_BASE_URL") or "https://api.deepseek.com"
DEEPSEEK_MODEL = _env("DEEPSEEK_MODEL", "deepseek-chat")

# ============ OpenAI（ChatGPT） ============
OPENAI_API_KEY = _env("OPENAI_API_KEY")
OPENAI_BASE_URL = _env("OPENAI_BASE_URL") or "https://api.openai.com/v1"
OPENAI_MODEL = _env("OPENAI_MODEL", "gpt-5.4")

# ============ Perplexity ============
PERPLEXITY_API_KEY = _env("PERPLEXITY_API_KEY")
PERPLEXITY_BASE_URL = "https://api.perplexity.ai"
PERPLEXITY_MODEL = _env("PERPLEXITY_MODEL", "sonar")

# ============ Anthropic（Claude） ============
ANTHROPIC_API_KEY = _env("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = _env("ANTHROPIC_MODEL", "claude-sonnet-4-6")

# ============ LLM 候选列表（解析失败时依次尝试） ============
# 格式: provider1,provider2,...  留空则不重试
REVIEW_FALLBACK_PROVIDERS = [
    p.strip() for p in _env("REVIEW_FALLBACK_PROVIDERS", "gemini,qwen,deepseek").split(",") if p.strip()
]

# ============ PPT 生成语言 ============
PPT_LANGUAGE = _env("PPT_LANGUAGE", "zh")

# ============ 项目路径 ============
PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"
THEMES_DIR = PROJECT_ROOT / "themes"

# ============ 爬虫配置 ============
MIN_CONTENT_BYTES = 1000
RETRY_WAIT_SECONDS = 15
CRAWL_MAX_RETRIES = 5

# ============ 内容截断限制 ============
# Step0 内容获取
RAW_LOAD_LIMIT = 130_000

# Step1 结构化分析
ANALYZE_CONTENT_LIMIT = 80_000

# Step2 大纲生成
OUTLINE_CONTENT_LIMIT = 60_000

# Step3 视觉资产
VISUAL_DESCRIPTION_LIMIT = 2_000

# ============ PPT 配置 ============
DEFAULT_THEME = "business"
DEFAULT_SLIDE_RANGE = None  # None = LLM 根据内容自行决策页数
MAX_BULLETS_PER_SLIDE = 5
MAX_CHARS_PER_BULLET = 25
CHART_DPI = 300
CHART_FIGSIZE = (19.2, 10.8)  # 1920x1080 at 100 DPI

# ============ 页数推算 ============
CHARS_PER_SLIDE = 1500           # 每页对应的原文字数基线
DATA_POINT_BONUS = 0.8           # 每个数据点额外加的页数
MAX_PAGES_PER_CHAPTER = 8        # 单章最大内容页数

# ============ 支持的文件扩展名 ============
TEXT_EXTENSIONS = {".txt", ".md", ".json", ".html"}
DOCUMENT_EXTENSIONS = {".docx", ".pdf"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}
TEMPLATE_EXTENSIONS = {".pptx", ".pdf"} | IMAGE_EXTENSIONS
INPUT_EXTENSIONS = TEXT_EXTENSIONS | DOCUMENT_EXTENSIONS

# ============ Cloubic 统一路由 ============
CLOUBIC_ENABLED = _env("CLOUBIC_ENABLED", "false").lower() in ("true", "1", "yes")
CLOUBIC_API_KEY = _env("CLOUBIC_API_KEY")
CLOUBIC_BASE_URL = _env("CLOUBIC_BASE_URL") or "https://api.cloubic.com/v1"
CLOUBIC_DEFAULT_PROVIDER = _env("CLOUBIC_DEFAULT_PROVIDER") or "deepseek"

# Cloubic 模型映射：provider -> 通过 Cloubic 调用时的模型 ID（2026-03-19 API 实测）
CLOUBIC_MODEL_MAP = {
    "openai": _env("CLOUBIC_OPENAI_MODEL") or "gpt-5.4",
    "claude": _env("CLOUBIC_CLAUDE_MODEL") or "claude-opus-4-6",
    "gemini": _env("CLOUBIC_GEMINI_MODEL") or "gemini-3.1-pro-preview",
    "deepseek": _env("CLOUBIC_DEEPSEEK_MODEL") or "deepseek-v3.2",
    "grok": _env("CLOUBIC_GROK_MODEL") or "grok-4-1-fast-non-reasoning",
    "qwen": _env("CLOUBIC_QWEN_MODEL") or "qwen3-max",
    "doubao": _env("CLOUBIC_DOUBAO_MODEL") or "doubao-seed-1-6-flash-250828",
    "minimax": _env("CLOUBIC_MINIMAX_MODEL") or "MiniMax-Hailuo-2.3",
    "kimi": _env("CLOUBIC_KIMI_MODEL") or "deepseek-v3.2",
    "glm": _env("CLOUBIC_GLM_MODEL") or "deepseek-v3.2",
    "perplexity": _env("CLOUBIC_PERPLEXITY_MODEL") or "deepseek-v3.2",
}

# Cloubic 推理模型映射：provider -> 推理/思考版模型 ID
CLOUBIC_REASONING_MODEL_MAP = {
    "openai": _env("CLOUBIC_OPENAI_REASONING_MODEL") or "o4-mini-2025-04-16",
    "claude": _env("CLOUBIC_CLAUDE_REASONING_MODEL") or "claude-sonnet-4-5-20250929-thinking",
    "gemini": _env("CLOUBIC_GEMINI_REASONING_MODEL") or "gemini-2.5-pro",  # 自带思考能力
    "deepseek": _env("CLOUBIC_DEEPSEEK_REASONING_MODEL") or "deepSeek-R1-0528",
    "grok": _env("CLOUBIC_GROK_REASONING_MODEL") or "grok-4-1-fast-reasoning",
    "qwen": _env("CLOUBIC_QWEN_REASONING_MODEL") or "qwen3-max",
    "doubao": _env("CLOUBIC_DOUBAO_REASONING_MODEL") or "doubao-seed-1-6-251015",
    "minimax": _env("CLOUBIC_MINIMAX_REASONING_MODEL") or "MiniMax-Hailuo-2.3",
}

# ============ 自动创建目录 ============
for d in (OUTPUT_DIR, THEMES_DIR):
    d.mkdir(parents=True, exist_ok=True)

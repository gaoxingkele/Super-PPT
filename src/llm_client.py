# -*- coding: utf-8 -*-
"""
统一 LLM 客户端：支持 Kimi、Gemini、Grok、MiniMax、GLM、Qwen、DeepSeek、OpenAI、Perplexity、Claude。
通过 LLM_PROVIDER 环境变量或 provider 参数切换。
"""
import os

import src  # noqa: F401  — 确保 PROJECT_ROOT 加入 sys.path

import time as _time

import httpx
from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception

from config import (
    LLM_PROVIDER,
    KIMI_API_KEY, KIMI_BASE_URL, KIMI_MODEL, KIMI_VISION_MODEL,
    GEMINI_API_KEY, GEMINI_MODEL,
    GROK_API_KEY, GROK_BASE_URL, GROK_MODEL,
    MINIMAX_API_KEY, MINIMAX_BASE_URL, MINIMAX_MODEL,
    GLM_API_KEY, GLM_BASE_URL, GLM_MODEL, GLM_VISION_MODEL,
    QWEN_API_KEY, QWEN_BASE_URL, QWEN_MODEL,
    DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, DEEPSEEK_MODEL,
    OPENAI_API_KEY, OPENAI_BASE_URL, OPENAI_MODEL,
    PERPLEXITY_API_KEY, PERPLEXITY_BASE_URL, PERPLEXITY_MODEL,
    ANTHROPIC_API_KEY, ANTHROPIC_MODEL,
    CLOUBIC_ENABLED, CLOUBIC_API_KEY, CLOUBIC_BASE_URL,
    CLOUBIC_DEFAULT_PROVIDER, CLOUBIC_MODEL_MAP, CLOUBIC_REASONING_MODEL_MAP,
)

HTTP_TIMEOUT = httpx.Timeout(60.0, read=600.0)

# 所有 OpenAI 兼容 provider 的配置
PROVIDER_CONFIG = {
    "kimi": {"key": KIMI_API_KEY, "base_url": KIMI_BASE_URL, "model": KIMI_MODEL},
    "grok": {"key": GROK_API_KEY, "base_url": GROK_BASE_URL, "model": GROK_MODEL},
    "minimax": {"key": MINIMAX_API_KEY, "base_url": MINIMAX_BASE_URL, "model": MINIMAX_MODEL},
    "glm": {"key": GLM_API_KEY, "base_url": GLM_BASE_URL, "model": GLM_MODEL},
    "qwen": {"key": QWEN_API_KEY, "base_url": QWEN_BASE_URL, "model": QWEN_MODEL},
    "deepseek": {"key": DEEPSEEK_API_KEY, "base_url": DEEPSEEK_BASE_URL, "model": DEEPSEEK_MODEL},
    "openai": {"key": OPENAI_API_KEY, "base_url": OPENAI_BASE_URL, "model": OPENAI_MODEL},
    "perplexity": {"key": PERPLEXITY_API_KEY, "base_url": PERPLEXITY_BASE_URL, "model": PERPLEXITY_MODEL},
    "claude": {"key": ANTHROPIC_API_KEY, "model": ANTHROPIC_MODEL},
    "gemini": {"key": GEMINI_API_KEY, "model": GEMINI_MODEL},
}

_OPENAI_COMPATIBLE = ("kimi", "grok", "minimax", "glm", "qwen", "deepseek", "openai", "perplexity")


def _is_retryable(exc: BaseException) -> bool:
    """判断异常是否值得重试：5xx、429、超时、连接错误。"""
    if isinstance(exc, (httpx.TimeoutException, httpx.ConnectError, ConnectionError, TimeoutError)):
        return True
    try:
        from openai import APIStatusError, APITimeoutError, APIConnectionError
        if isinstance(exc, (APITimeoutError, APIConnectionError)):
            return True
        if isinstance(exc, APIStatusError) and exc.status_code in (429, 500, 502, 503, 504):
            return True
    except ImportError:
        pass
    try:
        from anthropic import APIStatusError as AnthropicStatusError, APITimeoutError as AnthropicTimeout, APIConnectionError as AnthropicConnError
        if isinstance(exc, (AnthropicTimeout, AnthropicConnError)):
            return True
        if isinstance(exc, AnthropicStatusError) and exc.status_code in (429, 500, 502, 503, 504):
            return True
    except ImportError:
        pass
    if isinstance(exc, RuntimeError):
        msg = str(exc)
        if any(f"API {code}" in msg for code in ("429", "500", "502", "503", "504")):
            return True
    return False


def _log_retry(retry_state):
    exc = retry_state.outcome.exception()
    attempt = retry_state.attempt_number
    ts = _time.strftime("%H:%M:%S", _time.localtime())
    print(f"[{ts}] [LLM重试] 第 {attempt} 次失败: {type(exc).__name__}: {str(exc)[:200]}，即将重试...", flush=True)


_llm_retry = retry(
    retry=retry_if_exception(_is_retryable),
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=2, min=2, max=16),
    before_sleep=_log_retry,
    reraise=True,
)


@_llm_retry
def _openai_compatible_chat(provider: str, messages: list, model: str = None, max_tokens: int = 8192, temperature: float = 0.6) -> str:
    """OpenAI 兼容 API。"""
    cfg = PROVIDER_CONFIG.get(provider, PROVIDER_CONFIG["kimi"])
    key = cfg["key"]
    base_url = cfg.get("base_url")
    default_model = cfg.get("model", "gpt-5.4")
    if not key:
        raise ValueError(f"请设置 {provider.upper()}_API_KEY 或在 .env 中配置")
    client = OpenAI(api_key=key, base_url=base_url, http_client=httpx.Client(timeout=HTTP_TIMEOUT))
    m = model or default_model
    if provider == "kimi" and "k2" in m.lower():
        temperature = 1.0
    if provider == "deepseek" and max_tokens > 8192:
        max_tokens = 8192
    resp = client.chat.completions.create(
        model=m,
        messages=messages,
        max_tokens=max_tokens,
        temperature=temperature,
    )
    return (resp.choices[0].message.content or "").strip()


@_llm_retry
def _claude_chat(messages: list, model: str = None, max_tokens: int = 8192, temperature: float = 0.6) -> str:
    """Anthropic Claude API。"""
    if not ANTHROPIC_API_KEY:
        raise ValueError("请设置 ANTHROPIC_API_KEY 或在 .env 中配置")
    from anthropic import Anthropic
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    m = model or ANTHROPIC_MODEL
    system = ""
    msgs = []
    for item in messages:
        role = item.get("role", "")
        content = item.get("content", "")
        if isinstance(content, list):
            content = "\n".join(
                p.get("text", str(p)) for p in content
                if isinstance(p, dict) and ("text" in p or p.get("type") == "text")
            )
        if role == "system":
            system = content
        elif role in ("assistant", "user"):
            msgs.append({"role": role, "content": content})
    kwargs = {"model": m, "max_tokens": max_tokens, "temperature": temperature, "messages": msgs}
    if system:
        kwargs["system"] = system
    resp = client.messages.create(**kwargs)
    return (resp.content[0].text if resp.content else "").strip()


@_llm_retry
def _gemini_chat(messages: list, model: str = None, max_tokens: int = 8192, temperature: float = 0.6) -> str:
    """Google Gemini API。"""
    if not GEMINI_API_KEY:
        raise ValueError("请设置 GEMINI_API_KEY 或在 .env 中配置")
    import google.generativeai as genai
    genai.configure(api_key=GEMINI_API_KEY)
    m = model or GEMINI_MODEL
    parts = []
    for item in messages:
        role = item.get("role", "")
        content = item.get("content", "")
        if isinstance(content, list):
            content = "\n".join(
                p.get("text", str(p)) for p in content
                if isinstance(p, dict) and ("text" in p or p.get("type") == "text")
            )
        if role == "system":
            parts.append(f"[System]\n{content}")
        elif role == "user":
            parts.append(f"[User]\n{content}")
        elif role == "assistant":
            parts.append(f"[Assistant]\n{content}")
    prompt = "\n\n".join(parts) if parts else ""
    gen_model = genai.GenerativeModel(m)
    resp = gen_model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=max_tokens,
            temperature=temperature,
        ),
    )
    return (resp.text or "").strip()


@_llm_retry
def _cloubic_chat(provider: str, messages: list, model: str = None,
                  max_tokens: int = 8192, temperature: float = 0.6) -> str:
    """通过 Cloubic 统一路由调用任意模型（OpenAI 兼容协议）。"""
    if not CLOUBIC_API_KEY:
        raise ValueError("Cloubic 模式已启用但未设置 CLOUBIC_API_KEY，请在 .env.cloubic 中配置")
    m = model or CLOUBIC_MODEL_MAP.get(provider, "deepseek-chat")
    # Cloubic 调用不走系统代理（无代理模式）
    client = OpenAI(api_key=CLOUBIC_API_KEY, base_url=CLOUBIC_BASE_URL,
                    http_client=httpx.Client(timeout=HTTP_TIMEOUT, proxy=None))
    # Cloubic 是 OpenAI 兼容，system/user/assistant 消息直接传
    resp = client.chat.completions.create(
        model=m,
        messages=messages,
        max_tokens=max_tokens,
        temperature=temperature,
    )
    return (resp.choices[0].message.content or "").strip()


def _is_cloubic_mode() -> bool:
    """判断当前是否为 Cloubic 路由模式（运行时动态检查）。"""
    import config
    return config.CLOUBIC_ENABLED and bool(config.CLOUBIC_API_KEY)


def _should_route_via_cloubic(provider: str) -> bool:
    """判断该 provider 是否应走 Cloubic 路由。"""
    if not _is_cloubic_mode():
        return False
    if provider == "kimi":  # kimi 始终直连
        return False
    import config
    # 白名单为空 = 全部走 Cloubic
    if not config.CLOUBIC_ROUTED_PROVIDERS:
        return True
    return provider in config.CLOUBIC_ROUTED_PROVIDERS


def chat(
    messages: list,
    provider: str = None,
    model: str = None,
    max_tokens: int = 8192,
    temperature: float = 0.6,
) -> str:
    """
    统一对话接口。provider 未指定时使用环境变量 LLM_PROVIDER（默认 kimi）。
    当 Cloubic 路由启用时，所有请求通过 Cloubic 转发。
    """
    p = (provider or os.getenv("LLM_PROVIDER") or LLM_PROVIDER or "kimi").lower().strip()

    # 按白名单判断是否走 Cloubic（kimi 始终直连）
    if _should_route_via_cloubic(p):
        return _cloubic_chat(p, messages, model, max_tokens, temperature)

    # 直连模式
    if p == "claude":
        return _claude_chat(messages, model, max_tokens, temperature)
    if p == "gemini":
        return _gemini_chat(messages, model, max_tokens, temperature)
    if p in _OPENAI_COMPATIBLE:
        return _openai_compatible_chat(p, messages, model, max_tokens, temperature)
    return _openai_compatible_chat("kimi", messages, model, max_tokens, temperature)


def chat_vision(
    messages: list,
    provider: str = None,
    model: str = None,
    max_tokens: int = 8192,
    temperature: float = 0.3,
) -> str:
    """多模态对话（支持图片）。优先使用 Vision 模型。"""
    p = (provider or os.getenv("LLM_PROVIDER") or LLM_PROVIDER or "kimi").lower().strip()

    # 按白名单判断是否走 Cloubic
    if _should_route_via_cloubic(p):
        return _cloubic_chat(p, messages, model, max_tokens, temperature)

    # 直连模式
    if p == "kimi":
        model = model or KIMI_VISION_MODEL
        return _openai_compatible_chat("kimi", messages, model, max_tokens, temperature)
    if p in ("openai", "grok", "perplexity", "glm", "minimax", "qwen", "deepseek"):
        if p == "glm":
            model = model or GLM_VISION_MODEL
        return _openai_compatible_chat(p, messages, model, max_tokens, temperature)
    if p == "claude":
        return _claude_chat(messages, model, max_tokens, temperature)
    if p == "gemini":
        return _gemini_chat(messages, model, max_tokens, temperature)
    return _openai_compatible_chat("kimi", messages, model or KIMI_VISION_MODEL, max_tokens, temperature)


def chat_reasoning(
    messages: list,
    provider: str = None,
    model: str = None,
    max_tokens: int = 16384,
    temperature: float = 0.6,
) -> str:
    """
    使用推理/思考版模型。仅 Cloubic 模式支持。
    直连模式回退到普通 chat()。
    """
    p = (provider or os.getenv("LLM_PROVIDER") or LLM_PROVIDER or "kimi").lower().strip()

    if _should_route_via_cloubic(p):
        reasoning_model = model or CLOUBIC_REASONING_MODEL_MAP.get(p)
        if reasoning_model:
            return _cloubic_chat(p, messages, reasoning_model, max_tokens, temperature)

    # 直连模式无推理版，回退普通 chat
    return chat(messages, provider, model, max_tokens, temperature)

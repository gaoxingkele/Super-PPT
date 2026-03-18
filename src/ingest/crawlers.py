# -*- coding: utf-8 -*-
"""URL 爬取：使用 Pyppeteer/Playwright 提取网页正文。"""
import asyncio
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

import src  # noqa: F401
from config import MIN_CONTENT_BYTES, RETRY_WAIT_SECONDS, CRAWL_MAX_RETRIES


@dataclass
class CrawlResult:
    """爬取结果。"""
    text: str = ""
    images: List[str] = field(default_factory=list)
    title: str = ""


_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"


def _find_chrome_executable() -> str:
    """查找 Chrome/Edge 浏览器路径。"""
    import os
    candidates = [
        os.environ.get("PUPPETEER_EXECUTABLE_PATH"),
        os.environ.get("CHROME_PATH"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in candidates:
        if p and Path(p).exists():
            return p
    return ""


async def _crawl_pyppeteer(url: str) -> CrawlResult:
    """使用 Pyppeteer 爬取网页正文。"""
    from pyppeteer import launch

    opts = {
        "headless": True,
        "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
    }
    chrome = _find_chrome_executable()
    if chrome:
        opts["executablePath"] = chrome

    browser = await launch(**opts)
    try:
        page = await browser.newPage()
        await page.setViewport({"width": 1280, "height": 800})
        await page.setUserAgent(_USER_AGENT)

        result = CrawlResult()
        for attempt in range(1, CRAWL_MAX_RETRIES + 1):
            await page.goto(url, waitUntil="networkidle2", timeout=60000)
            await asyncio.sleep(2.5)

            # 提取标题和正文
            title = await page.evaluate("() => document.title || ''")
            body_text = await page.evaluate("""() => {
                const main = document.querySelector('main') || document.querySelector('article') || document.body;
                return main ? main.innerText || main.textContent || '' : '';
            }""")

            result = CrawlResult(text=(body_text or "").strip(), title=(title or "").strip())
            if len(result.text.encode("utf-8")) >= MIN_CONTENT_BYTES:
                break
            if attempt < CRAWL_MAX_RETRIES:
                await asyncio.sleep(RETRY_WAIT_SECONDS)

        return result
    finally:
        await browser.close()


async def _crawl_playwright(url: str) -> CrawlResult:
    """使用 Playwright 爬取网页正文。"""
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        try:
            page = await browser.new_page()
            await page.set_viewport_size({"width": 1280, "height": 800})
            await page.set_extra_http_headers({"User-Agent": _USER_AGENT})

            result = CrawlResult()
            for attempt in range(1, CRAWL_MAX_RETRIES + 1):
                await page.goto(url, wait_until="networkidle", timeout=60000)
                await asyncio.sleep(2.5)

                title = await page.evaluate("() => document.title || ''")
                body_text = await page.evaluate("""() => {
                    const main = document.querySelector('main') || document.querySelector('article') || document.body;
                    return main ? main.innerText || main.textContent || '' : '';
                }""")

                result = CrawlResult(text=(body_text or "").strip(), title=(title or "").strip())
                if len(result.text.encode("utf-8")) >= MIN_CONTENT_BYTES:
                    break
                if attempt < CRAWL_MAX_RETRIES:
                    await asyncio.sleep(RETRY_WAIT_SECONDS)

            return result
        finally:
            await browser.close()


def crawl_url(url: str) -> CrawlResult:
    """同步入口：爬取 URL 返回正文。首选 Pyppeteer，备用 Playwright。"""
    try:
        result = asyncio.run(_crawl_pyppeteer(url))
    except ImportError:
        print("[提示] Pyppeteer 未安装，尝试 Playwright...")
        result = asyncio.run(_crawl_playwright(url))
    except Exception as e:
        err_str = str(e).lower()
        if "executable" in err_str or "browser" in err_str or "chrome" in err_str:
            print("[提示] 未找到 Chrome/Edge，尝试 Playwright...")
        try:
            result = asyncio.run(_crawl_playwright(url))
        except Exception as e2:
            raise RuntimeError(
                f"爬虫失败。请安装浏览器：\n"
                f"  1. 首选: 安装 Chrome/Edge\n"
                f"  2. 备用: python -m playwright install chromium\n"
                f"错误: {e2}"
            ) from e2

    return result

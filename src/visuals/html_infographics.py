# -*- coding: utf-8 -*-
"""
AntV Infographic + Playwright 信息图渲染器。
LLM 生成 AntV DSL 语法 → 嵌入 HTML 模板 → Playwright 截图为 1920×1080 PNG。
支持全部 9 种信息图类型。
"""
import json
import logging
import re
import tempfile
import time
from datetime import datetime
from pathlib import Path

import src  # noqa: F401
from src.llm_client import chat

logger = logging.getLogger(__name__)

# 走 HTML 渲染的信息图类型（全部 9 种）
HTML_SUPPORTED_TYPES = {
    "process_flow", "timeline", "hierarchy", "comparison",
    "matrix", "cycle", "network", "pyramid", "stat_display",
}

# infographic_type → 推荐的 AntV 模板映射
_TYPE_TEMPLATE_MAP = {
    "process_flow": [
        "sequence-snake-steps-compact-card",
        "sequence-color-snake-steps-horizontal-icon-line",
        "sequence-zigzag-steps-underline-text",
        "sequence-ascending-steps",
        "list-row-horizontal-icon-arrow",
    ],
    "timeline": [
        "sequence-timeline-simple",
        "sequence-timeline-rounded-rect-node",
        "sequence-roadmap-vertical-simple",
    ],
    "hierarchy": [
        "hierarchy-tree-tech-style-badge-card",
        "hierarchy-tree-curved-line-rounded-rect-node",
        "hierarchy-tree-tech-style-capsule-item",
        "hierarchy-structure",
    ],
    "comparison": [
        "compare-binary-horizontal-badge-card-arrow",
        "compare-binary-horizontal-underline-text-vs",
        "compare-binary-horizontal-simple-fold",
        "compare-hierarchy-left-right-circle-node-pill-badge",
    ],
    "matrix": [
        "list-grid-badge-card",
        "list-grid-candy-card-lite",
        "list-grid-ribbon-card",
        "quadrant-quarter-simple-card",
    ],
    "cycle": [
        "sequence-circular-simple",
        "relation-circle-icon-badge",
        "relation-circle-circular-progress",
    ],
    "network": [
        "relation-circle-icon-badge",
        "relation-circle-circular-progress",
        "hierarchy-structure",
    ],
    "pyramid": [
        "sequence-pyramid-simple",
        "sequence-filter-mesh-simple",
        "sequence-ascending-stairs-3d-underline-text",
    ],
    "stat_display": [
        "list-grid-badge-card",
        "list-grid-candy-card-lite",
        "list-row-horizontal-icon-arrow",
        "list-column-vertical-icon-arrow",
    ],
}

# AntV HTML 模板
_ANTV_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>Infographic</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body { width: 1920px; height: 1080px; overflow: hidden; background: #fff; }
  #container { width: 1920px; height: 1080px; }
</style>
</head>
<body>
<div id="container"></div>
<script src="https://unpkg.com/@antv/infographic@latest/dist/infographic.min.js"></script>
<script>
const svgTextCache = new Map();
const pendingRequests = new Map();

AntVInfographic.registerResourceLoader(async (config) => {
    const { data, scene } = config;
    try {
        const key = `${scene}::${data}`;
        let svgText;
        if (svgTextCache.has(key)) {
            svgText = svgTextCache.get(key);
        } else if (pendingRequests.has(key)) {
            svgText = await pendingRequests.get(key);
        } else {
            const fetchPromise = (async () => {
                try {
                    let url;
                    if (scene === 'icon') {
                        url = `https://api.iconify.design/${data}.svg`;
                    } else if (scene === 'illus') {
                        url = `https://raw.githubusercontent.com/balazser/undraw-svg-collection/refs/heads/main/svgs/${data}.svg`;
                    } else return null;
                    if (!url) return null;
                    const response = await fetch(url, { referrerPolicy: 'no-referrer' });
                    if (!response.ok) return null;
                    const text = await response.text();
                    if (!text || !text.trim().startsWith('<svg')) return null;
                    svgTextCache.set(key, text);
                    return text;
                } catch (e) { return null; }
            })();
            pendingRequests.set(key, fetchPromise);
            try { svgText = await fetchPromise; }
            finally { pendingRequests.delete(key); }
        }
        if (!svgText) return null;
        const resource = AntVInfographic.loadSVGResource(svgText);
        return resource || null;
    } catch (e) { return null; }
});
</script>
<script>
const infographic = new AntVInfographic.Infographic({
    container: '#container',
    width: 1920,
    height: 1080,
});
const syntax = `SYNTAX_PLACEHOLDER`;
infographic.render(syntax);
// 等字体加载完后重新渲染
document.fonts?.ready.then(() => {
    infographic.render(syntax);
}).catch(() => {});
</script>
</body>
</html>"""


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _build_dsl_prompt(visual: dict, color_scheme: dict) -> str:
    """构建 LLM prompt，要求生成 AntV Infographic DSL 语法。"""
    infographic_type = visual.get("infographic_type", "process_flow")
    data = visual.get("data", {})
    description = visual.get("description", "")

    primary = color_scheme.get("primary", "#002060")
    accent = color_scheme.get("accent", "#C00000")
    secondary = color_scheme.get("secondary", "#4A90D9")

    data_text = json.dumps(data, ensure_ascii=False, indent=2, default=str) if data else ""
    templates = _TYPE_TEMPLATE_MAP.get(infographic_type, ["list-grid-badge-card"])
    template_list = "\n".join(f"  - {t}" for t in templates)

    prompt = f"""你是 AntV Infographic 信息图专家。请根据以下数据生成 AntV Infographic DSL 语法。

## 数据内容
{data_text}

## 内容描述
{description}

## 推荐模板（从中选择最合适的一个）
{template_list}

## 配色方案
- 主色: {primary}
- 辅色: {secondary}
- 强调色: {accent}

## AntV Infographic DSL 语法规范

语法格式如下（用两空格缩进）：
```
infographic <template-name>
theme
  palette {primary} {secondary} {accent}
data
  title 标题文字
  desc 描述文字
  items
    - label 标签1
      value 数值1
      desc 说明1
      icon mdi/icon-name
    - label 标签2
      value 数值2
      desc 说明2
      icon mdi/icon-name
```

## 关键规则
1. 第一行必须是 `infographic <模板名>`
2. 所有文字必须是简体中文
3. items 中每项用 `- ` 开头（注意减号后有空格）
4. 每项可包含 label(文字)、value(数值)、desc(说明)、icon(图标名)
5. icon 格式为 `mdi/<icon-name>`，如 mdi/chip、mdi/factory、mdi/shield-check、mdi/chart-line、mdi/cog、mdi/account-group
6. 对比类模板（compare-*）：构建两个根节点，每个根节点下用 children 放对比项
7. 层级类模板（hierarchy-*）：用 children 嵌套表示层级关系
8. theme 中用 palette 设置配色，多个颜色空格分隔
9. value 字段放纯数字或带单位的文本均可

## 输出要求
只输出 DSL 代码块，不要任何解释。用 ```plain 和 ``` 包裹。
"""
    return prompt


def _extract_dsl(llm_response: str) -> str:
    """从 LLM 响应中提取 AntV DSL 语法。"""
    # 尝试 ```plain ... ```
    match = re.search(r'```(?:plain|text)?\s*(infographic\s+.*?)\s*```', llm_response, re.DOTALL)
    if match:
        return match.group(1).strip()
    # 尝试直接匹配 infographic 开头的内容
    match = re.search(r'(infographic\s+\S+.*?)(?:\n\n\n|\Z)', llm_response, re.DOTALL)
    if match:
        return match.group(1).strip()
    return ""


def _build_html(dsl_syntax: str) -> str:
    """将 DSL 语法嵌入 HTML 模板。"""
    # 转义反引号，避免破坏 JS 模板字符串
    escaped = dsl_syntax.replace('\\', '\\\\').replace('`', '\\`').replace('${', '\\${')
    return _ANTV_HTML_TEMPLATE.replace('SYNTAX_PLACEHOLDER', escaped)


def _screenshot_html(html_content: str, output_path: Path) -> bool:
    """用 Playwright 将 HTML 截图为 PNG。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.error("playwright 未安装，无法截图")
        return False

    tmp_html = None
    try:
        tmp_html = tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8")
        tmp_html.write(html_content)
        tmp_html.close()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={"width": 1920, "height": 1080})
            page.goto(f"file:///{tmp_html.name.replace(chr(92), '/')}")
            # 等待 AntV 渲染完成（需要加载外部 JS + 图标）
            page.wait_for_timeout(5000)

            output_path.parent.mkdir(parents=True, exist_ok=True)
            page.screenshot(path=str(output_path), full_page=False)
            browser.close()

        if output_path.exists() and output_path.stat().st_size > 1000:
            logger.info("[%s] AntV 信息图截图成功: %s (%.1f KB)",
                        _ts(), output_path, output_path.stat().st_size / 1024)
            return True
        logger.warning("[%s] AntV 截图文件过小或为空: %s", _ts(), output_path)
        return False

    except Exception as exc:
        logger.warning("[%s] AntV 信息图截图失败: %s", _ts(), exc)
        return False
    finally:
        if tmp_html:
            try:
                Path(tmp_html.name).unlink(missing_ok=True)
            except Exception:
                pass


def render_html_infographic(visual: dict, color_scheme: dict, output_path: Path) -> bool:
    """
    用 LLM 生成 AntV DSL + Playwright 截图生成信息图。
    成功返回 True，失败返回 False（调用方可 fallback 到 AI 生图）。
    """
    infographic_type = visual.get("infographic_type", "process_flow")
    if infographic_type not in HTML_SUPPORTED_TYPES:
        logger.info("[%s] 类型 %s 不走 AntV 渲染", _ts(), infographic_type)
        return False

    output_path = Path(output_path)
    start = time.time()

    # 1) LLM 生成 AntV DSL
    prompt = _build_dsl_prompt(visual, color_scheme)
    logger.info("[%s] 调用 LLM 生成 AntV DSL | type=%s", _ts(), infographic_type)

    try:
        messages = [{"role": "user", "content": prompt}]
        response = chat(messages, max_tokens=4096, temperature=0.3)
    except Exception as exc:
        logger.warning("[%s] LLM 调用失败: %s", _ts(), exc)
        return False

    dsl = _extract_dsl(response)
    if not dsl:
        logger.warning("[%s] LLM 响应中未提取到有效 AntV DSL", _ts())
        return False

    logger.info("[%s] AntV DSL 生成成功 (%d 字符)\n%s", _ts(), len(dsl), dsl[:300])

    # 2) 嵌入 HTML 模板
    html = _build_html(dsl)

    # 3) Playwright 截图
    success = _screenshot_html(html, output_path)
    elapsed = time.time() - start
    if success:
        logger.info("[%s] AntV 信息图完成 | 耗时 %.1f 秒 | %s", _ts(), elapsed, output_path)
    else:
        logger.warning("[%s] AntV 信息图失败 | 耗时 %.1f 秒", _ts(), elapsed)

    return success

# -*- coding: utf-8 -*-
"""
Step1: 结构化分析。
LLM 深度分析内容，产出章节划分、关键数据点、核心概念、论证链。
"""
import json
from pathlib import Path

import src  # noqa: F401
from config import ANALYZE_CONTENT_LIMIT
from src.llm_client import chat
from src.prompts.analyze import ANALYZE_SYSTEM_PROMPT, build_analyze_user_prompt


def run_analyze(base: str, output_dir: Path) -> dict:
    """
    Step1 入口：对原始内容做结构化分析。

    Args:
        base: 项目名称
        output_dir: output/{base}/ 目录

    Returns:
        {
            "analysis_path": Path,
            "analysis": dict,      # 分析结果
        }
    """
    # 读取 Step0 输出
    raw_content = (output_dir / "raw_content.md").read_text(encoding="utf-8", errors="replace")
    raw_meta = json.loads((output_dir / "raw_meta.json").read_text(encoding="utf-8"))
    raw_tables = json.loads((output_dir / "raw_tables.json").read_text(encoding="utf-8"))

    # 截断内容
    content = raw_content[:ANALYZE_CONTENT_LIMIT]

    # 构建 prompt
    user_prompt = build_analyze_user_prompt(content, raw_meta, raw_tables)

    messages = [
        {"role": "system", "content": ANALYZE_SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ]

    print("[Step1] 正在进行结构化分析...", flush=True)
    response = chat(messages, max_tokens=16384, temperature=0.4)

    # 解析 JSON 响应
    analysis = _parse_analysis_response(response)

    # 补充元数据
    if not analysis.get("title") and raw_meta.get("title"):
        analysis["title"] = raw_meta["title"]

    # 保存
    analysis_path = output_dir / "analysis.json"
    analysis_path.write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")

    chapter_count = len(analysis.get("chapters", []))
    data_count = sum(len(ch.get("data_points", [])) for ch in analysis.get("chapters", []))
    concept_count = sum(len(ch.get("concepts", [])) for ch in analysis.get("chapters", []))
    content_type = analysis.get("content_type", "未识别")
    print(f"[Step1] 分析完成: 类型={content_type}, {chapter_count} 章节, "
          f"{data_count} 个数据点, {concept_count} 个概念", flush=True)

    # 打印视觉策划摘要
    if analysis.get("narrative_arc"):
        arc = analysis["narrative_arc"]
        print(f"[Step1] 叙事策划: 开场={arc.get('opening_strategy', '?')} → "
              f"高潮={arc.get('climax_chapter', '?')} → "
              f"收尾={arc.get('closing_strategy', '?')}", flush=True)
    if analysis.get("rhythm_plan"):
        print(f"[Step1] 节奏规划: {' → '.join(analysis['rhythm_plan'][:10])}...", flush=True)

    return {"analysis_path": analysis_path, "analysis": analysis}


def _parse_analysis_response(response: str) -> dict:
    """解析 LLM 返回的 JSON，容错处理。"""
    text = response.strip()

    # 提取 code block
    if "```json" in text:
        start = text.index("```json") + 7
        end = text.find("```", start)
        text = (text[start:end] if end != -1 else text[start:]).strip()
    elif "```" in text:
        start = text.index("```") + 3
        end = text.find("```", start)
        text = (text[start:end] if end != -1 else text[start:]).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    idx = text.find("{")
    if idx >= 0:
        last_brace = text.rfind("}")
        if last_brace > idx:
            try:
                return json.loads(text[idx:last_brace + 1])
            except json.JSONDecodeError:
                pass
        candidate = text[idx:]
        for suffix in ["}", "]}", "]}}", "]}]}}"]:
            try:
                return json.loads(candidate + suffix)
            except json.JSONDecodeError:
                continue

    return {"title": "", "chapters": [], "parse_error": True, "raw_response": response[:2000]}

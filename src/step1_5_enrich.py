# -*- coding: utf-8 -*-
"""
Step1.5: 联网数据补充。
对 analysis.json 中的数据点、关键论点进行联网检索，
补充最新的数据、统计结论、法案进展、重大事件。
将补充结果写入 raw_content.md 末尾 + 更新 analysis.json。
"""
import json
import time
from pathlib import Path

import src  # noqa: F401
from src.llm_client import chat


# ── 提取需要检索的关键词 ──
def _extract_search_queries(analysis: dict) -> list:
    """从 analysis.json 中提取需要联网检索的关键话题。"""
    queries = []

    title = analysis.get("title", "")
    core_thesis = analysis.get("core_thesis", "")

    # 全局话题
    if title:
        queries.append(f"{title} 最新进展 2025 2026")

    # 每章的关键数据点和概念
    for ch in analysis.get("chapters", []):
        ch_title = ch.get("title", "")
        key_points = ch.get("key_points", [])
        data_points = ch.get("data_points", [])
        concepts = ch.get("concepts", [])

        # 章节标题 + 最新
        if ch_title:
            queries.append(f"{ch_title} latest update 2025 2026")

        # 关键数据点
        for dp in data_points[:3]:
            label = dp.get("label", "") if isinstance(dp, dict) else str(dp)
            if label:
                queries.append(f"{label} 最新数据 2025 2026")

        # 核心概念
        for c in concepts[:3]:
            name = c.get("name", c) if isinstance(c, dict) else str(c)
            if name:
                queries.append(f"{name} 最新进展 2025 2026")

    # 去重 + 限制数量
    seen = set()
    unique = []
    for q in queries:
        q_key = q[:30]
        if q_key not in seen:
            seen.add(q_key)
            unique.append(q)
    return unique[:15]  # 最多15个查询


# ── 用 LLM 联网检索（通过 Perplexity 或有联网能力的模型） ──
def _search_latest_data(queries: list, analysis: dict) -> str:
    """
    用 LLM 对关键话题进行联网检索，返回最新数据补充文本。
    使用 Cloubic 路由的推理模型或联网模型。
    """
    # 构建批量检索 prompt
    query_list = "\n".join(f"{i+1}. {q}" for i, q in enumerate(queries))

    chapters_summary = ""
    for ch in analysis.get("chapters", []):
        ch_title = ch.get("title", "")
        ch_summary = ch.get("summary", "")[:100]
        chapters_summary += f"- {ch_title}: {ch_summary}\n"

    system_prompt = """你是一位资深的半导体产业与地缘政治研究员。
你的任务是针对一份2023年出版的研究报告，补充2024-2026年间的最新进展。

要求：
1. 只提供有据可查的事实更新，不编造数据
2. 每条更新标注时间（如"2025年3月"）和来源类型（如"美国商务部公告"、"台积电财报"）
3. 重点关注：法案进展、产能数据、市场份额变化、重大政策、技术突破、地缘事件
4. 用中文输出
5. 按章节对应关系组织，方便后续整合到PPT中"""

    user_prompt = f"""以下是一份2023年出版的报告的章节结构：

{chapters_summary}

请针对以下关键话题，检索并提供2024-2026年的最新数据更新：

{query_list}

请按以下格式输出：

# 最新数据补充（2024-2026）

## 对应章节：[章节标题]
- **[更新主题]**（[时间]）：[具体数据/事实]。来源：[来源类型]
- ...

每个章节至少提供3-5条关键更新。重点关注与原报告论点相关的数据变化。"""

    print("[Step1.5] 正在联网检索最新数据...", flush=True)
    response = chat(
        [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_tokens=8192,
        temperature=0.3,
    )
    return response


# ── 将补充数据整合到 analysis.json ──
def _merge_enrichment(analysis: dict, enrichment_text: str) -> dict:
    """让 LLM 将联网补充数据整合到 analysis.json 的 data_points 和 key_points 中。"""

    system_prompt = """你是一位数据整合专家。你需要将最新的联网检索结果整合到现有的分析结构中。

规则：
1. 在每个章节的 data_points 末尾追加新数据点（type="update", label="...", data={...}）
2. 在每个章节的 key_points 末尾追加包含最新数据的要点（以"【2024-2026更新】"开头）
3. 不修改已有内容，只追加
4. 输出完整的 JSON（与输入结构一致）"""

    user_prompt = f"""现有分析结构：
```json
{json.dumps(analysis, ensure_ascii=False, indent=2)[:12000]}
```

联网检索到的最新数据：
{enrichment_text[:6000]}

请将最新数据整合到分析结构中，输出完整 JSON。"""

    print("[Step1.5] 正在整合最新数据到分析结构...", flush=True)
    response = chat(
        [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_tokens=16384,
        temperature=0.2,
    )

    # 解析 JSON
    text = response.strip()
    if "```json" in text:
        start = text.index("```json") + 7
        end = text.find("```", start)
        text = text[start:end].strip() if end > start else text[start:].strip()
    elif "```" in text:
        start = text.index("```") + 3
        end = text.find("```", start)
        text = text[start:end].strip() if end > start else text[start:].strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # 找 { 开头
        idx = text.find("{")
        if idx >= 0:
            last = text.rfind("}")
            if last > idx:
                try:
                    return json.loads(text[idx:last+1])
                except json.JSONDecodeError:
                    pass
        print("[Step1.5] 警告: JSON 解析失败，保留原始分析结构", flush=True)
        return analysis


# ── 主入口 ──
def run_enrich(base: str, output_dir: Path) -> dict:
    """
    Step1.5 入口：联网数据补充。

    流程：
    1. 从 analysis.json 提取关键检索话题
    2. 用 LLM 联网检索最新数据（2024-2026）
    3. 将补充数据写入 raw_content.md 末尾
    4. 将结构化更新整合到 analysis.json
    """
    analysis_path = output_dir / "analysis.json"
    analysis = json.loads(analysis_path.read_text(encoding="utf-8"))

    t_start = time.time()

    # 1. 提取检索关键词
    queries = _extract_search_queries(analysis)
    print(f"[Step1.5] 提取 {len(queries)} 个检索话题", flush=True)
    for i, q in enumerate(queries[:5]):
        print(f"  {i+1}. {q[:50]}", flush=True)
    if len(queries) > 5:
        print(f"  ... 共 {len(queries)} 个", flush=True)

    # 2. 联网检索
    enrichment_text = _search_latest_data(queries, analysis)

    # 3. 保存联网结果
    enrich_path = output_dir / "enrichment_data.md"
    enrich_path.write_text(enrichment_text, encoding="utf-8")
    print(f"[Step1.5] 联网数据已保存: {enrich_path}", flush=True)

    # 4. 追加到 raw_content.md
    raw_content_path = output_dir / "raw_content.md"
    raw_content = raw_content_path.read_text(encoding="utf-8")
    separator = "\n\n" + "=" * 60 + "\n"
    separator += "# 最新数据补充（2024-2026年联网检索）\n"
    separator += f"# 检索时间: {time.strftime('%Y-%m-%d %H:%M')}\n"
    separator += "=" * 60 + "\n\n"
    raw_content_path.write_text(raw_content + separator + enrichment_text, encoding="utf-8")
    print(f"[Step1.5] 已追加到 raw_content.md", flush=True)

    # 5. 整合到 analysis.json
    enriched_analysis = _merge_enrichment(analysis, enrichment_text)

    # 备份原始 analysis
    backup_path = output_dir / "analysis_original.json"
    if not backup_path.exists():
        backup_path.write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")

    # 保存更新后的 analysis
    analysis_path.write_text(json.dumps(enriched_analysis, ensure_ascii=False, indent=2), encoding="utf-8")

    elapsed = time.time() - t_start
    # 统计新增数据量
    old_dp = sum(len(ch.get("data_points", [])) for ch in analysis.get("chapters", []))
    new_dp = sum(len(ch.get("data_points", [])) for ch in enriched_analysis.get("chapters", []))
    old_kp = sum(len(ch.get("key_points", [])) for ch in analysis.get("chapters", []))
    new_kp = sum(len(ch.get("key_points", [])) for ch in enriched_analysis.get("chapters", []))

    print(f"[Step1.5] 数据补充完成 ({elapsed:.1f}s)", flush=True)
    print(f"  数据点: {old_dp} → {new_dp} (+{new_dp - old_dp})", flush=True)
    print(f"  关键要点: {old_kp} → {new_kp} (+{new_kp - old_kp})", flush=True)

    return {
        "enrichment_path": enrich_path,
        "analysis_path": analysis_path,
        "queries_count": len(queries),
        "new_data_points": new_dp - old_dp,
        "new_key_points": new_kp - old_kp,
    }

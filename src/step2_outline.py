# -*- coding: utf-8 -*-
"""
Step2: 幻灯片大纲生成（两阶段模式）。
Phase A: 结构编排 — 决定页数、布局、节奏骨架
Phase B: 逐页设计 — 按章节批次填充 bullets / visual / notes
兼容旧的单次生成模式（fallback）。
"""
import json
from pathlib import Path

import src  # noqa: F401
from config import OUTLINE_CONTENT_LIMIT, DEFAULT_SLIDE_RANGE
from src.llm_client import chat
from src.prompts.outline import (
    OUTLINE_SYSTEM_PROMPT, build_outline_user_prompt,
    OUTLINE_SKELETON_PROMPT, OUTLINE_DETAIL_PROMPT,
    build_skeleton_user_prompt, build_detail_user_prompt,
)


# ── 每批次处理的最大页面数 ──
DETAIL_BATCH_SIZE = 8


def run_outline(base: str, output_dir: Path, style_profile: dict = None,
                slide_range: tuple = None, two_phase: bool = True) -> dict:
    """
    Step2 入口：生成幻灯片大纲。

    Args:
        base: 项目名称
        output_dir: output/{base}/ 目录
        style_profile: 风格配置（从参考模板提取，可选）
        slide_range: (min, max) 幻灯片数量范围
        two_phase: True=两阶段模式(推荐), False=兼容单次模式

    Returns:
        {"slide_plan_path": Path, "slide_plan": dict}
    """
    # 读取 Step1 输出
    analysis = json.loads((output_dir / "analysis.json").read_text(encoding="utf-8"))

    # 读取原始内容
    raw_content = (output_dir / "raw_content.md").read_text(encoding="utf-8", errors="replace")

    slide_range = slide_range or DEFAULT_SLIDE_RANGE

    if two_phase:
        slide_plan = _run_two_phase(analysis, raw_content, style_profile, slide_range, output_dir)
    else:
        slide_plan = _run_single_phase(analysis, raw_content, style_profile, slide_range)

    # ── 功能页自动校验与补全（借鉴 PPTAgent 规则引擎） ──
    slide_plan = _ensure_structural_pages(slide_plan, analysis)

    # 保存
    plan_path = output_dir / "slide_plan.json"
    plan_path.write_text(json.dumps(slide_plan, ensure_ascii=False, indent=2), encoding="utf-8")

    slide_count = len(slide_plan.get("slides", []))
    visual_count = sum(1 for s in slide_plan.get("slides", []) if s.get("visual"))
    print(f"[Step2] 大纲生成完成: {slide_count} 张幻灯片, {visual_count} 个视觉元素", flush=True)

    return {"slide_plan_path": plan_path, "slide_plan": slide_plan}


# ============================================================
# 两阶段模式
# ============================================================
def _run_two_phase(analysis: dict, raw_content: str,
                   style_profile: dict, slide_range: tuple,
                   output_dir: Path) -> dict:
    """Phase A 结构编排 + Phase B 逐页设计。"""

    # ── Phase A: 结构骨架 ──
    print("[Step2-A] 正在生成结构骨架...", flush=True)
    skeleton_prompt = build_skeleton_user_prompt(analysis, slide_range, style_profile)

    skeleton_response = chat(
        [
            {"role": "system", "content": OUTLINE_SKELETON_PROMPT},
            {"role": "user", "content": skeleton_prompt},
        ],
        max_tokens=8192,
        temperature=0.4,
    )
    skeleton = _parse_json_response(skeleton_response)

    # 保存中间结果
    (output_dir / "slide_skeleton.json").write_text(
        json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    slides = skeleton.get("slides", [])
    print(f"[Step2-A] 骨架完成: {len(slides)} 页", flush=True)

    if not slides:
        print("[Step2] 骨架为空，回退到单次模式", flush=True)
        return _run_single_phase(analysis, raw_content, style_profile, slide_range)

    # ── Phase B: 按章节批次填充内容 ──
    print("[Step2-B] 正在逐批填充页面内容...", flush=True)

    # 把页面按 chapter_ref 分组
    batches = _group_slides_into_batches(slides)
    chapter_map = _build_chapter_content_map(analysis, raw_content)

    filled_slides = {}

    for batch_idx, batch in enumerate(batches, 1):
        # 收集该批次涉及的章节内容
        chapter_refs = set(s.get("chapter_ref") or "" for s in batch)
        chapter_content_parts = []
        for ref in sorted(chapter_refs):
            if ref and ref in chapter_map:
                chapter_content_parts.append(chapter_map[ref])

        # 如果没有匹配的章节内容，用原始内容的前部分
        if not chapter_content_parts:
            chapter_content = raw_content[:4000]
        else:
            chapter_content = "\n\n".join(chapter_content_parts)
            # 限制长度避免超限
            if len(chapter_content) > 8000:
                chapter_content = chapter_content[:8000]

        print(f"[Step2-B] 批次 {batch_idx}/{len(batches)}: 填充 {len(batch)} 页 "
              f"(chapters: {', '.join(sorted(chapter_refs))})", flush=True)

        detail_prompt = build_detail_user_prompt(skeleton, batch, chapter_content, analysis)

        detail_response = chat(
            [
                {"role": "system", "content": OUTLINE_DETAIL_PROMPT},
                {"role": "user", "content": detail_prompt},
            ],
            max_tokens=16384,
            temperature=0.5,
        )
        detail_slides = _parse_json_response(detail_response)

        # detail_slides 可能是数组或 dict
        if isinstance(detail_slides, dict):
            # 可能返回了 {"slides": [...]} 的格式
            detail_slides = detail_slides.get("slides", [detail_slides])
        if not isinstance(detail_slides, list):
            detail_slides = [detail_slides]

        for ds in detail_slides:
            if isinstance(ds, dict) and "id" in ds:
                filled_slides[ds["id"]] = ds

    # ── 合并骨架 + 填充内容 ──
    merged_slides = []
    for skel_slide in slides:
        sid = skel_slide["id"]
        if sid in filled_slides:
            # 以填充结果为主，补充骨架中的字段
            merged = {**skel_slide, **filled_slides[sid]}
            merged_slides.append(merged)
        else:
            # 未被填充的页面（可能是 cover/end 等简单页面），保留骨架
            merged_slides.append(skel_slide)

    # 确保所有 slide 至少有基本字段
    for s in merged_slides:
        if "bullets" not in s:
            s["bullets"] = []
        if "notes" not in s:
            s["notes"] = ""
        if "takeaway" not in s:
            s["takeaway"] = ""
        # visual 字段：从 visual_type 推断（如果 Phase B 没填充）
        if "visual" not in s and s.get("visual_type") and s["visual_type"] != "null":
            s["visual"] = {"type": s["visual_type"]}

    slide_plan = {
        "meta": skeleton.get("meta", {}),
        "slides": merged_slides,
    }

    # 更新 total_slides
    slide_plan["meta"]["total_slides"] = len(merged_slides)

    return slide_plan


def _group_slides_into_batches(slides: list) -> list:
    """
    将页面分成批次。
    策略：按 chapter_ref 连续分组，每批不超过 DETAIL_BATCH_SIZE 页。
    """
    if not slides:
        return []

    batches = []
    current_batch = []
    current_chapter = None

    for s in slides:
        chapter = s.get("chapter_ref", "")

        # 换章节或批次已满，开始新批次
        if (chapter != current_chapter and current_batch) or len(current_batch) >= DETAIL_BATCH_SIZE:
            batches.append(current_batch)
            current_batch = []

        current_batch.append(s)
        current_chapter = chapter

    if current_batch:
        batches.append(current_batch)

    return batches


def _build_chapter_content_map(analysis: dict, raw_content: str) -> dict:
    """
    构建 chapter_id → 原始内容片段 的映射。
    策略：根据章节标题在原文中定位，提取对应段落。
    """
    chapters = analysis.get("chapters", [])
    if not chapters:
        return {}

    chapter_map = {}

    # 尝试按章节标题在原文中分割
    chapter_positions = []
    for ch in chapters:
        ch_id = ch.get("id", "")
        ch_title = ch.get("title", "")
        # 在原文中搜索章节标题
        pos = raw_content.find(ch_title)
        if pos == -1:
            # 尝试模糊匹配（取标题前几个字）
            short_title = ch_title[:8] if len(ch_title) > 8 else ch_title
            pos = raw_content.find(short_title)
        chapter_positions.append((ch_id, ch_title, pos, ch))

    # 按位置排序
    chapter_positions.sort(key=lambda x: x[2] if x[2] >= 0 else 999999)

    for i, (ch_id, ch_title, pos, ch) in enumerate(chapter_positions):
        if pos >= 0:
            # 找到下一个章节的位置
            next_pos = len(raw_content)
            for j in range(i + 1, len(chapter_positions)):
                if chapter_positions[j][2] >= 0:
                    next_pos = chapter_positions[j][2]
                    break
            content_slice = raw_content[pos:next_pos].strip()
        else:
            # 未在原文中找到，用分析中的摘要和要点
            parts = [f"## {ch_title}"]
            if ch.get("summary"):
                parts.append(ch["summary"])
            for kp in ch.get("key_points", []):
                parts.append(f"- {kp}")
            content_slice = "\n".join(parts)

        chapter_map[ch_id] = content_slice

    return chapter_map


# ============================================================
# 兼容的单次生成模式
# ============================================================
def _run_single_phase(analysis: dict, raw_content: str,
                      style_profile: dict, slide_range: tuple) -> dict:
    """原始的单次 LLM 调用生成完整 slide_plan。"""
    content = raw_content[:OUTLINE_CONTENT_LIMIT]
    user_prompt = build_outline_user_prompt(analysis, content, style_profile, slide_range)

    messages = [
        {"role": "system", "content": OUTLINE_SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ]

    if slide_range:
        print(f"[Step2] 正在生成幻灯片大纲 ({slide_range[0]}~{slide_range[1]} 张)...", flush=True)
    else:
        print("[Step2] 正在生成幻灯片大纲 (页数由内容决定)...", flush=True)

    response = chat(messages, max_tokens=16384, temperature=0.5)
    return _parse_json_response(response)


# ============================================================
# JSON 解析（容错）
# ============================================================
def _parse_json_response(response: str) -> dict:
    """解析 LLM 返回的 JSON，容错处理。"""
    text = response.strip()

    # 提取 code block 内容
    if "```json" in text:
        start = text.index("```json") + 7
        end = text.find("```", start)
        if end == -1:
            text = text[start:].strip()
        else:
            text = text[start:end].strip()
    elif "```" in text:
        start = text.index("```") + 3
        end = text.find("```", start)
        if end == -1:
            text = text[start:].strip()
        else:
            text = text[start:end].strip()

    # 尝试直接解析
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 尝试找到 { 或 [ 开头
    for opener, closer in [("{", "}"), ("[", "]")]:
        idx = text.find(opener)
        if idx >= 0:
            last = text.rfind(closer)
            if last > idx:
                try:
                    return json.loads(text[idx:last + 1])
                except json.JSONDecodeError:
                    pass
            # 尝试修复缺少结尾
            candidate = text[idx:]
            for suffix in [closer, f"{closer}{closer}", f"]{closer}", f"]{closer}{closer}"]:
                try:
                    return json.loads(candidate + suffix)
                except json.JSONDecodeError:
                    continue

    return {"meta": {}, "slides": [], "parse_error": True, "raw_response": response[:2000]}


# ============================================================
# 功能页自动校验与补全（借鉴 PPTAgent 规则引擎）
# ============================================================
def _ensure_structural_pages(slide_plan: dict, analysis: dict = None) -> dict:
    """
    校验并补全 PPT 必须的结构页面。

    借鉴 PPTAgent 的 functional layout insertion 规则：
    - 必须有 cover（封面）
    - 必须有 agenda（目录） — 5页以上的PPT
    - 必须有 end（结束页）
    - section_break 检查（章节过渡）

    如果 LLM 遗漏了这些页面，自动补充。
    """
    slides = slide_plan.get("slides", [])
    if not slides:
        return slide_plan

    meta = slide_plan.get("meta", {})
    title = meta.get("title", "演示文稿")
    subtitle = meta.get("subtitle", "")
    layouts = [s.get("layout", "") for s in slides]
    patched = False

    # ── 1. 检查封面 ──
    if "cover" not in layouts:
        cover = {
            "id": "s00",
            "layout": "cover",
            "title": title,
            "subtitle": subtitle,
            "bullets": [],
            "visual": {"type": "generate-image",
                       "prompt": "professional presentation cover, modern design, "
                                 "clean, 16:9 aspect ratio, high resolution"},
            "notes": f"欢迎各位。今天我将为大家介绍{title}。",
            "takeaway": "",
        }
        slides.insert(0, cover)
        patched = True
        print("[Step2] 自动补充: 封面页 (cover)", flush=True)

    # ── 2. 检查目录页（5页以上的PPT应有目录） ──
    if len(slides) >= 5 and "agenda" not in [s.get("layout") for s in slides]:
        # 从 analysis 或现有 section_break 提取章节标题
        chapter_bullets = []
        if analysis and analysis.get("chapters"):
            for ch in analysis["chapters"]:
                ch_title = ch.get("title", "")
                ch_summary = ch.get("summary", "")
                if ch_title:
                    bullet = f"{ch_title} — {ch_summary[:20]}" if ch_summary else ch_title
                    chapter_bullets.append(bullet)
        if not chapter_bullets:
            # 从 section_break 页面的标题提取
            for s in slides:
                if s.get("layout") == "section_break" and s.get("title"):
                    chapter_bullets.append(s["title"])

        if chapter_bullets:
            # 找到封面后的位置插入
            insert_pos = 1
            for i, s in enumerate(slides):
                if s.get("layout") == "cover":
                    insert_pos = i + 1
                    break

            agenda = {
                "id": "s_agenda",
                "layout": "agenda",
                "title": "目录",
                "subtitle": "CONTENTS",
                "bullets": chapter_bullets[:8],  # 最多8个章节
                "visual": None,
                "notes": f"本次报告共分为{len(chapter_bullets)}个部分。"
                         + "".join(f"第{i+1}部分讲述{b.split('—')[0].strip()}。"
                                   for i, b in enumerate(chapter_bullets[:5])),
                "takeaway": "",
            }
            slides.insert(insert_pos, agenda)
            patched = True
            print(f"[Step2] 自动补充: 目录页 (agenda, {len(chapter_bullets)}个章节)", flush=True)

    # ── 3. 检查结束页 ──
    if "end" not in [s.get("layout") for s in slides]:
        # 最后一页不是 end，补充
        max_id_num = 0
        for s in slides:
            sid = s.get("id", "")
            try:
                num = int(sid.replace("s", "").replace("_agenda", "99").replace("_end", "99"))
                if num > max_id_num:
                    max_id_num = num
            except (ValueError, AttributeError):
                pass

        end = {
            "id": f"s{max_id_num + 1:02d}",
            "layout": "end",
            "title": "感谢聆听",
            "subtitle": "THANK YOU",
            "bullets": [],
            "visual": None,
            "notes": f"以上就是{title}的全部内容。感谢各位的耐心聆听，欢迎提问和交流。",
            "takeaway": "",
        }
        slides.append(end)
        patched = True
        print("[Step2] 自动补充: 结束页 (end)", flush=True)

    # ── 4. 重新编号 ID（确保不重复） ──
    if patched:
        _renumber_slide_ids(slides)
        slide_plan["slides"] = slides
        slide_plan["meta"]["total_slides"] = len(slides)

    return slide_plan


def _renumber_slide_ids(slides: list):
    """重新编号 slide id，确保连续且不重复。"""
    for i, s in enumerate(slides):
        s["id"] = f"s{i + 1:02d}"

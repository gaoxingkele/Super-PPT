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
from config import CHARS_PER_SLIDE, DATA_POINT_BONUS, MAX_PAGES_PER_CHAPTER
from src.llm_client import chat
from src.prompts.outline import (
    OUTLINE_SYSTEM_PROMPT, build_outline_user_prompt,
    OUTLINE_SKELETON_PROMPT, OUTLINE_DETAIL_PROMPT,
    build_skeleton_user_prompt, build_detail_user_prompt,
    CHAPTER_SKELETON_SYSTEM_PROMPT, build_chapter_skeleton_user_prompt,
)


# ── 每批次处理的最大页面数 ──
DETAIL_BATCH_SIZE = 8


def estimate_slides(analysis: dict) -> tuple:
    """
    根据 analysis.json 内容自动推算 PPT 页数。

    Returns:
        (slide_range, per_chapter_targets)
        - slide_range: (min_pages, max_pages)
        - per_chapter_targets: {"ch01": 5, "ch02": 3, ...}
    """
    chapters = analysis.get("chapters", [])
    if not chapters:
        return (15, 20), {}

    base_pages = 4  # cover + agenda + summary + end
    section_breaks = min(len(chapters), 7)

    content_pages = 0
    per_chapter_targets = {}

    for ch in chapters:
        ch_id = ch.get("id", f"ch{len(per_chapter_targets)+1:02d}")

        # 字数基线：从 summary + key_points 估算内容量
        text_len = len(ch.get("summary", ""))
        for kp in ch.get("key_points", []):
            text_len += len(kp)
        raw = text_len / CHARS_PER_SLIDE

        # 数据点加成
        data_bonus = len(ch.get("data_points", [])) * DATA_POINT_BONUS

        # 权重调整
        weight = ch.get("weight", 3)
        weighted = (raw + data_bonus) * (weight / 3)

        # 限幅
        pages = max(1, min(MAX_PAGES_PER_CHAPTER, round(weighted)))
        per_chapter_targets[ch_id] = pages
        content_pages += pages

    total = base_pages + section_breaks + content_pages
    slide_range = (max(10, total - 2), total + 3)

    print(f"[Step2] 自动页数推算: {total} 页 (范围 {slide_range[0]}~{slide_range[1]}), "
          f"{len(chapters)} 章, 内容页 {content_pages}", flush=True)

    return slide_range, per_chapter_targets


def build_global_blueprint(analysis: dict, per_chapter_targets: dict) -> dict:
    """
    生成全局蓝图：预分配 slide ID、固定页、每章配置。
    纯规则计算，不调 LLM。

    Args:
        analysis: Step1 输出的 analysis.json
        per_chapter_targets: estimate_slides() 输出的每章页数分配

    Returns:
        global_blueprint dict
    """
    chapters = analysis.get("chapters", [])
    title = analysis.get("title", "演示文稿")
    rhythm_plan = analysis.get("rhythm_plan", [])

    # 计算总页数
    section_break_count = min(len(chapters), 7)
    content_pages = sum(per_chapter_targets.values())
    total = 4 + section_break_count + content_pages  # cover+agenda+summary+end + breaks + content

    # 预分配 slide ID
    sid = 1
    fixed_pages = [
        {"id": f"s{sid:02d}", "layout": "cover", "title": title},
    ]
    sid += 1
    fixed_pages.append({"id": f"s{sid:02d}", "layout": "agenda", "title": "目录"})
    sid += 1

    # 每章的蓝图
    chapter_blueprints = []
    rhythm_idx = 0

    for ch in chapters:
        ch_id = ch.get("id", f"ch{len(chapter_blueprints)+1:02d}")
        ch_title = ch.get("title", "")
        target = per_chapter_targets.get(ch_id, 2)

        # section_break ID
        sb_id = f"s{sid:02d}"
        sid += 1

        # 内容页 ID 范围
        slide_ids = [f"s{sid + i:02d}" for i in range(target)]
        sid += target

        # rhythm hint: 从全局 rhythm_plan 切片
        hint_parts = []
        for _ in range(target):
            if rhythm_idx < len(rhythm_plan):
                hint_parts.append(rhythm_plan[rhythm_idx])
                rhythm_idx += 1
            else:
                hint_parts.append("dense" if len(hint_parts) % 3 != 0 else "light")
        rhythm_hint = "-".join(hint_parts)

        chapter_blueprints.append({
            "chapter_id": ch_id,
            "chapter_title": ch_title,
            "section_break_id": sb_id,
            "target_content_pages": target,
            "slide_id_range": slide_ids,
            "weight": ch.get("weight", 3),
            "rhythm_hint": rhythm_hint,
        })

    # 尾部固定页
    summary_id = f"s{sid:02d}"
    sid += 1
    end_id_page = f"s{sid:02d}"

    fixed_pages.append({"id": summary_id, "layout": "summary", "title": "核心结论"})
    fixed_pages.append({"id": end_id_page, "layout": "end", "title": "谢谢观看"})

    # 提取配色方案（从 analysis 的 content_type 推断默认值）
    content_type = analysis.get("content_type", "")
    color_scheme = _default_color_scheme(content_type)

    blueprint = {
        "title": title,
        "total_slides": total,
        "fixed_pages": fixed_pages,
        "chapters": chapter_blueprints,
        "style_rules": {
            "color_scheme": color_scheme,
            "layout_distribution": "每章至少1个data_chart或infographic",
            "max_consecutive_dense": 3,
        },
    }

    print(f"[Step2] 全局蓝图: {total} 页, {len(chapter_blueprints)} 章", flush=True)
    return blueprint


def _default_color_scheme(content_type: str) -> dict:
    """根据内容类型返回默认配色方案。"""
    schemes = {
        "industry_report": {"primary": "#1B365D", "secondary": "#4A90D9", "accent": "#E8612D", "background": "#FFFFFF", "text": "#333333"},
        "academic_defense": {"primary": "#002060", "secondary": "#0060A8", "accent": "#C00000", "background": "#FFFFFF", "text": "#333333"},
        "competition_pitch": {"primary": "#2D3436", "secondary": "#6C5CE7", "accent": "#E17055", "background": "#FFFFFF", "text": "#333333"},
    }
    return schemes.get(content_type, {"primary": "#002060", "secondary": "#0060A8", "accent": "#C00000", "background": "#FFFFFF", "text": "#333333"})


def run_outline(base: str, output_dir: Path, style_profile: dict = None,
                slide_range: tuple = None, two_phase: bool = True) -> dict:
    """
    Step2 入口：生成幻灯片大纲。
    """
    analysis = json.loads((output_dir / "analysis.json").read_text(encoding="utf-8"))
    raw_content = (output_dir / "raw_content.md").read_text(encoding="utf-8", errors="replace")

    # 自动页数推算（用户未指定 --slides 时）
    per_chapter_targets = None
    if slide_range is None and two_phase:
        slide_range, per_chapter_targets = estimate_slides(analysis)
    elif slide_range is None:
        slide_range = DEFAULT_SLIDE_RANGE

    if two_phase:
        slide_plan = _run_two_phase(analysis, raw_content, style_profile, slide_range,
                                     output_dir, per_chapter_targets)
    else:
        slide_plan = _run_single_phase(analysis, raw_content, style_profile, slide_range)

    slide_plan = _ensure_structural_pages(slide_plan, analysis)

    plan_path = output_dir / "slide_plan.json"
    plan_path.write_text(json.dumps(slide_plan, ensure_ascii=False, indent=2), encoding="utf-8")

    slide_count = len(slide_plan.get("slides", []))
    visual_count = sum(1 for s in slide_plan.get("slides", []) if s.get("visual"))
    print(f"[Step2] 大纲生成完成: {slide_count} 张幻灯片, {visual_count} 个视觉元素", flush=True)

    return {"slide_plan_path": plan_path, "slide_plan": slide_plan}


# ============================================================
# 分章骨架生成（Phase A 改造版）
# ============================================================
def _run_chapter_skeleton_phase(analysis: dict, blueprint: dict,
                                 raw_content: str, output_dir: Path) -> list:
    """
    按章节分批生成骨架（Phase A 改造版）。
    每章独立调用一次 LLM，用全局蓝图协调。
    """
    chapter_map = _build_chapter_content_map(analysis, raw_content)
    chapters_analysis = {ch["id"]: ch for ch in analysis.get("chapters", []) if "id" in ch}
    bp_chapters = blueprint["chapters"]

    all_skeletons = []

    # 固定页：cover + agenda（蓝图前两个）
    for fp in blueprint["fixed_pages"]:
        if fp["layout"] in ("cover", "agenda"):
            all_skeletons.append({
                "id": fp["id"],
                "layout": fp["layout"],
                "title": fp["title"],
                "chapter_ref": "",
                "rhythm": "light",
                "visual_type": "generate-image" if fp["layout"] == "cover" else "null",
                "design_intent": "封面" if fp["layout"] == "cover" else "章节导航",
            })

    for ch_idx, ch_bp in enumerate(bp_chapters):
        ch_id = ch_bp["chapter_id"]
        print(f"[Step2-A] 章节 {ch_idx+1}/{len(bp_chapters)}: {ch_bp['chapter_title'][:30]} "
              f"({ch_bp['target_content_pages']} 页)...", flush=True)

        # section_break
        all_skeletons.append({
            "id": ch_bp["section_break_id"],
            "layout": "section_break",
            "title": ch_bp["chapter_title"],
            "subtitle": f"PART {ch_idx+1:02d}",
            "chapter_ref": ch_id,
            "rhythm": "light",
            "visual_type": "generate-image",
            "design_intent": "章节过渡，营造节奏感",
        })

        # 构建上下文
        adj = {}
        if ch_idx > 0:
            adj["prev"] = bp_chapters[ch_idx - 1]["chapter_title"]
        if ch_idx < len(bp_chapters) - 1:
            adj["next"] = bp_chapters[ch_idx + 1]["chapter_title"]

        global_context = {
            "ppt_title": blueprint["title"],
            "total_slides": blueprint["total_slides"],
            "adjacent_chapters": adj,
            "style_rules": blueprint["style_rules"],
        }

        chapter_content = chapter_map.get(ch_id, "")
        analysis_ch = chapters_analysis.get(ch_id, {})

        user_prompt = build_chapter_skeleton_user_prompt(
            chapter_content, ch_bp, global_context, analysis_ch
        )

        response = chat(
            [
                {"role": "system", "content": CHAPTER_SKELETON_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            max_tokens=4096,
            temperature=0.4,
        )

        batch_slides = _parse_json_response(response)
        if isinstance(batch_slides, dict):
            batch_slides = batch_slides.get("slides", [batch_slides])
        if not isinstance(batch_slides, list):
            batch_slides = [batch_slides]

        # 确保 chapter_ref 正确
        for s in batch_slides:
            if isinstance(s, dict):
                s["chapter_ref"] = ch_id
                all_skeletons.append(s)

    # 固定页：summary + end
    for fp in blueprint["fixed_pages"]:
        if fp["layout"] in ("summary", "end"):
            all_skeletons.append({
                "id": fp["id"],
                "layout": fp["layout"],
                "title": fp["title"],
                "chapter_ref": "",
                "rhythm": "light",
                "visual_type": "infographics" if fp["layout"] == "summary" else "generate-image",
                "design_intent": "总结回顾" if fp["layout"] == "summary" else "致谢收尾",
            })

    print(f"[Step2-A] 分章骨架完成: 共 {len(all_skeletons)} 页", flush=True)
    return all_skeletons


# ============================================================
# 两阶段模式
# ============================================================
def _run_two_phase(analysis: dict, raw_content: str,
                   style_profile: dict, slide_range: tuple,
                   output_dir: Path, per_chapter_targets: dict = None) -> dict:
    """Phase A 结构编排 + Phase B 逐页设计。"""

    # ── 判断是否使用分章模式 ──
    use_chapter_batch = per_chapter_targets is not None and len(per_chapter_targets) > 0

    if use_chapter_batch:
        # ── Phase A (分章模式): 按章节逐个生成骨架 ──
        blueprint = build_global_blueprint(analysis, per_chapter_targets)

        # 保存蓝图
        (output_dir / "global_blueprint.json").write_text(
            json.dumps(blueprint, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        slides = _run_chapter_skeleton_phase(analysis, blueprint, raw_content, output_dir)
    else:
        # ── Phase A (原有模式): 一次性生成全部骨架 ──
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
        skeleton_data = _parse_json_response(skeleton_response)
        slides = skeleton_data.get("slides", [])

    # 保存骨架中间结果
    skeleton = {"meta": {}, "slides": slides}
    if use_chapter_batch:
        skeleton["meta"] = {
            "title": analysis.get("title", ""),
            "total_slides": len(slides),
            "color_scheme": blueprint["style_rules"]["color_scheme"],
        }

    (output_dir / "slide_skeleton.json").write_text(
        json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"[Step2-A] 骨架完成: {len(slides)} 页", flush=True)

    if not slides:
        print("[Step2] 骨架为空，回退到单次模式", flush=True)
        return _run_single_phase(analysis, raw_content, style_profile, slide_range)

    # ── Phase B: 按章节批次填充内容（不变） ──
    print("[Step2-B] 正在逐批填充页面内容...", flush=True)

    batches = _group_slides_into_batches(slides)
    chapter_map = _build_chapter_content_map(analysis, raw_content)

    filled_slides = {}

    for batch_idx, batch in enumerate(batches, 1):
        chapter_refs = set(s.get("chapter_ref") or "" for s in batch)
        chapter_content_parts = []
        for ref in sorted(chapter_refs):
            if ref and ref in chapter_map:
                chapter_content_parts.append(chapter_map[ref])

        if not chapter_content_parts:
            chapter_content = raw_content[:4000]
        else:
            chapter_content = "\n\n".join(chapter_content_parts)
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

        if isinstance(detail_slides, dict):
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
            merged = {**skel_slide, **filled_slides[sid]}
            merged_slides.append(merged)
        else:
            merged_slides.append(skel_slide)

    for s in merged_slides:
        if "bullets" not in s:
            s["bullets"] = []
        if "notes" not in s:
            s["notes"] = ""
        if "takeaway" not in s:
            s["takeaway"] = ""
        if "visual" not in s and s.get("visual_type") and s["visual_type"] != "null":
            s["visual"] = {"type": s["visual_type"]}

    slide_plan = {
        "meta": skeleton.get("meta", {}),
        "slides": merged_slides,
    }
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

# 自动页数推算 + 按章节分批骨架生成 实施计划

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 根据文档内容量自动推算 PPT 页数，并按章节分批调用 LLM 生成骨架，降低上下文压力。

**Architecture:** 在 Step1→Step2 之间插入页数推算 + 全局蓝图生成（纯规则），然后改造 Phase A 为逐章 LLM 调用。Phase B 及后续步骤不变。

**Tech Stack:** Python, json, 现有 LLM client (`src/llm_client.chat`)

**Spec:** `docs/specs/2026-03-18-auto-pagecount-batch-skeleton-design.md`

---

## Chunk 1: 配置项 + 页数推算函数

### Task 1: 新增配置项

**Files:**
- Modify: `config.py:100-106`

- [ ] **Step 1: 在 config.py 的 PPT 配置区末尾新增3个常量**

```python
# ============ 页数推算 ============
CHARS_PER_SLIDE = 1500           # 每页对应的原文字数基线
DATA_POINT_BONUS = 0.8           # 每个数据点额外加的页数
MAX_PAGES_PER_CHAPTER = 8        # 单章最大内容页数
```

插入位置：`config.py` 第106行 `CHART_FIGSIZE` 之后。

- [ ] **Step 2: 验证 config 可导入**

Run: `cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "from config import CHARS_PER_SLIDE, DATA_POINT_BONUS, MAX_PAGES_PER_CHAPTER; print(CHARS_PER_SLIDE, DATA_POINT_BONUS, MAX_PAGES_PER_CHAPTER)"`

Expected: `1500 0.8 8`

- [ ] **Step 3: Commit**

```bash
git add config.py
git commit -m "feat: add page estimation config constants"
```

---

### Task 2: 实现 `estimate_slides()`

**Files:**
- Modify: `src/step2_outline.py` (在文件顶部 import 区和函数区新增)

- [ ] **Step 1: 在 `src/step2_outline.py` 顶部 import 区新增导入**

在第12行 `from config import OUTLINE_CONTENT_LIMIT, DEFAULT_SLIDE_RANGE` 后追加：

```python
from config import CHARS_PER_SLIDE, DATA_POINT_BONUS, MAX_PAGES_PER_CHAPTER
```

- [ ] **Step 2: 在 `DETAIL_BATCH_SIZE = 8` (第22行) 之后新增 `estimate_slides` 函数**

```python
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
```

- [ ] **Step 3: 验证函数可导入**

Run: `cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "from src.step2_outline import estimate_slides; print('ok')"`

Expected: `ok`

- [ ] **Step 4: 用实际 analysis.json 测试**

Run:
```bash
cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "
import sys, io, json; sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, '.')
from src.step2_outline import estimate_slides
analysis = json.load(open('output/japan_defense_50pages_v2/analysis.json', encoding='utf-8'))
slide_range, targets = estimate_slides(analysis)
print(f'slide_range={slide_range}')
for k,v in targets.items():
    print(f'  {k}: {v} pages')
print(f'total content pages: {sum(targets.values())}')
"
```

Expected: 输出合理的页数范围（应接近 40-50 页）和每章分配。

- [ ] **Step 5: Commit**

```bash
git add src/step2_outline.py
git commit -m "feat: add estimate_slides() for auto page count estimation"
```

---

## Chunk 2: 全局蓝图生成

### Task 3: 实现 `build_global_blueprint()`

**Files:**
- Modify: `src/step2_outline.py` (在 `estimate_slides` 之后新增)

- [ ] **Step 1: 在 `estimate_slides()` 函数之后新增 `build_global_blueprint()`**

```python
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
        start_id = f"s{sid:02d}"
        end_id = f"s{sid + target - 1:02d}"
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

    fixed_pages.append({"id": summary_id, "layout": "summary", "title": f"核心结论"})
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
```

- [ ] **Step 2: 用实际数据测试蓝图生成**

Run:
```bash
cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "
import sys, io, json; sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, '.')
from src.step2_outline import estimate_slides, build_global_blueprint
analysis = json.load(open('output/japan_defense_50pages_v2/analysis.json', encoding='utf-8'))
slide_range, targets = estimate_slides(analysis)
bp = build_global_blueprint(analysis, targets)
print(f'total_slides: {bp[\"total_slides\"]}')
print(f'fixed_pages: {len(bp[\"fixed_pages\"])}')
for ch in bp['chapters']:
    print(f'  {ch[\"chapter_id\"]}: {ch[\"chapter_title\"][:20]} -> {ch[\"target_content_pages\"]} pages, ids={ch[\"slide_id_range\"][0]}~{ch[\"slide_id_range\"][-1]}')
"
```

Expected: 合理的蓝图，slide ID 连续无重复，每章 1-8 页。

- [ ] **Step 3: Commit**

```bash
git add src/step2_outline.py
git commit -m "feat: add build_global_blueprint() for chapter-based ID pre-allocation"
```

---

## Chunk 3: 单章骨架 Prompt + Phase A 改造

### Task 4: 新增单章骨架 Prompt

**Files:**
- Modify: `src/prompts/outline.py` (在文件末尾新增)

- [ ] **Step 1: 在 `src/prompts/outline.py` 末尾新增系统 prompt 和 user prompt 构建函数**

```python
# ============================================================
# Phase A-Batch: 按章节生成骨架
# ============================================================
CHAPTER_SKELETON_SYSTEM_PROMPT = """你是一位顶级的演示文稿架构师。
你正在为一份大型PPT的**某一个章节**设计结构骨架。
你会收到：本章的原文摘要、目标页数、可用的 slide ID 段、全局上下文。
你只需输出本章的 slides 骨架（JSON数组），不包含 meta。

## 输出格式
[
  {
    "id": "s05",
    "layout": "title_content|data_chart|infographic|two_column|key_insight|table|quote|methodology|architecture",
    "title": "断言式标题（结论句，非主题标签）",
    "chapter_ref": "ch01",
    "rhythm": "dense|light",
    "visual_type": "generate-image|matplotlib|infographics|null",
    "design_intent": "一句话说明设计目的"
  }
]

## 规则
1. 必须使用指定的 slide ID（从 slide_id_range 中取）
2. 必须恰好产出 target_pages 个 slide
3. 标题必须是断言句（可验证的结论），不是主题标签
4. 连续不超过3页 dense（title_content/data_chart/table），之后必须有 light 页
5. 至少 60% 的页面有 visual_type（非 null）
6. 本章的第一页应呈现该章核心结论
7. layout 选择需多样化，不要全部用 title_content

## 布局类型（内容页可用）
- title_content: 常规内容页（左文右图）
- data_chart: 数据图表页（有具体数值时用）
- infographic: 概念/流程/对比页
- two_column: 对比/双方观点页
- key_insight: 关键发现/核心数据页
- table: 表格数据页
- quote: 引用/大字强调页
- methodology: 方法/技术路线页
- architecture: 系统架构/框架页"""


def build_chapter_skeleton_user_prompt(
    chapter_content: str,
    chapter_blueprint: dict,
    global_context: dict,
    analysis_chapter: dict,
) -> str:
    """构建单章骨架生成的 user prompt。"""
    import json

    ch = chapter_blueprint
    parts = []

    parts.append(f"## 全局信息")
    parts.append(f"PPT标题: {global_context['ppt_title']}")
    parts.append(f"PPT总页数: {global_context['total_slides']}")

    # 前后章标题（衔接参考）
    adj = global_context.get("adjacent_chapters", {})
    if adj.get("prev"):
        parts.append(f"上一章: {adj['prev']}")
    if adj.get("next"):
        parts.append(f"下一章: {adj['next']}")

    parts.append(f"\n## 本章配置")
    parts.append(f"章节ID: {ch['chapter_id']}")
    parts.append(f"章节标题: {ch['chapter_title']}")
    parts.append(f"目标页数: {ch['target_content_pages']}")
    parts.append(f"可用 slide ID: {json.dumps(ch['slide_id_range'])}")
    parts.append(f"节奏提示: {ch['rhythm_hint']}")
    parts.append(f"权重: {ch['weight']}")

    # 章节分析数据
    if analysis_chapter:
        parts.append(f"\n## 章节分析")
        parts.append(f"摘要: {analysis_chapter.get('summary', '')}")
        data_points = analysis_chapter.get("data_points", [])
        if data_points:
            parts.append(f"数据点 ({len(data_points)}):")
            for dp in data_points[:10]:
                parts.append(f"  - {dp}")
        concepts = analysis_chapter.get("concepts", [])
        if concepts:
            parts.append(f"核心概念: {', '.join(concepts[:8])}")
        key_points = analysis_chapter.get("key_points", [])
        if key_points:
            parts.append(f"关键要点:")
            for kp in key_points:
                parts.append(f"  - {kp}")

    # 原文内容
    content_limit = 6000
    content = chapter_content[:content_limit] if len(chapter_content) > content_limit else chapter_content
    parts.append(f"\n## 本章原文内容\n{content}")

    # 风格规则
    rules = global_context.get("style_rules", {})
    if rules:
        parts.append(f"\n## 风格规则")
        parts.append(f"配色: {json.dumps(rules.get('color_scheme', {}), ensure_ascii=False)}")
        parts.append(f"布局要求: {rules.get('layout_distribution', '')}")
        parts.append(f"最大连续密集页: {rules.get('max_consecutive_dense', 3)}")

    parts.append(f"\n请输出本章的 {ch['target_content_pages']} 个 slide 骨架（JSON数组），使用 ID: {ch['slide_id_range'][0]} ~ {ch['slide_id_range'][-1]}。")

    return "\n".join(parts)
```

- [ ] **Step 2: 验证 prompt 可导入**

Run: `cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "from src.prompts.outline import CHAPTER_SKELETON_SYSTEM_PROMPT, build_chapter_skeleton_user_prompt; print('ok')"`

Expected: `ok`

- [ ] **Step 3: Commit**

```bash
git add src/prompts/outline.py
git commit -m "feat: add chapter-level skeleton prompt for batched Phase A"
```

---

### Task 5: 改造 Phase A 为逐章循环

**Files:**
- Modify: `src/step2_outline.py:70-100` (改造 `_run_two_phase`)

- [ ] **Step 1: 在 `src/step2_outline.py` 顶部 import 区追加新 prompt 导入**

修改第14-18行，追加导入：

```python
from src.prompts.outline import (
    OUTLINE_SYSTEM_PROMPT, build_outline_user_prompt,
    OUTLINE_SKELETON_PROMPT, OUTLINE_DETAIL_PROMPT,
    build_skeleton_user_prompt, build_detail_user_prompt,
    CHAPTER_SKELETON_SYSTEM_PROMPT, build_chapter_skeleton_user_prompt,
)
```

- [ ] **Step 2: 新增 `_run_chapter_skeleton_phase()` 函数**

在 `_run_two_phase` 函数之前（约第67行）新增：

```python
def _run_chapter_skeleton_phase(analysis: dict, blueprint: dict,
                                 raw_content: str, output_dir: Path) -> list:
    """
    按章节分批生成骨架（Phase A 改造版）。
    每章独立调用一次 LLM，用全局蓝图协调。
    """
    chapter_map = _build_chapter_content_map(analysis, raw_content)
    chapters_analysis = {ch["id"]: ch for ch in analysis.get("chapters", [])}
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
                "design_intent": "封面/目录" if fp["layout"] == "cover" else "章节导航",
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
```

- [ ] **Step 3: 改造 `_run_two_phase()` 函数，增加分章模式**

替换 `_run_two_phase` 函数（第70-185行）为：

```python
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
        import json as _json
        (output_dir / "global_blueprint.json").write_text(
            _json.dumps(blueprint, ensure_ascii=False, indent=2), encoding="utf-8"
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
```

- [ ] **Step 4: 修改 `run_outline()` 入口，自动推算页数并传递**

替换 `run_outline` 函数（约第25-64行）：

```python
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
```

- [ ] **Step 5: 验证完整导入无报错**

Run: `cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "from src.step2_outline import run_outline, estimate_slides, build_global_blueprint; print('all imports ok')"`

Expected: `all imports ok`

- [ ] **Step 6: Commit**

```bash
git add src/step2_outline.py src/prompts/outline.py
git commit -m "feat: chapter-based Phase A skeleton generation with global blueprint"
```

---

## Chunk 4: CLI 集成 + 端到端验证

### Task 6: 确认 main.py 无需改动

**Files:**
- Review: `main.py:95-102`

- [ ] **Step 1: 确认 main.py 已兼容**

当前 `main.py` 第95-102行：
```python
slide_range = _parse_slide_range(getattr(args, "slides", None))
...
run_outline(base, output_dir, style_profile, slide_range)
```

当用户不传 `--slides` 时，`slide_range = None`，传入 `run_outline`。
`run_outline` 内部会自动调用 `estimate_slides()` → 无需改 main.py。

验证：阅读确认逻辑正确即可，不需要代码改动。

---

### Task 7: 端到端集成测试

**Files:**
- 无新文件

- [ ] **Step 1: 用已有的 analysis.json 跑完整 Step2**

Run:
```bash
cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "
import sys, io, json; sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, '.')
from pathlib import Path
from src.step2_outline import run_outline
result = run_outline('japan_defense_50pages_v2', Path('output/japan_defense_50pages_v2'))
print(f'slides: {len(result[\"slide_plan\"][\"slides\"])}')
"
```

Expected:
- 输出 `[Step2] 自动页数推算: XX 页 (范围 XX~XX)`
- 输出 `[Step2] 全局蓝图: XX 页, X 章`
- 输出多行 `[Step2-A] 章节 N/M: ...`
- 输出多行 `[Step2-B] 批次 N/M: ...`
- 最终 `[Step2] 大纲生成完成: XX 张`
- `output/japan_defense_50pages_v2/global_blueprint.json` 已生成
- `output/japan_defense_50pages_v2/slide_skeleton.json` 已生成
- `output/japan_defense_50pages_v2/slide_plan.json` 已更新

- [ ] **Step 2: 验证生成的 slide_plan 可正常构建 PPTX**

Run:
```bash
cd D:/BaiduSyncdisk/aicoding/Super-PPT && python -c "
import sys, io; sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, '.')
from pathlib import Path
from src.step4_build import run_build
result = run_build('japan_defense_50pages_v2', Path('output/japan_defense_50pages_v2'))
print(f'PPTX: {result[\"pptx_path\"]}')
"
```

Expected: PPTX 构建成功，无 WARN。

- [ ] **Step 3: 最终 Commit + Push**

```bash
git add -A
git commit -m "feat: complete auto page estimation + chapter-based skeleton batching"
git push origin main
```

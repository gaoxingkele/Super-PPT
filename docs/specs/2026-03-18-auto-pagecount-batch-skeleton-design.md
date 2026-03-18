# 设计文档：自动页数推算 + 按章节分批骨架生成

> 日期：2026-03-18
> 状态：已批准，待实施

## 背景与目标

当前 Super-PPT 管线的页数需要用户手动指定 `--slides 48-52`，或完全交给 LLM 自行决定（不可控）。骨架生成（Phase A）是一次 LLM 调用生成全部页面结构，对于长文档（100+ 页）上下文压力大。

**目标：**
1. 根据原始文档的内容量、章节数、数据密度自动推算合理页数
2. 按章节分批调用 LLM 生成骨架，降低单次上下文压力，支持超长文档

## 核心设计决策

| 决策点 | 选择 | 理由 |
|--------|------|------|
| 页数推算策略 | 规则公式算基线 + LLM 在范围内自主决定 | 纯规则易翻车，纯 LLM 不稳定，组合最稳 |
| 骨架分批粒度 | 按章节切（每章一次 Phase A 调用） | 章节是天然叙事单元，section_break 天然隔开 |
| 跨批协调 | 全局蓝图（纯规则，不调 LLM） | 预分配 slide ID、rhythm hint，保证一致性 |

## 模块设计

### 1. 页数推算函数 `estimate_slides()`

位置：`src/step2_outline.py`

```python
def estimate_slides(analysis: dict) -> tuple[tuple[int,int], dict]:
    """
    输入：analysis.json 内容
    输出：(slide_range, per_chapter_targets)
      - slide_range: (min, max) 总页数范围
      - per_chapter_targets: {chapter_id: target_pages}
    """
```

算法：
```
base_pages = 4  (封面1 + 目录1 + 总结1 + 结尾1)
section_breaks = min(len(chapters), 7)

content_pages = 0
per_chapter_targets = {}
for ch in chapters:
    raw = len(ch.content) / CHARS_PER_SLIDE         # 字数基线 (默认1500)
    data_bonus = len(ch.data_points) * DATA_POINT_BONUS  # 默认0.8
    weighted = raw * (ch.weight / 3)                  # weight=5 放大
    pages = clamp(weighted + data_bonus, min=1, max=MAX_PAGES_PER_CHAPTER)
    per_chapter_targets[ch.id] = round(pages)
    content_pages += round(pages)

total = base_pages + section_breaks + content_pages
slide_range = (total - 2, total + 3)
```

可调参数（config.py）：
- `CHARS_PER_SLIDE = 1500`
- `DATA_POINT_BONUS = 0.8`
- `MAX_PAGES_PER_CHAPTER = 8`

### 2. 全局蓝图 `build_global_blueprint()`

位置：`src/step2_outline.py`

纯规则计算，不调 LLM，毫秒级完成。

```python
def build_global_blueprint(analysis: dict, per_chapter_targets: dict) -> dict:
    """
    输出 global_blueprint:
    {
      "title": "PPT标题",
      "total_slides": 42,
      "fixed_pages": [
        {"id": "s01", "layout": "cover"},
        {"id": "s02", "layout": "agenda"},
        {"id": "sN-1", "layout": "summary"},
        {"id": "sN", "layout": "end"}
      ],
      "chapters": [
        {
          "chapter_id": "ch01",
          "chapter_title": "章节标题",
          "section_break_id": "s03",
          "target_content_pages": 5,
          "slide_id_range": ["s04", "s08"],
          "weight": 5,
          "rhythm_hint": "dense-light-dense-light-dense"
        },
        ...
      ],
      "style_rules": {
        "color_scheme": {...},
        "layout_distribution": "每章至少1个data_chart + 1个infographic",
        "max_consecutive_dense": 3
      }
    }
    """
```

关键点：
- **slide_id 预分配**：每章提前划好 ID 段，避免分批后 ID 冲突
- **rhythm_hint**：根据 analysis.json 的 rhythm_plan 拆到每章
- 持久化为 `output/{base}/global_blueprint.json`

### 3. 按章节分批骨架生成（改造 Phase A）

位置：`src/step2_outline.py` 的 `_run_skeleton_phase()`

```python
# 改造后的流程
def _run_skeleton_phase(analysis, blueprint, ...):
    all_skeletons = list(blueprint["fixed_pages"])  # 固定页

    for chapter in blueprint["chapters"]:
        # section_break 页
        all_skeletons.append({
            "id": chapter["section_break_id"],
            "layout": "section_break",
            "title": chapter["chapter_title"]
        })

        # 每章独立调用一次 LLM
        skeleton_batch = call_llm_skeleton(
            chapter_content=该章的原文摘要,
            chapter_blueprint=章节蓝图配置,
            global_context={
                "ppt_title": blueprint["title"],
                "total_slides": blueprint["total_slides"],
                "adjacent_chapters": 前后章标题,
                "style_rules": blueprint["style_rules"]
            },
            slide_ids=chapter["slide_id_range"],
            target_pages=chapter["target_content_pages"]
        )
        all_skeletons.extend(skeleton_batch)

        # 断点续传：每章完成后更新 progress
        save_progress(chapter["chapter_id"], "skeleton_done")

    return all_skeletons
```

### 4. 新增 Prompt `build_chapter_skeleton_prompt()`

位置：`src/prompts/outline.py`

```
系统提示：你是顶级演示文稿架构师，正在为一个 {total_slides} 页 PPT 的第 {chapter_index} 章生成骨架。

用户提示包含：
- 本章原文摘要
- 本章目标页数、可用 slide_id 段
- 本章 rhythm_hint
- 全局风格规则
- 前后章标题（衔接参考）

输出：本章的 slide 骨架列表（同原有格式）
```

原有的 `build_skeleton_system_prompt()` 保留，供 `--no-batch` 兼容模式使用。

## 数据流

```
Step1 → analysis.json
     ↓
estimate_slides(analysis) → slide_range + per_chapter_targets
     ↓
build_global_blueprint(analysis, targets) → global_blueprint.json
     ↓
Phase A (改造): for each chapter
     → call_llm_skeleton(chapter) → chapter_skeleton
     → 断点续传 checkpoint
     ↓
merge all → slide_skeleton.json
     ↓
Phase B (不变): 按8页分批填充 → slide_plan.json
     ↓
Step3 (不变) → assets/
     ↓
Step4 (不变) → .pptx
```

## 文件改动范围

| 文件 | 改动类型 | 说明 |
|------|----------|------|
| `src/step2_outline.py` | **重点改** | 新增 `estimate_slides()`、`build_global_blueprint()`；改造 `_run_skeleton_phase()` 为逐章循环 |
| `src/prompts/outline.py` | **重点改** | 新增 `build_chapter_skeleton_prompt()`；原有 prompt 保留兼容 |
| `config.py` | 微调 | 新增 `CHARS_PER_SLIDE`、`DATA_POINT_BONUS`、`MAX_PAGES_PER_CHAPTER` |
| `main.py` | 微调 | 用户未指定 `--slides` 时调用 `estimate_slides()` |

**不改的文件：**
- `step0_ingest.py`、`step1_analyze.py`、`step3_visuals.py`、`step4_build.py`、`pptx_engine.py`

## 兼容性

- **短文档（≤15页）**：仍走逐章生成，chapters 少开销小
- **用户手动 `--slides`**：覆盖公式结果，按比例重新分配每章页数
- **`--no-batch` 参数**：回退到原有一次性骨架生成
- **断点续传**：`global_blueprint.json` + `progress.json` 记录每章骨架完成状态

## 新增配置项（config.py）

```python
# 页数推算
CHARS_PER_SLIDE = 1500       # 每页对应的原文字数
DATA_POINT_BONUS = 0.8       # 每个数据点额外加的页数
MAX_PAGES_PER_CHAPTER = 8    # 单章最大内容页数
```

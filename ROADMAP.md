# Super-PPT 技术规划路线

> 从任意来源（URL、PDF、Word、Markdown、文件夹）自动生成视觉丰富、专业美观的 PPT。
> 支持参考模板样式（PPTX/PDF/图片），按特定风格和专业度创建。
> 模仿 chatgpt-document 项目架构，多步骤管线，步步为营。

---

## 一、系统全景

```
输入源                    管线步骤                              输出
──────          ──────────────────────────────          ──────
URL         ─┐
PDF         ─┤   Step0    Step1    Step2    Step3    Step4
Word/DOCX   ─┼→ 内容获取 → 结构化 → 幻灯片  → 视觉  → PPTX  → .pptx
Markdown    ─┤           分析      大纲     资产     装配
文件夹      ─┘                    生成     生成

可选输入:
  参考模板 (.pptx/.pdf/.png) → 风格提取 → 注入到 Step2/Step4
```

---

## 二、管线步骤详解

### Step0: 内容获取与统一化 (`src/step0_ingest.py`)

将各种来源统一为纯文本 + 元数据。

| 来源 | 处理方式 | 依赖 |
|------|----------|------|
| **URL** | Pyppeteer/Playwright 爬取，提取正文 | pyppeteer, playwright |
| **PDF** | pypdf 提取文本；pdfplumber 提取表格；pdf2image 提取图片 | pypdf, pdfplumber, pdf2image |
| **DOCX** | python-docx 提取文本/表格/图片 | python-docx |
| **Markdown** | 直接读取，解析前置 YAML 元数据 | - |
| **文件夹** | 递归扫描，按文件类型分别处理后合并 | 复用上述所有 |

**输出**：
```
output/{base}/
├── raw_content.md          # 统一的纯文本内容
├── raw_meta.json           # 元数据 (标题、来源、关键词、章节结构)
├── raw_tables.json         # 提取的表格数据 [{headers, rows, caption}, ...]
└── raw_images/             # 提取的内嵌图片
    ├── img_001.png
    └── ...
```

**关键设计**：
- 表格单独提取为结构化 JSON，后续 Step2 可直接映射为图表数据源
- 图片保留原始分辨率，后续 Step4 可选择复用到 PPT 中
- URL 爬取复用 chatgpt-document 的爬虫能力（Unicode 数学字符处理等）

---

### Step1: 结构化分析 (`src/step1_analyze.py`)

LLM 深度分析内容，产出结构化理解。

**输入**: `raw_content.md` + `raw_meta.json` + `raw_tables.json`

**LLM 任务**：
1. **主题提炼** — 标题、核心论点、目标受众
2. **章节划分** — 将内容分为 3~8 个逻辑章节
3. **关键数据提取** — 识别所有可视化的数据点（数字、趋势、对比、占比）
4. **核心概念提取** — 识别需要图文表达的概念、流程、关系
5. **论证链提取** — 核心观点 → 论据 → 结论的推理链

**输出**: `output/{base}/analysis.json`
```json
{
  "title": "报告标题",
  "subtitle": "副标题建议",
  "audience": "目标受众",
  "core_thesis": "核心论点（一句话）",
  "chapters": [
    {
      "id": "ch01",
      "title": "章节标题",
      "summary": "章节摘要",
      "key_points": ["要点1", "要点2"],
      "data_points": [
        {"type": "trend", "label": "市场规模", "data": {"2023": 120, "2024": 185}, "unit": "亿元"},
        {"type": "comparison", "items": ["方案A", "方案B"], "metrics": {...}}
      ],
      "concepts": [
        {"name": "技术架构", "type": "process_flow", "description": "数据采集→预处理→训练→部署"},
        {"name": "竞争格局", "type": "hierarchy", "description": "..."}
      ],
      "reasoning": {"claim": "...", "evidence": ["..."], "conclusion": "..."}
    }
  ],
  "global_data": [
    {"type": "kpi", "items": [{"label": "准确率", "value": "97.2%"}, ...]}
  ]
}
```

**为什么需要这一步**：
- 不直接从文本生成幻灯片大纲，而是先做深度理解
- 结构化的 data_points / concepts 为后续视觉资产生成提供精确输入
- 分离"理解内容"和"设计幻灯片"两个认知任务，各步骤质量更高

---

### Step2: 幻灯片大纲生成 (`src/step2_outline.py`)

基于 Step1 的结构化分析 + 可选的风格模板，生成带视觉指令的幻灯片大纲。

**输入**: `analysis.json` + 可选的 `style_profile.json`（从参考模板提取）

**LLM 任务**：
1. 规划幻灯片数量和顺序（15~30 张）
2. 为每张幻灯片选择最佳布局和视觉类型
3. 将 data_points 映射为具体图表规格
4. 将 concepts 映射为信息图描述
5. 为 AI 图片生成写 prompt
6. 控制文字量（每页 bullet ≤ 5 条，每条 ≤ 25 字）

**输出**: `output/{base}/slide_plan.json`
```json
{
  "meta": {
    "title": "报告标题",
    "subtitle": "副标题",
    "total_slides": 22,
    "theme": "business",
    "color_scheme": {"primary": "#1B365D", "accent": "#E8612D", "bg": "#FFFFFF"}
  },
  "slides": [
    {
      "id": "s01",
      "layout": "cover",
      "title": "标题",
      "subtitle": "副标题 | 2026年3月",
      "visual": {
        "type": "generate-image",
        "prompt": "abstract professional background, blue gradient, data visualization motif",
        "position": "background"
      },
      "notes": "开场：介绍报告背景和核心问题"
    },
    {
      "id": "s02",
      "layout": "agenda",
      "title": "目录",
      "items": ["第一章: ...", "第二章: ...", "第三章: ..."],
      "visual": null
    },
    {
      "id": "s05",
      "layout": "data_chart",
      "title": "市场规模增长趋势",
      "takeaway": "年复合增长率达 60%",
      "visual": {
        "type": "matplotlib",
        "chart": "bar",
        "data": {"labels": ["2023","2024","2025"], "values": [120,185,310]},
        "ylabel": "市场规模（亿元）",
        "highlight": [2],
        "annotation": {"index": 2, "text": "+68%"}
      },
      "notes": "强调2025年的爆发式增长"
    },
    {
      "id": "s08",
      "layout": "infographic",
      "title": "技术架构全景",
      "visual": {
        "type": "infographics",
        "infographic_type": "process_flow",
        "description": "数据采集→预处理→模型训练→推理部署，四阶段流程图，每阶段标注关键技术栈",
        "style": "tech",
        "data": {
          "stages": [
            {"name": "数据采集", "detail": "API + 爬虫"},
            {"name": "预处理", "detail": "清洗 + 标注"},
            {"name": "模型训练", "detail": "GPU 集群"},
            {"name": "推理部署", "detail": "K8s + TensorRT"}
          ]
        }
      },
      "notes": "解释每个阶段的技术选型理由"
    },
    {
      "id": "s12",
      "layout": "two_column",
      "title": "方案对比分析",
      "left": {"heading": "方案A", "bullets": ["优势1", "优势2"]},
      "right": {"heading": "方案B", "bullets": ["优势1", "优势2"]},
      "visual": {
        "type": "matplotlib",
        "chart": "radar",
        "data": {
          "categories": ["性能","成本","易用性","扩展性","安全性"],
          "series": {"方案A": [8,6,9,7,8], "方案B": [9,4,6,9,7]}
        }
      },
      "notes": "引导听众关注方案A在易用性上的优势"
    },
    {
      "id": "s15",
      "layout": "key_insight",
      "title": "核心发现",
      "quote": "关键结论的一句话提炼",
      "visual": {
        "type": "infographics",
        "infographic_type": "stat_display",
        "description": "3个核心KPI大数字展示",
        "style": "data-focused",
        "data": {
          "kpis": [
            {"label": "准确率", "value": "97.2%", "trend": "up"},
            {"label": "延迟", "value": "<50ms", "trend": "down"},
            {"label": "成本", "value": "-40%", "trend": "down"}
          ]
        }
      },
      "notes": "这是全场最关键的一页"
    },
    {
      "id": "s22",
      "layout": "summary",
      "title": "总结与展望",
      "bullets": ["结论1", "结论2", "结论3"],
      "call_to_action": "下一步行动建议",
      "visual": null
    }
  ]
}
```

**布局类型完整列表**：

| layout | 说明 | 视觉占比 | 适用场景 |
|---|---|---|---|
| `cover` | 封面 | 图片全屏 + 文字叠加 | 第一页 |
| `agenda` | 目录/议程 | 纯文字编号列表 | 第二页 |
| `section_break` | 章节过渡页 | 大标题 + 背景色/图 | 章节切换 |
| `title_content` | 标题 + 要点 | 左文右图 (6:4) | 常规内容页 |
| `data_chart` | 数据图表页 | 图表占 70% + 标注 | 数据展示 |
| `infographic` | 信息图页 | 图片居中占 80% | 流程/架构/概念 |
| `two_column` | 双栏对比 | 左右均分 + 底部图表 | 方案对比 |
| `key_insight` | 核心发现 | 大字引用 + KPI 图 | 关键结论 |
| `table` | 表格页 | 原生 Table 对象 | 详细数据 |
| `image_full` | 全屏图片 | 图片 100% | 效果展示/截图 |
| `quote` | 引用页 | 大字居中 + 来源 | 名言/关键论断 |
| `timeline` | 时间线 | 信息图 80% | 发展历程/里程碑 |
| `summary` | 总结页 | 要点回顾 + CTA | 最后一页 |

---

### Step3: 视觉资产生成 (`src/step3_visuals.py`)

并行生成所有幻灯片所需的视觉资产。

**输入**: `slide_plan.json`

**三条并行渲染管线**：

#### 管线 A: matplotlib/seaborn 统计图表 (`src/visuals/charts.py`)

本地渲染，快速精确。

```python
CHART_RENDERERS = {
    "bar": render_bar,              # 柱状图
    "grouped_bar": render_grouped,  # 分组柱状图
    "stacked_bar": render_stacked,  # 堆叠柱状图
    "line": render_line,            # 折线图
    "area": render_area,            # 面积图
    "pie": render_pie,              # 饼图 / 环形图
    "donut": render_donut,          # 环形图
    "radar": render_radar,          # 雷达图
    "heatmap": render_heatmap,      # 热力图
    "scatter": render_scatter,      # 散点图
    "waterfall": render_waterfall,  # 瀑布图
    "funnel": render_funnel,        # 漏斗图
    "treemap": render_treemap,      # 矩形树图
    "gauge": render_gauge,          # 仪表盘
}
```

**统一渲染规范**：
- 尺寸: 1920×1080px (16:9)，图表区域占 70%
- DPI: 300
- 字体: 中文 SimHei / 英文 Arial
- 配色: 从 `slide_plan.meta.color_scheme` 读取，保持与 PPT 主题一致
- 背景: 透明 PNG
- 数据标注: 自动添加数值标签
- 高亮: 支持 `highlight` 参数突出特定数据

#### 管线 B: infographics skill 信息图 (`src/visuals/infographics.py`)

调用 Claude Code infographics skill (Nano Banana Pro)。

**支持的信息图类型**：
- `process_flow` — 流程图（线性/分支）
- `timeline` — 时间线
- `hierarchy` — 层级结构/组织架构
- `comparison` — 对比图
- `stat_display` — KPI 大数字展示
- `cycle` — 循环图
- `matrix` — 矩阵图
- `network` — 关系网络图
- `pyramid` — 金字塔图
- `venn` — 维恩图

**渲染规范**：
- 分辨率: 4K (3840×2160)
- 风格: 与 PPT 主题匹配 (corporate/tech/academic/minimal)
- 配色: 从 `color_scheme` 传入
- 文字: 中文准确渲染

#### 管线 C: generate-image skill AI 插图 (`src/visuals/ai_images.py`)

调用 FLUX / Nano Banana 2 生成概念性视觉。

**用途**：
- 封面/章节过渡页的背景图
- 抽象概念的视觉隐喻
- 场景/行业插图

**生成策略**：
- prompt 使用英文（效果最佳）
- 添加风格后缀: "professional, clean, modern, high quality"
- 封面图生成 2~3 张备选，LLM 选最佳

**输出**: `output/{base}/assets/`
```
assets/
├── s01_cover.png           # 封面背景
├── s05_chart_bar.png       # matplotlib 柱状图
├── s08_infographic.png     # 信息图
├── s12_chart_radar.png     # matplotlib 雷达图
├── s15_infographic.png     # KPI 信息图
└── manifest.json           # 资产清单 {slide_id: asset_path, status: ok/failed/skipped}
```

**容错机制**：
- AI 图片生成失败 → 降级为纯色背景 + 文字
- 信息图生成失败 → 降级为 matplotlib 替代图表
- manifest.json 记录每个资产状态，支持单独重试

---

### Step4: PPTX 装配 (`src/step4_build.py`)

将幻灯片大纲 + 视觉资产 + 主题模板装配为最终 PPTX。

**输入**: `slide_plan.json` + `assets/` + 主题模板

**装配引擎** (`src/utils/pptx_engine.py`)：

```python
class PPTXBuilder:
    def __init__(self, template_path: Path, color_scheme: dict):
        self.prs = Presentation(template_path)
        self.colors = color_scheme

    def add_slide(self, slide_spec: dict, asset_path: Path | None):
        layout_fn = LAYOUT_HANDLERS[slide_spec["layout"]]
        layout_fn(self, slide_spec, asset_path)

    def save(self, output_path: Path):
        self.prs.save(output_path)
```

**各布局处理器的核心逻辑**：

| 布局 | 处理 |
|------|------|
| `cover` | 设置背景图 → 添加半透明遮罩 → 叠加标题/副标题 |
| `data_chart` | 插入图表 PNG (居中 70%) → 添加 takeaway 文字框 |
| `infographic` | 插入信息图 PNG (居中 80%) → 添加标题 |
| `two_column` | 左右文字框 → 底部或中间插入图表 |
| `table` | 创建原生 Table 对象 → 设置样式 → 填入数据 |
| `key_insight` | 大字号引用 → KPI 图片 → 强调色背景 |

**特殊处理**：
- 公式 `$...$`: matplotlib mathtext → PNG → 行内插入
- 超链接: 保留可点击
- 演讲者备注: `slide_spec.notes` → `slide.notes_slide.notes_text_frame`
- 页码: 自动编号（跳过封面）

---

## 三、参考模板风格系统

### 模板来源支持

| 格式 | 提取方式 | 提取内容 |
|------|----------|----------|
| **.pptx** | python-pptx 解析 | 配色方案、字体、布局结构、母版样式 |
| **.pdf** | pdf2image + LLM 视觉分析 | 配色、排版风格、视觉元素风格 |
| **.png/.jpg** | LLM 视觉分析 | 配色、排版风格、设计语言 |

### 风格提取流程 (`src/style_extractor.py`)

```
参考模板 → 解析/截图 → LLM 视觉分析 → style_profile.json
```

**style_profile.json**：
```json
{
  "source": "template.pptx",
  "color_scheme": {
    "primary": "#1B365D",
    "secondary": "#4A90D9",
    "accent": "#E8612D",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "typography": {
    "title_font": "微软雅黑",
    "body_font": "微软雅黑",
    "title_size_pt": 36,
    "body_size_pt": 18,
    "title_bold": true
  },
  "layout_style": {
    "margin_cm": 2.0,
    "content_alignment": "left",
    "visual_weight": "heavy",
    "decoration": "minimal"
  },
  "design_language": "corporate-modern",
  "extracted_layouts": ["cover", "title_content", "two_column", "section_break"]
}
```

### 内置主题 (`themes/`)

```
themes/
├── business.pptx          # 商务：深蓝渐变，无衬线，干净利落
├── academic.pptx          # 学术：白底，深色标题，衬线体
├── tech.pptx              # 科技：深色背景，亮色强调，几何元素
├── minimal.pptx           # 极简：大留白，黑白+单一强调色
├── consulting.pptx        # 咨询：麦肯锡/BCG 风格，蓝灰色系
└── creative.pptx          # 创意：大胆配色，不规则布局
```

### 风格应用优先级

1. 用户提供参考模板 → 提取 style_profile → 直接作为母版使用（如果是 .pptx）或套用到内置主题
2. 用户指定 `--theme` → 使用内置主题
3. 都没有 → 根据内容自动选择（LLM 在 Step1 根据受众/领域推荐）

---

## 四、项目文件结构

```
Super-PPT/
├── main.py                         # CLI 入口 (argparse 子命令)
├── config.py                       # 配置中心 (env + 默认值)
├── requirements.txt                # 依赖清单
├── .env.example                    # 环境变量模板
├── ROADMAP.md                      # 本文件
│
├── src/
│   ├── __init__.py                 # PROJECT_ROOT + sys.path 设置
│   ├── llm_client.py               # 统一 LLM 调用 (多 provider 支持)
│   │
│   ├── step0_ingest.py             # Step0: 内容获取与统一化
│   ├── step1_analyze.py            # Step1: 结构化分析
│   ├── step2_outline.py            # Step2: 幻灯片大纲生成
│   ├── step3_visuals.py            # Step3: 视觉资产并行生成 (调度器)
│   ├── step4_build.py              # Step4: PPTX 装配
│   │
│   ├── style_extractor.py          # 参考模板风格提取
│   │
│   ├── ingest/                     # 内容获取模块
│   │   ├── __init__.py
│   │   ├── crawlers.py             # URL 爬取 (Pyppeteer/Playwright)
│   │   ├── pdf_reader.py           # PDF 解析 (文本+表格+图片)
│   │   ├── docx_reader.py          # DOCX 解析
│   │   ├── md_reader.py            # Markdown 解析
│   │   └── dir_scanner.py          # 文件夹递归扫描
│   │
│   ├── visuals/                    # 视觉资产生成
│   │   ├── __init__.py
│   │   ├── charts.py               # matplotlib/seaborn 统计图表
│   │   ├── infographics.py         # infographics skill 调度
│   │   ├── ai_images.py            # generate-image skill 调度
│   │   └── formula.py              # LaTeX 公式 → PNG
│   │
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── pptx_engine.py          # python-pptx 装配引擎
│   │   ├── progress.py             # 断点续传 (JSON checkpoint)
│   │   ├── fonts.py                # 中英文字体管理
│   │   └── color_utils.py          # 配色方案工具
│   │
│   └── prompts/                    # LLM 提示词
│       ├── analyze.py              # Step1 分析提示词
│       ├── outline.py              # Step2 大纲生成提示词
│       └── style.py                # 风格提取提示词
│
├── themes/                         # 内置主题母版
│   ├── business.pptx
│   ├── academic.pptx
│   ├── tech.pptx
│   ├── minimal.pptx
│   ├── consulting.pptx
│   └── creative.pptx
│
└── output/                         # 生成结果 (自动创建)
    └── {base}/                     # 每个项目一个子目录
        ├── raw_content.md
        ├── raw_meta.json
        ├── raw_tables.json
        ├── raw_images/
        ├── analysis.json
        ├── slide_plan.json
        ├── style_profile.json
        ├── assets/
        │   ├── s01_cover.png
        │   ├── s05_chart_bar.png
        │   └── manifest.json
        ├── {base}_slides.pptx      # 最终输出
        └── {base}_progress.json    # 断点续传
```

---

## 五、CLI 接口设计

```bash
# ========== 一键生成（全管线） ==========
python main.py generate <source> [options]

# 从 URL
python main.py generate https://example.com/article

# 从文件
python main.py generate report.pdf
python main.py generate report.docx
python main.py generate report.md

# 从文件夹（合并所有文件）
python main.py generate ./research_papers/

# 指定主题
python main.py generate report.md --theme consulting

# 使用参考模板
python main.py generate report.md --template client_brand.pptx
python main.py generate report.md --template reference_style.pdf
python main.py generate report.md --template design_sample.png

# 控制幻灯片数量
python main.py generate report.md --slides 15-20

# 指定 LLM provider
python main.py generate report.md -p gemini

# 跳过 AI 图片（快速预览模式）
python main.py generate report.md --no-ai-images

# 指定输出名称
python main.py generate report.md -o my_presentation

# 断点续传（默认启用）
python main.py generate report.md              # 自动续传
python main.py generate report.md --no-resume  # 强制从头

# ========== 单步执行 ==========
python main.py ingest <source>                         # 仅 Step0
python main.py analyze <base>                          # 仅 Step1
python main.py outline <base> [--template ...]         # 仅 Step2
python main.py visuals <base> [--no-ai-images]         # 仅 Step3
python main.py build <base> [--theme ...]              # 仅 Step4

# ========== 工具命令 ==========
python main.py extract-style <template_file>           # 提取模板风格
python main.py list-themes                             # 列出内置主题
python main.py retry-asset <base> <slide_id>           # 重试单个视觉资产
```

---

## 六、核心技术方案

### 6.1 LLM 抽象层 (`src/llm_client.py`)

复用 chatgpt-document 的多 provider 架构：

```python
# 统一接口
def chat(messages, provider=None, model=None, max_tokens=8192, temperature=0.6) -> str

# 视觉分析（用于模板风格提取）
def chat_vision(messages, image_paths, provider=None) -> str

# 支持的 providers
PROVIDERS = {
    "kimi": {"base_url": "...", "model": "kimi-k2.5"},
    "gemini": {"model": "gemini-2.5-pro"},
    "grok": {"model": "grok-3"},
    "claude": {"model": "claude-opus-4-6"},
    "deepseek": {"model": "deepseek-chat"},
    "openai": {"model": "gpt-4o"},
    # ...
}
```

### 6.2 断点续传 (`src/utils/progress.py`)

```python
# 每个 step 完成后保存进度
save_progress(base, "step0_ingest")
save_progress(base, "step1_analyze")
save_progress(base, "step2_outline")
save_progress(base, "step3_visuals", extra={"completed_assets": ["s01", "s05"]})
save_progress(base, "step4_build")

# Step3 支持资产级细粒度续传
# 已生成的资产不会重新生成，仅补充 failed/missing 的
```

### 6.3 图表渲染统一规范

```python
# src/visuals/charts.py

CHART_DEFAULTS = {
    "figsize": (19.2, 10.8),    # 1920×1080 at 100 DPI
    "dpi": 300,
    "font_family_zh": "SimHei",
    "font_family_en": "Arial",
    "bg_color": "transparent",
    "grid_alpha": 0.3,
    "label_fontsize": 14,
    "title_fontsize": 20,
}

def render_chart(chart_spec: dict, color_scheme: dict, output_path: Path) -> Path:
    """统一入口：根据 chart_spec.chart 类型分发到具体渲染器。"""
    renderer = CHART_RENDERERS[chart_spec["chart"]]
    fig = renderer(chart_spec["data"], color_scheme, chart_spec.get("options", {}))
    fig.savefig(output_path, dpi=300, transparent=True, bbox_inches="tight")
    return output_path
```

### 6.4 公式渲染

```python
# src/visuals/formula.py

def render_latex_to_png(latex: str, output_path: Path, fontsize: int = 24) -> Path:
    """使用 matplotlib mathtext 将 LaTeX 公式渲染为 PNG。"""
    fig, ax = plt.subplots(figsize=(0.01, 0.01))
    ax.text(0, 0, f"${latex}$", fontsize=fontsize,
            transform=ax.transAxes, va="center", ha="center")
    ax.axis("off")
    fig.savefig(output_path, dpi=300, transparent=True, bbox_inches="tight")
    plt.close(fig)
    return output_path
```

---

## 七、集成的 Claude Code Skills

| Skill | 用途 | 对应步骤 |
|---|---|---|
| **pptx** | PPTX 模板解析与高级布局 | Step4, style_extractor |
| **scientific-slides** | 研究演讲设计原则指导 LLM | Step2 prompt |
| **infographics** | Nano Banana Pro 信息图生成 | Step3 |
| **generate-image** | FLUX AI 封面/概念插图生成 | Step3 |
| **scientific-visualization** | 出版级多面板图表编排 | Step3 |
| **matplotlib** | 底层统计图表渲染 | Step3 |
| **seaborn** | 快速统计图表 | Step3 |
| **markitdown** | 文件转 Markdown 预处理 | Step0 |
| **pdf** | PDF 解析与内容提取 | Step0 |

---

## 八、依赖清单

```
# === 核心 ===
python-pptx>=1.0.0            # PPTX 生成
matplotlib>=3.8.0              # 统计图表渲染
seaborn>=0.13.0                # 统计图表（美观默认样式）
Pillow>=10.0.0                 # 图片处理

# === 内容获取 ===
pypdf>=4.0.0                   # PDF 文本提取
pdfplumber>=0.11.0             # PDF 表格提取
pdf2image>=1.17.0              # PDF → 图片（用于风格提取）
python-docx>=1.1.0             # DOCX 解析
pyppeteer>=1.0.2               # URL 爬取（首选）
playwright>=1.40.0             # URL 爬取（备用）

# === LLM ===
openai>=1.0.0                  # OpenAI 兼容 API (Kimi/Grok/DeepSeek/...)
python-dotenv>=1.0.0           # .env 环境变量
httpx>=0.27.0                  # HTTP 客户端
tenacity>=8.2.0                # 重试机制

# === 可选 ===
# anthropic                    # Claude provider
# google-generativeai          # Gemini provider
# squarify                     # treemap 图表
# plotly>=5.18.0               # 交互式图表（可选）
```

---

## 九、质量保障

### 自动检查 (`src/utils/quality_check.py`)

装配完成后自动检查：

| 检查项 | 标准 |
|--------|------|
| 幻灯片数量 | 15~30 张 |
| 每页文字量 | bullet ≤ 5 条，每条 ≤ 30 字 |
| 字号下限 | 标题 ≥ 24pt，正文 ≥ 16pt |
| 图片分辨率 | ≥ 720p |
| 视觉覆盖率 | ≥ 60% 的页面包含视觉元素 |
| 配色一致性 | 所有元素颜色在 color_scheme 范围内 |
| 文件大小 | ≤ 30MB |

### LLM 审阅（可选 Step4b）

生成后用 LLM 视觉能力审阅 PPTX 截图，给出改进建议。

---

## 十、实现路线与优先级

### Phase 1 — MVP（✅ 已完成）
- [x] 项目骨架: main.py + config.py + src/__init__.py
- [x] `step0_ingest.py`: MD/DOCX/PDF 文件读取
- [x] `step1_analyze.py`: LLM 结构化分析
- [x] `step2_outline.py`: LLM 幻灯片大纲生成
- [x] `visuals/charts.py`: matplotlib 基础图表 (bar/line/pie/radar)
- [x] `step4_build.py` + `pptx_engine.py`: 基础 PPTX 装配
- [x] 内置 4 个主题 (各含 16 种 layout)
- [x] CLI: `python main.py generate <file>`

### Phase 2 — 视觉丰富（✅ 已完成）
- [x] `visuals/infographics.py`: 9 种信息图类型
- [x] `visuals/ai_images.py`: Gemini API + 回退方案
- [x] `visuals/charts.py`: 扩展图表类型 (heatmap/waterfall/funnel)
- [x] `visuals/formula.py`: LaTeX 公式渲染
- [x] 资产并行生成 (ThreadPoolExecutor)
- [x] 资产失败降级策略

### Phase 3 — 模板系统（⚠️ 部分完成）
- [ ] `style_extractor.py`: PPTX 模板风格提取（骨架待完善）
- [ ] `style_extractor.py`: PDF/图片模板 LLM 视觉分析
- [x] 内置 4 个主题母版（16 种 layout 完整）
- [x] style_profile.json → 渲染管线全链路应用

### Phase 4 — 输入扩展（✅ 已完成）
- [x] `ingest/crawlers.py`: URL 爬取
- [x] `ingest/dir_scanner.py`: 文件夹递归合并
- [x] `progress.py`: 断点续传
- [x] CLI: 所有单步命令 + retry-asset

### Phase 5 — 质量闭环（🚧 进行中）
- [ ] `quality_check.py`: 自动质量检查
- [ ] LLM 视觉审阅（截图 → 改进建议）
- [x] 日志系统与错误报告

---

## 当前状态（2026-03-16）

**整体完成度：~85%**

核心功能（Step0-4）已全部实现，进入测试验证阶段。

### 已验证功能
- ✅ 内容获取：支持 MD/DOCX/PDF/URL/文件夹
- ✅ 结构化分析：完整的 prompt 工程
- ✅ 两阶段大纲生成：借鉴 PPTAgent 论文
- ✅ PPTX 装配：16 种布局处理器
- ✅ 视觉生成：图表 + 信息图 + AI 图片

### 待完善功能
- 🔄 质量检查系统（quality_check.py）
- 📋 模板提取工具（extract_template）
- 🧪 端到端测试验证

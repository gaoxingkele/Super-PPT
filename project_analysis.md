# Super-PPT 项目分析报告

> 生成时间：2026-03-16  
> 分析范围：代码骨架、DEV_PLAN、ROADMAP

---

## 一、项目概况

**Super-PPT** 是一个从任意来源（URL、PDF、Word、Markdown、文件夹）自动生成视觉丰富、专业美观 PPT 的工具。

### 核心架构（5 步管线）

```
Step0: 内容获取与统一化 (step0_ingest.py) ✅ 完整实现
       ↓
Step1: LLM 结构化分析 (step1_analyze.py) ⚠️ 骨架/占位
       ↓
Step2: 幻灯片大纲生成 (step2_outline.py) ⚠️ 骨架/占位
       ↓
Step3: 视觉资产生成 (step3_visuals.py) ⚠️ 骨架/占位
       ↓
Step4: PPTX 装配 (step4_build.py) ⚠️ 骨架/占位
       ↓
    output/{base}/{base}_slides.pptx
```

---

## 二、当前状态盘点

### ✅ 已完成的模块

| 模块 | 文件路径 | 完成度 | 说明 |
|------|----------|--------|------|
| CLI 入口 | `main.py` | 100% | 8 个子命令 + help-all |
| 配置中心 | `config.py` | 100% | 10 provider + 路径 + 限制 |
| LLM 抽象层 | `src/llm_client.py` | 100% | 多 provider + retry 机制 |
| Step0 内容获取 | `src/step0_ingest.py` | 100% | URL/文件/目录分发处理 |
| 爬虫模块 | `src/ingest/crawlers.py` | 100% | Pyppeteer + Playwright |
| 文件读取器 | `src/ingest/*.py` | 100% | PDF/DOCX/MD/目录扫描 |
| 图表渲染 | `src/visuals/charts.py` | 100% | 10 种 matplotlib 图表 |
| 公式渲染 | `src/visuals/formula.py` | 100% | LaTeX → PNG |
| PPTX 引擎 | `src/utils/pptx_engine.py` | 100% | 13 种布局处理器 |
| 断点续传 | `src/utils/progress.py` | 100% | JSON checkpoint |
| Prompt 模板 | `src/prompts/*.py` | 100% | analyze/outline/style |

### ⚠️ 骨架/占位状态的模块

| 模块 | 文件路径 | 状态 | 主要问题 |
|------|----------|------|----------|
| Step1 结构化分析 | `src/step1_analyze.py` | ⚠️ 骨架 | 需调试 prompt，处理 JSON 解析容错 |
| Step2 大纲生成 | `src/step2_outline.py` | ⚠️ 骨架 | 需调试 prompt，确保 visual.data 格式正确 |
| Step3 视觉资产 | `src/step3_visuals.py` | ⚠️ 骨架 | 需验证并行调度，修复中文字体渲染 |
| Step4 PPTX 装配 | `src/step4_build.py` | ⚠️ 骨架 | 需验证 13 种布局，修复文字溢出/图片偏移 |
| 信息图渲染 | `src/visuals/infographics.py` | ⚠️ 部分 | process_flow + stat_display 实现，其余占位 |
| AI 图片生成 | `src/visuals/ai_images.py` | ⚠️ 占位 | 仅渐变背景降级方案 |
| 风格提取 | `src/style_extractor.py` | ⚠️ 骨架 | PPTX 基础 + PDF/图片占位 |

### ❌ 完全未实现的模块

| 模块 | 说明 |
|------|------|
| `themes/` 目录 | 空目录，需要制作 6 个内置主题母版 |
| `src/utils/quality_check.py` | 自动质量检查系统 |
| `src/utils/fonts.py` | 中英文字体管理工具 |
| `src/utils/color_utils.py` | 配色方案工具 |

---

## 三、分阶段改进计划

### 🚧 Phase 1：MVP 跑通（优先级最高）

**目标**：`python main.py generate report.md` 能端到端生成可用的 PPTX

```
□ Step1 结构化分析调试
  ├── 测试: python main.py analyze {base}
  ├── 检查 analysis.json 质量（chapters 3~8，data_points 含具体数字）
  ├── 迭代优化 prompts/analyze.py（增加 few-shot 示例）
  └── 处理 JSON 解析失败的容错

□ Step2 大纲生成调试
  ├── 测试: python main.py outline {base}
  ├── 检查 slide_plan.json 质量（幻灯片 15~30 张，visual 覆盖率 ≥60%）
  ├── 迭代优化 prompts/outline.py
  └── 添加 slide_plan 校验函数

□ Step3 视觉资产调试
  ├── 测试: python main.py visuals {base} --no-ai-images
  ├── 逐个验证 10 种 matplotlib 图表
  ├── 修复中文字体渲染问题
  └── 确认资产 manifest.json 正确记录状态

□ Step4 PPTX 装配调试
  ├── 测试: python main.py build {base}
  ├── 逐个验证 13 种布局处理器
  ├── 修复文字溢出、图片位置偏移
  └── 验证演讲者备注是否正确添加

□ 主题系统最低限度实现
  └── 制作 1 个 business.pptx 母版（至少包含 13 种 slide layout）
```

### 🔮 Phase 2：视觉质量提升

```
□ 信息图渲染器扩展
  ├── timeline（横向时间轴 + 节点）
  ├── hierarchy（树状结构）
  ├── comparison（左右对比 + 指标条）
  ├── cycle（环形箭头）
  ├── matrix（2×2 矩阵）
  └── pyramid（金字塔层级）

□ 图表渲染增强
  ├── stacked_bar 堆叠柱状图
  ├── area 面积图
  ├── treemap 矩形树图（需 squarify）
  └── gauge 仪表盘图

□ 集成 Claude Code Skills
  ├── infographics skill 调用
  └── generate-image skill 调用
```

### 🔮 Phase 3：模板系统完善

```
□ 内置主题制作
  ├── business.pptx（深蓝渐变、无衬线）
  ├── academic.pptx（白底、衬线体）
  ├── tech.pptx（深色背景、亮色强调）
  ├── minimal.pptx（大留白、黑白）
  ├── consulting.pptx（麦肯锡风格）
  └── creative.pptx（大胆配色）

□ 风格提取增强
  ├── PPTX：提取 theme.xml 配色、字体方案、布局结构
  ├── PDF：pdf2image 截图 + chat_vision 分析
  └── 图片：chat_vision 分析配色和排版
```

### 🔮 Phase 4：输入扩展

```
□ URL 爬取优化
  ├── 测试各类网站（新闻、博客、论文、产品页）
  ├── 添加常用平台特化选择器
  └── 处理 SPA 页面（等待渲染）

□ 高级文件处理
  ├── 支持 .xlsx 表格输入
  ├── 支持 .csv 数据输入
  └── 支持 .html 网页文件

□ 断点续传细化
  ├── Step3 资产级续传（只重新生成 failed 的资产）
  └── slide_plan 编辑后仅重新生成变更的幻灯片
```

### 🔮 Phase 5：质量闭环

```
□ 自动质量检查
  ├── 幻灯片数量 15~30
  ├── 每页 bullet ≤ 5 条，每条 ≤ 30 字
  ├── 字号下限检查（标题 ≥24pt，正文 ≥16pt）
  ├── 图片分辨率检查（≥720p）
  ├── 视觉覆盖率 ≥60%
  └── 文件大小 ≤30MB

□ LLM 视觉审阅
  ├── PPTX → 截图（每页）
  ├── chat_vision 审阅截图
  └── 输出改进建议
```

---

## 四、当前最大阻塞点

### 🔴 阻塞 MVP 的问题

1. **Step 1-4 都是骨架状态**
   - 只有函数签名和空实现
   - 需要逐一填充逻辑并调试

2. **Prompt 工程是核心但未验证**
   - Step1 和 Step2 的 prompt 质量直接决定最终 PPT 质量
   - 需要多次迭代 + 多类型输入测试

3. **数据格式契约未打通**
   - Step1 输出的 `data_points` 必须能被 Step2 映射为 `visual.data`
   - Step2 输出的 `visual.data` 必须能被 Step3 的 charts.py 消费
   - 需要端到端验证

4. **主题系统完全缺失**
   - `themes/` 目录是空的
   - Step4 装配 PPTX 时没有母版可用

### 🟡 次要但重要的问题

- 中文字体渲染问题（matplotlib 图表）
- AI 图片生成失败时的降级策略
- 错误处理和用户友好提示

---

## 五、推荐下一步行动

### 选项 A：快速 MVP 路线（推荐）

1. **制作最小主题模板**
   ```bash
   # 先创建一个包含 13 种布局的 business.pptx
   # 可使用 PowerPoint/WPS 手动创建或使用 python-pptx 程序化生成
   ```

2. **按顺序调试各 Step**
   - 准备 1 个测试用的 Markdown 文件
   - 从 Step0 开始，逐个运行并调试
   - 每步验证 JSON 输出质量

3. **端到端打通**
   - 修复各 Step 间的数据格式问题
   - 确保 `python main.py generate test.md` 能跑完

### 选项 B：并行推进路线

可以并行处理互不依赖的任务：

| 任务 | 负责人 | 依赖 |
|------|--------|------|
| 制作 business.pptx 主题 | 子任务 A | 无 |
| 调试 Step1 analyze | 子任务 B | 无 |
| 调试 Step2 outline | 子任务 C | Step1 |
| 调试 Step3 visuals | 子任务 D | Step2 |
| 调试 Step4 build | 子任务 E | Step3 + 主题 |

### 选项 C：深度优先路线

先深度完善某个特定功能：
- 例如：先把图表渲染做到极致（10 种类型全部验证）
- 或先把信息图渲染器完整实现

---

## 六、测试素材建议

准备以下类型的测试文件：

1. **商业分析报告 (.md)** — 测试 business 主题，含数据和趋势
2. **学术论文 (.pdf)** — 测试 academic 主题，含公式和引用
3. **技术文档 (.docx)** — 测试 tech 主题，含流程和架构图
4. **数据密集型内容** — 测试图表渲染能力
5. **流程/架构类内容** — 测试信息图渲染能力

---

## 七、参考资源

- **chatgpt-document**: 同架构管线，可复用 LLM 调用模式
- **scientific-slides skill**: 幻灯片设计最佳实践
- **infographics skill**: 信息图生成参考

---

*文档最后更新：2026-03-16*


---

## 八、模板系统详细要求

### 8.1 13 种 Slide Layout 规格

每种母版必须包含以下所有布局：

| layout | 用途 | 视觉占比 | Placeholder 要求 |
|--------|------|----------|------------------|
| `cover` | 封面 | 图片全屏 + 文字叠加 | 标题、副标题、日期 |
| `agenda` | 目录/议程 | 纯文字 | 标题 + 编号列表 |
| `section_break` | 章节过渡页 | 大标题 + 背景 | 章节号、章节标题 |
| `title_content` | 标题 + 要点 | 左文右图 (6:4) | 标题、bullet 列表 |
| `data_chart` | 数据图表页 | 图表占 70% | 标题、图表区域、takeaway 文字框 |
| `infographic` | 信息图页 | 图片居中占 80% | 标题、大图区域 |
| `two_column` | 双栏对比 | 左右均分 | 标题、左栏标题+内容、右栏标题+内容 |
| `key_insight` | 核心发现 | 大字引用 + KPI | 标题、quote、KPI 区域 |
| `table` | 表格页 | 原生 Table | 标题、表格区域 |
| `image_full` | 全屏图片 | 图片 100% | 图片区域 |
| `quote` | 引用页 | 大字居中 | 引用文字、来源 |
| `timeline` | 时间线 | 信息图 80% | 标题、时间轴区域 |
| `summary` | 总结页 | 要点 + CTA | 标题、bulllet 列表、call_to_action |

### 8.2 配色方案规范

每个主题需要定义 6 种颜色：

```json
{
  "color_scheme": {
    "primary": "#1B365D",      // 主色（标题、重点）
    "secondary": "#4A90D9",    // 次要色（辅助元素）
    "accent": "#E8612D",       // 强调色（高亮、CTA）
    "background": "#FFFFFF",   // 背景色
    "text": "#333333",         // 正文颜色
    "text_light": "#666666"    // 次要文字颜色
  }
}
```

### 8.3 字体规范

```json
{
  "typography": {
    "title_font": "微软雅黑",      // 标题字体
    "body_font": "微软雅黑",       // 正文字体
    "title_size_pt": 36,           // 标题字号
    "body_size_pt": 18,            // 正文字号
    "title_bold": true             // 标题是否加粗
  }
}
```

### 8.4 6 个内置主题风格定义

| 主题 | 风格描述 | 配色特点 |
|------|----------|----------|
| **business** | 商务正式 | 深蓝渐变(#1B365D)、无衬线字体、干净利落 |
| **academic** | 学术报告 | 白底(#FFFFFF)、深色标题、衬线体(宋体/Times) |
| **tech** | 科技现代 | 深色背景(#1A1A2E)、亮色强调(#00D9FF)、几何元素 |
| **minimal** | 极简主义 | 大留白、黑白(#000000)+单一强调色 |
| **consulting** | 咨询风格 | 麦肯锡/BCG 风格、蓝灰色系(#2E5C8A, #7F8C8D) |
| **creative** | 创意设计 | 大胆配色（紫/橙/青）、不规则布局 |

### 8.5 技术实现规范

#### 8.5.1 文件格式
- 使用 `.pptx` 格式
- 在母版视图（Slide Master）中定义
- 每个 layout 必须有唯一的 `name` 属性

#### 8.5.2 Placeholder 命名规范

```python
# python-pptx 中使用的 placeholder 索引或名称
PLACEHOLDERS = {
    "title": 0,        # 标题
    "subtitle": 1,     # 副标题
    "body": 2,         # 正文/ bullet
    "chart": 3,        # 图表区域（图片占位符）
    "image": 4,        # 图片占位符
}
```

#### 8.5.3 尺寸规范
- 标准 16:9 比例
- 尺寸：33.867 cm × 19.05 cm (13.333 × 7.5 英寸)
- 边距：2 cm

### 8.6 MVP 最低要求（简化版）

如果你现在就想开始，**先只做一个 `business.pptx`**：

**使用 PowerPoint / WPS 创建步骤：**

1. 打开 PowerPoint → 「视图」→「幻灯片母版」
2. 在母版视图中添加 13 个 layout
3. 每个 layout 设置：
   - 标题 placeholder（必需）
   - 内容 placeholder（根据类型配置）
4. 设置母版配色：
   - 主色：深蓝 #1B365D
   - 强调色：橙色 #E8612D
   - 背景：白色 #FFFFFF
5. 设置字体：微软雅黑
6. 保存到 `themes/business.pptx`

**更简化的版本（仅 5 个核心 layout）：**

如果 13 个太多，可以先做这 5 个最常用：
- `cover` - 封面
- `title_content` - 标题+内容（最常用）
- `data_chart` - 数据图表
- `section_break` - 章节过渡
- `summary` - 总结页

这样 MVP 就能跑起来，后续再补充其他 layout。

---

*文档最后更新：2026-03-16*


---

## 九、「提取模板工具」开发计划

### 9.1 工具概述

**目标**：将用户提供的任意 PPTX 文件，自动提炼为符合 Super-PPT 标准的母版模板。

**输入**：用户提供的 `.pptx` 文件（参考模板）  
**输出**：`themes/{theme_name}.pptx` + `style_profile.json`

---

### 9.2 功能架构

```
用户 PPTX
    ↓
[步骤1: 内容解析]
    ├── 提取 theme.xml（配色方案）
    ├── 提取字体方案
    ├── 提取所有 slide layout
    └── 截图每页用于视觉分析
    ↓
[步骤2: 智能分析]
    ├── LLM 视觉分析每页截图
    ├── 分类 layout 类型（匹配13种标准）
    ├── 提取设计模式（边距、对齐方式）
    └── 生成 style_profile.json
    ↓
[步骤3: 母版生成]
    ├── 创建新的母版 PPTX
    ├── 生成 13 种标准 layout
    ├── 应用提取的配色/字体
    └── 复制/适配用户的设计元素
    ↓
输出: business.pptx + style_profile.json
```

---

### 9.3 详细开发步骤

#### 步骤1: 内容解析模块 (`extractor/parser.py`)

**功能**：用 python-pptx 提取原始信息

```python
# 核心功能
class PPTXParser:
    def parse(self, pptx_path: Path) -> ParsedPPTX:
        return {
            "color_scheme": self._extract_theme_colors(),      # 6种核心颜色
            "fonts": self._extract_fonts(),                     # 标题/正文字体
            "layouts": self._extract_layouts(),                 # 所有 layout 结构
            "slides": self._extract_slide_screenshots(),        # 每页截图
            "placeholders": self._extract_placeholder_positions()  # 占位符位置
        }
```

**技术要点**：
- 解析 `pptx/ppt/theme/theme.xml` 提取配色
- 遍历所有 slide，记录每个 placeholder 的位置和大小
- 用 `python-pptx` + `PIL` 生成每页截图

**开发时间**：1-2 天

---

#### 步骤2: 智能分析模块 (`extractor/analyzer.py`)

**功能**：LLM 视觉分析，分类 layout 类型

```python
class LayoutAnalyzer:
    def analyze(self, parsed: ParsedPPTX) -> LayoutMapping:
        """
        将用户的 layout 映射到 13 种标准 layout
        """
        mapping = {}
        for slide_img in parsed["slides"]:
            # 用 LLM vision 分析这页是什么类型
            layout_type = self._classify_with_llm(slide_img)
            mapping[layout_type] = {
                "source_slide_idx": idx,
                "design_elements": self._extract_design_elements(slide_img),
                "confidence": 0.92
            }
        return mapping
```

**LLM Prompt 示例**：

```
分析这张 PPT 截图，判断它属于以下哪种类型：
- cover: 封面，大标题，可能有副标题和日期
- agenda: 目录，列表形式
- title_content: 标题+正文内容
- data_chart: 包含图表
- infographic: 信息图、流程图
- two_column: 左右两栏布局
- key_insight: 核心发现，大字强调
- table: 表格
- image_full: 全屏图片
- quote: 引用页
- timeline: 时间线
- section_break: 章节过渡页
- summary: 总结页

返回 JSON：{"layout_type": "xxx", "reason": "..."}
```

**技术要点**：
- 使用 `chat_vision` 进行视觉分析
- 支持多 provider（Gemini、Claude、GPT-4V）
- 置信度低于 0.7 时标记为"待确认"

**开发时间**：2-3 天

---

#### 步骤3: 母版生成模块 (`extractor/generator.py`)

**功能**：生成标准化的母版 PPTX

```python
class MasterGenerator:
    def generate(self, 
                 color_scheme: dict,
                 fonts: dict,
                 layout_mapping: LayoutMapping,
                 output_path: Path):
        """
        生成包含 13 种 layout 的母版文件
        """
        prs = Presentation()
        
        # 设置母版属性
        self._apply_color_scheme(prs, color_scheme)
        self._apply_fonts(prs, fonts)
        
        # 生成 13 种标准 layout
        for layout_type in STANDARD_LAYOUTS:
            if layout_type in layout_mapping:
                # 复制用户的设计
                self._create_layout_from_reference(
                    prs, layout_type, layout_mapping[layout_type]
                )
            else:
                # 使用默认模板生成
                self._create_default_layout(prs, layout_type, color_scheme)
        
        prs.save(output_path)
```

**Layout 创建策略**：

| 情况 | 处理方式 |
|------|----------|
| 用户 PPT 有匹配的 layout | 复制 placeholder 位置和样式 |
| 用户 PPT 缺少某些 layout | 用默认模板生成，但应用用户配色/字体 |
| 多个 slide 匹配同类型 | 选设计最完整的一个 |

**技术要点**：
- 使用 `python-pptx` 创建母版
- 精确控制 placeholder 位置和大小
- 复制背景图片/渐变效果

**开发时间**：3-4 天

---

#### 步骤4: 风格配置文件生成 (`extractor/style_profile.py`)

**功能**：生成 `style_profile.json`

```python
def generate_style_profile(parsed: ParsedPPTX, 
                          layout_mapping: LayoutMapping) -> dict:
    return {
        "source": "reference.pptx",
        "color_scheme": parsed["color_scheme"],
        "typography": parsed["fonts"],
        "layout_style": {
            "margin_cm": detect_margin(parsed),
            "content_alignment": detect_alignment(parsed),
            "visual_weight": detect_visual_weight(parsed),
            "decoration": detect_decoration_style(parsed)
        },
        "design_language": classify_design_language(parsed),
        "extracted_layouts": list(layout_mapping.keys()),
        "layout_details": layout_mapping,
        "generation_metadata": {
            "timestamp": "2026-03-16T...",
            "llm_provider": "gemini",
            "confidence_scores": {...}
        }
    }
```

**开发时间**：1 天

---

### 9.4 CLI 接口设计

```bash
# 基础用法
python -m tools.extract_template reference.pptx --name business

# 高级选项
python -m tools.extract_template reference.pptx \
    --name business \
    --output-dir themes/ \
    --provider gemini \
    --interactive  # 交互式确认 layout 分类

# 输出
# themes/business.pptx          # 生成的母版
# themes/business_profile.json  # 风格配置文件
```

---

### 9.5 项目文件结构

```
tools/
├── extract_template/
│   ├── __init__.py
│   ├── __main__.py          # CLI 入口
│   ├── parser.py            # 步骤1: 内容解析
│   ├── analyzer.py          # 步骤2: 智能分析
│   ├── generator.py         # 步骤3: 母版生成
│   ├── style_profile.py     # 步骤4: 配置文件生成
│   └── utils.py             # 辅助函数
```

---

### 9.6 开发里程碑

| 阶段 | 任务 | 时间 | 输出 |
|------|------|------|------|
| **MVP** | 解析 + 基础生成 | 3 天 | 能提取配色/字体，生成基础母版 |
| **v0.2** | LLM 视觉分析 | +2 天 | 自动分类 layout 类型 |
| **v0.3** | 完整 13 layout | +2 天 | 生成完整标准母版 |
| **v0.4** | 交互式确认 | +1 天 | 低置信度时人工确认 |
| **v0.5** | 优化迭代 | +2 天 | 处理复杂背景、渐变等 |

**总计：约 10 天完成完整功能**

---

### 9.7 技术风险与对策

| 风险 | 影响 | 对策 |
|------|------|------|
| LLM 分类准确率低 | layout 匹配错误 | 置信度阈值 + 人工确认机制 |
| 复杂背景提取失败 | 母版外观差异大 | 降级为纯色背景，保留配色 |
| 字体缺失 | 生成 PPT 字体不对 | 映射到系统可用字体 |
| Placeholder 位置不精确 | 内容错位 | 提供手动微调工具 |

---

### 9.8 与你的协作流程

如果你提供参考 PPT，我的处理流程：

```
你提供 reference.pptx
        ↓
我运行提取工具（MVP 版本）
        ↓
输出: business.pptx + 分析报告
        ↓
你检查 13 种 layout 的质量
        ↓
反馈问题 → 我调整代码 → 重新生成
        ↓
确认 OK，合并到 themes/
```

**MVP 版本可以先只做**：
1. 提取配色和字体
2. 生成 5 个核心 layout（cover, title_content, data_chart, section_break, summary）
3. 其他 layout 用默认模板填充

这样 2-3 天就能有一个可用版本。

---

*文档最后更新：2026-03-16*

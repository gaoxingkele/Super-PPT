# Super-PPT

> 从 URL、PDF、Word、Markdown、文件夹等任意来源，自动生成视觉丰富、结构完整、适合交付的 PPT。

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 项目目标

Super-PPT 的核心目标不是“把文本塞进模板”，而是做一条完整的自动化演示文稿生产管线：

- 自动获取和清洗原始内容
- 用 LLM 理解主题、结构、节奏和关键数据
- 生成带视觉指令的幻灯片大纲
- 按页面职责自动选择图表、信息图、AI 图片等视觉方案
- 最终装配为 `.pptx`

项目特别强调三件事：

- 叙事完整：不是孤立页面，而是一套能讲的汇报
- 视觉分层：封面、过渡页、信息图、图表、总结页各走不同策略
- 可恢复：支持 checkpoint、分步执行、审阅后局部重生图

## 主要能力

- 多源输入：支持 `URL / PDF / DOCX / Markdown / 文件夹`
- 两阶段大纲生成：先编排结构，再逐页填充内容
- 多模型接入：统一通过 `src/llm_client.py` 调用不同 Provider
- 风格提取：支持从 `pptx / pdf / 图片` 抽取参考风格
- 数据可视化：支持 `matplotlib` 图表与信息图
- AI 视觉：支持普通配图、高品质主视觉、HTML 信息图
- 中文优化：CJK 文本长度控制、字号和布局自适应
- 断点续传：每一步可恢复，避免全流程重跑
- 审阅迭代：支持多轮 review 后只重跑受影响页面

## 快速开始

### 安装

```bash
git clone <repository-url>
cd Super-PPT
pip install -r requirements.txt
```

### 配置环境变量

复制 `.env.example` 到 `.env`，按需填写；如果使用 Cloubic 路由，再补充 `.env.cloubic`。

最常用的是这些：

```bash
# 文本模型
LLM_PROVIDER=doubao
DOUBAO_API_KEY=your_doubao_key

# 普通图片直连豆包
DOUBAO_IMAGE_MODEL=seedream-5-0-260128

# Cloubic（用于高品质图片 / 统一路由）
CLOUBIC_ENABLED=true
CLOUBIC_API_KEY=your_cloubic_key

# 高品质图片模型（banana pro / 高质量图）
CLOUBIC_HIGH_QUALITY_IMAGE_MODEL=gemini-3-pro-image-preview
```

### 一键生成

```bash
python main.py generate report.md
python main.py generate paper.pdf --theme academic
python main.py generate https://example.com/article --template ref.pptx
```

### 分步执行

```bash
python main.py ingest report.md
python main.py analyze report
python main.py outline report
python main.py visuals report
python main.py build report --theme business
```

### 先确认大纲，再继续生成

如果希望流程更严谨，可以启用确认式生成：

```bash
python main.py generate report.md --require-outline-confirm
```

这会让流程在 `Step2` 后暂停，并输出两份中间结果：

- `output/<base>/slide_plan.json`
- `output/<base>/outline.md`

其中：

- `slide_plan.json` 面向程序执行
- `outline.md` 面向人工审阅与编辑

你可以直接修改 `outline.md` 中的每页标题、要点、视觉说明、备注、`layout`、`density`、`template_variant`，然后执行：

```bash
python main.py outline-import <base> output/<base>/outline.md
```

导回后，重新执行原来的 `generate ... --require-outline-confirm` 命令即可继续后续管线。

### 补充的分步命令

```bash
python main.py outline-export report
python main.py outline-import report output/report/outline.md
```

## 技术栈

项目当前不是单一“PPT 模板生成器”，而是一条由内容处理、LLM 编排、视觉生成和 PPTX 装配组成的多模块系统。下面按实际代码结构说明。

### 1. 语言与运行环境

- Python 3.10+
- 命令行入口：`main.py`
- 配置管理：`python-dotenv` + `config.py`

### 2. 内容获取与解析

对应模块：

- `src/step0_ingest.py`
- `src/ingest/crawlers.py`
- `src/ingest/pdf_reader.py`
- `src/ingest/docx_reader.py`
- `src/ingest/md_reader.py`
- `src/ingest/dir_scanner.py`

核心依赖：

- `pypdf`：PDF 文本提取
- `python-docx`：DOCX 读取
- `httpx`：URL 内容抓取与 API 请求
- 可选浏览器抓取：`playwright`

能力说明：

- 支持 URL、PDF、DOCX、Markdown、文件夹
- 将不同来源统一转成后续 LLM 可消费的文本内容
- 为 Step1 输出标准化原始内容

### 3. LLM 接入与模型路由

对应模块：

- `src/llm_client.py`
- `config.py`

核心依赖：

- `openai`：作为 OpenAI 兼容接口客户端
- `httpx`：直接访问部分图片与模型接口
- `tenacity`：重试机制

已接入的 Provider：

- Doubao
- DeepSeek
- Qwen
- Gemini
- OpenAI
- Claude
- Grok
- Kimi
- MiniMax
- GLM
- Perplexity

能力说明：

- 抽象统一 `chat()` 接口
- 支持直连与 Cloubic 路由两种模式
- 允许文本模型、推理模型、图片模型分开配置

### 4. 大纲生成与内容编排

对应模块：

- `src/step1_analyze.py`
- `src/step2_outline.py`
- `src/prompts/analyze.py`
- `src/prompts/outline.py`

核心思路：

- Step1 用 LLM 做结构化理解
- Step2 用两阶段策略生成 PPT 大纲
- 通过章节权重、节奏规划、数据点估算页数
- 给每页补足 `layout / bullets / visual / notes / takeaway`
- 同时补充 `density / template_variant / content_summary`
- 导出双轨中间层：
  - `slide_plan.json` 给程序消费
  - `outline.md` 给人工审阅和编辑

这部分的技术栈重点不在第三方库，而在提示词系统和中间 JSON 结构设计。

### 5. 视觉生成

对应模块：

- `src/step3_visuals.py`
- `src/visuals/charts.py`
- `src/visuals/ai_images.py`
- `src/visuals/infographics.py`
- `src/visuals/html_infographics.py`
- `src/visuals/pptx_charts.py`
- `src/visuals/pptx_infographics.py`

核心依赖：

- `matplotlib`
- `seaborn`
- `numpy`
- `Pillow`
- `playwright`

视觉能力拆分：

- 图表：`matplotlib` 生成柱状图、折线图、饼图、雷达图等
- 信息图：AntV Infographic DSL + HTML + Playwright 截图
- AI 图片：普通图走豆包直连，高品质图走 Cloubic 高质量模型
- PPT 内原生图表/信息图：在部分场景下由 `python-pptx` 原生绘制

### 6. 视觉路由层

对应模块：

- `src/step3_visuals.py`

这是这次更新后非常关键的一层，负责在真正渲染前判断：

- 这页是不是典型信息图
- 这页是不是高品质主视觉
- 这页是不是普通配图

再把任务分发到：

- AntV Infographic
- Cloubic 高品质图片模型
- 豆包 Seedream 直连
- 各自 fallback 链

### 7. PPTX 装配

对应模块：

- `src/step4_build.py`
- `src/utils/pptx_engine.py`
- `src/style_extractor.py`

核心依赖：

- `python-pptx`
- `Pillow`

能力说明：

- 根据页面布局放置标题、正文、图片、图表、信息图
- 支持主题模板与参考风格注入
- 管理字体、颜色、留白、对齐和图文混排
- 根据 `density / template_variant` 对关键布局做轻量模板化调整
- 当前已在 `title_content / data_chart / infographic / summary` 等布局中消费这些结构字段

### 8. 断点续传与安全写入

对应模块：

- `src/utils/progress.py`
- `src/utils/safe_write.py`

能力说明：

- 记录每一步完成状态
- 支持中断后从中间继续
- 降低长流程生成时的重复计算成本
- 额外记录 `outline_confirmed`，支持“Step2 后暂停确认，再继续后续管线”

### 9. 审阅与质量控制

对应模块：

- `src/step5_review.py`
- `src/visual_inspector.py`

能力说明：

- 多轮 review 评分
- 排版异常检测
- 局部修改后只重跑受影响页面
- 重生图阶段复用与 Step3 相同的视觉路由逻辑

### 10. 主题与风格系统

对应资源：

- `themes/*.pptx`
- `src/style_extractor.py`

能力说明：

- 支持内置主题
- 支持从参考 PPTX / PDF / 图片提取风格
- 将风格信息注入到大纲生成和 PPTX 装配阶段

## 命令说明

### 全流程

```bash
python main.py generate <source>
```

常见参数：

- `--theme business|academic|tech|minimal|consulting|creative`
- `--template ref.pptx`
- `--slides 15-25`
- `--require-outline-confirm`
- `--no-ai-images`
- `--review`
- `--review-rounds 5`
- `--review-target 9.0`
- `--enrich`
- `--cloubic`
- `--direct`

### 单步命令

```bash
python main.py ingest <source>
python main.py analyze <base>
python main.py outline <base>
python main.py outline-export <base>
python main.py outline-import <base> <outline.md>
python main.py visuals <base>
python main.py build <base>
python main.py review <base>
python main.py extract-style <template>
python main.py list-themes
```

## 处理管线

```text
输入源
  URL / PDF / Word / Markdown / 文件夹 / 参考模板
        ↓
Step0  内容获取与统一化
        ↓
Step1  结构化分析
        ↓
Step2  幻灯片大纲生成（两阶段）
        ↓
Step3  视觉资产生成与路由
        ↓
Step4  PPTX 装配
        ↓
Step5  审阅迭代（可选）
```

### Step0 `src/step0_ingest.py`

负责把不同来源统一成可处理的中间内容：

- URL 抓取正文
- PDF / DOCX 提取文本
- Markdown / 文本文件直接读取
- 文件夹递归收集文本型内容

输出通常会写入 `output/<base>/raw_content.md` 等中间文件。

### Step1 `src/step1_analyze.py`

负责用 LLM 生成结构化理解结果，典型内容包括：

- 标题、副标题
- 章节划分
- 章节权重
- 关键结论
- 数据点
- 节奏建议
- 推荐主题风格

这一步的结果会影响后面页数估算、布局分配、视觉类型选择。

### Step2 `src/step2_outline.py`

大纲生成采用两阶段策略：

1. Phase A：生成全局蓝图
2. Phase B：逐页补充细节

核心逻辑：

- 先根据章节权重、字数、数据点估算页数
- 固定插入 `cover / agenda / summary / end`
- 控制 `section_break` 数量和节奏
- 对每页生成：
- `layout`
- `title / subtitle`
- `bullets`
- `visual`
- `notes`
- `takeaway`
- `density`
- `template_variant`
- `content_summary`

Step2 完成后会同时输出两份中间结果：

- `slide_plan.json`
  - 面向程序执行
  - 给 Step3 / Step4 / Step5 消费
- `outline.md`
  - 面向人工审阅
  - 支持直接编辑后再回填到 `slide_plan.json`

如果启用 `--require-outline-confirm`，流程会在这里暂停，直到执行 `outline-import` 后才继续。

页面设计遵循金字塔原则：

- 标题是结论，不是栏目名
- 页面文字尽量少
- 连续高密度页后要插视觉缓冲页

### Step3 `src/step3_visuals.py`

这是项目目前最关键的一层之一。它不只是“生成图片”，而是先做视觉任务路由，再调用不同渲染器。

当前支持三类视觉资产：

- `matplotlib` 图表
- `infographics` 信息图
- `generate-image` AI 图片

其中 AI 图片和信息图现在会进一步走统一的智能分流逻辑。

## 视觉路由逻辑

这次更新后，视觉不再简单按 `visual.type` 直接执行，而是先判断“这页视觉的职责是什么”，再选择最合适的生成方式。

### 1. 什么时候走 AntV Infographic

满足以下任一条件，会优先走信息图渲染：

- `visual.type == "infographics"`
- 有 `infographic_type`
- `layout` 属于 `infographic / timeline / key_insight / summary`
- 原始是 `generate-image`，但带有明显结构化数据
- `design_intent / description / prompt / title` 中出现流程、时间线、对比、矩阵、层级、网络、循环、金字塔、KPI、架构图等特征

适合的页面：

- 流程页
- 时间线
- 架构图
- KPI 展示
- 对比页
- 总结页

执行路径：

- `src/visuals/html_infographics.py`
- 先由 LLM 生成 AntV DSL
- 再嵌入 HTML
- 最后用 Playwright 截图为 PNG

如果 HTML 信息图失败，再走 AI 信息图 fallback，最后回落到 `matplotlib`。

### 2. 什么时候走高品质图片

满足以下任一条件，会判定为高品质主视觉：

- `layout` 属于 `cover / section_break / image_full / end`
- `visual.role` 是 `hero` 或 `cover`
- `visual.quality == "high"`
- `visual.image_route == "high_quality"`
- `design_intent / prompt / title` 中出现封面、主视觉、冲击力、海报感、电影感、hero、poster、cinematic 等特征

适合的页面：

- 封面页
- 章节过渡页
- 全屏主视觉页
- 收尾致谢页

执行路径：

- 优先走 Cloubic 高品质图片模型
- 当前配置项为 `CLOUBIC_HIGH_QUALITY_IMAGE_MODEL`
- 适合接 banana pro 这类高质量模型

### 3. 什么时候走普通图片

没有被判定成信息图，也没有被判定成高品质主视觉的普通配图，默认都走普通图片链路。

适合的页面：

- 常规左文右图
- 概念插图
- 背景配图
- 图表下方氛围背景图

执行路径：

- 优先 `doubao-seedream-5-0-260128` 直连
- 失败后回退 Cloubic 备选模型
- 最后回退到 `matplotlib` 渐变背景

## 图片生成链路

### 普通图片

路径如下：

1. 豆包直连 `DOUBAO_IMAGE_MODEL`
2. Cloubic `wan2.6-t2i`
3. Cloubic `qwen-image-edit-plus`
4. `matplotlib` fallback

### 高品质图片

路径如下：

1. Cloubic 高品质图片模型 `CLOUBIC_HIGH_QUALITY_IMAGE_MODEL`
2. 豆包直连 `DOUBAO_IMAGE_MODEL`
3. Cloubic `wan2.6-t2i`
4. Cloubic `qwen-image-edit-plus`
5. `matplotlib` fallback

### 信息图

路径如下：

1. AntV HTML + Playwright
2. 豆包直连生成信息图
3. Cloubic `wan2.6-t2i`
4. Cloubic `qwen-image-edit-plus`
5. Gemini 直连
6. `matplotlib` 信息图回退

## Step4 `src/step4_build.py`

负责把前面所有结构化结果和视觉资产装配成 PPTX，核心工作包括：

- 选择布局处理器
- 放置标题、正文、图表、信息图、图片
- 套用主题样式
- 处理字体、颜色、留白、对齐
- 消费 `density / template_variant`
- 对关键布局做轻量的后处理排版优化

底层 PPTX 相关逻辑集中在：

- `src/utils/pptx_engine.py`

## Step5 `src/step5_review.py`

可选的审阅迭代阶段，目标不是“重跑全部”，而是：

- 让多个 Agent 从逻辑、受众、排版等角度打分
- 识别有问题的页面
- 只对改动页面重生成视觉

现在 Step5 的局部重渲染已经复用了 Step3 的统一视觉路由，所以审阅后的结果不会和正常生成走出两套策略。

## 配置说明

核心配置在 [config.py](config.py)。

### 常用文本模型配置

```bash
LLM_PROVIDER=doubao
DOUBAO_API_KEY=your_key
DOUBAO_MODEL=doubao-seed-2-0-pro-260215
```

### 图片相关配置

```bash
# 普通图片
DOUBAO_IMAGE_MODEL=seedream-5-0-260128

# 是否启用 Cloubic
CLOUBIC_ENABLED=true
CLOUBIC_API_KEY=your_cloubic_key
CLOUBIC_BASE_URL=https://api.cloubic.com/v1

# 普通 Cloubic 图片模型（备选）
CLOUBIC_IMAGE_MODEL=gemini-3-pro-image-preview

# 高品质图片模型
CLOUBIC_HIGH_QUALITY_IMAGE_MODEL=gemini-3-pro-image-preview
```

### 路由模式切换

```bash
# 强制启用 Cloubic
python main.py generate report.md --cloubic

# 强制直连
python main.py generate report.md --direct
```

### 其他常用配置

```bash
PPT_LANGUAGE=zh
DEFAULT_THEME=business
REVIEW_FALLBACK_PROVIDERS=deepseek,qwen,gemini
```

## 项目结构

```text
Super-PPT/
├── main.py
├── config.py
├── requirements.txt
├── src/
│   ├── step0_ingest.py
│   ├── step1_analyze.py
│   ├── step1_5_enrich.py
│   ├── step2_outline.py
│   ├── step3_visuals.py
│   ├── step4_build.py
│   ├── step5_review.py
│   ├── llm_client.py
│   ├── style_extractor.py
│   ├── visual_inspector.py
│   ├── utils/
│   │   ├── progress.py
│   │   └── pptx_engine.py
│   ├── visuals/
│   │   ├── ai_images.py
│   │   ├── charts.py
│   │   ├── html_infographics.py
│   │   ├── infographics.py
│   │   └── pptx_infographics.py
│   └── prompts/
├── themes/
└── output/
```

## 使用建议

### 什么时候用 `--no-ai-images`

适合：

- 快速预览结构
- 调试 Step2 / Step4
- API Key 尚未配置完整

不适合：

- 想验证最终视觉效果
- 需要高品质封面和章节过渡页

### 什么时候加 `--template`

适合：

- 你已经有公司/学校/客户的参考模板
- 需要贴近既有品牌样式
- 需要让生成结果更稳定

### 什么时候开 `--review`

适合：

- 正式交付版
- 页数较多
- 对视觉和逻辑要求高

## 已知限制

- HTML 信息图依赖 Playwright
- 高品质图片依赖 Cloubic 和对应图片模型
- 普通图片依赖豆包图片 API
- 无可用图片服务时会回退到本地生成的占位视觉
- 中文字体显示效果依赖系统字体环境

## 相关文档

- [DEV_PLAN.md](DEV_PLAN.md)
- [DEBUG_PLAN.md](DEBUG_PLAN.md)
- [ROADMAP.md](ROADMAP.md)
- [THEMES_STATUS.md](THEMES_STATUS.md)
- [CHANGELOG.md](CHANGELOG.md)

## 许可证

MIT License

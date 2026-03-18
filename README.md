# Super-PPT

> 从任意来源（URL、PDF、Word、Markdown、文件夹）自动生成视觉丰富、专业美观的 PPT。

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## ✨ 特性

- **📝 多源输入**：支持 URL、PDF、DOCX、Markdown、文件夹
- **🎨 智能设计**：自动匹配主题风格，16 种专业布局
- **📊 数据可视化**：自动提取数据并生成图表
- **🖼️ AI 视觉**：支持 AI 生成封面和信息图
- **📱 中文优化**：CJK 文字自适应，微软雅黑字体
- **⚡ 两阶段生成**：结构编排 + 逐页填充，确保叙事流畅

## 🚀 快速开始

### 安装

```bash
# 克隆仓库
git clone <repository-url>
cd Super-PPT

# 安装依赖
pip install -r requirements.txt

# 配置环境变量
cp .env.example .env
# 编辑 .env，添加你的 API Keys
```

### 一键生成

```bash
# 从 Markdown 文件生成
python main.py generate report.md --theme yili_power

# 从 PDF 生成
python main.py generate paper.pdf --theme xmu_graph

# 从 URL 生成
python main.py generate https://example.com/article --theme epri_nature
```

### 分步执行

```bash
# Step 0: 内容获取
python main.py ingest report.md

# Step 1: 结构化分析
python main.py analyze report

# Step 2: 幻灯片大纲生成
python main.py outline report --two-phase

# Step 3: 视觉资产生成
python main.py visuals report --no-ai-images

# Step 4: PPTX 装配
python main.py build report --theme yili_power
```

## 📚 文档

- [开发计划](DEV_PLAN.md) - 详细开发路线图
- [调试计划](DEBUG_PLAN.md) - 测试验证步骤
- [主题状态](THEMES_STATUS.md) - 可用主题列表
- [技术规划](ROADMAP.md) - 架构设计文档

## 🎨 内置主题

| 主题 | 风格 | 适用场景 |
|------|------|----------|
| `yili_power` | 商务深蓝 | 企业汇报、项目申报 |
| `xmu_graph` | 学术青绿 | 学术答辩、研究报告 |
| `epri_nature` | 科研蓝绿 | 基金申请、科研报告 |
| `zhinang_qa` | 科技渐变 | 产品演示、技术分享 |

## 🏗️ 架构

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

## 🛠️ 技术栈

- **核心**：Python 3.10+
- **PPTX 处理**：python-pptx
- **图表渲染**：matplotlib, seaborn
- **AI 生成**：Google Gemini API
- **内容提取**：PyPDF2, python-docx, beautifulsoup4

## 📁 项目结构

```
Super-PPT/
├── main.py                    # CLI 入口
├── config.py                  # 配置中心
├── requirements.txt           # 依赖清单
├── src/
│   ├── step0_ingest.py        # 内容获取
│   ├── step1_analyze.py       # 结构化分析
│   ├── step2_outline.py       # 大纲生成（两阶段）
│   ├── step3_visuals.py       # 视觉资产生成
│   ├── step4_build.py         # PPTX 装配
│   ├── llm_client.py          # LLM 抽象层
│   ├── utils/
│   │   └── pptx_engine.py     # PPTX 装配引擎
│   ├── visuals/
│   │   ├── charts.py          # 图表渲染
│   │   ├── infographics.py    # 信息图生成
│   │   └── ai_images.py       # AI 图片生成
│   └── prompts/               # LLM 提示词
├── themes/                    # 主题模板
│   ├── yili_power.pptx
│   ├── xmu_graph.pptx
│   ├── epri_nature.pptx
│   └── zhinang_qa.pptx
└── output/                    # 生成结果
```

## 🎯 核心特性详解

### 两阶段大纲生成

借鉴 PPTAgent 论文，采用结构编排 + 逐页填充的两阶段策略：

1. **Phase A - 结构编排**：决定页数、章节划分、布局类型
2. **Phase B - 逐页填充**：填充 bullets、视觉指令、演讲备注

### 金字塔原则

- **断言式标题**：每页标题是完整结论，不是标签
- **结论先行**：先给结论，再展开证据
- **页面密度**：每页正文不超过 70 字

### CJK 文字自适应

- 自动计算中文字符宽度
- 根据布局自动截断过长 bullet
- 微软雅黑 + Times New Roman 混排

## ⚙️ 配置

### 环境变量 (.env)

```bash
# LLM API Keys（至少配置一个）
KIMI_API_KEY=your_kimi_key
GEMINI_API_KEY=your_gemini_key
GROK_API_KEY=your_grok_key
DEEPSEEK_API_KEY=your_deepseek_key

# 可选：配置默认提供商
DEFAULT_PROVIDER=gemini
```

### 生成参数

```bash
# 指定幻灯片数量
python main.py generate report.md --slides 15-25

# 跳过 AI 图片生成（快速预览）
python main.py generate report.md --no-ai-images

# 指定主题
python main.py generate report.md --theme xmu_graph

# 使用参考模板
python main.py generate report.md --template ref.pptx
```

## 🧪 测试

```bash
# 创建测试文档
echo "# 测试报告
## 数据
- 2023年营收：1000万元
- 2024年营收：1500万元

## 结论
业务增长迅速。
" > test.md

# 运行测试
python main.py generate test.md --theme yili_power --no-ai-images
```

## 📌 已知限制

- 需要配置 LLM API Key 才能使用完整功能
- AI 图片生成依赖 Gemini API（无 API key 时回退到渐变背景）
- 中文字体依赖系统安装的微软雅黑

## 🤝 贡献

欢迎提交 Issue 和 PR！

## 📄 许可证

MIT License

---

*Made with ❤️ for better presentations*

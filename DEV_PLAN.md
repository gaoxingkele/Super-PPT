# Super-PPT 开发详细计划与执行步骤

## 当前状态：代码骨架已搭建

已完成的文件（可运行的结构，部分含占位实现）：

```
Super-PPT/
├── main.py                    ✅ CLI 入口（8 子命令 + help-all）
├── config.py                  ✅ 完整配置（10 provider + 路径 + 限制）
├── requirements.txt           ✅ 依赖清单
├── .env.example               ✅ 环境变量模板
├── .gitignore                 ✅
├── CLAUDE.md                  ✅ 项目说明
├── ROADMAP.md                 ✅ 技术规划路线
├── DEV_PLAN.md                ✅ 本文件
├── src/
│   ├── __init__.py            ✅ sys.path 设置
│   ├── llm_client.py          ✅ 完整（10 provider + retry）
│   ├── step0_ingest.py        ✅ 完整（URL/文件/目录分发）
│   ├── step1_analyze.py       ✅ 骨架（需调试 prompt）
│   ├── step2_outline.py       ✅ 骨架（需调试 prompt）
│   ├── step3_visuals.py       ✅ 骨架（并行调度器）
│   ├── step4_build.py         ✅ 骨架（调用 pptx_engine）
│   ├── style_extractor.py     ✅ 骨架（PPTX 基础 + PDF/图片占位）
│   ├── ingest/
│   │   ├── crawlers.py        ✅ 完整（Pyppeteer + Playwright）
│   │   ├── md_reader.py       ✅ 完整（YAML + 表格 + 图片）
│   │   ├── docx_reader.py     ✅ 完整（文本 + 表格 + 图片）
│   │   ├── pdf_reader.py      ✅ 完整（文本 + 表格 + 图片）
│   │   └── dir_scanner.py     ✅ 完整（递归扫描 + 合并）
│   ├── visuals/
│   │   ├── charts.py          ✅ 完整（10 种图表渲染器）
│   │   ├── infographics.py    ✅ 部分（process_flow + stat_display + 占位）
│   │   ├── ai_images.py       ✅ 占位（渐变背景）
│   │   └── formula.py         ✅ 完整（LaTeX → PNG）
│   ├── utils/
│   │   ├── pptx_engine.py     ✅ 完整（13 种布局处理器）
│   │   └── progress.py        ✅ 完整（断点续传）
│   └── prompts/
│       ├── analyze.py         ✅ 完整（system + user prompt 模板）
│       ├── outline.py         ✅ 完整（system + user prompt 模板）
│       └── style.py           ✅ 完整（风格提取 prompt）
└── themes/                    ⬜ 空（需制作主题模板）
```

---

## Phase 1：MVP 跑通（优先级最高）

**目标**：`python main.py generate report.md` 能端到端生成可用的 PPTX。

### 1.1 环境搭建与基础验证
```
□ git init && 首次提交
□ cp .env.example .env && 配置至少一个 LLM API Key
□ pip install -r requirements.txt
□ 验证: python main.py --help
□ 验证: python main.py list-themes
```

### 1.2 Step0 内容获取调试
```
□ 准备测试文件: 1 个 MD、1 个 DOCX、1 个 PDF
□ 测试: python main.py ingest test_report.md
□ 检查输出: output/{base}/raw_content.md, raw_meta.json, raw_tables.json
□ 修复: 表格提取准确性
□ 修复: 图片提取路径
□ 测试: python main.py ingest ./test_dir/  (文件夹)
```

### 1.3 Step1 结构化分析调试
```
□ 测试: python main.py analyze {base}
□ 检查 analysis.json 质量:
  - chapters 数量是否 3~8
  - data_points 是否包含具体数字
  - concepts 描述是否足够生成信息图
□ 迭代优化 prompts/analyze.py:
  - 增加 few-shot 示例
  - 调整 JSON schema 约束
  - 测试不同 provider 效果
□ 处理 JSON 解析失败的容错
```

### 1.4 Step2 大纲生成调试
```
□ 测试: python main.py outline {base}
□ 检查 slide_plan.json 质量:
  - 幻灯片数量是否在 15~30
  - visual 覆盖率是否 ≥ 60%
  - matplotlib data 字段是否包含完整数据
  - infographics description 是否足够详细
□ 迭代优化 prompts/outline.py:
  - 增加布局选择的 few-shot 示例
  - 约束 visual.data 的数据格式
  - 确保生成的 chart data 可直接被 charts.py 消费
□ 添加 slide_plan 校验函数
```

### 1.5 Step3 视觉资产调试
```
□ 测试: python main.py visuals {base} --no-ai-images
□ 逐个验证 10 种 matplotlib 图表:
  - bar, line, pie, donut, radar, heatmap
  - scatter, grouped_bar, waterfall, funnel
□ 修复中文字体渲染问题
□ 修复: 数据格式不匹配导致的渲染失败
□ 验证 process_flow 和 stat_display 信息图
□ 确认资产 manifest.json 正确记录状态
```

### 1.6 Step4 PPTX 装配调试
```
□ 测试: python main.py build {base}
□ 逐个验证 13 种布局:
  - cover, agenda, section_break, title_content
  - data_chart, infographic, two_column, key_insight
  - table, image_full, quote, timeline, summary
□ 修复: 文字溢出、图片位置偏移
□ 修复: 演讲者备注是否正确添加
□ 验证: 无主题模板时的空白模板效果
```

### 1.7 全管线端到端测试
```
□ 测试: python main.py generate test_report.md
□ 用 PowerPoint/WPS 打开验证
□ 检查: 视觉覆盖率、文字可读性、配色一致性
□ 修复所有阻塞问题
□ git commit: "feat: MVP 全管线跑通"
```

---

## Phase 2：视觉质量提升

### 2.1 信息图渲染器扩展
```
□ 实现 timeline 渲染器（横向时间轴 + 节点）
□ 实现 hierarchy 渲染器（树状结构）
□ 实现 comparison 渲染器（左右对比 + 指标条）
□ 实现 cycle 渲染器（环形箭头）
□ 实现 matrix 渲染器（2×2 矩阵）
□ 实现 pyramid 渲染器（金字塔层级）
□ 统一信息图样式: 圆角、阴影、配色方案
```

### 2.2 图表渲染增强
```
□ 添加 stacked_bar 堆叠柱状图
□ 添加 area 面积图
□ 添加 treemap 矩形树图（需 squarify）
□ 添加 gauge 仪表盘图
□ 图表自适应: 根据数据量调整尺寸和标注密度
□ 添加动画效果标记（供 PPTX 动画使用）
```

### 2.3 集成 infographics skill
```
□ 研究 infographics skill 的调用方式
□ 实现 visuals/infographics.py 的 skill 调用逻辑
□ 添加 AI 信息图 ↔ matplotlib 降级切换
□ 配色方案传递给 skill
□ 质量检查: Gemini 3 Pro 审核
```

### 2.4 集成 generate-image skill
```
□ 研究 generate-image skill 的调用方式
□ 实现 visuals/ai_images.py 的 skill 调用逻辑
□ 封面生成 2~3 张备选 + LLM 选最佳
□ prompt 工程: 确保风格统一
□ 添加降级: AI 生成失败 → 渐变色背景
```

---

## Phase 3：模板系统完善

### 3.1 内置主题制作
```
□ 制作 business.pptx 母版（深蓝渐变、无衬线）
□ 制作 academic.pptx 母版（白底、衬线体）
□ 制作 tech.pptx 母版（深色背景、亮色强调）
□ 制作 minimal.pptx 母版（大留白、黑白）
□ 制作 consulting.pptx 母版（麦肯锡风格）
□ 制作 creative.pptx 母版（大胆配色）
□ 每个母版包含 13 种 slide layout
```

### 3.2 PPTX 模板风格深度提取
```
□ 提取母版配色主题 (theme.xml)
□ 提取字体方案
□ 提取布局结构 (layout 名称 + placeholder 位置)
□ 用户 PPTX 模板直接作为母版使用
```

### 3.3 PDF/图片模板风格分析
```
□ pdf2image 将 PDF 转为截图
□ chat_vision 分析配色和排版
□ 输出 style_profile.json
□ 图片模板同理
```

---

## Phase 4：输入扩展

### 4.1 URL 爬取优化
```
□ 测试各类网站: 新闻、博客、论文、产品页
□ 添加常用平台特化选择器
□ 处理 SPA 页面（等待渲染）
□ 支持自动检测语言
```

### 4.2 高级文件处理
```
□ 支持 .xlsx 表格输入（直接提取数据）
□ 支持 .csv 数据输入
□ 支持 .html 网页文件
□ 大文件分块处理 + 摘要合并
```

### 4.3 断点续传细化
```
□ Step3 资产级续传: 只重新生成 failed 的资产
□ slide_plan 编辑后仅重新生成变更的幻灯片
□ 进度百分比显示
```

---

## Phase 5：质量闭环

### 5.1 自动质量检查
```
□ 实现 quality_check.py:
  - 幻灯片数量 15~30
  - 每页文字量限制
  - 字号下限检查
  - 图片分辨率检查
  - 视觉覆盖率 ≥ 60%
  - 文件大小 ≤ 30MB
□ 检查结果写入 quality_report.json
□ 检查不通过时给出具体修复建议
```

### 5.2 LLM 视觉审阅
```
□ PPTX → 截图（每页）
□ chat_vision 审阅截图
□ 输出改进建议
□ 可选: 根据建议自动修正
```

### 5.3 日志与错误处理
```
□ 统一日志格式 (带时间戳)
□ 每步耗时统计
□ LLM token 用量统计
□ 错误分类与友好提示
```

---

## 开发注意事项

### Prompt 工程是核心
- Step1 (analyze) 和 Step2 (outline) 的 prompt 质量直接决定最终 PPT 质量
- 每次调整 prompt 后需用 3+ 不同类型的输入测试
- 记录每次 prompt 迭代的效果

### 数据格式契约
- Step1 输出的 `data_points` 必须能被 Step2 直接映射为 `visual.data`
- Step2 输出的 `visual.data` 必须能被 Step3 的 charts.py 直接消费
- 中间 JSON 格式变更需同步更新所有消费方

### 测试素材
准备以下类型的测试文件:
1. 商业分析报告 (.md) — 测试 business 主题
2. 学术论文 (.pdf) — 测试 academic 主题
3. 技术文档 (.docx) — 测试 tech 主题
4. 数据密集型内容 — 测试图表渲染
5. 流程/架构类内容 — 测试信息图渲染

### 参考项目
- chatgpt-document: 同架构管线，可复用 LLM 调用模式
- scientific-slides skill: 幻灯片设计最佳实践
- infographics skill: 信息图生成

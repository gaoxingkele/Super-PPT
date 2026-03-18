# 更新日志

## [Unreleased] - 2026-03-16

### 新增

#### 核心功能
- **Step0 内容获取**：支持 URL、PDF、DOCX、Markdown、文件夹多种输入源
- **Step1 结构化分析**：LLM 深度分析，提取章节、数据点、概念、论证链
- **Step2 大纲生成**：两阶段模式（结构编排 + 逐页填充），借鉴 PPTAgent 论文
- **Step3 视觉资产**：并行生成 matplotlib 图表、信息图、AI 图片
- **Step4 PPTX 装配**：16 种布局处理器，支持 CJK 文字自适应

#### 高级特性
- **金字塔原则**：断言式标题、结论先行、页面密度控制
- **功能页自动校验**：自动补充 cover/agenda/end 页面
- **智能图文拆页**：图片面积利用率低时自动拆分为两页
- **CJK 文字自适应**：根据布局自动截断/调整 bullet 文字
- **富文本支持**：`**bold**` 自动渲染为深红色加粗

#### 视觉生成
- **图表渲染**：10 种 matplotlib 图表（bar/line/pie/radar/heatmap/waterfall/funnel...）
- **信息图生成**：9 种类型（process_flow/timeline/hierarchy/comparison/matrix/network/pyramid/cycle/stat_display）
  - 优先使用 Gemini API 生成
  - 失败时回退到 matplotlib 本地渲染
- **AI 图片生成**：
  - 使用 Gemini API 生成高质量图片
  - 支持 prompt 增强（根据内容类型自动选择风格）
  - 失败时回退到渐变背景

#### 主题模板
- **yili_power**：商务深蓝风格（#024177）
- **xmu_graph**：学术青绿风格（#178F95）
- **epri_nature**：科研蓝绿风格（#0070C0）
- **zhinang_qa**：科技渐变风格（#00479D）
每个主题包含完整的 16 种 layout。

#### CLI 工具
- `generate`：一键生成 PPT（全管线）
- `ingest/analyze/outline/visuals/build`：单步执行
- `list-themes`：列出内置主题
- `review`：三角色迭代审阅
- `extract-style`：提取参考模板风格
- `retry-asset`：重试单个视觉资产

### 技术亮点

1. **两阶段大纲生成**
   - Phase A：结构编排（决定页数、布局、节奏）
   - Phase B：逐页填充（bullets、visual、notes）
   - 借鉴 PPTAgent 论文的最佳实践

2. **JSON 容错解析**
   - 多种 fallback 策略处理 LLM 返回的非标准 JSON
   - 自动提取 code block 内容
   - 智能修复缺少的括号

3. **原生图表渲染**
   - 在 PPTX 中直接渲染 matplotlib 图表（非图片插入）
   - 支持交互式数据标签和样式

4. **智能字体设置**
   - 中文使用微软雅黑
   - 英文/数字使用 Times New Roman
   - 通过 XML 直接设置字体属性

### 项目结构

```
Super-PPT/
├── main.py                    # CLI 入口
├── config.py                  # 配置中心
├── src/
│   ├── step0_ingest.py        # 内容获取与统一化
│   ├── step1_analyze.py       # 结构化分析
│   ├── step2_outline.py       # 幻灯片大纲生成
│   ├── step3_visuals.py       # 视觉资产生成
│   ├── step4_build.py         # PPTX 装配
│   ├── llm_client.py          # 多 provider LLM 抽象层
│   ├── prompts/               # LLM 提示词工程
│   ├── utils/
│   │   └── pptx_engine.py     # PPTX 装配引擎（16 种布局）
│   └── visuals/
│       ├── charts.py          # matplotlib 图表渲染
│       ├── infographics.py    # 信息图生成
│       └── ai_images.py       # AI 图片生成
├── themes/                    # 4 个完整主题模板
└── output/                    # 生成结果
```

### 待办事项

- [ ] 质量检查系统（quality_check.py）
- [ ] 模板提取工具（tools/extract_template）
- [ ] 端到端测试验证
- [ ] 用户使用文档完善

---

## 开发历程

### 2026-03-16
- ✅ 完成 4 个主题模板的 16 种 layout
- ✅ 创建项目 README.md
- ✅ 更新 ROADMAP.md 完成状态
- ✅ 编写 CHANGELOG.md

### 2026-03-15
- ✅ 分析 source-doc 中的 4 个 PPTX 案例
- ✅ 手工创建 4 个主题模板（初版 5 layout）
- ✅ 编写模板提取工具开发计划

### 2026-03-14
- ✅ 完成 Step2 两阶段大纲生成（Phase A + Phase B）
- ✅ 实现功能页自动校验（cover/agenda/end 补全）
- ✅ 完善 prompts/analyze.py 和 prompts/outline.py

### 2026-03-13
- ✅ 完成 Step4 PPTX 装配引擎
- ✅ 实现 15 种布局处理器
- ✅ 添加 CJK 文字自适应
- ✅ 实现图文拆页策略

### 2026-03-12
- ✅ 完成 Step3 视觉资产生成
- ✅ 实现 matplotlib 图表渲染
- ✅ 实现信息图生成（9 种类型）
- ✅ 实现 AI 图片生成（Gemini API）

### 2026-03-11
- ✅ 完成 Step1 结构化分析
- ✅ 编写详细 prompt 工程
- ✅ 实现 JSON 容错解析

### 2026-03-10
- ✅ 完成 Step0 内容获取
- ✅ 实现多源输入支持
- ✅ 搭建项目骨架

---

## 参考资源

- [PPTAgent: Generating Professional Presentations](https://github.com/icip-cas/pptagent) - 两阶段生成策略灵感来源
- [chatgpt-document](https://github.com/.../chatgpt-document) - 架构参考
- [python-pptx](https://python-pptx.readthedocs.io/) - PPTX 处理库
- [matplotlib](https://matplotlib.org/) - 图表渲染库

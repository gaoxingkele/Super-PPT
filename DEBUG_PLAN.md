# Super-PPT 调试验证计划

> 状态：Step0-4 代码已完整，进入测试验证阶段
> 时间：2026-03-16

---

## 一、当前状态评估

### 1.1 代码完成度

| 模块 | 完成度 | 状态 |
|------|--------|------|
| Step0 内容获取 | 100% | ✅ 完整实现 |
| Step1 结构化分析 | 100% | ✅ 完整实现 + 详细 prompt |
| Step2 大纲生成 | 100% | ✅ 两阶段模式 + 功能页校验 |
| Step3 视觉资产 | 100% | ✅ 并行生成 + manifest |
| Step4 PPTX 装配 | 100% | ✅ 15 种布局处理器 |
| 图表渲染 | 100% | ✅ 10 种 matplotlib 图表 |
| 信息图渲染 | 100% | ✅ 9 种类型 + Gemini 回退 |
| AI 图片生成 | 100% | ✅ Gemini API + 渐变回退 |
| 主题模板 | 80% | ⚠️ 4 个手工模板（5 layout） |

### 1.2 关键特性已实现

- ✅ **两阶段大纲生成**：Phase A 骨架 + Phase B 逐页填充
- ✅ **功能页自动校验**：自动补充 cover/agenda/end
- ✅ **CJK 文字自适应**：根据布局自动截断/调整 bullet
- ✅ **图文拆页策略**：图片面积利用率低时自动拆为两页
- ✅ **原生图表渲染**：pptx_charts 直接在 PPTX 中渲染
- ✅ **智能字体设置**：中文微软雅黑 + 英文 Times New Roman
- ✅ **富文本支持**：**bold** 自动渲染为深红色加粗
- ✅ **JSON 容错解析**：多种 fallback 策略

---

## 二、测试验证计划

### 阶段 1：单步测试（第 1 天）

#### 测试 1.1: Step0 内容获取
```bash
# 准备测试文件
echo "# 测试报告
## 背景
这是一个测试文档，用于验证 Super-PPT 的端到端流程。

## 数据
- 2023年营收：1000万元
- 2024年营收：1500万元
- 增长率：50%

## 结论
业务增长迅速，前景良好。
" > test_report.md

# 运行 Step0
python main.py ingest test_report.md
```

**验证点**：
- [ ] 正确生成 `raw_content.md`
- [ ] 正确生成 `raw_meta.json`
- [ ] 正确生成 `raw_tables.json`

#### 测试 1.2: Step1 结构化分析
```bash
python main.py analyze test_report
```

**验证点**：
- [ ] `analysis.json` 格式正确
- [ ] 提取出 3~8 个章节
- [ ] 识别出数据点（1000万、1500万、50%）
- [ ] content_type 判断合理
- [ ] narrative_arc 和 rhythm_plan 已填充

#### 测试 1.3: Step2 大纲生成
```bash
python main.py outline test_report --two-phase
```

**验证点**：
- [ ] `slide_plan.json` 格式正确
- [ ] 生成 15~30 页幻灯片
- [ ] visual 覆盖率 ≥60%
- [ ] 包含 cover + agenda + end
- [ ] 每页都有 notes（100-300字）
- [ ] data_chart 页的 visual.data 格式正确

#### 测试 1.4: Step3 视觉资产
```bash
python main.py visuals test_report --no-ai-images
```

**验证点**：
- [ ] `assets/manifest.json` 生成
- [ ] matplotlib 图表渲染成功
- [ ] 信息图渲染成功（如有）

#### 测试 1.5: Step4 PPTX 装配
```bash
python main.py build test_report --theme yili_power
```

**验证点**：
- [ ] PPTX 文件生成
- [ ] 能在 PowerPoint/WPS 中打开
- [ ] 15 种布局渲染正常
- [ ] 配色和字体正确

### 阶段 2：端到端测试（第 2 天）

#### 测试 2.1: 一键生成
```bash
python main.py generate test_report.md --theme yili_power -o test_output
```

**验证点**：
- [ ] 全流程无报错
- [ ] 输出 PPTX 可用

#### 测试 2.2: 真实案例测试
使用 `source-doc/` 中的 4 个 PPTX 对应的原始文档（如有）或创建简化版：

| 案例 | 类型 | 预期页数 |
|------|------|----------|
| 亿力科技项目汇报 | 项目申报 | 30-40页 |
| 厦大图理论研究 | 学术报告 | 25-35页 |
| 电科院自然基金 | 科研申报 | 30-40页 |
| 输配作业智囊 | 产品演示 | 15-25页 |

### 阶段 3：质量检查（第 3 天）

#### 检查清单

**内容质量**：
- [ ] 章节逻辑清晰
- [ ] 每页标题是断言句（非标签）
- [ ] bullets 包含具体数字
- [ ] notes 详细且有用

**视觉质量**：
- [ ] 图表配色协调
- [ ] 字体大小合适（标题≥24pt，正文≥16pt）
- [ ] 无文字溢出
- [ ] 图片位置正确

**技术质量**：
- [ ] 文件大小 ≤30MB
- [ ] 打开速度正常
- [ ] 页码正确

---

## 三、已知限制

### 3.1 主题模板限制
- 当前只有 5 个核心 layout，缺少 8 个进阶 layout
- 需要手动创建或完善提取工具

### 3.2 API 依赖
- Gemini API key 需要配置才能生成 AI 图片
- 无 API key 时会回退到 matplotlib 渐变背景

### 3.3 中文处理
- 依赖系统安装的微软雅黑字体
- 某些系统可能需要手动安装字体

---

## 四、调试工具

### 4.1 日志查看
```bash
# 查看详细日志
python main.py generate test.md --verbose

# 查看某一步的输出
cat output/test/analysis.json | jq '.chapters | length'
cat output/test/slide_plan.json | jq '.slides | length'
```

### 4.2 单步重试
```bash
# 只重新运行某一步
python main.py outline test --two-phase --no-resume
python main.py visuals test --no-ai-images
python main.py build test --theme xmu_graph
```

### 4.3 快速验证
```bash
# 快速验证（跳过 AI 图片）
python main.py generate test.md --no-ai-images --slides 10-15
```

---

## 五、问题记录表

| 问题描述 | 严重度 | 状态 | 解决方案 |
|----------|--------|------|----------|
|          |        |      |          |
|          |        |      |          |

---

## 六、下一步行动

1. **立即开始**：运行测试 1.1-1.5 验证单步功能
2. **修复问题**：记录并修复发现的问题
3. **完整测试**：运行 4 个真实案例的端到端测试
4. **文档更新**：根据测试结果更新使用文档

---

*计划制定：2026-03-16*

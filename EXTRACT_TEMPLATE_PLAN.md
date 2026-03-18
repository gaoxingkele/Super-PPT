# 提取模板工具开发计划

> 目标：开发 `tools/extract_template/` 模块，将任意 PPTX 提炼为 Super-PPT 标准母版  
> 预计工期：10 天（MVP 3 天）

---

## 一、项目结构

```
tools/
└── extract_template/
    ├── __init__.py
    ├── __main__.py          # CLI 入口
    ├── parser.py            # 步骤1: PPTX 内容解析
    ├── analyzer.py          # 步骤2: LLM 视觉分析
    ├── generator.py         # 步骤3: 母版生成
    ├── style_builder.py     # 步骤4: 风格配置文件生成
    ├── constants.py         # 常量定义（13种layout规范）
    └── utils.py             # 辅助函数
```

---

## 二、Phase 1: MVP 基础版（3天）

### Day 1: 内容解析模块 (`parser.py`)

**目标**：提取 PPTX 的基本信息

```python
class PPTXParser:
    def parse(self, pptx_path: Path) -> ParsedData:
        return {
            "color_scheme": {...},      # 6种核心颜色
            "fonts": {...},              # 标题/正文字体
            "layouts": [...],            # 原始 layout 列表
            "slides": [...],             # 每页截图
        }
```

**任务清单**：
```
□ 实现 theme.xml 解析，提取配色方案
    ├── 读取 pptx/ppt/theme/theme.xml
    ├── 提取 clrScheme 中的 6 种颜色
    └── 映射到标准命名（primary/secondary/accent/bg/text/text_light）

□ 实现字体提取
    ├── 提取 majorFont（标题字体）
    ├── 提取 minorFont（正文字体）
    └── 中文字体名映射处理

□ 实现 layout 结构提取
    ├── 遍历所有 slide master 的 layouts
    ├── 记录每个 layout 的 name 和 placeholder 数量
    └── 生成每页截图（用于 LLM 分析）

□ 单元测试
    └── 用 3 个不同风格的 PPTX 测试解析准确性
```

**验收标准**：
- 能正确提取 90% 以上 PPTX 的配色和字体
- 生成清晰的每页截图

---

### Day 2: 母版生成模块 (`generator.py`) - 基础版

**目标**：生成包含 5 个核心 layout 的母版

**5 个核心 layout**：
1. `cover` - 封面
2. `title_content` - 标题+内容（最常用）
3. `data_chart` - 数据图表
4. `section_break` - 章节过渡
5. `summary` - 总结页

**任务清单**：
```
□ 实现母版创建基础框架
    ├── 创建 Presentation 对象
    ├── 设置 slide width/height（16:9）
    └── 设置默认文本样式

□ 实现 cover layout 生成
    ├── 添加标题 placeholder（顶部居中）
    ├── 添加副标题 placeholder
    ├── 添加日期/作者区域
    └── 应用配色方案

□ 实现 title_content layout 生成
    ├── 添加标题 placeholder（顶部）
    ├── 添加正文 placeholder（左侧 60%）
    ├── 添加图片 placeholder（右侧 40%）
    └── 设置 bullet 样式

□ 实现 data_chart layout 生成
    ├── 添加标题 placeholder
    ├── 添加图表区域 placeholder（居中 70%）
    └── 添加 takeaway 文字框

□ 实现 section_break layout 生成
    ├── 添加章节号 placeholder
    ├── 添加章节标题 placeholder（大字）
    └── 设置背景色/渐变

□ 实现 summary layout 生成
    ├── 添加标题 placeholder
    ├── 添加要点列表 placeholder
    └── 添加 CTA 区域

□ 单元测试
    └── 生成测试母版，用 PowerPoint 打开验证
```

**验收标准**：
- 生成的 PPTX 能在 PowerPoint/WPS 中正常打开
- 5 个 layout 的 placeholder 位置合理
- 配色和字体正确应用

---

### Day 3: CLI 集成与 MVP 验证

**目标**：打通从输入 PPTX 到输出母版的完整流程

**任务清单**：
```
□ 实现 __main__.py CLI 入口
    ├── 参数解析：input, --name, --output-dir
    ├── 调用 parser 提取信息
    ├── 调用 generator 生成母版
    └── 保存到 themes/{name}.pptx

□ 实现基础 style_profile.json 生成
    ├── 记录提取的 color_scheme
    ├── 记录提取的 fonts
    └── 记录生成的 layout 列表

□ MVP 端到端测试
    ├── 准备 3 个不同风格的测试 PPTX
    ├── 运行：python -m tools.extract_template test1.pptx --name theme1
    ├── 检查生成的 theme1.pptx
    └── 在 PowerPoint 中打开验证

□ 编写 MVP 使用文档
    └── README_MVP.md（快速开始指南）
```

**MVP 验收标准**：
```bash
# 命令能成功执行
python -m tools.extract_template reference.pptx --name business

# 输出文件
# themes/business.pptx          ✅ 存在且可打开
# themes/business_profile.json  ✅ 包含配色和字体信息

# 母版包含 5 个核心 layout，配色字体应用正确
```

---

## 三、Phase 2: 智能分析（+2天）

### Day 4: LLM 视觉分析模块 (`analyzer.py`)

**目标**：自动识别每页 PPT 的 layout 类型

**任务清单**：
```
□ 实现 layout 分类器
    ├── 准备 13 种 layout 的详细描述 prompt
    ├── 调用 chat_vision 分析每页截图
    └── 返回 layout_type + confidence

□ 实现智能匹配逻辑
    ├── 遍历用户 PPT 的每一页
    ├── 用 LLM 判断属于哪种标准 layout
    ├── 记录置信度分数
    └── 处理冲突（一页匹配多个类型）

□ 实现设计元素提取
    ├── 检测背景类型（纯色/渐变/图片）
    ├── 检测是否有装饰元素
    └── 记录特殊布局特征

□ 集成到主流程
    └── analyzer 在 parser 之后、generator 之前调用
```

**Prompt 模板**：
```python
LAYOUT_CLASSIFICATION_PROMPT = """
分析这张 PPT 截图，判断它属于以下哪种标准类型：

1. cover - 封面页，通常包含大标题、副标题、日期
2. agenda - 目录页，显示章节列表
3. title_content - 标题+正文内容，最常见的内容页
4. data_chart - 包含图表（柱状图、折线图、饼图等）
5. infographic - 信息图、流程图、架构图
6. two_column - 左右两栏布局
7. key_insight - 核心发现页，大字强调某个结论
8. table - 表格数据页
9. image_full - 全屏图片
10. quote - 引用页，显示名言或重要论断
11. timeline - 时间线/发展历程
12. section_break - 章节过渡页，大字章节标题
13. summary - 总结页，要点回顾

返回 JSON 格式：
{
    "layout_type": "xxx",
    "confidence": 0.95,
    "reason": "判断理由...",
    "key_features": ["特征1", "特征2"]
}
"""
```

**验收标准**：
- 对 5 个核心 layout 的分类准确率 > 85%
- 置信度低于 0.7 时标记为"待确认"

---

### Day 5: 智能母版生成增强

**目标**：根据 LLM 分析结果，智能复制用户的设计

**任务清单**：
```
□ 实现参考布局复制
    ├── 当检测到用户有匹配的 layout 时
    ├── 复制 placeholder 的精确位置和大小
    ├── 复制背景样式（颜色/渐变/图片）
    └── 复制特殊设计元素

□ 实现自适应布局生成
    ├── 当用户缺少某些 layout 时
    ├── 用已有 layout 的设计风格推导
    └── 生成风格一致的新 layout

□ 实现置信度处理
    ├── 高置信度 (>0.8): 自动处理
    ├── 中置信度 (0.5-0.8): 记录日志
    └── 低置信度 (<0.5): 使用默认模板

□ 集成测试
    └── 用 5 个不同 PPTX 测试智能提取效果
```

**验收标准**：
- 用户 layout 能被正确识别并复制到标准母版
- 生成的母版风格与用户参考 PPT 一致

---

## 四、Phase 3: 完整 13 Layout（+2天）

### Day 6-7: 扩展全部 Layout 类型

**目标**：支持完整的 13 种标准 layout

**新增 8 个 layout**：
```
□ agenda - 目录/议程
□ infographic - 信息图页
□ two_column - 双栏对比
□ key_insight - 核心发现
□ table - 表格页
□ image_full - 全屏图片
□ quote - 引用页
□ timeline - 时间线
```

**每个 layout 的实现内容**：
```python
LAYOUT_SPECS = {
    "agenda": {
        "placeholders": [
            {"type": "title", "left": 0.1, "top": 0.1, "width": 0.8, "height": 0.15},
            {"type": "body", "left": 0.1, "top": 0.3, "width": 0.8, "height": 0.6},
        ],
        "bullet_style": "numbered"
    },
    "infographic": {
        "placeholders": [
            {"type": "title", "left": 0.1, "top": 0.05, "width": 0.8, "height": 0.1},
            {"type": "picture", "left": 0.1, "top": 0.2, "width": 0.8, "height": 0.7},
        ]
    },
    # ... 其他 layout
}
```

**任务清单**：
```
□ 在 constants.py 中定义 13 种 layout 的完整规范
□ 为每个 layout 实现生成函数
□ 实现 two_column 的特殊布局（左右均分）
□ 实现 table 的原生表格 placeholder
□ 实现 timeline 的时间轴区域
□ 全量测试：确保 13 个 layout 都能正确生成
```

**验收标准**：
- 生成的母版包含完整的 13 个 layout
- 每个 layout 的 placeholder 位置合理
- 能用 Super-PPT 的 Step4 正确消费

---

## 五、Phase 4: 交互式确认（+1天）

### Day 8: 人工确认机制

**目标**：低置信度时让用户确认

**任务清单**：
```
□ 实现交互式确认模式
    ├── --interactive 参数
    ├── 显示截图和 LLM 判断结果
    └── 让用户选择/修改 layout 类型

□ 实现预览功能
    ├── 生成母版前预览 layout 分配
    ├── 显示哪些使用了用户设计，哪些是默认
    └── 允许调整配置

□ 实现增量更新
    ├── 已有母版时只更新变更部分
    └── 保留用户手动调整
```

**交互流程**：
```bash
$ python -m tools.extract_template ref.pptx --name business --interactive

正在分析第 3/15 页...
[显示截图]
LLM 判断: title_content (置信度: 0.65)
请选择:
  1. 确认 (title_content)
  2. 改为 cover
  3. 改为 data_chart
  4. 跳过此页
> 
```

---

## 六、Phase 5: 优化与完善（+2天）

### Day 9: 高级功能

**任务清单**：
```
□ 实现复杂背景处理
    ├── 渐变背景提取和复制
    ├── 背景图片提取
    └── 透明度处理

□ 实现装饰元素识别
    ├── 检测形状、线条、图标
    ├── 判断是否复制到母版
    └── 处理页眉页脚元素

□ 实现字体回退机制
    ├── 检测系统可用字体
    ├── 缺失字体时寻找替代
    └── 警告用户字体变更
```

### Day 10: 测试与文档

**任务清单**：
```
□ 编写完整测试套件
    ├── 单元测试（各模块）
    ├── 集成测试（端到端）
    └── 用 10+ 个真实 PPTX 测试

□ 编写用户文档
    ├── 使用方法
    ├── 参数说明
    ├── 常见问题
    └── 最佳实践

□ 代码审查与优化
    ├── 添加类型注解
    ├── 完善错误处理
    └── 性能优化
```

---

## 七、CLI 完整接口

```bash
# 基础用法
python -m tools.extract_template reference.pptx --name business

# 高级用法
python -m tools.extract_template reference.pptx \
    --name business \
    --output-dir themes/ \
    --provider gemini \
    --interactive \
    --verbose

# 仅分析不生成
python -m tools.extract_template reference.pptx --analyze-only

# 更新已有母版
python -m tools.extract_template reference.pptx --name business --update
```

---

## 八、输出文件规范

### 8.1 母版文件 (`themes/{name}.pptx`)

**要求**：
- 包含 13 种标准 layout（MVP 5 种）
- 每个 layout 有正确的 name 属性
- 配色和字体统一应用
- 16:9 比例，标准尺寸

### 8.2 风格配置文件 (`themes/{name}_profile.json`)

```json
{
  "source": "reference.pptx",
  "theme_name": "business",
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
    "body_size_pt": 18
  },
  "layout_style": {
    "margin_cm": 2.0,
    "content_alignment": "left",
    "visual_weight": "heavy",
    "decoration": "minimal"
  },
  "extracted_layouts": ["cover", "title_content", "data_chart", "section_break", "summary"],
  "layout_mapping": {
    "cover": {"source_slide": 1, "confidence": 0.95},
    "title_content": {"source_slide": 3, "confidence": 0.88}
  },
  "generation_metadata": {
    "timestamp": "2026-03-16T10:00:00Z",
    "tool_version": "0.1.0",
    "llm_provider": "gemini",
    "total_slides": 15,
    "matched_layouts": 5
  }
}
```

---

## 九、风险与对策

| 风险 | 概率 | 影响 | 对策 |
|------|------|------|------|
| python-pptx 无法读取某些 PPTX | 中 | 高 | 使用 zipfile 直接解析 XML 作为 fallback |
| LLM 视觉分析成本高 | 中 | 中 | 添加本地缓存，重复利用结果 |
| 复杂布局无法准确复制 | 高 | 中 | 提供手动编辑指南，记录已知限制 |
| 生成的母版 Super-PPT 无法消费 | 低 | 高 | 严格遵循 Step4 的 layout 规范 |

---

## 十、里程碑检查点

| 里程碑 | 时间 | 检查项 |
|--------|------|--------|
| **MVP 完成** | Day 3 | 5 个 layout 能生成可用母版 |
| **智能分析完成** | Day 5 | LLM 分类准确率 > 85% |
| **完整 Layout 完成** | Day 7 | 13 个 layout 全部可用 |
| **交互功能完成** | Day 8 | 交互式确认流程跑通 |
| **正式发布** | Day 10 | 测试通过，文档完整 |

---

*计划制定：2026-03-16*  
*预计完成：2026-03-26（MVP: 2026-03-19）*

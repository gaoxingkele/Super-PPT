# 贡献指南

感谢你对 Super-PPT 项目的关注！

## 项目架构

Super-PPT 采用**多步骤管线架构**，每个步骤独立运行，通过 JSON 文件传递数据：

```
Step0 → Step1 → Step2 → Step3 → Step4
(ingest) (analyze) (outline) (visuals) (build)
```

### 核心模块

| 模块 | 文件 | 职责 |
|------|------|------|
| CLI 入口 | `main.py` | 参数解析，子命令分发 |
| 内容获取 | `src/step0_ingest.py` | 多源输入统一化 |
| 结构分析 | `src/step1_analyze.py` | LLM 深度分析内容 |
| 大纲生成 | `src/step2_outline.py` | 两阶段幻灯片规划 |
| 视觉生成 | `src/step3_visuals.py` | 并行生成图表/信息图/AI图片 |
| PPTX 装配 | `src/step4_build.py` | 组装最终 PPTX |
| 装配引擎 | `src/utils/pptx_engine.py` | 16 种布局处理器 |

## 开发规范

### 代码风格

- **类型注解**：所有函数参数和返回值必须添加类型注解
- **文档字符串**：所有公共函数必须包含 docstring
- **错误处理**：使用 try-except 捕获预期异常，记录日志

### 模块首行

所有模块文件首行必须包含：

```python
import src  # noqa: F401
```

这是为了确保 `src` 目录在 Python 路径中。

### LLM Prompt 工程

- **System Prompt**：定义角色、任务、输出格式
- **User Prompt**：提供上下文、数据、约束条件
- **Few-shot**：复杂任务提供 1-2 个示例
- **JSON Schema**：明确指定输出 JSON 结构

示例：

```python
SYSTEM_PROMPT = """你是一位顶级的演示文稿设计师。

## 任务
将内容转化为专业 PPT 大纲。

## 输出格式
严格的 JSON：
{
  "meta": {...},
  "slides": [...]
}

## 规则
1. 每页标题必须是断言句
2. 视觉覆盖率 ≥60%
"""
```

### 数据契约

步骤间的数据传递必须遵守**数据契约**：

**Step1 → Step2**：
```python
# analysis.json
{
  "chapters": [
    {
      "id": "ch01",
      "title": "...",
      "data_points": [...],  # Step2 映射为 visual.data
      "concepts": [...]      # Step2 映射为 infographics
    }
  ]
}
```

**Step2 → Step3**：
```python
# slide_plan.json
{
  "slides": [
    {
      "visual": {
        "type": "matplotlib",
        "chart": "bar",
        "data": {...}  # 必须可被 charts.py 消费
      }
    }
  ]
}
```

## 如何添加新 Layout

在 `src/utils/pptx_engine.py` 中添加：

1. **实现布局处理器**：

```python
def _add_new_layout(builder: PPTXBuilder, spec: dict, asset_path: Optional[Path]):
    """新布局处理器。"""
    slide = builder._add_blank_slide()
    
    # 1. 添加标题
    _add_title_textbox(slide, spec.get("title", ""), builder)
    
    # 2. 添加内容
    # ... 自定义布局逻辑
    
    # 3. 添加页码
    _add_page_number(slide, builder, builder._slide_count + 1)
```

2. **注册到处理器字典**：

```python
_LAYOUT_HANDLERS = {
    # ... 现有 layout
    "new_layout": _add_new_layout,
}
```

3. **更新主题模板**：

创建包含新 layout 的 PPTX 文件，放入 `themes/` 目录。

## 如何添加新图表类型

在 `src/visuals/charts.py` 中添加：

1. **实现渲染函数**：

```python
def render_new_chart(data: dict, color_scheme: dict, options: dict = None) -> plt.Figure:
    """新图表渲染器。"""
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE)
    
    # 图表绘制逻辑
    # ...
    
    fig.tight_layout()
    return fig
```

2. **注册到渲染器字典**：

```python
CHART_RENDERERS = {
    # ... 现有图表
    "new_chart": render_new_chart,
}
```

3. **更新 Prompt**：

在 `src/prompts/outline.py` 中更新图表类型说明。

## 测试

### 本地测试

```bash
# 单步测试
python main.py ingest test.md
python main.py analyze test

# 端到端测试
python main.py generate test.md --no-ai-images
```

### 调试技巧

1. **查看中间输出**：
```bash
cat output/test/analysis.json | jq '.chapters | length'
cat output/test/slide_plan.json | jq '.slides[0]'
```

2. **查看日志**：
```bash
python main.py generate test.md --verbose 2>&1 | tee log.txt
```

3. **断点续传**：
```bash
# 删除某一步的进度，重新运行
rm output/test/analysis.json
python main.py analyze test
```

## 提交 PR

1. **Fork 仓库** 并创建分支
2. **编写代码** 并确保通过测试
3. **更新文档**（README.md、CHANGELOG.md）
4. **提交 PR** 并描述改动

## 代码审查清单

- [ ] 代码符合 PEP 8 规范
- [ ] 添加了适当的类型注解
- [ ] 包含文档字符串
- [ ] 错误处理完善
- [ ] 更新了相关文档
- [ ] 通过了本地测试

## 联系

如有问题，欢迎提交 Issue 或联系维护者。

---

感谢你的贡献！

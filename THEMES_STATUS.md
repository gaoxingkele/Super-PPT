# Super-PPT 主题模板状态

> 更新时间：2026-03-16

## 主题模板列表

| 主题文件 | 显示名称 | 主色调 | 布局数量 |
|----------|----------|--------|----------|
| `yili_power.pptx` | 亿力科技-电力生产安全 | 深蓝 #024177 | 16 |
| `xmu_graph.pptx` | 厦门大学-图理论研究 | 青绿 #178F95 | 16 |
| `epri_nature.pptx` | 电科院-自然基金 | 蓝色 #0070C0 | 16 |
| `zhinang_qa.pptx` | 输配作业智囊 | 深蓝 #00479D | 16 |

## 16 种标准 Layout

每个主题模板都包含以下全部 layout：

| 序号 | Layout 名称 | 用途说明 |
|------|-------------|----------|
| 1 | `cover` | 封面页 |
| 2 | `agenda` | 目录/议程页 |
| 3 | `section_break` | 章节过渡页 |
| 4 | `title_content` | 标题+内容页（最常用） |
| 5 | `data_chart` | 数据图表页 |
| 6 | `infographic` | 信息图页 |
| 7 | `two_column` | 双栏对比页 |
| 8 | `key_insight` | 核心发现页 |
| 9 | `table` | 表格页 |
| 10 | `image_full` | 全屏图片页 |
| 11 | `quote` | 引用页 |
| 12 | `timeline` | 时间线页 |
| 13 | `summary` | 总结页 |
| 14 | `methodology` | 研究方法/技术路线页 |
| 15 | `architecture` | 系统架构页 |
| 16 | `end` | 结束页 |

## 使用方法

### 命令行使用

```bash
# 使用特定主题生成 PPT
python main.py generate report.md --theme yili_power

# 列出所有可用主题
python main.py list-themes
```

### 在代码中使用

```python
from src.utils.pptx_engine import PPTXBuilder

builder = PPTXBuilder(
    template_path="themes/yili_power.pptx",
    color_scheme={...},
    assets_dir=...,
)
```

## 主题配色方案

### yili_power（商务深蓝）
```json
{
  "primary": "#024177",
  "secondary": "#005D7F",
  "accent": "#E8612D",
  "background": "#FFFFFF",
  "text": "#333333"
}
```

### xmu_graph（学术青绿）
```json
{
  "primary": "#178F95",
  "secondary": "#4472C4",
  "accent": "#E8612D",
  "background": "#FFFFFF",
  "text": "#333333"
}
```

### epri_nature（科研蓝绿）
```json
{
  "primary": "#0070C0",
  "secondary": "#4C9857",
  "accent": "#E8612D",
  "background": "#FFFFFF",
  "text": "#333333"
}
```

### zhinang_qa（科技渐变）
```json
{
  "primary": "#00479D",
  "secondary": "#5B9BD5",
  "accent": "#F2F2F2",
  "background": "#FFFFFF",
  "text": "#333333"
}
```

## 模板设计规范

### 尺寸
- 标准 16:9 宽屏比例
- 尺寸：13.333" × 7.5"（33.87cm × 19.05cm）

### 字体
- 中文字体：微软雅黑
- 英文字体：Times New Roman（在装配引擎中设置）

### 布局结构
每个 layout 都遵循一致的结构：
1. 标题区域（顶部）
2. 装饰线（标题下方）
3. 内容区域（中部）
4. 底部色条（页脚）

## 未来扩展

### 计划添加的主题
- [ ] `business` - 通用商务风格
- [ ] `academic` - 学术答辩风格
- [ ] `tech` - 科技深色风格
- [ ] `minimal` - 极简风格
- [ ] `consulting` - 咨询报告风格
- [ ] `creative` - 创意设计风格

### 提取工具
未来可开发 `tools/extract_template` 模块，自动从参考 PPT 提取风格并生成模板。

---

*当前所有 4 个主题均已包含完整的 16 种 layout，可直接用于 Super-PPT 生成。*

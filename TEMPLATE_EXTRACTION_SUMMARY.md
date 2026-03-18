# 模板提取工具开发总结

> 基于 4 个真实 PPTX 案例的手工分析与模板生成经验  
> 时间：2026-03-16

---

## 一、已完成工作

### 1.1 案例分析

分析了 `source-doc/` 目录下的 4 个 PPTX 文件：

| 文件 | 页数 | 主色调 | 风格特点 |
|------|------|--------|----------|
| 国网信通亿力科技-多模态项目 | 37页 | 深蓝 #024177 + 青色 #005D7F | 商务正式、政府项目汇报 |
| 厦门大学-图理论研究 | 44页 | 青绿 #178F95 + 蓝 #4472C4 | 学术、研究型、图表丰富 |
| 电科院-自然基金 | 40页 | 蓝 #0070C0 + 绿 #4C9857 | 科研、双色调、数据密集 |
| 输配作业智囊 | 14页 | 深蓝 #00479D + 天蓝 #5B9BD5 | 科技、简洁、产品展示 |

**共同特点**：
- 全部使用 **微软雅黑** 字体
- 16:9 宽屏比例
- 标题+内容是最主要布局类型
- 表格页用于展示数据和对比

### 1.2 手工创建模板

基于分析结果，创建了 4 个独立模板：

```
themes/
├── yili_power.pptx          # 亿力科技-商务深蓝
├── yili_power_profile.json
├── xmu_graph.pptx           # 厦门大学-学术青绿
├── xmu_graph_profile.json
├── epri_nature.pptx         # 电科院-科研蓝绿
├── epri_nature_profile.json
├── zhinang_qa.pptx          # 输配智囊-科技渐变
└── zhinang_qa_profile.json
```

每个模板包含 **5 个核心 layout**：
1. `cover` - 封面
2. `title_content` - 标题+内容（最常用）
3. `section_break` - 章节过渡
4. `data_chart` - 数据图表
5. `summary` - 总结页

---

## 二、关键发现与经验

### 2.1 颜色提取的挑战

**问题**：python-pptx 无法直接读取 `theme.xml` 中的配色方案

**实际做法**：
- 遍历前 5 页的所有形状
- 统计 `shape.fill.fore_color.rgb` 出现频率
- 取 Top 5 作为配色候选

**局限**：
- 只能提取实际使用的颜色，而非定义的配色方案
- 渐变色只能取到起始色
- 图片背景无法分析

**工具开发建议**：
```python
# 需要直接解析 XML 获取完整配色
def extract_theme_colors(pptx_path):
    """从 theme.xml 提取完整配色"""
    import zipfile
    from xml.etree import ElementTree as ET
    
    with zipfile.ZipFile(pptx_path) as z:
        xml_content = z.read('ppt/theme/theme1.xml')
    
    # 解析 a:clrScheme 获取完整配色定义
    # 返回: {primary, secondary, accent1-6, bg, text}
```

### 2.2 字体提取的挑战

**问题**：无法直接获取母版字体设置

**实际做法**：
- 遍历所有文本框的 `run.font.name`
- 统计出现频率
- 区分标题（>24pt）和正文

**发现**：
- 所有 4 个案例都用 **微软雅黑**
- 部分嵌入了 Times New Roman（英文/数字）
- 字体大小层次分明：标题 32-44pt，正文 16-20pt

**工具开发建议**：
```python
def extract_fonts(prs):
    """提取字体并智能分类"""
    fonts_by_size = defaultdict(list)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text_frame'):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name and run.font.size:
                            size_pt = run.font.size.pt
                            fonts_by_size[size_pt].append(run.font.name)
    
    # 聚类分析找出标题和正文字体
    title_font = most_common_in_range(fonts_by_size, (28, 48))
    body_font = most_common_in_range(fonts_by_size, (14, 24))
```

### 2.3 布局分类的挑战

**问题**：如何自动判断一页 PPT 属于哪种标准 layout？

**手工分析时的启发**：

| 判断依据 | layout 类型 |
|----------|-------------|
| 第 1 页，大标题 | `cover` |
| 有 MSO_SHAPE_TYPE.CHART | `data_chart` |
| 有 MSO_SHAPE_TYPE.TABLE | `table` |
| 有图片 + 少量文字 | `image_full` |
| 多文本框 + 复杂结构 | `title_content` |
| 背景色块 + 大字 | `section_break` |

**工具开发建议**：
- 先用规则分类（准确率约 70%）
- 再用 LLM 视觉分析验证和修正
- 低置信度时提示用户确认

```python
class LayoutClassifier:
    def classify(self, slide, slide_idx):
        # 规则分类
        rule_based = self._rule_based_classify(slide, slide_idx)
        
        # LLM 视觉分析
        screenshot = self._capture_slide(slide)
        vision_result = self._llm_classify(screenshot)
        
        # 融合决策
        if rule_based == vision_result['type']:
            return rule_based, 0.95
        elif vision_result['confidence'] > 0.8:
            return vision_result['type'], vision_result['confidence']
        else:
            return rule_based, 0.6  # 低置信度
```

### 2.4 布局复用的复杂性

**观察**：用户 PPT 中的 layout 很难直接复用，因为：
1. 形状位置是硬编码的，不是 placeholder
2. 背景和装饰元素与内容混合
3. 不同页面尺寸/比例可能不同

**更实用的做法**：
- **分析**用户 PPT 的设计元素（配色、字体、边距、风格）
- **新建**标准化的 13 个 layout
- **应用**提取的设计元素到标准 layout

这样生成的母版既能保持用户风格，又符合 Super-PPT 的标准。

---

## 三、提取工具架构建议

基于手工制作经验，工具应该分为 3 个阶段：

### Phase 1: 解析器 (`parser.py`)

```python
class TemplateParser:
    def parse(self, pptx_path: Path) -> ParsedData:
        return {
            # 直接提取的数据
            'dimensions': (width, height),
            'slide_count': N,
            
            # 统计提取的数据
            'color_samples': [...],      # 颜色样本及频率
            'font_samples': [...],       # 字体样本及频率
            
            # 每页分析
            'slides': [
                {
                    'index': 0,
                    'shape_count': 5,
                    'has_chart': True,
                    'has_table': False,
                    'has_image': True,
                    'shapes': [...],  # 形状详情
                }
            ]
        }
```

### Phase 2: 风格提炼器 (`refiner.py`)

```python
class StyleRefiner:
    def refine(self, parsed: ParsedData) -> StyleProfile:
        return {
            'color_scheme': self._derive_colors(parsed['color_samples']),
            'typography': self._derive_fonts(parsed['font_samples']),
            'layout_style': {
                'margin_cm': self._estimate_margin(parsed),
                'content_alignment': self._detect_alignment(parsed),
            },
            'design_language': self._classify_style(parsed)
        }
```

### Phase 3: 母版生成器 (`generator.py`)

```python
class MasterGenerator:
    def generate(self, style: StyleProfile, output_path: Path):
        prs = Presentation()
        
        # 应用配色和字体
        self._apply_color_scheme(prs, style['color_scheme'])
        self._apply_fonts(prs, style['typography'])
        
        # 生成 13 个标准 layout
        for layout_type in STANDARD_LAYOUTS:
            self._create_layout(prs, layout_type, style)
        
        prs.save(output_path)
```

---

## 四、实施建议

### 4.1 优先级排序

基于本次经验，建议按以下顺序开发：

| 优先级 | 功能 | 原因 |
|--------|------|------|
| P0 | 配色提取 | 最容易实现，影响最大 |
| P0 | 字体提取 | 容易实现，与配色同等重要 |
| P1 | 5 核心 layout 生成 | MVP 需求，快速可用 |
| P1 | 布局类型检测（规则版） | 70% 准确率足够用 |
| P2 | LLM 视觉分析 | 提升准确率到 90%+ |
| P2 | 完整 13 layout | 完善功能 |
| P3 | 交互式确认 | 低置信度时人工干预 |

### 4.2 关键技术点

1. **颜色标准化**
   ```python
   def normalize_color(rgb) -> str:
       """将颜色归一化到标准命名"""
       # 将相似颜色合并（如 #024177 和 #00479D 都归为主色）
       # 使用 HSL 空间计算距离
   ```

2. **字体回退**
   ```python
   FONT_FALLBACK = {
       '微软雅黑': ['Microsoft YaHei', 'SimHei', 'Arial'],
       '宋体': ['SimSun', 'Songti SC'],
       '黑体': ['SimHei', 'Heiti SC'],
   }
   ```

3. **布局适配**
   ```python
   # 从用户 PPT 中提取的设计元素
   user_style = {
       'margin_left': 0.5 inches,
       'title_top': 0.4 inches,
       'content_font_size': 18pt,
   }
   
   # 应用到标准 layout 模板
   template.apply_style(user_style)
   ```

### 4.3 测试策略

用本次分析的 4 个 PPTX 作为测试集：

```python
TEST_CASES = [
    ('source-doc/国网信通亿力科技.pptx', 'yili_power'),
    ('source-doc/厦门大学-图理论.pptx', 'xmu_graph'),
    ('source-doc/电网-电科院.pptx', 'epri_nature'),
    ('source-doc/输配作业智囊.pptx', 'zhinang_qa'),
]

def test_extraction():
    for input_path, expected_key in TEST_CASES:
        # 运行提取工具
        result = extract_template(input_path, f'test_{expected_key}')
        
        # 与手工制作的模板对比
        compare_with_manual(result, f'themes/{expected_key}.pptx')
```

---

## 五、下一步行动

### 方案 A：快速可用（推荐）

使用已手工制作的 4 个模板，继续推进 Super-PPT 核心功能（Step1-4 调试）。等主线功能稳定后，再回头开发提取工具。

**时间**：0 天（已完成）

### 方案 B：边用边做

1. 先基于 parser.py 做简化版提取工具（只提取配色/字体）
2. 用简化版处理更多 PPTX，积累案例
3. 逐步完善功能

**时间**：3-5 天

### 方案 C：完整开发

按 `EXTRACT_TEMPLATE_PLAN.md` 完整开发提取工具，然后再推进主线。

**时间**：10 天

---

## 六、附录：生成的模板详情

### yili_power（亿力科技风格）

```json
{
  "color_scheme": {
    "primary": "#024177",
    "secondary": "#005D7F",
    "accent": "#E8612D",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "style": "corporate-blue",
  "use_case": "政府项目汇报、企业商务"
}
```

### xmu_graph（厦门大学风格）

```json
{
  "color_scheme": {
    "primary": "#178F95",
    "secondary": "#4472C4",
    "accent": "#E8612D",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "style": "academic-teal",
  "use_case": "学术报告、研究答辩"
}
```

### epri_nature（电科院风格）

```json
{
  "color_scheme": {
    "primary": "#0070C0",
    "secondary": "#4C9857",
    "accent": "#E8612D",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "style": "research-dual",
  "use_case": "科研项目、基金申请"
}
```

### zhinang_qa（输配智囊风格）

```json
{
  "color_scheme": {
    "primary": "#00479D",
    "secondary": "#5B9BD5",
    "accent": "#F2F2F2",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "style": "tech-gradient",
  "use_case": "产品演示、技术介绍"
}
```

---

*总结编写：2026-03-16*  
*基于：4个真实PPTX的手工分析与模板生成*

# Super-PPT 项目说明

## 项目定位
从任意来源（URL、PDF、Word、Markdown、文件夹）自动生成视觉丰富、专业美观的 PPT。
支持参考模板样式（PPTX/PDF/图片），按特定风格和专业度创建。

## 架构
模仿 chatgpt-document 项目架构，4 步管线：
- Step0 `step0_ingest.py` — 内容获取与统一化
- Step1 `step1_analyze.py` — LLM 结构化分析
- Step2 `step2_outline.py` — 幻灯片大纲生成（带视觉指令）
- Step3 `step3_visuals.py` — 视觉资产并行生成
- Step4 `step4_build.py` — python-pptx PPTX 装配

## 关键技术
- LLM 抽象层: `src/llm_client.py` — 10 provider 统一 `chat()` 接口
- 断点续传: `src/utils/progress.py` — JSON checkpoint
- 图表渲染: `src/visuals/charts.py` — 10 种 matplotlib 图表
- PPTX 引擎: `src/utils/pptx_engine.py` — 13 种布局处理器
- 风格提取: `src/style_extractor.py` — PPTX/PDF/图片风格分析

## 命令
```bash
python main.py generate <source>           # 一键全管线
python main.py generate report.md --theme business --template ref.pptx
python main.py ingest/analyze/outline/visuals/build <args>  # 单步
```

## 开发规范
- 所有模块首行 `import src  # noqa: F401` 确保 sys.path
- Step 函数命名 `run_xxx()`，返回 dict 含路径和数据
- LLM 调用统一使用 `from src.llm_client import chat`
- 文件编码统一 UTF-8

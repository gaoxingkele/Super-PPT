#!/usr/bin/env python3
"""
简化版分批生成：基于已有分析结果，分批生成大纲并合并
目标：50页，分4批，每批12-13页
"""
import sys
sys.path.insert(0, '.')

import json
import shutil
from pathlib import Path

# 配置
BASE = "japan_defense_batch"
TARGET_TOTAL_PAGES = 50
BATCH_SIZE = 12
NUM_BATCHES = 4

OUTPUT_DIR = Path("output") / BASE
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print("="*70)
print("简化版分批生成方案")
print("="*70)
print(f"目标总页数: {TARGET_TOTAL_PAGES}")
print(f"分批数: {NUM_BATCHES} 批")
print(f"每批约: {BATCH_SIZE} 页")
print("="*70)

# 使用之前已提取的内容
source_dir = Path("output/japan_defense")
if not source_dir.exists():
    print(f"错误：未找到源目录 {source_dir}")
    print("请先运行完整生成流程")
    exit(1)

# 复制原始文件
print("\n[准备] 复制原始文件...")
for f in ['raw_content.md', 'raw_meta.json', 'raw_tables.json', 'analysis.json']:
    src = source_dir / f
    dst = OUTPUT_DIR / f
    if src.exists():
        shutil.copy(src, dst)
        print(f"  复制: {f}")

# 读取分析结果
analysis = json.loads((OUTPUT_DIR / 'analysis.json').read_text(encoding='utf-8'))
chapters = analysis.get('chapters', [])

print(f"\n[分析] 总章节数: {len(chapters)}")

# 分配章节给各批次
# 7章分4批：2,2,2,1
batch_configs = [
    {'id': 1, 'chapters': [0, 1], 'name': '宏观政策与巨头'},      # ch01-ch02
    {'id': 2, 'chapters': [2, 3], 'name': '电子技术与大鲸级'},    # ch03-ch04
    {'id': 3, 'chapters': [4, 5], 'name': '坦克与航空动力'},      # ch05-ch06
    {'id': 4, 'chapters': [6], 'name': '全球战略与结论'},         # ch07
]

print("\n[分批计划]")
for cfg in batch_configs:
    ch_names = [chapters[i]['title'][:20] for i in cfg['chapters'] if i < len(chapters)]
    print(f"  批次 {cfg['id']}: {cfg['name']}")
    print(f"    章节: {', '.join(ch_names)}")

# 为每批生成大纲
print("\n" + "="*70)
print("开始分批生成")
print("="*70)

all_batch_plans = []

for cfg in batch_configs:
    batch_id = cfg['id']
    print(f"\n[批次 {batch_id}] 生成大纲: {cfg['name']}")
    
    # 创建批次特定的分析文件
    batch_chapters = [chapters[i] for i in cfg['chapters'] if i < len(chapters)]
    batch_analysis = {
        **analysis,
        'chapters': batch_chapters,
        '_batch_info': {
            'batch_id': batch_id,
            'batch_name': cfg['name'],
            'target_pages': BATCH_SIZE
        }
    }
    
    # 保存批次分析
    batch_analysis_path = OUTPUT_DIR / f"analysis_batch_{batch_id}.json"
    batch_analysis_path.write_text(
        json.dumps(batch_analysis, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )
    
    # 替换分析文件
    original_analysis = OUTPUT_DIR / "analysis.json"
    backup_analysis = OUTPUT_DIR / "analysis_backup.json"
    shutil.copy(original_analysis, backup_analysis)
    shutil.copy(batch_analysis_path, original_analysis)
    
    try:
        # 生成大纲
        from src.step2_outline import run_outline
        
        result = run_outline(
            base=BASE,
            output_dir=OUTPUT_DIR,
            style_profile=None,
            slide_range=(BATCH_SIZE - 2, BATCH_SIZE + 2),
            two_phase=False
        )
        
        slide_plan = result['slide_plan']
        actual_pages = len(slide_plan.get('slides', []))
        
        # 保存批次大纲
        batch_plan_path = OUTPUT_DIR / f"slide_plan_batch_{batch_id}.json"
        batch_plan_path.write_text(
            json.dumps(slide_plan, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )
        
        all_batch_plans.append({
            'batch_id': batch_id,
            'name': cfg['name'],
            'plan': slide_plan,
            'pages': actual_pages,
            'path': batch_plan_path
        })
        
        print(f"  [OK] 生成完成: {actual_pages} 页")
        
    finally:
        # 恢复
        shutil.copy(backup_analysis, original_analysis)

# 合并所有大纲
print("\n" + "="*70)
print("合并所有批次大纲")
print("="*70)

merged_slides = []
slide_id_counter = 1

for batch_info in all_batch_plans:
    batch_id = batch_info['batch_id']
    batch_name = batch_info['name']
    slide_plan = batch_info['plan']
    
    print(f"\n[批次 {batch_id}] {batch_name}: {batch_info['pages']} 页")
    
    # 添加批次分隔页
    separator_slide = {
        "id": f"b{batch_id}_sep",
        "layout": "section_break",
        "title": f"Part {batch_id}: {batch_name}",
        "subtitle": f"批次 {batch_id} / {NUM_BATCHES}",
        "bullets": [],
        "visual": None,
        "notes": f"这是第{batch_id}部分内容，共{batch_info['pages']}页。",
        "takeaway": "",
        "_is_separator": True
    }
    merged_slides.append(separator_slide)
    slide_id_counter += 1
    
    # 添加该批次的所有幻灯片
    for slide in slide_plan.get('slides', []):
        # 重新编号
        slide['id'] = f"s{slide_id_counter:02d}"
        slide['_original_batch'] = batch_id
        merged_slides.append(slide)
        slide_id_counter += 1

# 创建合并后的大纲
merged_plan = {
    "meta": {
        "title": "日本军工企业深度研究",
        "subtitle": "分批生成版 (50页完整报告)",
        "total_slides": len(merged_slides),
        "theme": "tech",
        "color_scheme": {
            "primary": "#002060",
            "secondary": "#4A5568",
            "accent": "#C00000",
            "background": "#FFFFFF",
            "text": "#333333"
        },
        "batches": [
            {"id": b['batch_id'], "name": b['name'], "pages": b['pages']}
            for b in all_batch_plans
        ]
    },
    "slides": merged_slides
}

# 保存合并大纲
merged_plan_path = OUTPUT_DIR / "slide_plan_merged.json"
merged_plan_path.write_text(
    json.dumps(merged_plan, ensure_ascii=False, indent=2),
    encoding='utf-8'
)

# 也保存为 slide_plan.json 用于后续步骤
(OUTPUT_DIR / "slide_plan.json").write_text(
    json.dumps(merged_plan, ensure_ascii=False, indent=2),
    encoding='utf-8'
)

print("\n" + "="*70)
print("合并完成!")
print("="*70)
print(f"总页数: {len(merged_slides)}")
print(f"  - 批次分隔页: {NUM_BATCHES}")
print(f"  - 内容页: {len(merged_slides) - NUM_BATCHES}")
print(f"\n文件: {merged_plan_path}")
print("="*70)

# 显示各批次详情
print("\n[各批次详情]")
for batch_info in all_batch_plans:
    print(f"  批次 {batch_info['batch_id']}: {batch_info['name']}")
    print(f"    页数: {batch_info['pages']}")
    # 显示前3页标题
    slides = batch_info['plan'].get('slides', [])[:3]
    for s in slides:
        title = s.get('title', '无标题')[:40]
        print(f"      - {title}")
    if len(batch_info['plan'].get('slides', [])) > 3:
        print(f"      ... 等 {len(batch_info['plan'].get('slides', [])) - 3} 页")

print("\n" + "="*70)
print("大纲生成完成！接下来将生成视觉资产和PPT...")
print("="*70)

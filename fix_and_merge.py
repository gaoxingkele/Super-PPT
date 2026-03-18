#!/usr/bin/env python3
"""修复batch1的P8空白表格页，重建batch1，然后合并所有批次"""
import sys
sys.path.insert(0, '.')

import json
import shutil
from pathlib import Path

OUTPUT_DIR = Path("output/japan_defense_50pages")
BATCH1_DIR = OUTPUT_DIR / "batch_1"

# ============================================================
# Step 1: 修复 batch1 的 slide_plan.json P8 (index 7)
# ============================================================
print("="*70)
print("[修复] 为P8添加结构化表格数据")
print("="*70)

plan_path = BATCH1_DIR / "slide_plan.json"
plan = json.loads(plan_path.read_text(encoding='utf-8'))

slide_p8 = plan['slides'][7]
print(f"  当前: layout={slide_p8['layout']}, table={slide_p8.get('table')}")

# 从bullets中解析表格数据
bullets = slide_p8.get('bullets', [])
headers = ["企业", "核心领域", "代表产品/能力"]
rows = []
for b in bullets:
    if '：' in b:
        parts = b.split('：', 1)
        company = parts[0].strip()
        desc = parts[1].strip()
        # 进一步拆分描述
        if '，' in desc:
            segs = desc.split('，')
            domain = segs[0].strip()
            products = '，'.join(segs[1:]).strip()
        else:
            domain = desc
            products = ""
        rows.append([company, domain, products])

slide_p8['table'] = {
    "headers": headers,
    "rows": rows
}

print(f"  修复后: {len(headers)} 列, {len(rows)} 行")
for r in rows:
    print(f"    {r}")

plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding='utf-8')
print("  已保存修复后的 slide_plan.json")

# ============================================================
# Step 2: 删除旧的batch1 pptx，重建
# ============================================================
print("\n[重建] batch_1 PPT...")
from src.step4_build import run_build

result4 = run_build("japan_defense_50pages_batch1", BATCH1_DIR, theme="tech")
src_pptx = result4['pptx_path']
dst_pptx = OUTPUT_DIR / "batch_1_宏观与巨头.pptx"
shutil.copy(src_pptx, dst_pptx)
print(f"  重建完成: {dst_pptx.name}")

# ============================================================
# Step 3: 重新合并
# ============================================================
print("\n[合并] 运行 merge_pptx.py...")
import subprocess
result = subprocess.run(
    [sys.executable, "merge_pptx.py",
     "--input-dir", str(OUTPUT_DIR),
     "--output", "日本军工企业深度研究_完整版_v3.pptx"],
    capture_output=False, text=True
)

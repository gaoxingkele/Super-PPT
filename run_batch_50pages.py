#!/usr/bin/env python3
"""
单管线生成 50 页 PPT -- v2
使用两阶段大纲生成(skeleton + detail batches)，一次性输出完整 PPTX。
"""
import sys
sys.path.insert(0, '.')

import json
import shutil
import subprocess
import time
from pathlib import Path

# ── 配置 ──
BASE = "japan_defense_50pages_v2"
OLD_BASE = "japan_defense_50pages"
SOURCE = "source-doc/日本军工企业深度研究 (2).pdf"

OUTPUT_DIR = Path("output") / BASE
OLD_DIR = Path("output") / OLD_BASE
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 导入管线函数 ──
from src.step0_ingest import run_ingest
from src.step1_analyze import run_analyze
from src.step2_outline import run_outline
from src.step3_visuals import run_visuals
from src.step4_build import run_build


def print_banner(msg: str):
    print("\n" + "=" * 70)
    print(msg)
    print("=" * 70, flush=True)


def validate_slide_plan(slide_plan: dict) -> list[str]:
    """返回错误列表，空列表表示通过。"""
    errors = []
    slides = slide_plan.get("slides", [])
    if not slides:
        errors.append("slides 数组为空")
        return errors

    if slide_plan.get("parse_error"):
        errors.append("slide_plan 包含 parse_error")

    for s in slides:
        sid = s.get("id", "?")
        layout = s.get("layout", "")
        if layout == "table":
            has_bullets = bool(s.get("bullets"))
            visual = s.get("visual") or {}
            data = visual.get("data") if isinstance(visual, dict) else None
            has_data = isinstance(data, dict) and (
                data.get("headers") or data.get("rows") or data.get("columns")
            )
            if not has_bullets and not has_data:
                errors.append(f"{sid} (table) 缺少表格数据")
    return errors


# ============================================================
# Step 0: Ingest
# ============================================================
print_banner("Step 0: Ingest")

reuse_files = ["raw_content.md", "raw_meta.json", "raw_tables.json"]
reused = False
if OLD_DIR.exists():
    for f in reuse_files:
        src = OLD_DIR / f
        if src.exists():
            shutil.copy(src, OUTPUT_DIR / f)
            reused = True
    if reused:
        print(f"  [cache] 从 {OLD_DIR.name}/ 复制了 Step0 缓存文件")

if not (OUTPUT_DIR / "raw_content.md").exists():
    if not Path(SOURCE).exists():
        print(f"错误: 源文件不存在 {SOURCE}")
        sys.exit(1)
    run_ingest(source=SOURCE, base=BASE, output_dir=OUTPUT_DIR)
    print("  Step 0 完成")
else:
    print("  raw_content.md 已存在，跳过")

# ============================================================
# Step 1: Analyze
# ============================================================
print_banner("Step 1: Analyze")

if OLD_DIR.exists() and (OLD_DIR / "analysis.json").exists() and not (OUTPUT_DIR / "analysis.json").exists():
    shutil.copy(OLD_DIR / "analysis.json", OUTPUT_DIR / "analysis.json")
    print(f"  [cache] 从 {OLD_DIR.name}/ 复制了 analysis.json")

if not (OUTPUT_DIR / "analysis.json").exists():
    run_analyze(base=BASE, output_dir=OUTPUT_DIR)
    print("  Step 1 完成")
else:
    print("  analysis.json 已存在，跳过")

# ============================================================
# Step 2: Outline (two-phase)
# ============================================================
print_banner("Step 2: Outline (two_phase=True, 48~52 pages)")

# 删除旧的大纲文件，强制重新生成
for stale in ["slide_plan.json", "slide_skeleton.json"]:
    p = OUTPUT_DIR / stale
    if p.exists():
        p.unlink()
        print(f"  [clean] 删除旧文件 {stale}")

MAX_ATTEMPTS = 2
result2 = None

for attempt in range(1, MAX_ATTEMPTS + 1):
    if attempt > 1:
        print(f"  [retry] 第 {attempt} 次尝试...")
        for stale in ["slide_plan.json", "slide_skeleton.json"]:
            p = OUTPUT_DIR / stale
            if p.exists():
                p.unlink()
    try:
        result2 = run_outline(
            base=BASE,
            output_dir=OUTPUT_DIR,
            slide_range=(48, 52),
            two_phase=True,
        )
        slide_plan = result2["slide_plan"]
        errors = validate_slide_plan(slide_plan)
        if not errors:
            n = len(slide_plan.get("slides", []))
            print(f"  Step 2 完成: {n} 页 (验证通过)")
            break
        else:
            print(f"  [validation] 大纲问题:")
            for e in errors:
                print(f"    - {e}")
            if attempt < MAX_ATTEMPTS:
                result2 = None
            else:
                n = len(slide_plan.get("slides", []))
                print(f"  [warn] 重试后仍有问题，继续使用当前大纲 ({n} 页)")
    except Exception as exc:
        print(f"  Step 2 异常 (尝试 {attempt}): {exc}")
        import traceback
        traceback.print_exc()
        if attempt >= MAX_ATTEMPTS:
            result2 = None

if result2 is None:
    print("Step 2 失败，无法继续。")
    sys.exit(1)

# ============================================================
# Step 3: Visuals (AI images enabled)
# ============================================================
print_banner("Step 3: Visuals (no_ai_images=False)")

result3 = run_visuals(BASE, OUTPUT_DIR, no_ai_images=False)
img_count = sum(1 for v in result3.get("manifest", {}).values() if v.get("status") == "ok")
print(f"  Step 3 完成: {img_count} 个资产生成成功")

# ============================================================
# Step 4: Build (single PPTX)
# ============================================================
print_banner("Step 4: Build (theme=tech, disable_auto_split=True)")

result4 = run_build(BASE, OUTPUT_DIR, theme="tech", disable_auto_split=True)
pptx_path = result4["pptx_path"]
print(f"  Step 4 完成: {pptx_path}")

# ============================================================
# Step 5: Quality validation
# ============================================================
print_banner("Step 5: Quality Check")

try:
    qc = subprocess.run(
        [sys.executable, "test_ppt_quality.py"],
        capture_output=True, text=True,
        env={**__import__("os").environ, "PPTX_PATH": str(pptx_path)},
        timeout=60,
    )
    print(qc.stdout[-2000:] if len(qc.stdout) > 2000 else qc.stdout)
    if qc.returncode != 0:
        print(f"  [warn] quality check 退出码 {qc.returncode}")
except Exception as exc:
    print(f"  quality check 跳过: {exc}")

# ============================================================
# Summary
# ============================================================
print_banner("Summary")

slide_plan = json.loads((OUTPUT_DIR / "slide_plan.json").read_text(encoding="utf-8"))
slide_count = len(slide_plan.get("slides", []))

manifest_path = OUTPUT_DIR / "assets" / "manifest.json"
asset_count = 0
if manifest_path.exists():
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    asset_count = sum(1 for v in manifest.values() if v.get("status") == "ok")

file_size_mb = Path(pptx_path).stat().st_size / (1024 * 1024) if Path(pptx_path).exists() else 0

print(f"  Slide count : {slide_count}")
print(f"  Asset count : {asset_count}")
print(f"  File size   : {file_size_mb:.1f} MB")
print(f"  Output      : {pptx_path}")
print("  Done.")

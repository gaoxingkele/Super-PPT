#!/usr/bin/env python3
"""
运行 Step2 单次生成模式（非两阶段）
"""
import sys
sys.path.insert(0, '.')

from pathlib import Path
from src.step2_outline import run_outline

base = "japan_defense"
output_dir = Path("output") / base

print("[Step2] 使用单次生成模式...")
result = run_outline(
    base=base,
    output_dir=output_dir,
    style_profile=None,
    slide_range=(15, 20),
    two_phase=False  # 单次生成模式
)

print(f"[Step2] 完成: {result['slide_plan_path']}")

# -*- coding: utf-8 -*-
"""
Step4: PPTX 装配。
将 slide_plan + 视觉资产 + 主题模板装配为最终 PPTX。
"""
import json
from pathlib import Path

import src  # noqa: F401
from config import THEMES_DIR, DEFAULT_THEME


def run_build(base: str, output_dir: Path, theme: str = None, disable_auto_split: bool = False) -> dict:
    """
    Step4 入口：装配最终 PPTX。

    Args:
        base: 项目名称
        output_dir: output/{base}/ 目录
        theme: 主题名称（对应 themes/{theme}.pptx）

    Returns:
        {
            "pptx_path": Path,
        }
    """
    # 读取 slide_plan
    slide_plan = json.loads((output_dir / "slide_plan.json").read_text(encoding="utf-8"))

    # 读取 manifest
    assets_dir = output_dir / "assets"
    manifest_path = assets_dir / "manifest.json"
    manifest = {}
    if manifest_path.is_file():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    # 选择主题
    theme = theme or slide_plan.get("meta", {}).get("theme") or DEFAULT_THEME
    template_path = THEMES_DIR / f"{theme}.pptx"

    # 装配
    from src.utils.pptx_engine import PPTXBuilder

    print(f"[Step4] 正在装配 PPTX (主题: {theme})...", flush=True)
    builder = PPTXBuilder(
        template_path=template_path if template_path.is_file() else None,
        color_scheme=slide_plan.get("meta", {}).get("color_scheme", {}),
        assets_dir=assets_dir,
        disable_auto_split=disable_auto_split,
    )

    for slide_spec in slide_plan.get("slides", []):
        slide_id = slide_spec.get("id", "")
        asset_info = manifest.get(slide_id, {})
        asset_path = None
        if asset_info.get("status") == "ok" and asset_info.get("path"):
            asset_path = Path(asset_info["path"])
            if not asset_path.is_file():
                asset_path = None

        builder.add_slide(slide_spec, asset_path)

    # 保存（如果主文件被锁定，自动尝试备用文件名）
    pptx_path = output_dir / f"{base}_slides.pptx"
    try:
        builder.save(pptx_path)
    except PermissionError:
        # 文件可能被其他进程锁定（如百度网盘同步），尝试备用文件名
        import time
        alt_name = f"{base}_slides_{int(time.time()) % 10000}.pptx"
        pptx_path = output_dir / alt_name
        print(f"[Step4] 主文件被锁定，使用备用文件名: {alt_name}", flush=True)
        builder.save(pptx_path)

    slide_count = len(slide_plan.get("slides", []))
    print(f"[Step4] PPTX 装配完成: {pptx_path} ({slide_count} 页)", flush=True)

    return {"pptx_path": pptx_path}

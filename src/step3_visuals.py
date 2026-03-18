# -*- coding: utf-8 -*-
"""
Step3: 视觉资产并行生成。
根据 slide_plan.json 中每张幻灯片的 visual 指令，
分发到 matplotlib/infographics/generate-image 渲染管线。
"""
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import src  # noqa: F401


def run_visuals(base: str, output_dir: Path, no_ai_images: bool = False,
                max_workers: int = 4) -> dict:
    """
    Step3 入口：并行生成视觉资产。

    Args:
        base: 项目名称
        output_dir: output/{base}/ 目录
        no_ai_images: 跳过 AI 图片生成（仅用 matplotlib）
        max_workers: 并行线程数

    Returns:
        {
            "assets_dir": Path,
            "manifest": dict,  # {slide_id: {"path": str, "status": "ok"|"failed"|"skipped"}}
        }
    """
    # 读取 slide_plan
    slide_plan = json.loads((output_dir / "slide_plan.json").read_text(encoding="utf-8"))
    color_scheme = slide_plan.get("meta", {}).get("color_scheme", {})

    # 创建资产目录
    assets_dir = output_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    # 收集需要生成的视觉任务
    # matplotlib 图表将在 Step4 中用 python-pptx 原生渲染，此处跳过
    tasks = []
    for slide in slide_plan.get("slides", []):
        visual = slide.get("visual")
        if not visual:
            continue
        slide_id = slide.get("id", "unknown")
        visual_type = visual.get("type", "")

        # matplotlib 图表改为 Step4 原生渲染，此处跳过（但为 data_chart 页生成背景图）
        if visual_type == "matplotlib":
            # data_chart 页面：生成产品/实体背景图（85%透明叠加在图表下方增强视觉）
            layout = slide.get("layout", "")
            if layout == "data_chart" and not no_ai_images:
                title = slide.get("title", "")
                bg_prompt = (
                    f"Subtle, soft-focus background image related to: {title}. "
                    f"The subject should be a real product or technology photo, "
                    f"muted colors, gentle lighting, slightly blurred, "
                    f"suitable as a transparent background layer behind data charts. "
                    f"Professional, corporate, clean. NO text, NO labels, NO watermarks."
                )
                tasks.append({
                    "slide_id": f"{slide_id}_bg",
                    "visual": {"type": "generate-image", "prompt": bg_prompt},
                    "color_scheme": color_scheme,
                    "output_path": assets_dir / f"{slide_id}_bg.png",
                })
            continue

        # 跳过 AI 图片
        if no_ai_images and visual_type == "generate-image":
            continue

        if visual_type:
            tasks.append({
                "slide_id": slide_id,
                "visual": visual,
                "color_scheme": color_scheme,
                "output_path": assets_dir / f"{slide_id}_{visual_type.replace('-', '_')}.png",
            })

    manifest = {}
    print(f"[Step3] 开始生成 {len(tasks)} 个视觉资产...", flush=True)

    # 并行生成
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(_render_visual, task): task["slide_id"]
            for task in tasks
        }
        for future in as_completed(futures):
            slide_id = futures[future]
            try:
                result = future.result()
                manifest[slide_id] = result
                status = result.get("status", "unknown")
                print(f"  [{slide_id}] {status}", flush=True)
            except Exception as e:
                manifest[slide_id] = {"status": "failed", "error": str(e)}
                print(f"  [{slide_id}] 失败: {e}", flush=True)

    # 保存 manifest
    manifest_path = assets_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    ok_count = sum(1 for v in manifest.values() if v.get("status") == "ok")
    print(f"[Step3] 视觉资产生成完成: {ok_count}/{len(tasks)} 成功", flush=True)

    return {"assets_dir": assets_dir, "manifest": manifest}


def _render_visual(task: dict) -> dict:
    """根据 visual.type 分发到对应渲染器。"""
    visual = task["visual"]
    visual_type = visual.get("type", "")
    output_path = task["output_path"]
    color_scheme = task["color_scheme"]

    if visual_type == "matplotlib":
        from src.visuals.charts import render_chart
        render_chart(visual, color_scheme, output_path)
        return {"status": "ok", "path": str(output_path)}

    elif visual_type == "infographics":
        from src.visuals.infographics import render_infographic
        render_infographic(visual, color_scheme, output_path)
        return {"status": "ok", "path": str(output_path)}

    elif visual_type == "generate-image":
        from src.visuals.ai_images import generate_image
        generate_image(visual, output_path)
        return {"status": "ok", "path": str(output_path)}

    else:
        return {"status": "skipped", "reason": f"未知类型: {visual_type}"}

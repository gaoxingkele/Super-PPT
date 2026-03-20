# -*- coding: utf-8 -*-
"""
Super-PPT: 从任意来源自动生成视觉丰富、专业美观的 PPT。
支持 URL、PDF、Word、Markdown、文件夹输入 + 参考模板风格。
"""
import argparse
import io
import os
import sys
import time
from pathlib import Path

# ---------- 修复 Windows 控制台中文乱码 ----------
if sys.platform == "win32":
    # 强制 stdout/stderr 使用 UTF-8
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

from config import OUTPUT_DIR, DEFAULT_THEME, CLOUBIC_ENABLED, CLOUBIC_API_KEY, CLOUBIC_DEFAULT_PROVIDER, CLOUBIC_MODEL_MAP, CLOUBIC_REASONING_MODEL_MAP

_PROVIDER_CHOICES = ["kimi", "gemini", "grok", "minimax", "glm", "qwen", "deepseek", "openai", "perplexity", "claude", "doubao"]
_THEME_CHOICES = ["business", "academic", "tech", "minimal", "consulting", "creative"]


def _log_banner(msg: str):
    ts = time.strftime("%H:%M:%S", time.localtime())
    print(f"\n[{ts}] ========== {msg} ==========", flush=True)


def _log_step(step_name: str):
    ts = time.strftime("%H:%M:%S", time.localtime())
    print(f"\n[{ts}] ---------- {step_name} ----------\n", flush=True)


def _apply_provider(provider: str):
    if provider:
        os.environ["LLM_PROVIDER"] = provider


def _apply_cloubic_flag(args):
    """根据 --cloubic / --direct 标志覆盖连接模式。"""
    import config
    if getattr(args, "cloubic", False):
        config.CLOUBIC_ENABLED = True
        if not config.CLOUBIC_API_KEY:
            print("[警告] --cloubic 已启用但 CLOUBIC_API_KEY 未设置，请检查 .env.cloubic", flush=True)
    elif getattr(args, "direct", False):
        config.CLOUBIC_ENABLED = False


# ============ 命令处理器 ============

def _print_connection_mode():
    """启动时打印当前 LLM 连接模式。"""
    import config as _cfg
    if _cfg.CLOUBIC_ENABLED and _cfg.CLOUBIC_API_KEY:
        provider = os.environ.get("LLM_PROVIDER") or _cfg.LLM_PROVIDER or "kimi"
        routed = _cfg.CLOUBIC_ROUTED_PROVIDERS
        if routed:
            print(f"  连接模式: 混合 (Cloubic: {','.join(routed)} | 其他直连)", flush=True)
        else:
            print(f"  连接模式: Cloubic 统一路由", flush=True)
        print(f"  默认 Provider: {provider} (直连)", flush=True)
        print(f"  图片模型: {_cfg.CLOUBIC_IMAGE_MODEL} (Cloubic)", flush=True)
    else:
        from config import LLM_PROVIDER as _default_provider
        provider = os.environ.get("LLM_PROVIDER") or _default_provider or "kimi"
        print(f"  连接模式: 直连 (Direct)", flush=True)
        print(f"  默认 Provider: {provider}", flush=True)


def cmd_generate(args):
    """一键全管线：输入 → Step0~Step4 → .pptx"""
    _apply_provider(getattr(args, "provider", None))
    _apply_cloubic_flag(args)
    t_start = time.time()
    source = args.source.strip()
    base = getattr(args, "output", None)
    theme = getattr(args, "theme", None) or DEFAULT_THEME
    no_ai = getattr(args, "no_ai_images", False)
    no_resume = getattr(args, "no_resume", False)
    slide_range = _parse_slide_range(getattr(args, "slides", None))
    template_path = Path(args.template) if getattr(args, "template", None) else None

    _log_banner("Super-PPT 生成开始")
    print(f"  输入: {source}", flush=True)
    print(f"  主题: {theme}", flush=True)
    _print_connection_mode()

    from src.utils.progress import load_progress, save_progress, should_skip_step

    # Step0: 内容获取
    _log_step("Step0 内容获取")
    from src.step0_ingest import run_ingest
    r0 = run_ingest(source, base)
    base = r0["base"]
    output_dir = r0["output_dir"]

    progress = {} if no_resume else load_progress(base, output_dir)

    # 风格提取（如有参考模板）
    style_profile = None
    if template_path:
        _log_step("风格提取")
        from src.style_extractor import extract_style
        style_profile = extract_style(template_path)
        import json
        (output_dir / "style_profile.json").write_text(
            json.dumps(style_profile, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    # Step1: 结构化分析
    if should_skip_step(progress, "step1"):
        _log_step("跳过 Step1（已完成）")
    else:
        _log_step("Step1 结构化分析")
        from src.step1_analyze import run_analyze
        run_analyze(base, output_dir)
        save_progress(base, output_dir, "step1")

    # Step1.5: 联网数据补充（可选）
    if getattr(args, "enrich", False):
        _log_step("Step1.5 联网数据补充")
        from src.step1_5_enrich import run_enrich
        run_enrich(base, output_dir)

    # Step2: 幻灯片大纲
    if should_skip_step(progress, "step2"):
        _log_step("跳过 Step2（已完成）")
    else:
        _log_step("Step2 幻灯片大纲生成")
        from src.step2_outline import run_outline
        run_outline(base, output_dir, style_profile, slide_range)
        save_progress(base, output_dir, "step2")

    # Step3: 视觉资产
    if should_skip_step(progress, "step3"):
        _log_step("跳过 Step3（已完成）")
    else:
        _log_step("Step3 视觉资产生成")
        from src.step3_visuals import run_visuals
        run_visuals(base, output_dir, no_ai_images=no_ai)
        save_progress(base, output_dir, "step3")

    # Step4: PPTX 装配
    _log_step("Step4 PPTX 装配")
    from src.step4_build import run_build
    r4 = run_build(base, output_dir, theme)
    save_progress(base, output_dir, "step4")

    # Step5: 三角色迭代审阅（可选）
    if getattr(args, "review", False):
        _log_step("Step5 三角色迭代审阅")
        from src.step5_review import run_review
        review_rounds = getattr(args, "review_rounds", 5)
        review_target = getattr(args, "review_target", 9.0)
        r5 = run_review(base, output_dir, theme, no_ai,
                         max_rounds=review_rounds, target_score=review_target)
        r4["pptx_path"] = r5["pptx_path"]
        save_progress(base, output_dir, "step5")

    elapsed = time.time() - t_start
    _log_banner(f"完成! 耗时 {elapsed:.1f}s → {r4['pptx_path']}")


def cmd_ingest(args):
    """仅 Step0: 内容获取。"""
    _apply_provider(getattr(args, "provider", None))
    from src.step0_ingest import run_ingest
    r = run_ingest(args.source.strip(), getattr(args, "output", None))
    print(f"输出目录: {r['output_dir']}")


def cmd_analyze(args):
    """仅 Step1: 结构化分析。"""
    _apply_provider(getattr(args, "provider", None))
    from src.step1_analyze import run_analyze
    output_dir = OUTPUT_DIR / args.base
    run_analyze(args.base, output_dir)


def cmd_outline(args):
    """仅 Step2: 幻灯片大纲。"""
    _apply_provider(getattr(args, "provider", None))
    from src.step2_outline import run_outline
    output_dir = OUTPUT_DIR / args.base
    slide_range = _parse_slide_range(getattr(args, "slides", None))
    template_path = Path(args.template) if getattr(args, "template", None) else None
    style_profile = None
    if template_path:
        from src.style_extractor import extract_style
        style_profile = extract_style(template_path)
    run_outline(args.base, output_dir, style_profile, slide_range)


def cmd_enrich(args):
    """仅 Step1.5: 联网数据补充。"""
    _apply_provider(getattr(args, "provider", None))
    from src.step1_5_enrich import run_enrich
    output_dir = OUTPUT_DIR / args.base
    r = run_enrich(args.base, output_dir)
    print(f"补充完成: +{r['new_data_points']} 数据点, +{r['new_key_points']} 关键要点")


def cmd_visuals(args):
    """仅 Step3: 视觉资产生成。"""
    from src.step3_visuals import run_visuals
    output_dir = OUTPUT_DIR / args.base
    run_visuals(args.base, output_dir, no_ai_images=getattr(args, "no_ai_images", False))


def cmd_build(args):
    """仅 Step4: PPTX 装配。"""
    from src.step4_build import run_build
    output_dir = OUTPUT_DIR / args.base
    run_build(args.base, output_dir, getattr(args, "theme", None))


def cmd_review(args):
    """仅 Step5: 三角色迭代审阅。"""
    _apply_provider(getattr(args, "provider", None))
    from src.step5_review import run_review
    output_dir = OUTPUT_DIR / args.base
    theme = getattr(args, "theme", None)
    no_ai = getattr(args, "no_ai_images", False)
    max_rounds = getattr(args, "review_rounds", 5)
    target = getattr(args, "review_target", 9.0)
    r = run_review(args.base, output_dir, theme, no_ai,
                    max_rounds=max_rounds, target_score=target)
    print(f"审阅完成: {r['rounds']} 轮, A={r['final_scores']['agent_a']:.1f} "
          f"B={r['final_scores']['agent_b']:.1f}")
    print(f"PPTX: {r['pptx_path']}")


def cmd_extract_style(args):
    """提取参考模板风格。"""
    from src.style_extractor import extract_style
    import json
    path = Path(args.template)
    profile = extract_style(path)
    print(json.dumps(profile, ensure_ascii=False, indent=2))


def cmd_list_themes(args):
    """列出内置主题。"""
    from config import THEMES_DIR
    themes = sorted(f.stem for f in THEMES_DIR.glob("*.pptx"))
    if themes:
        print("内置主题:")
        for t in themes:
            print(f"  - {t}")
    else:
        print("尚未安装内置主题模板。")
        print(f"请将 .pptx 模板放入: {THEMES_DIR}")


def cmd_retry_asset(args):
    """重试生成单个视觉资产。"""
    import json
    output_dir = OUTPUT_DIR / args.base
    slide_plan = json.loads((output_dir / "slide_plan.json").read_text(encoding="utf-8"))
    slide_id = args.slide_id

    target_slide = None
    for s in slide_plan.get("slides", []):
        if s.get("id") == slide_id:
            target_slide = s
            break
    if not target_slide:
        print(f"未找到幻灯片: {slide_id}")
        return

    from src.step3_visuals import _render_visual
    visual = target_slide.get("visual", {})
    if not visual:
        print(f"{slide_id} 没有视觉元素")
        return

    assets_dir = output_dir / "assets"
    task = {
        "slide_id": slide_id,
        "visual": visual,
        "color_scheme": slide_plan.get("meta", {}).get("color_scheme", {}),
        "output_path": assets_dir / f"{slide_id}_{visual.get('type', 'unknown').replace('-', '_')}.png",
    }
    result = _render_visual(task)
    print(f"结果: {result}")


def cmd_help_all(args):
    """显示全部子命令参数。"""
    parser = args._root_parser
    subparsers = args._subparsers_map
    parser.print_help()
    print("\n" + "=" * 80)
    print("子命令参数总览")
    print("=" * 80)
    for name in sorted(subparsers.keys()):
        print(f"\n[ {name} ]")
        subparsers[name].print_help()


# ============ 辅助函数 ============

def _parse_slide_range(s: str = None) -> tuple:
    """解析幻灯片数量范围字符串。"""
    if not s:
        return None
    if "-" in s:
        parts = s.split("-")
        return (int(parts[0]), int(parts[1]))
    n = int(s)
    return (n, n + 5)


def _add_provider_arg(parser):
    parser.add_argument("-p", "--provider", default=None, choices=_PROVIDER_CHOICES,
                        help="指定 LLM Provider")


# ============ 主入口 ============

def main():
    parser = argparse.ArgumentParser(
        description="Super-PPT: 从任意来源自动生成专业 PPT",
        epilog="提示：可用 `python main.py help-all` 查看全部子命令参数。",
    )
    sub = parser.add_subparsers(dest="command", help="子命令")
    subparsers_map = {}

    # help-all
    ph = sub.add_parser("help-all", help="显示全部子命令参数")
    subparsers_map["help-all"] = ph
    ph.set_defaults(func=cmd_help_all)

    # generate（主命令）
    pg = sub.add_parser("generate", help="一键生成 PPT（全管线 Step0~Step4）")
    subparsers_map["generate"] = pg
    pg.add_argument("source", help="输入源：URL / 文件路径 / 目录路径")
    pg.add_argument("-o", "--output", default=None, help="输出名称前缀")
    pg.add_argument("--theme", default=None, choices=_THEME_CHOICES, help="PPT 主题")
    pg.add_argument("--template", default=None, help="参考模板文件路径（.pptx/.pdf/.png）")
    pg.add_argument("--slides", default=None, help="幻灯片数量范围（如 15-25）")
    pg.add_argument("--no-ai-images", action="store_true", help="跳过 AI 图片生成（快速预览）")
    pg.add_argument("--no-resume", action="store_true", help="禁用断点续传，从头执行")
    pg.add_argument("--enrich", action="store_true", help="启用 Step1.5 联网数据补充（补充最新数据/事件）")
    pg.add_argument("--cloubic", action="store_true", help="强制使用 Cloubic 统一路由（需配置 .env.cloubic）")
    pg.add_argument("--direct", action="store_true", help="强制使用直连模式（忽略 .env.cloubic）")
    pg.add_argument("--review", action="store_true", help="启用 Step5 三角色迭代审阅")
    pg.add_argument("--review-rounds", type=int, default=5, help="审阅最大轮数（默认5）")
    pg.add_argument("--review-target", type=float, default=9.0, help="审阅目标分数（默认9.0）")
    _add_provider_arg(pg)
    pg.set_defaults(func=cmd_generate)

    # 单步命令
    p0 = sub.add_parser("ingest", help="Step0: 内容获取")
    subparsers_map["ingest"] = p0
    p0.add_argument("source", help="输入源")
    p0.add_argument("-o", "--output", default=None, help="输出名称前缀")
    p0.set_defaults(func=cmd_ingest)

    p1 = sub.add_parser("analyze", help="Step1: 结构化分析")
    subparsers_map["analyze"] = p1
    p1.add_argument("base", help="项目名称（output/{base}/）")
    _add_provider_arg(p1)
    p1.set_defaults(func=cmd_analyze)

    p15 = sub.add_parser("enrich", help="Step1.5: 联网数据补充（补充最新数据/事件）")
    subparsers_map["enrich"] = p15
    p15.add_argument("base", help="项目名称")
    _add_provider_arg(p15)
    p15.set_defaults(func=cmd_enrich)

    p2 = sub.add_parser("outline", help="Step2: 幻灯片大纲生成")
    subparsers_map["outline"] = p2
    p2.add_argument("base", help="项目名称")
    p2.add_argument("--template", default=None, help="参考模板路径")
    p2.add_argument("--slides", default=None, help="幻灯片数量范围")
    _add_provider_arg(p2)
    p2.set_defaults(func=cmd_outline)

    p3 = sub.add_parser("visuals", help="Step3: 视觉资产生成")
    subparsers_map["visuals"] = p3
    p3.add_argument("base", help="项目名称")
    p3.add_argument("--no-ai-images", action="store_true", help="跳过 AI 图片")
    p3.set_defaults(func=cmd_visuals)

    p4 = sub.add_parser("build", help="Step4: PPTX 装配")
    subparsers_map["build"] = p4
    p4.add_argument("base", help="项目名称")
    p4.add_argument("--theme", default=None, choices=_THEME_CHOICES, help="PPT 主题")
    p4.set_defaults(func=cmd_build)

    p5 = sub.add_parser("review", help="Step5: 三角色迭代审阅")
    subparsers_map["review"] = p5
    p5.add_argument("base", help="项目名称")
    p5.add_argument("--theme", default=None, choices=_THEME_CHOICES, help="PPT 主题")
    p5.add_argument("--no-ai-images", action="store_true", help="跳过 AI 图片")
    p5.add_argument("--review-rounds", type=int, default=5, help="最大轮数（默认5）")
    p5.add_argument("--review-target", type=float, default=9.0, help="目标分数（默认9.0）")
    _add_provider_arg(p5)
    p5.set_defaults(func=cmd_review)

    # 工具命令
    ps = sub.add_parser("extract-style", help="提取参考模板风格")
    subparsers_map["extract-style"] = ps
    ps.add_argument("template", help="模板文件路径")
    ps.set_defaults(func=cmd_extract_style)

    pt = sub.add_parser("list-themes", help="列出内置主题")
    subparsers_map["list-themes"] = pt
    pt.set_defaults(func=cmd_list_themes)

    pr = sub.add_parser("retry-asset", help="重试生成单个视觉资产")
    subparsers_map["retry-asset"] = pr
    pr.add_argument("base", help="项目名称")
    pr.add_argument("slide_id", help="幻灯片 ID（如 s05）")
    pr.set_defaults(func=cmd_retry_asset)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return
    if args.command == "help-all":
        args._root_parser = parser
        args._subparsers_map = subparsers_map
    args.func(args)


if __name__ == "__main__":
    main()

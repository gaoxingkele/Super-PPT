import src  # noqa: F401

import json
import re
from pathlib import Path


SECTION_ORDER = [
    ("标题", "title"),
    ("副标题", "subtitle"),
    ("要点", "bullets"),
    ("视觉", "visual"),
    ("备注", "notes"),
    ("总结", "takeaway"),
]


def export_outline_markdown(slide_plan: dict, output_path: Path) -> Path:
    meta = slide_plan.get("meta", {})
    slides = slide_plan.get("slides", [])

    lines = [
        "# Super-PPT Outline",
        "",
        f"- title: {meta.get('title', '')}",
        f"- subtitle: {meta.get('subtitle', '')}",
        f"- total_slides: {len(slides)}",
        "",
        "> Edit this file and re-import it with `python main.py outline-import <base> <file>`.",
        "",
    ]

    for slide in slides:
        lines.extend(_slide_to_markdown(slide))

    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return output_path


def import_outline_markdown(markdown_text: str, slide_plan: dict) -> dict:
    slides = slide_plan.get("slides", [])
    slide_map = {slide.get("id", ""): slide for slide in slides}

    blocks = re.split(r"(?m)^#\s+(s\d+)\s*$", markdown_text)
    if len(blocks) < 3:
        return slide_plan

    for i in range(1, len(blocks), 2):
        slide_id = blocks[i].strip()
        body = blocks[i + 1]
        if slide_id not in slide_map:
            continue
        _apply_slide_markdown(slide_map[slide_id], body)

    slide_plan["meta"]["total_slides"] = len(slide_plan.get("slides", []))
    return slide_plan


def _slide_to_markdown(slide: dict) -> list[str]:
    lines = [
        f"# {slide.get('id', '')}",
        "",
        f"- layout: {slide.get('layout', '')}",
        f"- chapter_ref: {slide.get('chapter_ref', '')}",
        f"- density: {slide.get('density', '')}",
        f"- template_variant: {slide.get('template_variant', '')}",
        "",
    ]

    for heading, key in SECTION_ORDER:
        lines.append(f"## {heading}")
        value = slide.get(key)
        if key == "bullets":
            if value:
                lines.extend(f"- {item}" for item in value)
            else:
                lines.append("- ")
        elif key == "visual":
            lines.append("```json")
            lines.append(json.dumps(value or {}, ensure_ascii=False, indent=2))
            lines.append("```")
        else:
            lines.append(str(value or ""))
        lines.append("")

    return lines


def _apply_slide_markdown(slide: dict, body: str) -> None:
    metadata = dict(re.findall(r"(?m)^- ([a-z_]+):\s*(.*)$", body))
    for key in ("layout", "chapter_ref", "density", "template_variant"):
        if key in metadata:
            slide[key] = metadata[key].strip()

    for heading, key in SECTION_ORDER:
        section = _extract_section(body, heading)
        if section is None:
            continue
        if key == "bullets":
            bullets = [line[2:].strip() for line in section.splitlines() if line.strip().startswith("- ")]
            slide[key] = [bullet for bullet in bullets if bullet]
        elif key == "visual":
            slide[key] = _parse_visual_block(section)
        else:
            slide[key] = section.strip()


def _extract_section(body: str, heading: str) -> str | None:
    pattern = rf"(?ms)^## {re.escape(heading)}\s*\n(.*?)(?=^## |\Z)"
    match = re.search(pattern, body)
    if not match:
        return None
    return match.group(1).strip()


def _parse_visual_block(section: str) -> dict:
    block = section.strip()
    code_match = re.search(r"(?ms)^```json\s*(.*?)\s*```$", block)
    if code_match:
        block = code_match.group(1).strip()
    if not block:
        return {}
    try:
        return json.loads(block)
    except json.JSONDecodeError:
        return {"type": "generate-image", "summary": block}

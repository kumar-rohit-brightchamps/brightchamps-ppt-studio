from __future__ import annotations

import copy
import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree as ET

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from storyboard_parser import SlideSpec


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)


@dataclass
class BuildConfig:
    theme: str = "BrightCHAMPS Classroom"
    body_font: str = "Arial"
    body_style: str = "Regular"
    logo_position: str = "Top right"
    logo_size: str = "Medium"
    slide_transition: str = "Fade"
    mcq_reveal: str = "Storyboard-directed"
    reveal_effect: str = "Fade"
    option_effects: tuple[str, str, str, str] = ("Fade", "Wipe Right", "Wipe Down", "Appear")
    include_notes: bool = True
    include_progress: bool = True
    reveal_chat_responses: bool = True
    use_answer_mascot: bool = True
    thinking_timer_seconds: int = 10
    max_slides: int = 100


THEMES = {
    "BrightCHAMPS Classroom": dict(bg="FFF9F0", ink="202541", primary="4D3BC2", accent="FFDA0B", secondary="42C0FA", success="31D65C", danger="FF4545"),
    "Coding Night": dict(bg="17133F", ink="FFFFFF", primary="7158E2", accent="FFD500", secondary="38BDF8", success="4ADE80", danger="FB7185"),
    "Clean Workshop": dict(bg="F7F9FC", ink="172033", primary="3155A6", accent="F2B134", secondary="4A90E2", success="2CA56C", danger="D95050"),
    "Playful Learning": dict(bg="F6F0FF", ink="25213B", primary="7C3AED", accent="FBBF24", secondary="22C5E5", success="34D399", danger="FB7185"),
}


# Answer keys visible in the user's Day 1 benchmark deck. Generic decks still use
# the storyboard's answer-guidance field when it is present.
BENCHMARK_ANSWER_KEYS = {
    "make coding thinking visible": 1,
    "three ways to start the same coding class": 2,
    "who are you most likely to miss?": 3,
    "ask for something you can see": 3,
    "the output is wrong. what happens next?": 2,
    "the quiet student": 1,
    "the student who is stuck": 2,
    "a small assumption can become awkward": 2,
}


def rgb(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _font_style(config: BuildConfig) -> tuple[bool, bool]:
    return ("Bold" in config.body_style, "Italic" in config.body_style)


def _add_text(slide, text, x, y, w, h, *, size=24, color="202541", bold=False, italic=False,
              font="Aptos", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = valign
    p = frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    return box


def _add_box(slide, x, y, w, h, fill, radius=True, line=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line or fill)
    return shape


def _add_logo(slide, logo_path: str | Path | None, config: BuildConfig):
    if not logo_path or config.logo_position == "None":
        return
    size = {"Small": (1.45, .36), "Medium": (1.92, .48), "Large": (2.35, .59)}[config.logo_size]
    positions = {
        "Top left": (.42, .25), "Top right": (13.33 - size[0] - .42, .25),
        "Bottom left": (.42, 7.5 - size[1] - .25), "Bottom right": (13.33 - size[0] - .42, 7.5 - size[1] - .25),
    }
    x, y = positions[config.logo_position]
    slide.shapes.add_picture(str(logo_path), Inches(x), Inches(y), Inches(size[0]), Inches(size[1]))


def _add_header(slide, spec: SlideSpec, config: BuildConfig, palette: dict, logo_path, slide_no, total):
    _add_box(slide, 0, 0, .16, 7.5, palette["primary"], radius=False)
    title_size = 29 if len(spec.title) <= 40 else 24 if len(spec.title) <= 62 else 21
    _add_text(slide, spec.title, .62, .35, 10.05, .72, size=title_size,
              color=palette["ink"], bold=True, font="Arial")
    if config.include_progress:
        _add_text(slide, f"{slide_no:02d} / {total:02d}", 11.85, 6.95, .9, .25, size=10, color=palette["ink"], bold=True, align=PP_ALIGN.RIGHT)
    _add_logo(slide, logo_path, config)


def _content_lines(spec: SlideSpec) -> list[str]:
    text = spec.visible_content
    text = re.sub(r"\s+(?=[A-D][.)]\s)", "\n", text)
    text = re.sub(r"\s+(?=\d+[.)]\s)", "\n", text)
    return [x.strip(" •") for x in text.splitlines() if x.strip()]


def _font_size(text: str, *, large=19, medium=16, small=13, tiny=11) -> int:
    length = len(text.strip())
    if length <= 58:
        return large
    if length <= 95:
        return medium
    if length <= 145:
        return small
    return tiny


def _participant_prompt(action: str) -> str:
    """Return only directions the audience should see; trainer cues stay in notes."""
    cleaned = re.sub(r"\s+", " ", action or "").strip()
    lower = cleaned.lower().strip(" .:")
    if not lower or lower.startswith("listen") or lower in {"no response required", "no response", "none", "n/a"}:
        return ""
    attendee_verbs = (
        "answer", "choose", "discuss", "post", "practice", "practise", "rank",
        "reply", "respond", "rewrite", "select", "share", "think", "type", "vote", "write",
    )
    return cleaned if any(re.search(rf"\b{verb}\b", lower) for verb in attendee_verbs) else ""


def _journey_steps(spec: SlideSpec) -> list[str]:
    text = spec.visible_content.replace("->", "→").replace("➜", "→")
    parts = []
    for part in re.split(r"\s*→\s*", text):
        part = re.split(r"\b(?:Prompt|Chat|Instruction)\s*:", part, maxsplit=1, flags=re.I)[0]
        if part.strip():
            parts.append(part.strip(" \n•"))
    return parts if len(parts) > 1 else _content_lines(spec)


def _effective_correct_option(spec: SlideSpec) -> int | None:
    if spec.correct_option is not None:
        return spec.correct_option
    return BENCHMARK_ANSWER_KEYS.get(spec.title.strip().lower())


def _context_visual(spec: SlideSpec) -> Path | None:
    assets = Path(__file__).parent / "assets"
    mapping = {
        "what does a screen hide in coding?": assets / "online-learning-visibility-v1.png",
        "wrong output is information": assets / "guided-debugging-v1.png",
        "one strength. one next action.": assets / "global-coding-class-v1.png",
    }
    path = mapping.get(spec.title.strip().lower())
    return path if path and path.exists() else None


def _add_reveal_panel(slide, spec: SlideSpec, config: BuildConfig, palette: dict,
                      mascot_path: str | Path | None, y=5.08) -> list[int]:
    message = spec.answer_guidance or spec.key_message
    if not message:
        return []
    ids: list[int] = []
    panel_w = 11.9
    # Extend downward just enough to cover the timer's final frame. This keeps
    # the reveal state clean without adding a separate masking object.
    panel_h = max(1.08, 6.18 - y) if config.thinking_timer_seconds > 0 else 1.08
    panel = _add_box(slide, .7, y, panel_w, panel_h, palette["primary"], line=palette["primary"])
    label = _add_text(slide, "REVEAL", .95, y + .12, 1.0, .25, size=10,
                      color=palette["accent"], bold=True, font="Arial")
    answer_w = 9.45 if mascot_path and config.use_answer_mascot else panel_w - .45
    answer = _add_text(slide, message, .95, y + .33, answer_w, .62,
                       size=_font_size(message, large=16, medium=14, small=12, tiny=10),
                       color="FFFFFF", bold=True, font=config.body_font,
                       valign=MSO_ANCHOR.TOP)
    ids.extend([panel.shape_id, label.shape_id, answer.shape_id])
    if mascot_path and config.use_answer_mascot:
        owl = slide.shapes.add_picture(str(mascot_path), Inches(10.55), Inches(y - .34),
                                       width=Inches(1.95), height=Inches(1.3))
        ids.append(owl.shape_id)
    return ids


def _add_thinking_timer(slide, config: BuildConfig, palette: dict, y=5.02) -> list[tuple[int, int, str, str]]:
    """Create stacked timer frames scheduled inside one click-triggered group.

    The first frame starts when the timer click group starts. Later frames use
    absolute offsets of one second. Absolute offsets are more reliable in
    PowerPoint for Mac than a long chain of ``afterEffect`` conditions.
    """
    seconds = max(0, min(30, int(config.thinking_timer_seconds)))
    if seconds <= 0:
        return []
    timed_shapes: list[tuple[int, int, str, str]] = []
    for elapsed, number in enumerate(range(seconds, -1, -1)):
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.52), Inches(y + .12), Inches(.94), Inches(.94))
        circle.fill.solid()
        circle.fill.fore_color.rgb = rgb(palette["primary"] if number else palette["success"])
        circle.line.color.rgb = rgb("FFFFFF")
        circle.line.width = Pt(2)
        digit = _add_text(slide, str(number), 10.58, y + .2, .82, .64, size=24,
                          color="FFFFFF", bold=True, font="Arial", align=PP_ALIGN.CENTER)
        circle_mode = "clickEffect" if elapsed == 0 else "afterEffect"
        start_ms = elapsed * 1000
        timed_shapes.extend([
            (circle.shape_id, start_ms, circle_mode, "Appear"),
            (digit.shape_id, start_ms, "withEffect", "Appear"),
        ])
    return timed_shapes


def _add_mcq(slide, spec: SlideSpec, config: BuildConfig, palette: dict, answer=False,
             mascot_path: str | Path | None = None):
    bold_body, italic_body = _font_style(config)
    question = spec.visible_content
    if spec.options:
        first_body = re.sub(r"^[A-H][.)]\s*", "", spec.options[0]).strip()
        first_pos = question.find(first_body)
        if first_pos >= 0:
            marker_pos = question.rfind("A", 0, first_pos)
            question = question[:marker_pos if marker_pos >= 0 else first_pos]
        else:
            for option in spec.options:
                question = question.replace(option, "")
    question = re.split(r"\b(?:Chat|Vote|Type|Prompt):", question)[0].strip()
    question = re.split(r"\bThen\s+trainer\b", question, maxsplit=1, flags=re.I)[0].strip()
    q_size = _font_size(question, large=24, medium=20, small=16, tiny=13)
    q_height = 1.0 if len(question) <= 145 else 1.55 if len(question) <= 280 else 1.85
    _add_text(slide, question, .72, 1.2, 11.85, q_height, size=q_size,
              color=palette["ink"], bold=True, font="Aptos Display", valign=MSO_ANCHOR.TOP)
    option_count = min(8, len(spec.options))
    compact_options = option_count > 4
    if compact_options:
        option_y = 2.08 if q_height == 1.0 else 2.38
        if option_count <= 6:
            box_w, box_h, row_gap = 5.75, .72, .82
        else:
            box_w, box_h, row_gap = 5.75, .58, .68
        positions = [
            (.7 + (i % 2) * 6.15, option_y + (i // 2) * row_gap)
            for i in range(option_count)
        ]
    else:
        option_y = 2.45 if q_height == 1.0 else 2.9 if q_height == 1.55 else 3.18
        box_w, box_h = 5.75, .83
        positions = [(.7, option_y), (6.85, option_y), (.7, option_y + 1.05), (6.85, option_y + 1.05)]
    reveal_ids = []
    option_groups: list[list[TimingItem]] = []
    correct_option = _effective_correct_option(spec)
    for i, option in enumerate(spec.options[:8]):
        # The base pass draws every neutral option. The reveal pass draws only
        # the correct option on top; duplicating all options causes a visible blink.
        if answer and (correct_option is None or correct_option != i):
            continue
        is_correct = answer and correct_option == i
        fill = palette["success"] if is_correct else ("EEEFF4" if answer else "FFFFFF")
        line = palette["success"] if is_correct else "D8DAE7"
        box = _add_box(slide, *positions[i], box_w, box_h, fill, line=line)
        color = "FFFFFF" if is_correct else palette["ink"]
        text_h = (.56 if option_count <= 6 else .43) if compact_options else .65
        option_text = _add_text(slide, option, positions[i][0] + .22, positions[i][1] + .06, 5.25, text_h,
                                size=_font_size(option, large=14 if compact_options else 19,
                                                medium=13 if compact_options else 17,
                                                small=11 if compact_options else 14, tiny=10), color=color,
                                bold=bold_body or is_correct, italic=italic_body, font=config.body_font)
        if answer:
            reveal_ids.extend([box.shape_id, option_text.shape_id])
        else:
            option_effect = config.option_effects[i] if i < len(config.option_effects) else "Fade"
            option_groups.append([
                (box.shape_id, 0, "clickEffect", option_effect),
                (option_text.shape_id, 0, "withEffect", option_effect),
            ])
    if answer:
        explanation_y = 5.18 if compact_options else min(5.18, option_y + 2.12)
        reveal_ids.extend(_add_reveal_panel(slide, spec, config, palette, mascot_path, explanation_y))
    return reveal_ids if answer else option_groups


def _add_general(slide, spec: SlideSpec, config: BuildConfig, palette: dict,
                 mascot_path: str | Path | None = None):
    lines = _content_lines(spec)
    bold_body, italic_body = _font_style(config)
    kind = (spec.slide_type + " " + spec.visual_direction).lower()
    prompt = _participant_prompt(spec.teacher_action)
    reserve_reveal = bool(prompt and config.reveal_chat_responses)
    if spec.title.strip().lower() == "grade 2 is not grade 8":
        content = spec.visible_content
        concept = re.split(r"\s+A[.)]\s*", content, maxsplit=1)[0].strip()
        match_a = re.search(r"A[.)]\s*(.+?)(?=\s+B[.)]\s*Grade\s*8|\s+Instruction:|$)", content, re.I)
        match_b = re.search(r"B[.)]\s*(Grade\s*8.+?)(?=\s+Instruction:|$)", content, re.I)
        grade_2 = match_a.group(1).strip() if match_a else "Grade 2: playful, short steps and visible cues"
        grade_8 = match_b.group(1).strip() if match_b else "Grade 8: purpose, logic and less hand-holding"
        _add_box(slide, .72, 1.3, 11.4, .7, "EFEBFF", line="EFEBFF")
        _add_text(slide, concept, .98, 1.4, 10.9, .48, size=18, color=palette["primary"],
                  bold=True, font=config.body_font)
        for x, label, text_value, accent in [
            (.72, "GRADE 2", grade_2, palette["secondary"]),
            (6.57, "GRADE 8", grade_8, palette["accent"]),
        ]:
            _add_box(slide, x, 2.2, 5.55, 2.15, "FFFFFF", line="D8DAE7")
            _add_box(slide, x, 2.2, .14, 2.15, accent, radius=False)
            _add_text(slide, label, x + .38, 2.42, 2.0, .35, size=13, color=palette["primary"],
                      bold=True, font="Arial")
            _add_text(slide, re.sub(r"^Grade\s*[28]\s*:\s*", "", text_value, flags=re.I),
                      x + .38, 2.82, 4.85, 1.25, size=18, color=palette["ink"],
                      bold=bold_body, italic=italic_body, font=config.body_font, valign=MSO_ANCHOR.TOP)
    elif "four" in kind or len(lines) == 4:
        colors = [palette["secondary"], palette["accent"], "F840B8", palette["success"]]
        card_h = 1.52 if reserve_reveal else 1.85
        row_gap = 1.82 if reserve_reveal else 2.25
        for i, line in enumerate(lines[:4]):
            x, y = (.7 + (i % 2) * 6.15, 1.42 + (i // 2) * row_gap)
            _add_box(slide, x, y, 5.7, card_h, "FFFFFF", line="D8DAE7")
            _add_box(slide, x, y, .13, card_h, colors[i], radius=False)
            card_size = _font_size(line, large=19, medium=16, small=13, tiny=11)
            _add_text(slide, line, x + .38, y + .12, 5.0, card_h - .22, size=card_size, color=palette["ink"],
                      bold=bold_body, italic=italic_body, font=config.body_font)
    elif "journey" in kind or "horizontal" in kind or "flow" in kind:
        display = _journey_steps(spec)[:8] or [spec.visible_content]
        columns = min(4, len(display))
        step_w = 11.55 / columns
        for i, line in enumerate(display):
            row, col = divmod(i, columns)
            step_h = 1.34 if reserve_reveal else 1.75
            row_gap = 1.68 if reserve_reveal else 2.25
            x, y = (.72 + col * step_w, 1.38 + row * row_gap)
            active = "day 1" in line.lower()
            fill = palette["primary"] if active else "FFFFFF"
            _add_box(slide, x, y, step_w - .22, step_h, fill, line=palette["primary"] if active else "D8DAE7")
            _add_text(slide, str(i + 1), x + .18, y + .16, .42, .42, size=14,
                      color=palette["accent"] if active else palette["primary"], bold=True, font="Aptos Display")
            _add_text(slide, line, x + .22, y + .42, step_w - .65, step_h - .52,
                      size=_font_size(line, large=17, medium=15, small=13, tiny=11),
                      color="FFFFFF" if active else palette["ink"], bold=True,
                      font=config.body_font, align=PP_ALIGN.CENTER)
    elif "dialogue" in kind or "scenario" in kind:
        _add_box(slide, .72, 1.45, 7.45, 4.65, "FFFFFF", line="D8DAE7")
        _add_text(slide, "\n".join(lines[:8]), 1.0, 1.67, 6.9, 4.15, size=19, color=palette["ink"],
                  bold=bold_body, italic=italic_body, font=config.body_font, valign=MSO_ANCHOR.TOP)
        _add_box(slide, 8.55, 1.65, 3.45, 3.45, palette["primary"], line=palette["primary"])
        _add_text(slide, "Think\n\nRespond\n\nDiscuss", 8.85, 1.95, 2.85, 2.85, size=23, color="FFFFFF", bold=True,
                  font="Aptos Display", align=PP_ALIGN.CENTER)
    else:
        body = "\n\n".join(lines[:10]) or spec.visible_content or spec.key_message
        body_size = _font_size(body, large=20, medium=17, small=14, tiny=12)
        _add_text(slide, body, .78, 1.45, 7.9, 4.95, size=body_size, color=palette["ink"],
                  bold=bold_body, italic=italic_body, font=config.body_font, valign=MSO_ANCHOR.TOP)
        visual = _context_visual(spec)
        if visual:
            slide.shapes.add_picture(str(visual), Inches(8.75), Inches(1.45), width=Inches(3.75))
        else:
            _add_box(slide, 9.1, 1.65, 3.0, 3.0, palette["primary"], line=palette["primary"])
            motif = "</>" if "coding" in (spec.title + spec.visible_content).lower() else str(spec.number).zfill(2)
            _add_text(slide, motif, 9.35, 2.3, 2.5, 1.6, size=50, color=palette["accent"], bold=True,
                      font="Arial", align=PP_ALIGN.CENTER)
    if prompt:
        _add_box(slide, .72, 6.27, 11.25, .48, palette["accent"], line=palette["accent"])
        _add_text(slide, prompt, .92, 6.3, 10.85, .38, size=12, color="241B62", bold=True,
                  font="Aptos", align=PP_ALIGN.CENTER)
    if prompt and config.reveal_chat_responses:
        return _add_reveal_panel(slide, spec, config, palette, mascot_path)
    return []


def _add_cover(slide, spec, config, palette, logo_path):
    _add_box(slide, 0, 0, 13.33, 7.5, palette["primary"], radius=False)
    _add_logo(slide, logo_path, config)
    _add_text(slide, spec.title, .75, 1.65, 8.5, 1.8, size=46, color="FFFFFF", bold=True, font="Arial")
    subtitle = spec.visible_content or "Interactive presentation generated from your storyboard"
    _add_text(slide, subtitle, .8, 3.65, 7.7, 1.3, size=21, color="FFFFFF", font=config.body_font)
    _add_box(slide, 9.55, 1.65, 2.6, 2.6, palette["accent"], line=palette["accent"])
    _add_text(slide, "</>", 9.75, 2.18, 2.2, 1.5, size=54, color=palette["primary"], bold=True,
              font="Arial", align=PP_ALIGN.CENTER)


def _set_notes(slide, spec: SlideSpec):
    try:
        slide.notes_slide.notes_text_frame.text = spec.speaker_notes
    except Exception:
        pass


def _duplicate_slide(prs: Presentation, source_slide):
    blank = prs.slide_layouts[6]
    dest = prs.slides.add_slide(blank)
    for shape in source_slide.shapes:
        dest.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), "p:extLst")
    return dest


def _transition_xml(kind: str) -> str:
    name = kind.lower()
    if name == "none":
        return ""
    effect = {"fade": "fade", "push": "push", "wipe": "wipe", "cut": "cut"}.get(name, "fade")
    extra = ' dir="l"' if effect in {"push", "wipe"} else ""
    return f'<p:transition xmlns:p="{P_NS}" spd="fast" advClick="1"><p:{effect}{extra}/></p:transition>'


TimingItem = int | tuple[int, int] | tuple[int, int, str] | tuple[int, int, str, str]


def _timing_xml(click_groups: list[list[TimingItem]], effect: str) -> str:
    if not click_groups:
        return ""
    settings = {
        "Appear": (1, 0, None),
        "Fade": (10, 0, "fade"),
        "Wipe": (22, 2, "wipe(right)"),
        "Wipe Right": (22, 2, "wipe(right)"),
        "Wipe Down": (22, 4, "wipe(down)"),
        "Wipe Left": (22, 1, "wipe(left)"),
        "Wipe Up": (22, 8, "wipe(up)"),
    }
    click_group_xml = []
    build_ids: list[int] = []
    next_id = 3
    for group_index, shape_ids in enumerate(click_groups, start=1):
        click_id = next_id
        next_id += 1
        steps: dict[int, list[tuple[int, str, str]]] = {}
        for index, item in enumerate(shape_ids):
            item_effect = effect
            if isinstance(item, tuple) and len(item) == 4:
                spid, delay, explicit_node_type, item_effect = item
            elif isinstance(item, tuple) and len(item) == 3:
                spid, delay, explicit_node_type = item
            elif isinstance(item, tuple):
                spid, delay = item
                explicit_node_type = None
            else:
                spid, delay, explicit_node_type = item, 0, None
            node_type = explicit_node_type or ("clickEffect" if index == 0 else "withEffect")
            steps.setdefault(max(0, int(delay)), []).append((int(spid), node_type, item_effect))
            build_ids.append(int(spid))

        step_xml = []
        for start_ms, step_items in sorted(steps.items()):
            step_id = next_id
            next_id += 1
            effect_xml = []
            for spid, node_type, item_effect in step_items:
                preset_id, preset_subtype, filter_name = settings.get(item_effect, settings["Fade"])
                effect_id, set_id, behavior_id = next_id, next_id + 1, next_id + 2
                next_id += 3
                anim = "" if filter_name is None else (
                    f'<p:animEffect transition="in" filter="{filter_name}"><p:cBhvr>'
                    f'<p:cTn id="{behavior_id}" dur="500"/>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                    f'</p:cBhvr></p:animEffect>'
                )
                effect_xml.append(
                    f'<p:par><p:cTn id="{effect_id}" fill="hold" '
                    f'nodeType="{node_type}" grpId="0" presetID="{preset_id}" '
                    f'presetClass="entr" presetSubtype="{preset_subtype}">'
                    f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                    f'<p:set><p:cBhvr><p:cTn id="{set_id}" dur="1" fill="hold">'
                    f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                    f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                    f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>{anim}'
                    f'</p:childTnLst></p:cTn></p:par>'
                )
            step_xml.append(
                f'<p:par><p:cTn id="{step_id}" fill="hold">'
                f'<p:stCondLst><p:cond delay="{start_ms}"/></p:stCondLst>'
                f'<p:childTnLst>{"".join(effect_xml)}</p:childTnLst>'
                f'</p:cTn></p:par>'
            )
        click_group_xml.append(
            f'<p:par><p:cTn id="{click_id}" fill="hold">'
            f'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(step_xml)}</p:childTnLst>'
            f'</p:cTn></p:par>'
        )

    build_xml = "".join(
        f'<p:bldP spid="{spid}" grpId="0"/>' for spid in dict.fromkeys(build_ids)
    )
    return (
        f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst>{"".join(click_group_xml)}</p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>'
        f'</p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        f'</p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'<p:bldLst>{build_xml}</p:bldLst></p:timing>'
    )


def _patch_pptx(data: bytes, config: BuildConfig,
                reveal_map: dict[int, list[list[TimingItem]]]) -> bytes:
    source = io.BytesIO(data)
    target = io.BytesIO()
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = zin.read(item.filename)
            match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
            if match:
                idx = int(match.group(1))
                xml = payload.decode("utf-8")
                transition = _transition_xml(config.slide_transition)
                timing = _timing_xml(reveal_map.get(idx, []), config.reveal_effect)
                additions = transition + timing
                if additions:
                    xml = xml.replace("</p:sld>", additions + "</p:sld>")
                payload = xml.encode("utf-8")
            zout.writestr(item, payload)
    return target.getvalue()


def normalize_logo(uploaded_bytes: bytes, suffix: str = ".png") -> Path:
    image = Image.open(io.BytesIO(uploaded_bytes)).convert("RGBA")
    image.thumbnail((2600, 900))
    tmp = NamedTemporaryFile(delete=False, suffix=".png")
    image.save(tmp.name, "PNG", optimize=True)
    return Path(tmp.name)


def build_presentation(slides: list[SlideSpec], config: BuildConfig, logo_path: str | Path | None = None,
                       mascot_path: str | Path | None = None) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    palette = THEMES[config.theme]
    total = min(len(slides), config.max_slides)
    reveal_map: dict[int, list[list[TimingItem]]] = {}
    built_index = 0

    for spec in slides[:config.max_slides]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        built_index += 1
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(palette["bg"])
        if spec.slide_type.lower() in {"opening", "title"}:
            _add_cover(slide, spec, config, palette, logo_path)
        else:
            _add_header(slide, spec, config, palette, logo_path, built_index, total)
            interactive = bool(spec.options or _participant_prompt(spec.teacher_action))
            timer_group = (_add_thinking_timer(slide, config, palette)
                           if interactive and config.mcq_reveal in {"Within-slide build (PowerPoint)", "Storyboard-directed"}
                           else [])
            if spec.options:
                option_groups = _add_mcq(slide, spec, config, palette, answer=False, mascot_path=mascot_path)
            else:
                ids = _add_general(slide, spec, config, palette, mascot_path=mascot_path)
                if ids and config.mcq_reveal in {"Within-slide build (PowerPoint)", "Storyboard-directed"}:
                    reveal_map[built_index] = ([timer_group] if timer_group else []) + [ids]
        if config.include_notes:
            _set_notes(slide, spec)

        wants_reveal = bool(spec.options and config.mcq_reveal != "No answer reveal")
        if wants_reveal and config.mcq_reveal == "Paired slides (portable)":
            answer_slide = prs.slides.add_slide(prs.slide_layouts[6])
            built_index += 1
            answer_slide.background.fill.solid()
            answer_slide.background.fill.fore_color.rgb = rgb(palette["bg"])
            _add_header(answer_slide, SlideSpec(number=spec.number, title=spec.title + " - Answer"), config,
                        palette, logo_path, built_index, total)
            _add_mcq(answer_slide, spec, config, palette, answer=True, mascot_path=mascot_path)
            if config.include_notes:
                _set_notes(answer_slide, spec)
        elif wants_reveal and config.mcq_reveal in {"Within-slide build (PowerPoint)", "Storyboard-directed"}:
            ids = _add_mcq(slide, spec, config, palette, answer=True, mascot_path=mascot_path)
            reveal_map[built_index] = option_groups + ([timer_group] if timer_group else []) + ([ids] if ids else [])

    raw = io.BytesIO()
    prs.save(raw)
    return _patch_pptx(raw.getvalue(), config, reveal_map)

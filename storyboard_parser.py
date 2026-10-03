from __future__ import annotations

import io
import re
from dataclasses import asdict, dataclass, field
from typing import Iterable

from pypdf import PdfReader


FIELD_LABELS = {
    "slide_type": "Slide Type",
    "visible_content": "What is Visible on Slide",
    "visual_direction": "Visual / Layout Direction",
    "animation": "Animation / Reveal Sequence",
    "teacher_action": "Teacher Action",
    "trainer_action": "Speaker Notes - Trainer Action",
    "context": "Speaker Notes - Explanation / Context",
    "answer_guidance": "Speaker Notes - Expected Responses / Answer Guidance",
    "watch_out": "Speaker Notes - Trainer Watch-out",
    "key_message": "Speaker Notes - Key Message to Land",
    "transition_notes": "Speaker Notes - Transition to Next Slide",
    "assets_source": "Assets / Source",
}

SLIDE_TYPE_PREFIXES = sorted([
    "Principle + 1:1 vs 3-4", "Orientation Retrieval + Concept", "Foundation Framework",
    "Open Response + Reveal", "Two Quick Scenarios", "Reflection + Close", "Compare + Discussion",
    "Framework + Quick Check", "Activity Instructions", "Rapid Chat Activity", "Chat + Round Robin",
    "A/B/C/D Scenario", "OTO vs 3-4 Decision", "Rank / First Move", "Rewrite Activity", "Scaffold Choice",
    "Breakout Practice", "Triad Practice", "Mini Practice", "Think and Build", "Think + Chat",
    "Journey Context", "Norms + Poll", "Compare + Choice", "Weak vs Better", "Open Response",
    "OTM Decision", "OTM Reveal", "Teach-Back Set-up", "Set-up", "Trigger", "Experience",
    "Scenario", "Principle", "Framework", "Break", "Opening", "Debrief",
], key=len, reverse=True)


@dataclass
class SlideSpec:
    number: int
    title: str
    start_minute: int | None = None
    end_minute: int | None = None
    slide_type: str = "Content"
    visible_content: str = ""
    visual_direction: str = ""
    animation: str = ""
    teacher_action: str = ""
    trainer_action: str = ""
    context: str = ""
    answer_guidance: str = ""
    watch_out: str = ""
    key_message: str = ""
    transition_notes: str = ""
    assets_source: str = ""
    options: list[str] = field(default_factory=list)
    correct_option: int | None = None

    def to_dict(self) -> dict:
        return asdict(self)

    @property
    def speaker_notes(self) -> str:
        sections = [
            ("Trainer Action", self.trainer_action),
            ("Explanation / Context", self.context),
            ("Expected Responses / Answer Guidance", self.answer_guidance),
            ("Trainer Watch-out", self.watch_out),
            ("Key Message to Land", self.key_message),
            ("Transition", self.transition_notes),
            ("Assets / Source", self.assets_source),
        ]
        return "\n\n".join(f"{name}:\n{text}" for name, text in sections if text.strip())


def extract_pdf_text(uploaded_file) -> str:
    if hasattr(uploaded_file, "getvalue"):
        stream = io.BytesIO(uploaded_file.getvalue())
    elif isinstance(uploaded_file, (bytes, bytearray)):
        stream = io.BytesIO(uploaded_file)
    else:
        stream = uploaded_file
    pdf = PdfReader(stream)
    # Browser-exported storyboards often contain positioned glyphs. Layout-mode
    # extraction can turn normal words into "Wha t" / "Studen t" and break every
    # field label. Plain extraction preserves the logical table reading order for
    # these files and is therefore the preferred source.
    plain = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if plain.count("Slide Type") >= 2 and plain.count("What is Visible on Slide") >= 2:
        return plain
    return "\n".join(page.extract_text(extraction_mode="layout") or "" for page in pdf.pages)


def _clean(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip()
    return value.replace("• ", "\n• ")


def _remove_pdf_furniture(text: str) -> str:
    lines = []
    for line in text.splitlines():
        compact = re.sub(r"\s+", " ", line).strip()
        if re.search(r"BrightCHAMPS Coding Teacher Training \| Day \d+ Storyboard", compact, re.I):
            continue
        if re.fullmatch(r"Trainer-ready (?:production )?handoff", compact, re.I):
            continue
        lines.append(line)
    return "\n".join(lines)


def _find_field(block: str, label: str, next_labels: Iterable[str]) -> str:
    labels = [re.escape(x) for x in next_labels]
    stop = "|".join(labels + [r"Slide\s+\d{1,2}\s*\|"])
    pattern = rf"{re.escape(label)}\s+(.*?)(?=\s+(?:{stop})|\Z)"
    match = re.search(pattern, block, flags=re.S | re.I)
    return _clean(match.group(1)) if match else ""


def _extract_options(content: str) -> list[str]:
    parsed = []
    marker = r"[A-H](?:[.)]\s*|\s+)"
    pattern = rf"(?:^|\s[/|;]?\s*)([A-H])(?:[.)]\s*|\s+)(.+?)(?=(?:\s[/|;]?\s*{marker})|\Z)"
    for letter, value in re.findall(pattern, content, re.S):
        cleaned = re.split(
            r"\b(?:Chat|Vote|Prompt|Type|Question|Instruction|Rules|Bottom prompt|Bottom line):"
            r"|\bUse these\b|\bThen\s+trainer\b|\bPrepare\s+\d+\s*(?:sec|second|min)",
            value, maxsplit=1, flags=re.I,
        )[0].strip()
        if cleaned:
            parsed.append((letter, cleaned))
    if len(parsed) < 2 or [letter for letter, _ in parsed] != list("ABCDEFGH")[:len(parsed)]:
        return []
    if any(len(value) < 3 for _, value in parsed):
        return []
    return [f"{letter}. {value}" for letter, value in parsed[:8]]


def _correct_index(guidance: str, options: list[str]) -> int | None:
    if not options:
        return None
    match = re.search(r"\b([A-D])\b\s*(?:strongest|correct|best|recommended|is the answer)", guidance, re.I)
    if not match:
        match = re.search(r"(?:answer|highlight|reveal)\s*[:=-]?\s*([A-D])\b", guidance, re.I)
    if not match:
        match = re.match(r"\s*([A-H])[.)]\s+", guidance)
    return ord(match.group(1).upper()) - 65 if match else None


def _recover_clipped_day2_slides(slides: list[SlideSpec], text: str) -> list[SlideSpec]:
    """Recover the three rows clipped from the supplied Day 2 browser PDF.

    The PDF's timing map contains slides 30-32, but the exported page jumps from
    the middle of slide 29 to the body of slide 32. These specs use the timing-map
    titles and the surrounding trainer-model debrief. Recovery is deliberately
    limited to this exact storyboard signature.
    """
    numbers = {slide.number for slide in slides}
    signature = "Where Did Learning Become Visible?" in text and "What Made the Close Speciﬁc?" in text
    if not signature:
        return slides
    # Slide 29 is cut mid-table and otherwise absorbs the visible tail of slide
    # 32. Restore the mapping activity from the surviving slide-29 content.
    for slide in slides:
        if slide.number == 29:
            slide.visible_content = (
                "Match the model moment to the ownership move. "
                "A. Student predicts what should happen "
                "B. Student builds or changes the code "
                "C. Student runs and checks the output "
                "D. Student explains why it works"
            )
            slide.options = [
                "A. Student predicts what should happen",
                "B. Student builds or changes the code",
                "C. Student runs and checks the output",
                "D. Student explains why it works",
            ]
            slide.answer_guidance = "A - Predict/decide; B - Build/change; C - Run/test; D - Explain."
            slide.key_message = "Student ownership appears through decisions, building, testing and explanation."
            slide.teacher_action = "Match A-D to Predict, Build, Test and Explain."
    if not {30, 31, 32}.isdisjoint(numbers):
        return slides
    recovered = [
        SlideSpec(
            number=30, start_minute=163, end_minute=168,
            title="Where Did Learning Become Visible?", slide_type="Evidence Mapping",
            visible_content=(
                "Choose the moments that gave evidence of learning. "
                "A. Student predicted what the code would do "
                "B. Student built or changed the code "
                "C. Student ran the project and checked the output "
                "D. Student explained what the result showed"
            ),
            visual_direction="Four equal evidence cards. Keep all options visually identical.",
            animation="Initial: A-D. Reveal: Learning became visible through prediction, building, testing and explanation.",
            teacher_action="Choose one moment and name the evidence.",
            trainer_action="Take two evidence-based responses before revealing the synthesis.",
            answer_guidance="Learning became visible through prediction, student action, run-test evidence and explanation.",
            key_message="Learning is visible when the student produces evidence, not when the teacher finishes explaining.",
            transition_notes="Now identify where the trainer adapted in response to that evidence.",
            assets_source="Recovered from the Day 2 timing map and trainer-model debrief after PDF export clipping.",
            options=["A. Student predicted what the code would do", "B. Student built or changed the code",
                     "C. Student ran the project and checked the output", "D. Student explained what the result showed"],
        ),
        SlideSpec(
            number=31, start_minute=168, end_minute=172,
            title="Where Did the Trainer Adapt?", slide_type="Adaptation Mapping",
            visible_content=(
                "Which move shows useful adaptation? "
                "A. Repeat the complete explanation "
                "B. Lower the entry point with one smaller question "
                "C. Take control of the student's mouse "
                "D. Skip the difficult step"
            ),
            visual_direction="Four equal cards. No correctness colour before reveal.",
            animation="Initial: A-D. Reveal: B with a short explanation.",
            teacher_action="Vote A-D and justify the choice.",
            trainer_action="Ask what evidence would tell the teacher to use the smaller question.",
            answer_guidance="B. Lowering the entry point preserves student ownership while responding to evidence.",
            key_message="Adapt the support while keeping the coding and thinking with the student.",
            transition_notes="Finish the debrief by examining the close and next step.",
            assets_source="Recovered from the Day 2 timing map and trainer-model debrief after PDF export clipping.",
            options=["A. Repeat the complete explanation", "B. Lower the entry point with one smaller question",
                     "C. Take control of the student's mouse", "D. Skip the difficult step"],
            correct_option=1,
        ),
        SlideSpec(
            number=32, start_minute=172, end_minute=175,
            title="What Made the Close Specific?", slide_type="Compare",
            visible_content=(
                "Compare feedback. "
                "A. Great class. You did very well today. "
                "B. Today you used sensing and a condition to make the score react when the player touched the fruit, "
                "and you corrected the logic after testing it. Next, practise explaining the IF condition in one sentence."
            ),
            visual_direction="Two equal feedback cards. Keep both visually identical before reveal.",
            animation="Initial: A/B. Reveal: Stronger feedback - B. Then reveal Evidence -> Next step.",
            teacher_action="Vote A/B and name one evidence phrase.",
            trainer_action="Take a quick vote, then ask what exact evidence appears in B.",
            context="Specific closure supports continuity without becoming detailed parent or PTM training.",
            answer_guidance="B. It names evidence from the class and gives a specific next step.",
            watch_out="Do not turn this into detailed parent or PTM training.",
            key_message="A strong close tells the student what changed today and what comes next.",
            transition_notes="End Day 2 by turning observation into a short DIY task.",
            assets_source="Recovered from the visible slide-32 table body after PDF export clipping.",
            options=["A. Great class. You did very well today.",
                     "B. Specific evidence from today's coding plus one next step."],
            correct_option=1,
        ),
    ]
    return sorted(slides + recovered, key=lambda slide: slide.number)


def parse_storyboard_text(text: str) -> list[SlideSpec]:
    text = _remove_pdf_furniture(text.replace("\u00a0", " ").replace("\r", "\n"))
    header_re = re.compile(
        r"Slide\s+(\d{1,2})\s*\|\s*(?:(\d+)\s*-\s*(\d+)\s*min\s*\|\s*)?([^\n]+)",
        re.I,
    )
    matches = list(header_re.finditer(text))
    slides: list[SlideSpec] = []
    labels = list(FIELD_LABELS.values())
    for idx, match in enumerate(matches):
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        block = text[match.end():end]
        kwargs = {}
        for key, label in FIELD_LABELS.items():
            kwargs[key] = _find_field(block, label, [x for x in labels if x != label])
        # PDF text extraction can merge the next row's content into the Slide Type row.
        # Recover the known type prefix and move the overflow back to visible content.
        raw_type = kwargs["slide_type"]
        for prefix in SLIDE_TYPE_PREFIXES:
            if raw_type.lower().startswith(prefix.lower()):
                overflow = raw_type[len(prefix):].strip()
                kwargs["slide_type"] = prefix
                if overflow:
                    kwargs["visible_content"] = _clean(overflow + " " + kwargs["visible_content"])
                break
        options = _extract_options(kwargs["visible_content"])
        slides.append(
            SlideSpec(
                number=int(match.group(1)),
                title=_clean(match.group(4)),
                start_minute=int(match.group(2)) if match.group(2) else None,
                end_minute=int(match.group(3)) if match.group(3) else None,
                **kwargs,
                options=options,
                correct_option=_correct_index(kwargs["answer_guidance"], options),
            )
        )
    # Remove duplicates caused by timing-map references before the detailed storyboard.
    detailed = {}
    for slide in slides:
        if slide.visible_content or slide.visual_direction:
            detailed[slide.number] = slide
    parsed = [detailed[k] for k in sorted(detailed)]
    return _recover_clipped_day2_slides(parsed, text)


def parse_storyboard_pdf(uploaded_file) -> list[SlideSpec]:
    return parse_storyboard_text(extract_pdf_text(uploaded_file))


def validate_slide_plan(slides: list[SlideSpec]) -> tuple[list[str], list[str]]:
    """Return blocking errors and non-blocking warnings before PPT generation."""
    errors: list[str] = []
    warnings: list[str] = []
    if not slides:
        return ["No storyboard slides were detected."], []
    numbers = [slide.number for slide in slides]
    duplicates = sorted({number for number in numbers if numbers.count(number) > 1})
    if duplicates:
        errors.append(f"Duplicate slide numbers: {', '.join(map(str, duplicates))}.")
    expected = set(range(min(numbers), max(numbers) + 1))
    missing = sorted(expected - set(numbers))
    if missing:
        errors.append(f"Missing slide numbers: {', '.join(map(str, missing))}.")
    for slide in slides:
        if not slide.title.strip():
            errors.append(f"Slide {slide.number} has no title.")
        if not slide.slide_type.strip():
            warnings.append(f"Slide {slide.number} has no detected slide type.")
        option_markers = len(re.findall(r"(?:^|\s)[A-H][.)]\s+", slide.visible_content))
        if option_markers >= 2 and not slide.options:
            warnings.append(f"Slide {slide.number} contains choice markers but no options were parsed.")
        if slide.options and not (slide.answer_guidance or slide.key_message):
            warnings.append(f"Slide {slide.number} has options but no answer or reveal guidance.")
    return errors, warnings


def parse_manual_content(text: str) -> list[SlideSpec]:
    """Parse friendly blocks separated by ---.

    Supported fields: Title, Type, Content, Options, Answer, Notes, Visual.
    """
    blocks = [b.strip() for b in re.split(r"\n\s*---+\s*\n", text.strip()) if b.strip()]
    slides = []
    for i, block in enumerate(blocks, 1):
        fields: dict[str, str] = {}
        current = None
        for line in block.splitlines():
            match = re.match(r"^(Title|Type|Content|Options|Answer|Notes|Visual|Animation)\s*:\s*(.*)$", line, re.I)
            if match:
                current = match.group(1).lower()
                fields[current] = match.group(2).strip()
            elif current:
                fields[current] += "\n" + line.rstrip()
        raw_options = [x.strip() for x in re.split(r"\n|\|", fields.get("options", "")) if x.strip()]
        options = [f"{chr(65+j)}. {re.sub(r'^[A-D][.)]\s*', '', x)}" for j, x in enumerate(raw_options[:4])]
        answer = fields.get("answer", "").strip().upper()
        correct = ord(answer[0]) - 65 if answer and answer[0] in "ABCD" else None
        slides.append(
            SlideSpec(
                number=i,
                title=fields.get("title", f"Slide {i}"),
                slide_type=fields.get("type", "Content"),
                visible_content=fields.get("content", ""),
                visual_direction=fields.get("visual", ""),
                animation=fields.get("animation", ""),
                context=fields.get("notes", ""),
                answer_guidance=f"Correct answer: {answer}" if answer else "",
                options=options,
                correct_option=correct,
            )
        )
    return slides

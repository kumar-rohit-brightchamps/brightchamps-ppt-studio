from __future__ import annotations

import io
import json
import re
from pathlib import Path

import streamlit as st

from ppt_builder import BuildConfig, build_presentation, normalize_logo
from storyboard_parser import SlideSpec, parse_manual_content, parse_storyboard_pdf


def validate_slide_plan(slides: list[SlideSpec]) -> tuple[list[str], list[str]]:
    """Validate a parsed plan locally so Streamlit reloads cannot split imports."""
    errors: list[str] = []
    warnings: list[str] = []
    if not slides:
        return ["No storyboard slides were detected."], []

    numbers = [slide.number for slide in slides]
    duplicates = sorted({number for number in numbers if numbers.count(number) > 1})
    if duplicates:
        errors.append(f"Duplicate slide numbers: {', '.join(map(str, duplicates))}.")
    missing = sorted(set(range(min(numbers), max(numbers) + 1)) - set(numbers))
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


APP_DIR = Path(__file__).parent
BUILTIN_LOGO = APP_DIR / "assets" / "brightchamps-logo-hires.png"
BUILTIN_MASCOT = APP_DIR / "assets" / "answer-owl-v1.png"

st.set_page_config(page_title="BrightCHAMPS PPT Studio", page_icon="🎞️", layout="wide")
st.markdown("""
<style>
  .stApp { background: radial-gradient(circle at 96% 2%, #e8f7ff 0, transparent 24%), #fbfaf7; }
  .block-container { max-width: 1380px; padding-top: 1.3rem; }
  h1, h2, h3 { color: #241b62; letter-spacing: -0.025em; }
  [data-testid="stMetric"] { background: white; border: 1px solid #eceaf5; border-radius: 16px; padding: 12px; }
  [data-testid="stFileUploader"] { background: white; border-radius: 18px; padding: 12px; }
  div.stButton > button, div.stDownloadButton > button { border-radius: 999px; font-weight: 750; }
  .preview { border-radius: 24px; min-height: 285px; padding: 30px; color: white;
             background: linear-gradient(135deg,#241b62,#4d3bc2); position: relative; overflow:hidden; }
  .preview:after { content:''; position:absolute; right:-60px; top:-80px; width:210px; height:210px;
                   border-radius:50%; background:#42c0fa; }
  .preview h2 { color:white; font-size:38px; max-width:70%; margin-top:50px; }
  .preview p { color:#eee9ff; font-size:18px; max-width:70%; }
  .eyebrow { display:inline-block; padding:7px 14px; border-radius:999px; background:#ffda0b; color:#241b62; font-weight:800; }
  .logo-pos { position:absolute; z-index:2; background:white; border-radius:12px; padding:8px 12px; color:#4d3bc2; font-weight:850; }
</style>
""", unsafe_allow_html=True)

st.title("BrightCHAMPS PPT Studio")
st.caption("Turn a trainer storyboard or structured content into a branded, interactive PowerPoint.")

if "slides" not in st.session_state:
    st.session_state.slides = []
if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None
if "validation" not in st.session_state:
    st.session_state.validation = ([], [])

input_tab, design_tab, interaction_tab, review_tab = st.tabs([
    "1  Content", "2  Brand & typography", "3  Interactivity", "4  Review & generate"
])

with input_tab:
    source_mode = st.radio("Content source", ["Upload storyboard PDF", "Write or paste content"], horizontal=True)
    if source_mode == "Upload storyboard PDF":
        uploaded_storyboard = st.file_uploader("Upload storyboard", type=["pdf"], help="The parser reads slide titles, visible content, layout direction, reveal sequence and speaker notes.")
        if uploaded_storyboard and st.button("Analyse storyboard", type="primary"):
            with st.spinner("Reading slide structure and trainer notes..."):
                st.session_state.slides = parse_storyboard_pdf(uploaded_storyboard)
                st.session_state.validation = validate_slide_plan(st.session_state.slides)
            errors, warnings = st.session_state.validation
            if errors:
                st.error("Storyboard needs attention before generation: " + " ".join(errors))
            else:
                st.success(f"Found and validated {len(st.session_state.slides)} detailed slides.")
            for warning in warnings[:8]:
                st.warning(warning)
    else:
        sample = """Title: Coding with Conditions
Type: Opening
Content: Learn how computers make decisions.
---
Title: Quick Check
Type: MCQ
Content: Which operator checks whether two values are equal?
Options: = | == | != | >=
Answer: B
Notes: Explain that = assigns a value while == compares values.
Animation: Reveal the answer after the vote."""
        manual = st.text_area("Presentation content", value=sample, height=360)
        if st.button("Build slide plan", type="primary"):
            st.session_state.slides = parse_manual_content(manual)
            st.session_state.validation = validate_slide_plan(st.session_state.slides)
            st.success(f"Created a {len(st.session_state.slides)}-slide plan.")

    if st.session_state.slides:
        c1, c2, c3 = st.columns(3)
        c1.metric("Slides found", len(st.session_state.slides))
        c2.metric("Interactive slides", sum(bool(x.animation and "click" in x.animation.lower()) for x in st.session_state.slides))
        c3.metric("MCQ / choice slides", sum(bool(x.options) for x in st.session_state.slides))
        with st.expander("Inspect parsed slide plan", expanded=False):
            for slide in st.session_state.slides:
                st.markdown(f"**{slide.number:02d}. {slide.title}**  ·  {slide.slide_type}")
                st.caption(slide.visible_content[:280] + ("..." if len(slide.visible_content) > 280 else ""))

with design_tab:
    left, right = st.columns([1, 1.05], gap="large")
    with left:
        theme = st.selectbox("Visual theme", ["BrightCHAMPS Classroom", "Coding Night", "Clean Workshop", "Playful Learning"])
        logo_mode = st.radio("Logo", ["Built-in BrightCHAMPS", "Upload my logo", "No logo"], horizontal=True)
        logo_upload = None
        if logo_mode == "Upload my logo":
            logo_upload = st.file_uploader("Upload a transparent PNG logo", type=["png", "jpg", "jpeg"])
        logo_position = st.selectbox("Logo position", ["Top right", "Top left", "Bottom right", "Bottom left", "None"], disabled=logo_mode == "No logo")
        logo_size = st.segmented_control("Logo size", ["Small", "Medium", "Large"], default="Medium")
        st.divider()
        st.subheader("Body typography")
        st.caption("Headers and subsection labels use presentation defaults chosen for readability.")
        body_font = st.selectbox("Content font", ["Aptos", "Arial", "Calibri", "Georgia", "Trebuchet MS", "Verdana", "Times New Roman"], index=1)
        body_style = st.selectbox("Content font style", ["Regular", "Bold", "Italic", "Bold Italic"])
        include_progress = st.toggle("Show slide progress", value=True)
    with right:
        pos_css = {"Top right":"right:18px;top:18px", "Top left":"left:18px;top:18px", "Bottom right":"right:18px;bottom:18px", "Bottom left":"left:18px;bottom:18px", "None":"display:none"}[logo_position]
        st.markdown(f"""
        <div class="preview">
          <div class="logo-pos" style="{pos_css}">Bright<span style="color:#ff8900">C</span><span style="color:#ffda0b">H</span><span style="color:#46de60">A</span><span style="color:#42c0fa">M</span><span style="color:#f840b8">P</span>S</div>
          <span class="eyebrow">LIVE PREVIEW</span>
          <h2>Interactive Coding Classroom</h2>
          <p style="font-family:'{body_font}';font-weight:{'700' if 'Bold' in body_style else '400'};font-style:{'italic' if 'Italic' in body_style else 'normal'}">Visible thinking, active participation and purposeful practice.</p>
        </div>
        """, unsafe_allow_html=True)
        st.info("The header font and section-label font adapt automatically. Your selected font applies to body content, options and explanations.")

with interaction_tab:
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.subheader("MCQ and answer reveal")
        mcq_reveal = "Within-slide build (PowerPoint)"
        st.success(
            "Default structure for every detected MCQ: question visible → options on separate clicks → "
            "automatic timer → answer reveal on the next click → next slide."
        )
        animation_setup = st.radio(
            "Option animation setup",
            ["Recommended defaults", "Choose manually"],
            horizontal=True,
            help="Recommended defaults apply automatically to every uploaded storyboard. Manual mode lets you choose A–D separately.",
        )
        reveal_effect = st.selectbox(
            "Answer entrance effect",
            ["Wipe Up", "Fade", "Wipe Right", "Wipe Down", "Appear"],
            help="The correct-answer highlight, explanation panel and owl use this effect together.",
        )
        effect_defaults = ("Fade", "Wipe Right", "Wipe Down", "Appear")
        if animation_setup == "Recommended defaults":
            option_effects = effect_defaults
            st.caption("Default option effects: A Fade · B Wipe Right · C Wipe Down · D Appear")
        else:
            effect_choices = ["Fade", "Wipe Right", "Wipe Down", "Wipe Left", "Wipe Up", "Appear"]
            effect_columns = st.columns(4)
            option_effects = tuple(
                effect_columns[i].selectbox(
                    f"Option {letter}", effect_choices,
                    index=effect_choices.index(effect_defaults[i]),
                    key=f"option_effect_{letter}",
                )
                for i, letter in enumerate("ABCD")
            )
        reveal_chat_responses = st.toggle("Reveal chat-response guidance on click", value=True,
                                          help="Shows the prompt first, then reveals the storyboard key message on the next click.")
        use_answer_mascot = st.toggle("Use owl guide on reveal", value=True,
                                      help="Adds a friendly visual cue only when the answer or trainer guidance appears.")
        thinking_timer_seconds = st.number_input("Thinking timer (seconds)", min_value=0, max_value=30, value=10,
                                                 help="After the options appear, one click starts this countdown. The following click reveals the answer.")
        st.info("Uses PowerPoint timing XML. Test in Microsoft PowerPoint; some browser viewers may ignore within-slide builds.")
    with col2:
        st.subheader("Slide transitions")
        slide_transition = st.selectbox("Default transition", ["Fade", "Push", "Wipe", "Cut", "None"])
        st.caption("Every transition advances on click. Individual storyboard reveal instructions remain available in speaker notes.")
        st.info("Trainer cues such as “Listen” stay in speaker notes. They do not create an audio player. A play button is only appropriate when an audio file is supplied.")
        include_notes = st.toggle("Include trainer speaker notes", value=True)
        max_slides = st.number_input("Maximum slides to generate", min_value=1, max_value=150, value=100)

with review_tab:
    if not st.session_state.slides:
        st.warning("Add a storyboard or manual content in the Content tab first.")
    else:
        validation_errors, validation_warnings = validate_slide_plan(st.session_state.slides)
        st.subheader("Generation summary")
        summary = {
            "source_slides": len(st.session_state.slides), "theme": theme, "body_font": body_font,
            "logo_position": logo_position, "mcq_structure": "Question → options → timer → answer",
            "option_effects": dict(zip("ABCD", option_effects)), "answer_effect": reveal_effect,
            "transition": slide_transition,
        }
        st.json(summary, expanded=False)
        if validation_errors:
            st.error("Generation is blocked until these issues are fixed: " + " ".join(validation_errors))
        elif validation_warnings:
            st.warning(f"Generation can continue, but {len(validation_warnings)} parser warning(s) were found. Review the slide plan.")
        if st.button("Generate PowerPoint", type="primary", use_container_width=True, disabled=bool(validation_errors)):
            logo_path = None
            if logo_mode == "Built-in BrightCHAMPS" and logo_position != "None":
                logo_path = BUILTIN_LOGO
            elif logo_mode == "Upload my logo" and logo_upload:
                logo_path = normalize_logo(logo_upload.getvalue(), Path(logo_upload.name).suffix)
            config = BuildConfig(
                theme=theme, body_font=body_font, body_style=body_style,
                logo_position="None" if logo_mode == "No logo" else logo_position,
                logo_size=logo_size or "Medium", slide_transition=slide_transition,
                mcq_reveal=mcq_reveal, reveal_effect=reveal_effect,
                option_effects=option_effects,
                include_notes=include_notes, include_progress=include_progress, max_slides=int(max_slides),
                reveal_chat_responses=reveal_chat_responses, use_answer_mascot=use_answer_mascot,
                thinking_timer_seconds=int(thinking_timer_seconds),
            )
            with st.spinner("Designing slides, adding notes and packaging PowerPoint animations..."):
                mascot_path = BUILTIN_MASCOT if use_answer_mascot else None
                st.session_state.ppt_bytes = build_presentation(st.session_state.slides, config, logo_path, mascot_path)
            st.success("PowerPoint generated.")
        if st.session_state.ppt_bytes:
            st.download_button(
                "Download generated PPTX", data=st.session_state.ppt_bytes,
                file_name="BrightCHAMPS_Interactive_Training.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary", use_container_width=True,
            )

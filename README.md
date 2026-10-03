# BrightCHAMPS PPT Studio

A Streamlit application that converts a structured trainer storyboard PDF or manually entered content into a branded, interactive PowerPoint.

## Features

- Parses storyboard fields into slide titles, visible content, layout guidance, animations and speaker notes.
- Chooses clean PDF text extraction automatically, removes repeated PDF headers and validates slide numbering before generation.
- Accepts manual content when no storyboard is available.
- Built-in BrightCHAMPS logo or uploaded logo.
- Top-left, top-right, bottom-left and bottom-right logo placement.
- Selectable body font and font style. Headers and subsection labels use automatic presentation defaults.
- Four visual themes and content-aware layouts.
- Every detected MCQ uses the staged PowerPoint sequence by default: question, separate option clicks, automatic timer, answer reveal, then slide advance.
- Fade, push, wipe, cut or no slide transition.
- Optional trainer notes and slide progress.
- Benchmark-inspired facilitation layouts with large prompts, answer strips and clearer activity states.
- One-click reveal groups for MCQ answers and open chat-response guidance.
- Separate PowerPoint-compatible entrance effects for options A, B, C and D, plus an independent answer-reveal effect.
- Supports compact A-H option activities while preserving one option group per click.
- Built-in illustrated teaching scenes and an optional owl answer guide.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Manual content format

Separate slides with `---` and use these optional fields:

```text
Title: Quick Check
Type: MCQ
Content: Which operator compares two values?
Options: = | == | != | >=
Answer: B
Notes: Explain that = assigns while == compares.
Animation: Reveal the answer after voting.
```

## Animation compatibility

- The Streamlit app always selects `Within-slide build (PowerPoint)` for uploaded storyboards and manual content. Microsoft PowerPoint provides the intended playback; web viewers may ignore the build.
- `Recommended defaults` automatically uses A Fade, B Wipe Right, C Wipe Down and D Appear. `Choose manually` exposes a separate selector for each option.
- A reveal is grouped as one click: the correct choice, trainer guidance and mascot appear together.
- With a thinking timer enabled, the within-slide sequence is: question visible → each option on its own click → one click starts the automatic countdown → the next click reveals the answer/highlight/explanation/mascot → the following click advances the slide.
- Slide transitions are stored in the PowerPoint package and advance on click.

## Storyboard rule handling

The parser keeps visible slide content separate from trainer notes. It does not treat production directions as participant-facing text. Selecting paired-slide reveals may increase the slide count, so the app displays a warning before generation.

from io import BytesIO
from xml.etree import ElementTree as ET
import zipfile

from ppt_builder import BuildConfig, build_presentation
from storyboard_parser import SlideSpec, parse_manual_content, parse_storyboard_text


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def test_manual_content_and_pptx():
    slides = parse_manual_content("""Title: Test\nType: MCQ\nContent: Pick one\nOptions: Red | Blue | Green | Yellow\nAnswer: B\nAnimation: reveal""")
    data = build_presentation(slides, BuildConfig(mcq_reveal="Paired slides (portable)"))
    assert data[:2] == b"PK"
    with zipfile.ZipFile(BytesIO(data)) as zf:
        assert "ppt/presentation.xml" in zf.namelist()
        assert sum(name.startswith("ppt/slides/slide") and name.endswith(".xml") for name in zf.namelist()) == 2


def test_storyboard_parser():
    text = """Slide 01 | 0-2 min | Welcome\nSlide Type Opening\nWhat is Visible on Slide Hello room\nVisual / Layout Direction Clean opening\nAnimation / Reveal Sequence All visible\nTeacher Action Listen\nSpeaker Notes - Trainer Action Welcome everyone\nSpeaker Notes - Explanation / Context Context here\nSpeaker Notes - Expected Responses / Answer Guidance No response\nSpeaker Notes - Trainer Watch-out Keep it short\nSpeaker Notes - Key Message to Land Students think\nSpeaker Notes - Transition to Next Slide Continue\nAssets / Source Internal"""
    slides = parse_storyboard_text(text)
    assert len(slides) == 1
    assert slides[0].title == "Welcome"
    assert "Hello room" in slides[0].visible_content


def test_mcq_click_groups_and_unique_timing_ids():
    slides = [
        SlideSpec(number=1, title="Animation Test", slide_type="Opening", visible_content="Timing reference"),
        SlideSpec(
            number=2,
            title="Quick Check",
            slide_type="MCQ",
            visible_content="Which answer is correct? A. One B. Two C. Three D. Four",
            options=["A. One", "B. Two", "C. Three", "D. Four"],
            correct_option=1,
            answer_guidance="B is correct.",
        ),
    ]
    data = build_presentation(
        slides,
        BuildConfig(
            mcq_reveal="Within-slide build (PowerPoint)",
            reveal_effect="Fade",
            thinking_timer_seconds=10,
        ),
    )
    with zipfile.ZipFile(BytesIO(data)) as zf:
        root = ET.fromstring(zf.read("ppt/slides/slide2.xml"))

    ns = {"p": P_NS}
    main_seq = root.find(".//p:cTn[@nodeType='mainSeq']", ns)
    assert main_seq is not None
    child_list = main_seq.find("p:childTnLst", ns)
    assert child_list is not None
    click_groups = []
    for par in child_list.findall("p:par", ns):
        delay = par.find("p:cTn/p:stCondLst/p:cond", ns)
        if delay is not None and delay.get("delay") == "indefinite":
            click_groups.append(par)
    assert len(click_groups) == 6

    option_presets = []
    for group in click_groups[:4]:
        first_effect = group.find(".//p:cTn[@presetClass='entr']", ns)
        assert first_effect is not None
        option_presets.append((first_effect.get("presetID"), first_effect.get("presetSubtype")))
    assert option_presets == [("10", "0"), ("22", "2"), ("22", "4"), ("1", "0")]

    timer_steps = click_groups[4].findall("p:cTn/p:childTnLst/p:par", ns)
    timer_delays = [
        int(step.find("p:cTn/p:stCondLst/p:cond", ns).get("delay"))
        for step in timer_steps
    ]
    assert timer_delays == list(range(0, 10001, 1000))

    timing = root.find("p:timing", ns)
    assert timing is not None
    ids = [node.get("id") for node in timing.findall(".//p:cTn", ns)]
    assert None not in ids
    assert len(ids) == len(set(ids))
    assert timing.find("p:bldLst", ns) is not None
    assert timing.find(".//p:cTn[@presetID='10']", ns) is not None
    assert timing.find(".//p:cTn[@presetID='9']", ns) is None

    seq = root.find(".//p:seq", ns)
    assert seq is not None
    assert seq.find("p:prevCondLst", ns) is not None
    assert seq.find("p:nextCondLst", ns) is not None
    assert main_seq.find("p:prevCondLst", ns) is None
    assert main_seq.find("p:nextCondLst", ns) is None


def test_first_storyboard_slide_can_use_mcq_sequence():
    slide = SlideSpec(
        number=1,
        title="Immediate Check",
        slide_type="MCQ",
        visible_content="Choose one. A. One B. Two C. Three D. Four",
        options=["A. One", "B. Two", "C. Three", "D. Four"],
        correct_option=0,
        answer_guidance="A is correct.",
    )
    data = build_presentation(
        [slide],
        BuildConfig(mcq_reveal="Within-slide build (PowerPoint)", thinking_timer_seconds=10),
    )
    with zipfile.ZipFile(BytesIO(data)) as zf:
        root = ET.fromstring(zf.read("ppt/slides/slide1.xml"))
    ns = {"p": P_NS}
    main_seq = root.find(".//p:cTn[@nodeType='mainSeq']", ns)
    assert main_seq is not None
    groups = []
    for par in main_seq.find("p:childTnLst", ns).findall("p:par", ns):
        delay = par.find("p:cTn/p:stCondLst/p:cond", ns)
        if delay is not None and delay.get("delay") == "indefinite":
            groups.append(par)
    assert len(groups) == 6

import json

from dissertation.gost import PRESETS, gost_7_32_2017, gost_r_7_0_11_2011
from dissertation.profile import FormattingProfile


def test_json_roundtrip(tmp_path):
    prof = gost_r_7_0_11_2011()
    path = tmp_path / "p.json"
    prof.save(path)
    loaded = FormattingProfile.load(path)
    assert loaded == prof


def test_from_dict_is_tolerant():
    """Старые/чужие профили: лишние ключи игнорируются, недостающие — по умолчанию."""
    data = {"name": "X", "unknown_key": 1, "page": {"margin_left_mm": 27, "junk": 2},
            "body": {"font_size_pt": 12}}
    prof = FormattingProfile.from_dict(data)
    assert prof.name == "X"
    assert prof.page.margin_left_mm == 27.0
    assert prof.page.margin_right_mm == 15.0
    assert prof.body.font_size_pt == 12.0
    assert prof.body.font_name == "Times New Roman"


def test_presets_match_standards():
    p = gost_7_32_2017()
    assert (p.page.margin_left_mm, p.page.margin_right_mm,
            p.page.margin_top_mm, p.page.margin_bottom_mm) == (30, 15, 20, 20)
    assert p.body.line_spacing == 1.5 and p.body.first_line_indent_cm == 1.25
    assert p.body.alignment == "justify"
    assert p.page_numbering.position == "footer" and p.page_numbering.hide_on_first_page
    assert p.chapter_heading.page_break_before and p.chapter_heading.bold

    d = gost_r_7_0_11_2011()
    assert (d.page.margin_left_mm, d.page.margin_right_mm) == (25, 10)
    assert d.page_numbering.position == "header"
    assert d.chapter_heading.alignment == "center"


def test_appendix_letters_skip_forbidden():
    p = gost_7_32_2017()
    letters = [p.appendix_letter(i) for i in range(len(p.appendix_letters))]
    for forbidden in "ЁЗЙОЧЪЫЬ":
        assert forbidden not in letters
    assert p.appendix_letter(0) == "А" and p.appendix_letter(7) == "И"
    assert p.appendix_letter(100) == "101"


def test_all_presets_serializable():
    for factory in PRESETS.values():
        json.loads(factory().to_json())

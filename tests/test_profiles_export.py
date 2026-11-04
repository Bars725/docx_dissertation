import json

from docx import Document

from dissertation.controller import AppController
from dissertation.document_builder import BlockType as B


def test_export_preset_edit_and_use(tmp_path):
    ctl = AppController(profiles_dir=tmp_path / "profiles")
    path = ctl.export_preset("gost-7.32-2017")
    assert path.exists() and path.parent == tmp_path / "profiles"
    assert ctl.export_preset("gost-7.32-2017") != path          # не перезаписывает
    data = json.loads(path.read_text(encoding="utf-8"))
    data["body"]["font_size_pt"] = 12                             # «ручная точечная правка»
    data["page"]["margin_left_mm"] = 25
    data["figure_caption_format"] = "Рис. {n}. {title}"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    key = f"file:{path}"
    assert key in [k for k, _ in ctl.available_profiles()]
    ctl.new_document(key)
    ctl.insert(B.CHAPTER, "Текст.", heading="Глава")
    out = ctl.save(tmp_path / "edited.docx")
    doc = Document(str(out))
    assert doc.styles["ГОСТ Основной текст"].font.size.pt == 12
    assert round(doc.sections[0].left_margin.mm) == 25


def test_export_current_document_profile(tmp_path, word_like_template):
    ctl = AppController(profiles_dir=tmp_path / "profiles")
    ctl.open_document(word_like_template)
    path = ctl.export_current_profile()
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["body"]["font_name"] == "Arial"                  # извлечено из документа

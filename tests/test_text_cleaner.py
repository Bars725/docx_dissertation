from dissertation.text_cleaner import CleanOptions, TextCleaner

C = TextCleaner()


def test_pdf_wraps_and_hyphenation_joined():
    raw = "Метод позволяет рас-\nпространить подход на\nновые задачи.\nВторой абзац начинается здесь."
    assert C.to_paragraphs(raw) == [
        "Метод позволяет распространить подход на новые задачи.",
        "Второй абзац начинается здесь.",
    ]


def test_hyphen_kept_for_particles_and_prepositions():
    assert C.to_paragraphs("Это из-\nза того, что-\nто случилось.") == ["Это из-за того, что-то случилось."]


def test_blank_line_mode_auto_detected():
    raw = "Первый абзац\nпродолжается Здесь.\n\nВторой абзац.\n\n\nТретий."
    assert C.to_paragraphs(raw) == ["Первый абзац продолжается Здесь.", "Второй абзац.", "Третий."]


def test_line_mode():
    c = TextCleaner(CleanOptions(paragraph_mode="line"))
    assert c.to_paragraphs("а\nб\n\nв") == ["а", "б", "в"]


def test_list_items_not_glued():
    raw = "Задачи исследования:\n1) изучить теорию;\n2) провести эксперимент."
    assert C.to_paragraphs(raw) == ["Задачи исследования:", "1) изучить теорию;", "2) провести эксперимент."]


def test_spaces_invisible_and_ligatures():
    raw = "Слово   второе​ ,  ﬁnal­ ( текст )"
    assert C.to_paragraphs(raw) == ["Слово второе, final (текст)"]


def test_windows_newlines_and_form_feed():
    assert C.to_paragraphs("Абзац один.\r\nАбзац два.\fАбзац три.") == [
        "Абзац один.", "Абзац два.", "Абзац три."]


def test_bibliography_numbers_stripped():
    raw = "1. Иванов И.И. Книга.\n[2] Петров П.П. Статья.\n3) Сидоров\n\nКозлов К.К."
    assert C.bibliography_entries(raw) == ["Иванов И.И. Книга.", "Петров П.П. Статья.", "Сидоров", "Козлов К.К."]


def test_heading_no_trailing_dot_but_keeps_ellipsis():
    assert C.heading("  Обзор   литературы. ") == "Обзор литературы"
    assert C.heading("Что дальше...") == "Что дальше..."


def test_typography_optional():
    c = TextCleaner(CleanOptions(typography=True))
    assert c.to_paragraphs('Он сказал "да" - и ушёл.') == ["Он сказал «да» – и ушёл."]


def test_title_lines_keep_empty_lines():
    assert C.lines("\nА\n\nБ\n\n", keep_empty=True) == ["А", "", "Б"]

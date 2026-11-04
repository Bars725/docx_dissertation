import importlib
import sys

from dissertation import envcheck


def test_current_environment_ok():
    problems = envcheck.check_environment(need_gui=False)
    assert not [p for p in problems if p.level == "error" and "docx" in p.title.lower()]


def test_detects_conflicting_python2_docx(tmp_path, monkeypatch):
    """Старый пакет «docx» — одиночный файл docx.py, падающий на import exceptions."""
    (tmp_path / "docx.py").write_text("from exceptions import PendingDeprecationWarning\n")
    monkeypatch.syspath_prepend(str(tmp_path))
    saved = {k: v for k, v in sys.modules.items() if k == "docx" or k.startswith("docx.")}
    for k in saved:
        monkeypatch.delitem(sys.modules, k)
    importlib.invalidate_caches()
    problems = envcheck.check_docx()
    assert len(problems) == 1 and problems[0].level == "error"
    assert "устаревший пакет «docx»" in problems[0].title
    assert "uninstall -y docx" in problems[0].fix
    report = envcheck.format_report(problems)
    assert "не может запуститься" in report


def test_version_parsing():
    assert envcheck._version_tuple("1.2.0") == (1, 2, 0)
    assert envcheck._version_tuple("1.1.0rc1") == (1, 1, 0)

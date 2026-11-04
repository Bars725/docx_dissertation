"""Точка входа.

    python -m dissertation                         — GUI
    python -m dissertation inspect образец.docx    — что извлекается из образца
    python -m dissertation profile образец.docx -o профиль.json
    python -m dissertation new out.docx [--preset gost-r-7.0.11-2011 | --profile p.json]
    python -m dissertation add doc.docx --type chapter --heading "Обзор" --text-file t.txt

CLI нужен для пакетной проверки на разных .docx (Word / LibreOffice / Google Docs)
и для воспроизведения ошибок без GUI.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .document_builder import BLOCK_LABELS, BlockType, DocumentBuilder
from .gost import PRESETS
from .profile import FormattingProfile
from .template_parser import TemplateParser


def _profile_from_args(args) -> FormattingProfile | None:
    if getattr(args, "profile", None):
        return FormattingProfile.load(args.profile)
    if getattr(args, "preset", None):
        return PRESETS[args.preset]()
    return None


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        from .gui.app import run
        run()
        return 0

    ap = argparse.ArgumentParser(prog="dissertation", description="Оформление диссертации по ГОСТ")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inspect", help="показать профиль, извлечённый из образца")
    p.add_argument("docx")

    p = sub.add_parser("profile", help="сохранить профиль образца в JSON")
    p.add_argument("docx")
    p.add_argument("-o", "--output", required=True)

    p = sub.add_parser("new", help="создать пустой документ по ГОСТ")
    p.add_argument("output")
    p.add_argument("--preset", choices=sorted(PRESETS))
    p.add_argument("--profile")

    p = sub.add_parser("add", help="добавить блок текста в документ")
    p.add_argument("docx")
    p.add_argument("--type", required=True, choices=[b.value for b in BlockType],
                   help="; ".join(f"{b.value} — {BLOCK_LABELS[b]}" for b in BlockType))
    p.add_argument("--heading", default="")
    p.add_argument("--text-file")
    p.add_argument("--raw", action="store_true", help="без очистки текста")
    p.add_argument("--preset", choices=sorted(PRESETS))
    p.add_argument("--profile")
    p.add_argument("-o", "--output", help="сохранить в другой файл")

    args = ap.parse_args(argv)
    if args.cmd == "inspect":
        prof = TemplateParser.from_path(args.docx).extract()
        print("\n".join(prof.extraction_report))
        return 0
    if args.cmd == "profile":
        TemplateParser.from_path(args.docx).extract().save(args.output)
        print(f"Профиль сохранён: {args.output}")
        return 0
    if args.cmd == "new":
        prof = _profile_from_args(args) or PRESETS["gost-7.32-2017"]()
        print(f"Создан: {DocumentBuilder.new(prof).save(args.output)}")
        return 0
    if args.cmd == "add":
        text = Path(args.text_file).read_text(encoding="utf-8") if args.text_file else sys.stdin.read()
        b = DocumentBuilder.open(args.docx, profile=_profile_from_args(args))
        res = b.insert_block(BlockType(args.type), text=text, heading=args.heading, clean=not args.raw)
        for w in res.warnings:
            print("! " + w)
        print(f"Добавлено абзацев: {len(res.paragraphs)}; сохранено: {b.save(args.output)}")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

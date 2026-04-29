#!/usr/bin/env python3
"""تَبديلُ علاماتِ التَّنصيصِ الفَرَنسيَّة « ↔ » في النَّصّ."""
import sys
from pathlib import Path

SENTINEL = "\u0001"  # حرف وسيط لا يظهر في النصّ

def swap_guillemets(text: str) -> str:
    text = text.replace("«", SENTINEL)
    text = text.replace("»", "«")
    text = text.replace(SENTINEL, "»")
    return text

def main():
    targets = [
        Path("غريب_في_طيبة.md"),
        Path("EDITORIAL_NOTES.md"),
    ]
    for path in targets:
        if not path.exists():
            print(f"[swap] تَخَطّي: {path} (غير موجود)")
            continue
        original = path.read_text(encoding="utf-8")
        before_open = original.count("«")
        before_close = original.count("»")
        swapped = swap_guillemets(original)
        after_open = swapped.count("«")
        after_close = swapped.count("»")
        path.write_text(swapped, encoding="utf-8")
        print(f"[swap] {path}")
        print(f"   قبل: « = {before_open}, » = {before_close}")
        print(f"   بعد: « = {after_open}, » = {after_close}")
        assert before_open == after_close and before_close == after_open, "تطابق الأعداد"

if __name__ == "__main__":
    main()

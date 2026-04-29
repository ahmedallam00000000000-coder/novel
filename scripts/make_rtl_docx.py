#!/usr/bin/env python3
"""
يُولِّد نسخة DOCX من المخطوط مع ضبط الاتّجاه العربيّ (من اليمين إلى اليسار)،
وخطٍّ عربيٍّ ملائم للنشر، وهوامش مناسبة لكتاب عربيّ مطبوع.

Generates a Word .docx from the manuscript with RTL paragraph direction,
Arabic-friendly fonts, and book-style page margins.

Usage:
    python3 scripts/make_rtl_docx.py
"""
from __future__ import annotations

import re
import shutil
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE_MD = ROOT / "غريب_في_طيبة.md"
TARGET_DOCX = ROOT / "غريب_في_طيبة.docx"
REFERENCE_DOCX = ROOT / "scripts" / "reference-rtl.docx"


# الخطوط المستعملة:
#   - "Amiri" أو "Traditional Arabic" — خطوط نسخ ملائمة لمتن الكتاب.
#   - "Cairo" أو "Tajawal" — خطوط حديثة للعناوين.
# نضع Amiri أوّلًا لأنّه مفتوح المصدر ومتاح، ثمّ Traditional Arabic كاحتياطيّ.
ARABIC_BODY_FONT = "Amiri"
ARABIC_HEADING_FONT = "Cairo"

# حجم الخطّ الافتراضيّ للمتن (بالأنصاف-نقاط — w:sz val="28" => 14pt)
BODY_FONT_SIZE_HALFPOINTS = "28"  # 14pt — مناسب لكتب الرواية بالعربيّة
HEADING1_SIZE_HALFPOINTS = "44"   # 22pt
HEADING2_SIZE_HALFPOINTS = "36"   # 18pt
HEADING3_SIZE_HALFPOINTS = "32"   # 16pt


def make_reference_docx() -> Path:
    """يأخذ reference.docx الافتراضيّ من pandoc ويعدّله ليصير عربيًّا RTL."""
    REFERENCE_DOCX.parent.mkdir(parents=True, exist_ok=True)
    # توليد reference.docx الافتراضيّ
    default_ref = ROOT / "scripts" / "_pandoc-default-reference.docx"
    subprocess.run(
        ["pandoc", "-o", str(default_ref), "--print-default-data-file", "reference.docx"],
        check=True,
    )

    # نفك الضغط، ونعدّل الملفّات الداخليّة، ثمّ نعيد الضغط
    work_dir = ROOT / "scripts" / "_refdocx_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)
    with zipfile.ZipFile(default_ref) as z:
        z.extractall(work_dir)

    patch_styles_xml(work_dir / "word" / "styles.xml")
    patch_settings_xml(work_dir / "word" / "settings.xml")

    # إعادة الضغط بترتيب الملفّات الأصليّ
    if REFERENCE_DOCX.exists():
        REFERENCE_DOCX.unlink()
    with zipfile.ZipFile(REFERENCE_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(work_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(work_dir).as_posix())

    # تنظيف
    shutil.rmtree(work_dir)
    default_ref.unlink()

    return REFERENCE_DOCX


def patch_styles_xml(path: Path) -> None:
    """يضيف bidi إلى الفقرات وخطًّا عربيًّا للنصوص (cs=complex script)."""
    text = path.read_text(encoding="utf-8")

    # 1) ضبط الخطّ الافتراضيّ للنصّ المركّب (Arabic) في docDefaults
    #    نستبدل سطر rFonts ليتضمّن w:cs (complex script font) باسم خطّنا العربيّ.
    new_rfonts = (
        f'<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" '
        f'w:hAnsiTheme="minorHAnsi" w:cs="{ARABIC_BODY_FONT}" w:cstheme="minorBidi" />'
    )
    text = re.sub(
        r'<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" '
        r'w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" />',
        new_rfonts,
        text,
        count=1,
    )

    # 2) في docDefaults > rPrDefault > rPr، نضبط حجم الخطّ المركّب
    text = re.sub(
        r'(<w:szCs w:val=")24(" />)',
        rf'\g<1>{BODY_FONT_SIZE_HALFPOINTS}\g<2>',
        text,
        count=1,
    )
    text = re.sub(
        r'(<w:sz w:val=")24(" />)',
        rf'\g<1>{BODY_FONT_SIZE_HALFPOINTS}\g<2>',
        text,
        count=1,
    )

    # 3) في pPrDefault، نضيف <w:bidi/> ليصبح كلّ الفقرات بالاتّجاه العربيّ
    text = text.replace(
        "<w:pPrDefault>\n      <w:pPr>\n        <w:spacing w:after=\"200\" />\n      </w:pPr>\n    </w:pPrDefault>",
        '<w:pPrDefault>\n      <w:pPr>\n        <w:bidi />\n        <w:spacing w:after="200" />\n        <w:jc w:val="both" />\n      </w:pPr>\n    </w:pPrDefault>',
    )

    # 4) في كلّ نمط (style) من نوع paragraph، نضيف <w:bidi/> داخل w:pPr إن لم يكن موجودًا.
    #    نستعمل خاصّيّة جانبيّة: إن كان للنمط w:pPr نضع bidi في أوّله؛ وإن لم يكن، نضيفه.
    def add_bidi_to_style(m: re.Match[str]) -> str:
        block = m.group(0)
        # إن وُجد bidi مسبقًا، لا تكرّره
        if "<w:bidi" in block:
            return block
        # إن وُجد <w:pPr>، أضف <w:bidi/> داخله
        if "<w:pPr>" in block:
            return block.replace("<w:pPr>", "<w:pPr>\n      <w:bidi />", 1)
        # وإلّا أضف <w:pPr><w:bidi/></w:pPr> قبل إغلاق w:style
        return block.replace(
            "</w:style>", "<w:pPr><w:bidi /></w:pPr></w:style>", 1
        )

    text = re.sub(
        r'<w:style w:type="paragraph"[^>]*>.*?</w:style>',
        add_bidi_to_style,
        text,
        flags=re.DOTALL,
    )

    # 5) ضبط محاذاة العناوين والمتن
    #    العنوان الرئيسيّ (Title) - وسط
    #    العناوين 1-3 - يمين
    #    المتن - مضبوط (justify) من الجهتين

    path.write_text(text, encoding="utf-8")


def patch_settings_xml(path: Path) -> None:
    """يضبط لغة الواجهة العربيّة وخصائص المستند."""
    text = path.read_text(encoding="utf-8")

    # ضبط ثيم اللغة العربيّة
    text = text.replace(
        '<w:themeFontLang w:val="en-US" />',
        '<w:themeFontLang w:val="en-US" w:bidi="ar-SA" />',
    )

    # تشغيل دعم الـ Bidi
    if "<w:bidi" not in text:
        text = text.replace(
            "</w:settings>",
            '<w:bidi />\n</w:settings>',
        )

    path.write_text(text, encoding="utf-8")


def patch_generated_docx(docx_path: Path) -> None:
    """بعد توليد الـdocx من pandoc، نضبط خصائص الصفحة (RTL، حجم، هوامش)."""
    work_dir = ROOT / "scripts" / "_outdocx_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(work_dir)

    patch_document_xml(work_dir / "word" / "document.xml")

    # إعادة الضغط
    docx_path.unlink()
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(work_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(work_dir).as_posix())

    shutil.rmtree(work_dir)


def patch_document_xml(path: Path) -> None:
    """يضبط sectPr (خصائص المقطع/الصفحة) لتكون عربيّة RTL وبهوامش كتاب."""
    text = path.read_text(encoding="utf-8")

    # هوامش مناسبة للنشر العربيّ (للكتب: حوالي ٢٫٥سم من كلّ جهة)
    # القيم بـ twips (1 inch = 1440 twips، 1 cm ≈ 567 twips)
    page_size = '<w:pgSz w:w="11906" w:h="16838" />'  # A4: 21.0 × 29.7 cm
    page_margin = (
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" '
        'w:left="1418" w:header="708" w:footer="708" w:gutter="0" />'
    )

    new_sect_pr = (
        f'<w:sectPr>'
        f'{page_size}'
        f'{page_margin}'
        f'<w:bidi />'  # اتّجاه المقطع من اليمين إلى اليسار
        f'<w:cols w:space="708" />'
        f'<w:docGrid w:linePitch="360" />'
        f'</w:sectPr>'
    )

    # إذا كان هناك sectPr مسبق (سواء self-closing أو لا)، استبدله؛ وإلّا أضفه قبل </w:body>
    if re.search(r"<w:sectPr\b[^>]*?/>", text):
        text = re.sub(r"<w:sectPr\b[^>]*?/>", new_sect_pr, text)
    elif "<w:sectPr" in text:
        text = re.sub(r"<w:sectPr\b.*?</w:sectPr>", new_sect_pr, text, flags=re.DOTALL)
    else:
        text = text.replace("</w:body>", f"{new_sect_pr}</w:body>")

    path.write_text(text, encoding="utf-8")


def generate() -> None:
    print("== توليد reference docx عربيّ RTL ==")
    ref = make_reference_docx()
    print(f"  reference: {ref.relative_to(ROOT)}")

    print("== تشغيل pandoc ==")
    subprocess.run(
        [
            "pandoc",
            str(SOURCE_MD),
            "-o", str(TARGET_DOCX),
            "--from", "markdown",
            "--to", "docx",
            f"--reference-doc={ref}",
        ],
        check=True,
    )
    print(f"  generated: {TARGET_DOCX.relative_to(ROOT)}")

    print("== ضبط خصائص الصفحة (RTL، A4، هوامش) ==")
    patch_generated_docx(TARGET_DOCX)
    print("  done")

    size_kb = TARGET_DOCX.stat().st_size / 1024
    print(f"\n✔ الناتج: {TARGET_DOCX.relative_to(ROOT)} ({size_kb:.1f} KB)")


if __name__ == "__main__":
    generate()

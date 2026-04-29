#!/usr/bin/env python3
"""
يُولِّد نسخة DOCX احترافيّة من المخطوط بتنسيق كتاب عربيّ منشور:
- اتّجاه نصّ وصفحة من اليمين إلى اليسار (RTL)
- خطّ عربيّ ملائم للنشر (Amiri)
- صفحة بحجم A5 (148 × 210 مم) — بمقاس روائيّ
- هوامش غير متماثلة مرآويّة (داخل/خارج) لكتاب مطبوع
- عنوان كبير في صفحة مستقلّة
- بداية كلّ قسم/فصل في صفحة جديدة، عناوين في المنتصف
- اقتباسات (epigraphs) في المنتصف بخطّ مائل
- مسافة بادئة لأوّل سطر في الفقرات (طريقة الكتب المنشورة)
- تباعد سطور 1.5 لراحة القراءة

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


# الخطّ الأساسيّ (يمكن تغييره — الخطوط البديلة المتوفّرة عادةً:
# Traditional Arabic, Cairo, Tajawal, Lateef, Scheherazade, Arabic Typesetting).
ARABIC_BODY_FONT = "Amiri"

# ================== أحجام الخطوط بالأنصاف-نقاط ==================
# (w:sz val="28" => 14pt)  أنصاف-نقاط = نقاط × 2
SZ_BODY = "26"        # 13pt للمتن — مناسب لكتاب A5
SZ_FIRST_PARA = "26"
SZ_BLOCK = "24"       # 12pt للاقتباسات
SZ_TITLE = "72"       # 36pt لعنوان الكتاب
SZ_SUBTITLE = "32"    # 16pt
SZ_AUTHOR = "32"      # 16pt
SZ_H1 = "44"          # 22pt للأقسام
SZ_H2 = "36"          # 18pt للفصول
SZ_H3 = "30"          # 15pt للعناوين الفرعيّة

# ================== أبعاد الصفحة بالـ twips (567 ≈ 1سم) ==================
# A5: 14.8 × 21.0 سم  (8392 × 11906 twips تقريبًا)
PAGE_W = "8391"       # عرض A5
PAGE_H = "11906"      # ارتفاع A5

# هوامش مرآويّة (داخل/خارج) — مناسبة للتجليد العربيّ
MARGIN_TOP = "1134"     # 2.0سم
MARGIN_BOTTOM = "1134"  # 2.0سم
MARGIN_INNER = "1418"   # 2.5سم (الجهة المجلَّدة — اليمين في الكتاب العربيّ)
MARGIN_OUTER = "1134"   # 2.0سم
MARGIN_HEADER = "709"
MARGIN_FOOTER = "709"


# ================== خصائص الفقرات ==================
# تباعد السطور (line spacing). 360 = 1.5 سطر
LINE_HEIGHT = "360"
# المسافة البادئة لأوّل سطر في فقرات المتن (twips)، 567 = 1سم
FIRST_LINE_INDENT = "567"


# ============================================================
# توليد القالب المرجعيّ (reference-rtl.docx)
# ============================================================

def make_reference_docx() -> Path:
    """يأخذ reference.docx الافتراضيّ من pandoc ويعدّله ليصير قالبًا عربيًّا احترافيًّا."""
    REFERENCE_DOCX.parent.mkdir(parents=True, exist_ok=True)
    default_ref = ROOT / "scripts" / "_pandoc-default-reference.docx"
    subprocess.run(
        ["pandoc", "-o", str(default_ref), "--print-default-data-file", "reference.docx"],
        check=True,
    )

    work_dir = ROOT / "scripts" / "_refdocx_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)
    with zipfile.ZipFile(default_ref) as z:
        z.extractall(work_dir)

    patch_styles_xml(work_dir / "word" / "styles.xml")
    patch_settings_xml(work_dir / "word" / "settings.xml")

    if REFERENCE_DOCX.exists():
        REFERENCE_DOCX.unlink()
    with zipfile.ZipFile(REFERENCE_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(work_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(work_dir).as_posix())

    shutil.rmtree(work_dir)
    default_ref.unlink()
    return REFERENCE_DOCX


def replace_style(text: str, style_id: str, new_block: str) -> str:
    """يستبدل كتلة style كاملة بتعريف جديد."""
    pattern = rf'<w:style\s+w:type="paragraph"[^>]*\bw:styleId="{re.escape(style_id)}"[^>]*>.*?</w:style>'
    if not re.search(pattern, text, flags=re.DOTALL):
        # إن لم يوجد، أضف قبل </w:styles>
        return text.replace("</w:styles>", new_block + "\n</w:styles>")
    return re.sub(pattern, lambda m: new_block, text, count=1, flags=re.DOTALL)


def font_run(size: str, *, bold: bool = False, italic: bool = False, color: str | None = None) -> str:
    """يبني <w:rPr> مع خطّ عربيّ وحجم وخصائص اختياريّة."""
    parts = [f'<w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" '
             f'w:cs="{ARABIC_BODY_FONT}" />']
    if bold:
        parts.append('<w:b />')
        parts.append('<w:bCs />')
    if italic:
        parts.append('<w:i />')
        parts.append('<w:iCs />')
    if color:
        parts.append(f'<w:color w:val="{color}" />')
    parts.append(f'<w:sz w:val="{size}" />')
    parts.append(f'<w:szCs w:val="{size}" />')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def patch_styles_xml(path: Path) -> None:
    """يستبدل أنماط Pandoc الافتراضيّة بأنماط احترافيّة لكتاب عربيّ."""
    text = path.read_text(encoding="utf-8")

    # -------- (1) docDefaults: خطّ افتراضيّ عربيّ + bidi افتراضيّ + تباعد --------
    new_doc_defaults = f'''<w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" w:cs="{ARABIC_BODY_FONT}" w:eastAsiaTheme="minorHAnsi" />
        <w:sz w:val="{SZ_BODY}" />
        <w:szCs w:val="{SZ_BODY}" />
        <w:lang w:val="ar-SA" w:eastAsia="ar-SA" w:bidi="ar-SA" />
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:bidi />
        <w:spacing w:before="0" w:after="120" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
        <w:jc w:val="both" />
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>'''
    text = re.sub(
        r"<w:docDefaults>.*?</w:docDefaults>",
        lambda m: new_doc_defaults,
        text,
        count=1,
        flags=re.DOTALL,
    )

    # -------- (2) Title: عنوان الكتاب — كبير، وسط، عريض، أسود --------
    title_style = f'''<w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="Author" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:keepLines />
      <w:pageBreakBefore />
      <w:spacing w:before="3600" w:after="800" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
    </w:pPr>
    {font_run(SZ_TITLE, bold=True, color="1F4E79")}
  </w:style>'''
    text = replace_style(text, "Title", title_style)

    # -------- (3) Subtitle --------
    subtitle_style = f'''<w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="Author" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:pageBreakBefore w:val="0" />
      <w:spacing w:before="600" w:after="400" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
    </w:pPr>
    {font_run(SZ_SUBTITLE, italic=True)}
  </w:style>'''
    text = replace_style(text, "Subtitle", subtitle_style)

    # -------- (4) Author --------
    author_style = f'''<w:style w:type="paragraph" w:customStyle="1" w:styleId="Author">
    <w:name w:val="Author" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="BodyText" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:pageBreakBefore w:val="0" />
      <w:spacing w:before="200" w:after="2400" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
    </w:pPr>
    {font_run(SZ_AUTHOR, bold=True)}
  </w:style>'''
    text = replace_style(text, "Author", author_style)

    # -------- (5) Heading1: الأقسام (#) — صفحة جديدة، وسط، كبير --------
    h1_style = f'''<w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="BodyText" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:keepLines />
      <w:pageBreakBefore />
      <w:spacing w:before="3600" w:after="1200" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
      <w:outlineLvl w:val="0" />
    </w:pPr>
    {font_run(SZ_H1, bold=True, color="1F4E79")}
  </w:style>'''
    text = replace_style(text, "Heading1", h1_style)

    # -------- (6) Heading2: الفصول (##) — صفحة جديدة، وسط --------
    h2_style = f'''<w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="FirstParagraph" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:keepLines />
      <w:pageBreakBefore />
      <w:spacing w:before="2400" w:after="800" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
      <w:outlineLvl w:val="1" />
    </w:pPr>
    {font_run(SZ_H2, bold=True, color="1F4E79")}
  </w:style>'''
    text = replace_style(text, "Heading2", h2_style)

    # -------- (7) Heading3: العناوين الفرعيّة (###) — وسط، مائل --------
    h3_style = f'''<w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="FirstParagraph" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:keepLines />
      <w:spacing w:before="600" w:after="240" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
      <w:outlineLvl w:val="2" />
    </w:pPr>
    {font_run(SZ_H3, bold=True, italic=True, color="2E75B6")}
  </w:style>'''
    text = replace_style(text, "Heading3", h3_style)

    # -------- (8) BodyText: متن الرواية — مضبوط مع مسافة بادئة --------
    body_style = f'''<w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="BodyTextChar" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:spacing w:before="0" w:after="0" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:ind w:firstLine="{FIRST_LINE_INDENT}" />
      <w:jc w:val="both" />
    </w:pPr>
    {font_run(SZ_BODY)}
  </w:style>'''
    text = replace_style(text, "BodyText", body_style)

    # -------- (9) FirstParagraph: أوّل فقرة بعد عنوان (لا مسافة بادئة) --------
    first_para_style = f'''<w:style w:type="paragraph" w:customStyle="1" w:styleId="FirstParagraph">
    <w:name w:val="First Paragraph" />
    <w:basedOn w:val="BodyText" />
    <w:next w:val="BodyText" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:spacing w:before="240" w:after="0" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:ind w:firstLine="0" />
      <w:jc w:val="both" />
    </w:pPr>
    {font_run(SZ_FIRST_PARA)}
  </w:style>'''
    text = replace_style(text, "FirstParagraph", first_para_style)

    # -------- (10) BlockText: الاقتباسات/الإهداء — وسط، مائل --------
    block_text_style = f'''<w:style w:type="paragraph" w:styleId="BlockText">
    <w:name w:val="Block Text" />
    <w:basedOn w:val="Normal" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:spacing w:before="240" w:after="240" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:ind w:left="567" w:right="567" />
      <w:jc w:val="center" />
    </w:pPr>
    {font_run(SZ_BLOCK, italic=True, color="3F3F3F")}
  </w:style>'''
    text = replace_style(text, "BlockText", block_text_style)

    # -------- (10b) GroupHeading: عناوين فرعيّة في الفهرس/دليل الشخصيّات --------
    group_heading_style = f'''<w:style w:type="paragraph" w:customStyle="1" w:styleId="GroupHeading">
    <w:name w:val="Group Heading" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="Normal" />
    <w:qFormat />
    <w:pPr>
      <w:bidi />
      <w:keepNext />
      <w:keepLines />
      <w:spacing w:before="480" w:after="240" w:line="{LINE_HEIGHT}" w:lineRule="auto" />
      <w:jc w:val="center" />
    </w:pPr>
    {font_run(SZ_H3, bold=True, color="1F4E79")}
  </w:style>'''
    text = replace_style(text, "GroupHeading", group_heading_style)

    # -------- (11) إضافة <w:bidi/> للأنماط المتبقّية --------
    def add_bidi_to_style(m: re.Match[str]) -> str:
        block = m.group(0)
        if "<w:bidi" in block:
            return block
        if "<w:pPr>" in block:
            return block.replace("<w:pPr>", "<w:pPr>\n      <w:bidi />", 1)
        return block.replace("</w:style>", "<w:pPr><w:bidi /></w:pPr></w:style>", 1)

    text = re.sub(
        r'<w:style w:type="paragraph"[^>]*>.*?</w:style>',
        add_bidi_to_style,
        text,
        flags=re.DOTALL,
    )

    path.write_text(text, encoding="utf-8")


def patch_settings_xml(path: Path) -> None:
    """يضبط لغة الواجهة العربيّة وخصائص المستند."""
    text = path.read_text(encoding="utf-8")
    text = text.replace(
        '<w:themeFontLang w:val="en-US" />',
        '<w:themeFontLang w:val="en-US" w:bidi="ar-SA" />',
    )
    if "<w:bidi" not in text:
        text = text.replace("</w:settings>", "<w:bidi />\n</w:settings>")
    path.write_text(text, encoding="utf-8")


# ============================================================
# تعديل docx الناتج (لا يتأثّر بالقالب المرجعيّ)
# ============================================================

def patch_generated_docx(docx_path: Path) -> None:
    """بعد توليد الـdocx من pandoc، نضبط خصائص الصفحة (RTL، حجم، هوامش، تجليد)."""
    work_dir = ROOT / "scripts" / "_outdocx_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(work_dir)

    add_headers_and_footers(work_dir)
    patch_document_xml(work_dir / "word" / "document.xml")

    docx_path.unlink()
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(work_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(work_dir).as_posix())
    shutil.rmtree(work_dir)


# ============================================================
# الرؤوس والتذييلات وأرقام الصفحات
# ============================================================

# عنوان الكتاب — يَظهَرُ في رَأسِ الصَّفحة في كلِّ صفحات الكتاب ما عدا الأَوائل.
HEADER_TITLE = "غريب في طيبة"

# قالب رأس الصفحة (عنوان الكتاب وسط، خطّ Amiri صغير، اتّجاه RTL، حدّ سُفليّ خفيف)
HEADER_DEFAULT_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:bidi />
      <w:jc w:val="center" />
      <w:pBdr>
        <w:bottom w:val="single" w:sz="4" w:space="4" w:color="999999" />
      </w:pBdr>
      <w:spacing w:before="0" w:after="120" />
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" w:cs="{ARABIC_BODY_FONT}" />
        <w:i />
        <w:iCs />
        <w:color w:val="666666" />
        <w:sz w:val="20" />
        <w:szCs w:val="20" />
      </w:rPr>
      <w:t xml:space="preserve">{HEADER_TITLE}</w:t>
    </w:r>
  </w:p>
</w:hdr>'''

# رأس صفحة الـفصل/العنوان (فارغ — لا نريد رأسًا فوق صفحة بداية الفصل)
HEADER_EMPTY_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:bidi /><w:jc w:val="center" /></w:pPr>
  </w:p>
</w:hdr>'''

# تَذييل افتراضيّ — رَقم الصَّفحة في المُنتَصَف بِالأَرقام العربيّة-الهنديّة (١٢٣)
FOOTER_DEFAULT_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:bidi />
      <w:jc w:val="center" />
      <w:spacing w:before="120" w:after="0" />
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" w:cs="{ARABIC_BODY_FONT}" />
        <w:color w:val="666666" />
        <w:sz w:val="20" />
        <w:szCs w:val="20" />
      </w:rPr>
      <w:fldChar w:fldCharType="begin" />
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" w:cs="{ARABIC_BODY_FONT}" />
        <w:color w:val="666666" />
        <w:sz w:val="20" />
        <w:szCs w:val="20" />
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" w:cs="{ARABIC_BODY_FONT}" />
        <w:color w:val="666666" />
        <w:sz w:val="20" />
        <w:szCs w:val="20" />
      </w:rPr>
      <w:fldChar w:fldCharType="end" />
    </w:r>
  </w:p>
</w:ftr>'''

# تَذييل الصَّفحة الأُولى (صفحة العنوان) — فارغ
FOOTER_EMPTY_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:bidi /><w:jc w:val="center" /></w:pPr>
  </w:p>
</w:ftr>'''


def add_headers_and_footers(work_dir: Path) -> None:
    """يُضيفُ ملفَّات الرأس/التذييل ويسجِّلُها في الـrelationships وContent_Types.

    لا نُضيفُ مراجعَ في sectPr هنا — يَفعَلُه patch_document_xml بعد ذلك،
    لأنّه يَستبدِلُ sectPr بأكمَلِه.
    """
    word_dir = work_dir / "word"
    rels_dir = word_dir / "_rels"

    (word_dir / "header1.xml").write_text(HEADER_DEFAULT_XML, encoding="utf-8")
    (word_dir / "header2.xml").write_text(HEADER_EMPTY_XML, encoding="utf-8")
    (word_dir / "footer1.xml").write_text(FOOTER_DEFAULT_XML, encoding="utf-8")
    (word_dir / "footer2.xml").write_text(FOOTER_EMPTY_XML, encoding="utf-8")

    # تَسجيلٌ في word/_rels/document.xml.rels
    rels_path = rels_dir / "document.xml.rels"
    rels_text = rels_path.read_text(encoding="utf-8")
    extra_rels = (
        '<Relationship Id="rId100" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml" />'
        '<Relationship Id="rId101" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header2.xml" />'
        '<Relationship Id="rId102" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer1.xml" />'
        '<Relationship Id="rId103" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer2.xml" />'
    )
    rels_text = rels_text.replace("</Relationships>", extra_rels + "</Relationships>")
    rels_path.write_text(rels_text, encoding="utf-8")

    # تَسجيلٌ في [Content_Types].xml
    ct_path = work_dir / "[Content_Types].xml"
    ct_text = ct_path.read_text(encoding="utf-8")
    extra_ct = (
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" />'
        '<Override PartName="/word/header2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" />'
        '<Override PartName="/word/footer1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" />'
        '<Override PartName="/word/footer2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" />'
    )
    ct_text = ct_text.replace("</Types>", extra_ct + "</Types>")
    ct_path.write_text(ct_text, encoding="utf-8")


def _apply_group_heading(text: str) -> str:
    """يطبّق نمط GroupHeading على الفقرات التي محتواها كلّه run واحد بنصّ عريض.

    يلتقط الفقرات على شكل:
        <w:p><w:pPr>…</w:pPr><w:r><w:rPr><w:b /><w:bCs/></w:rPr><w:t>…</w:t></w:r></w:p>
    حيث لا يوجد إلّا run واحد في الفقرة، ويستبدل pStyle (إن وُجد) بـ GroupHeading.
    """
    pattern = re.compile(
        r'<w:p>'
        r'(?P<ppr><w:pPr>(?:(?!</w:pPr>).)*</w:pPr>)?'
        r'<w:r>'
        r'<w:rPr>(?P<rpr>(?:\s*<w:b\s*/>\s*|\s*<w:bCs\s*/>\s*)+)</w:rPr>'
        r'<w:t(?P<tattr>[^>]*)>(?P<txt>[^<]+)</w:t>'
        r'</w:r>'
        r'</w:p>',
        flags=re.DOTALL,
    )

    def repl(m: re.Match[str]) -> str:
        # نتجاهل الفقرات داخل عناصر القائمة (التي لها numPr) — لكن أصلًا الـ regex
        # يلتقط فقرات بـ run واحد فقط، فالقوائم (التي تحوي bullet + text) لن تُلتقط.
        return (
            f'<w:p><w:pPr><w:pStyle w:val="GroupHeading" /></w:pPr>'
            f'<w:r><w:t{m.group("tattr")}>{m.group("txt")}</w:t></w:r></w:p>'
        )

    return pattern.sub(repl, text)


def _style_title_page(text: str) -> str:
    """يطبّق نمطَي Subtitle/Author على صفحة العنوان (الفقرتين بعد عنوان الكتاب)."""
    # «رواية» (مائلة) → Subtitle
    text = re.sub(
        r'<w:p>(?:<w:pPr>(?:(?!</w:pPr>).)*</w:pPr>)?<w:r><w:rPr>'
        r'(?:\s*<w:i\s*/>\s*|\s*<w:iCs\s*/>\s*)+</w:rPr>'
        r'<w:t([^>]*)>([^<]*رواية[^<]*)</w:t></w:r></w:p>',
        lambda m: (
            f'<w:p><w:pPr><w:pStyle w:val="Subtitle" /></w:pPr>'
            f'<w:r><w:t{m.group(1)}>{m.group(2)}</w:t></w:r></w:p>'
        ),
        text,
        count=1,
        flags=re.DOTALL,
    )
    # ملاحظة: «أحمد علام» (عريضة) سَيلتقطه _apply_group_heading قبل هذه الدالة،
    # فنُعيد استبدال GroupHeading بـ Author لأنّه على صفحة العنوان.
    text = re.sub(
        r'<w:p><w:pPr><w:pStyle w:val="GroupHeading" /></w:pPr>'
        # نطابق «أحمد عل» فقط (بدون آخر الكلمة) لأنّ المصدر قد يحتوي
        # على تشكيل (فتحة U+064E + شدّة U+0651) بين «عل» و«ام».
        r'<w:r><w:t([^>]*)>([^<]*أحمد عل[^<]*)</w:t></w:r></w:p>',
        lambda m: (
            f'<w:p><w:pPr><w:pStyle w:val="Author" /></w:pPr>'
            f'<w:r><w:t{m.group(1)}>{m.group(2)}</w:t></w:r></w:p>'
        ),
        text,
        count=1,
    )
    return text


def _add_half_title_page(text: str) -> str:
    """يَحقُنُ صفحة شطر العنوان (Half-title) قبل صفحة العنوان الرئيسة.

    شطرُ العنوان عُرفٌ نَشريٌّ كلاسيكيّ: صفحةٌ مُنفَردة تَحوي اسمَ الكتاب وحدَه،
    تَأتي قبل صفحة العنوان الكاملة (العنوان + المؤلف + الناشر).
    """
    half_title = (
        '<w:p>'
        '<w:pPr>'
        '<w:bidi />'
        '<w:jc w:val="center" />'
        f'<w:spacing w:before="4800" w:after="0" w:line="{LINE_HEIGHT}" w:lineRule="auto" />'
        '</w:pPr>'
        '<w:r>'
        '<w:rPr>'
        f'<w:rFonts w:ascii="{ARABIC_BODY_FONT}" w:hAnsi="{ARABIC_BODY_FONT}" '
        f'w:cs="{ARABIC_BODY_FONT}" />'
        '<w:b /><w:bCs />'
        '<w:color w:val="1F4E79" />'
        '<w:sz w:val="56" /><w:szCs w:val="56" />'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{HEADER_TITLE}</w:t>'
        '</w:r>'
        '</w:p>'
    )
    # نَحقُنُها مباشرةً بعد <w:body>. نمطُ Title له pageBreakBefore فيَدفَعُ
    # العنوانَ الرئيسَ إلى صفحةٍ جديدة.
    text = text.replace("<w:body>", "<w:body>" + half_title, 1)
    return text


def _split_title_section(text: str) -> str:
    """يَحقُنُ فاصلَ مَقطَعٍ (sectPr) بعد فقرة Author مباشرةً.

    النَّتيجة: صفحتا الافتتاح (شطر العنوان + صفحة العنوان الرئيسة) تَكونان
    في مَقطَعٍ مُستَقِلٍّ بلا رأسٍ ولا تَذييل ولا أرقامِ صَفحات؛ ويُستَأنَفُ
    التَّرقيم من ١ في المَقطَع التالي.
    """
    page_size = f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}" />'
    page_margin = (
        f'<w:pgMar w:top="{MARGIN_TOP}" w:right="{MARGIN_INNER}" '
        f'w:bottom="{MARGIN_BOTTOM}" w:left="{MARGIN_OUTER}" '
        f'w:header="{MARGIN_HEADER}" w:footer="{MARGIN_FOOTER}" w:gutter="0" />'
    )

    # مَقطَع صفحات الافتتاح: لا رؤوس ولا تَذييلات ولا ترقيم
    front_section_pr = (
        '<w:pPr><w:sectPr>'
        f'{page_size}{page_margin}'
        '<w:bidi />'
        '<w:cols w:space="708" />'
        '<w:docGrid w:linePitch="360" />'
        '</w:sectPr></w:pPr>'
    )

    # نَجِدُ فقرة Author ونَحقُنُ فاصلَ مَقطَعٍ بَعدَها مباشرةً.
    # شكلُ الفقرة كما يَخرُجُ بَعد _style_title_page:
    # <w:p><w:pPr><w:pStyle w:val="Author" /></w:pPr>...</w:p>
    author_pattern = re.compile(
        r'(<w:p><w:pPr><w:pStyle w:val="Author"\s*/></w:pPr>'
        r'<w:r><w:t[^>]*>[^<]+</w:t></w:r></w:p>)',
    )
    section_break_para = f'<w:p>{front_section_pr}</w:p>'
    text = author_pattern.sub(r'\1' + section_break_para, text, count=1)
    return text


def _strip_redundant_horizontal_rules(text: str) -> str:
    """يحذف فقرات الخطّ الأفقيّ (---) التي تَكونُ وحيدةً أو زائدةً.

    حالات الحذف:
    1) خطّ أفقيّ يَلي مباشرةً فقرةَ نَمَطِ صفحة العنوان (Title/Subtitle/Author):
       صفحة العنوان مُهَنْدَسَةٌ بحيث لا تَستوعِبُ الفاصِل تَحتَ الاسم،
       والمُتَبَقّي يُدفَعُ إلى الصفحة التَّالية فيَيتَمُ الفاصلُ في رأسها.
    2) خطّ أفقيّ يَسبِقُ عنوانًا يَبدَأُ صفحة جديدة (Title/Heading1/Heading2):
       العنوان نفسُه يَكسِرُ الصفحة فيُصبِحُ الفاصلُ وحيدًا في صفحةٍ سابقة.
    """
    # Pattern لخطّ أفقيّ من Pandoc (يستعمل v:rect ضمن w:pict)
    hr_pattern = r'<w:p><w:r><w:pict>(?:(?!</w:pict>).)*</w:pict></w:r></w:p>'

    # 1) خطّ أفقيّ يَتبَعُ مباشرةً فقرة Author (تَيتُمُه أعلى الصفحة التالية)
    author_then_hr = re.compile(
        r'(<w:p><w:pPr><w:pStyle w:val="Author"\s*/></w:pPr>'
        r'<w:r><w:t[^>]*>[^<]+</w:t></w:r></w:p>)' + hr_pattern,
        flags=re.DOTALL,
    )
    text = author_then_hr.sub(r'\1', text)

    # 2) خطّ أفقيّ يَسبِقُ عنوانًا يَكسِرُ الصفحة
    next_pagebreak = (
        r'(?=<w:p><w:pPr><w:pStyle w:val="(?:Title|Heading1|Heading2)"\s*/>)'
    )
    text = re.sub(hr_pattern + next_pagebreak, "", text, flags=re.DOTALL)
    return text


def patch_document_xml(path: Path) -> None:
    """يضبط sectPr وأنماط الفقرات الخاصّة (عناوين فرعيّة، صفحة عنوان…)."""
    text = path.read_text(encoding="utf-8")

    # --- (1) فقرات «عنوان مجموعة» (bold-only في المنتصف) ---
    # نُعطي أيّ فقرة محتواها كلّه run واحد بـ <w:b/> نمطَ GroupHeading.
    text = _apply_group_heading(text)

    # --- (2) صفحة العنوان: «رواية» و«أحمد علام» في الوسط ---
    text = _style_title_page(text)

    # --- (3) إزالة الخطوط الفاصلة الزائدة قبل العناوين التي تبدأ صفحة جديدة ---
    text = _strip_redundant_horizontal_rules(text)

    # --- (4) إضافة صفحة شطر العنوان (Half-title) قبل صفحة العنوان الرَّئيسة ---
    text = _add_half_title_page(text)

    # --- (5) فاصلٌ مَقطَعيّ بعد صفحة العنوان حتّى تَخلوَ صفحتا الافتتاح
    #         (شطر العنوان + العنوان الرئيس) من رأسٍ وتَذييل وأرقامِ صَفحات ---
    text = _split_title_section(text)

    page_size = f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}" />'
    # هوامش مرآويّة: w:gutter ليس مفيدًا هنا — نضبط top/right/bottom/left
    # في كتاب عربيّ مجلَّد على اليمين: العمود اليمين = داخل (inner)، اليسار = خارج (outer)
    page_margin = (
        f'<w:pgMar w:top="{MARGIN_TOP}" w:right="{MARGIN_INNER}" '
        f'w:bottom="{MARGIN_BOTTOM}" w:left="{MARGIN_OUTER}" '
        f'w:header="{MARGIN_HEADER}" w:footer="{MARGIN_FOOTER}" w:gutter="0" />'
    )

    # مَراجع الرَّأس والتَّذييل (أَنشَأناها في add_headers_and_footers).
    # type="default" للصفحات العاديّة، type="first" لصفحة العنوان (يُفعَّل بـ<w:titlePg/>).
    header_footer_refs = (
        '<w:headerReference r:id="rId100" w:type="default" />'
        '<w:headerReference r:id="rId101" w:type="first" />'
        '<w:footerReference r:id="rId102" w:type="default" />'
        '<w:footerReference r:id="rId103" w:type="first" />'
    )

    # أرقام الصَّفحات بالأَرقام العربيّة-الهنديّة (١٢٣)
    page_num = '<w:pgNumType w:fmt="hindiNumbers" />'

    new_sect_pr = (
        f'<w:sectPr>'
        f'{header_footer_refs}'
        f'{page_size}'
        f'{page_margin}'
        f'{page_num}'
        f'<w:bidi />'
        f'<w:cols w:space="708" />'
        f'<w:titlePg />'
        f'<w:docGrid w:linePitch="360" />'
        f'</w:sectPr>'
    )

    if re.search(r"<w:sectPr\b[^>]*?/>", text):
        text = re.sub(r"<w:sectPr\b[^>]*?/>", new_sect_pr, text)
    elif "<w:sectPr" in text:
        text = re.sub(r"<w:sectPr\b.*?</w:sectPr>", new_sect_pr, text, flags=re.DOTALL)
    else:
        text = text.replace("</w:body>", f"{new_sect_pr}</w:body>")

    path.write_text(text, encoding="utf-8")


# ============================================================
# نقطة الدخول
# ============================================================

def generate() -> None:
    print("== توليد القالب المرجعيّ (reference-rtl.docx) ==")
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

    print("== ضبط خصائص الصفحة (RTL، A5، هوامش كتاب) ==")
    patch_generated_docx(TARGET_DOCX)
    print("  done")

    size_kb = TARGET_DOCX.stat().st_size / 1024
    print(f"\n✔ الناتج: {TARGET_DOCX.relative_to(ROOT)} ({size_kb:.1f} KB)")


if __name__ == "__main__":
    generate()

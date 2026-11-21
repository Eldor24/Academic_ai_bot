"""
ilmiy_shablon.py — «👨‍🔬 Ilmiy izlanuvchi» bo'limi Word hujjatlari uchun yagona chiroyli shablon
═══════════════════════════════════════════════════════════════════
Barcha hujjatlarda bir xil ko'rinish:
  • yuqorida yupqa sarlavha qatori (bo'lim nomi · sana), pastda sahifa raqami va manba
  • to'q ko'k sarlavha, kursiv izoh, kulrang fonli «Parametrlar» jadvali
  • bo'lim sarlavhalari — pastki chiziqli, raqamlangan
  • kartochkalar (adabiyot, jurnal) — sarlavha, ma'lumot qatori, belgilar, izoh, havola
  • jadvallar — to'q ko'k sarlavha qatori (oq matn), navbatma-navbat och fon
Shrift: Times New Roman (OTM talablari), A4, chekkalar 2,5 / 1,5 / 2 / 2 sm.
"""
from __future__ import annotations

import datetime

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

KOK = RGBColor(0x1F, 0x3A, 0x5F)        # sarlavhalar
KOK_HEX = "1F3A5F"
OCH_HEX = "EEF3F8"                      # jadval navbatma-navbat fon
KULRANG = RGBColor(0x59, 0x59, 0x59)
YASHIL = RGBColor(0x1E, 0x7B, 0x34)
TOQ_SARIQ = RGBColor(0x9A, 0x6A, 0x00)

_AL = {"c": WD_ALIGN_PARAGRAPH.CENTER, "l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT,
       "j": WD_ALIGN_PARAGRAPH.JUSTIFY}
_BOT = {"uz": "Academic Bot · Ilmiy izlanuvchi", "ru": "Academic Bot · Научному соискателю",
        "en": "Academic Bot · Researchers"}
_BET = {"uz": "bet", "ru": "стр.", "en": "page"}


def _shd(el, color):
    pr = el.get_or_add_tcPr() if hasattr(el, "get_or_add_tcPr") else el.get_or_add_pPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear")
    s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), color)
    pr.append(s)


def _border_bottom(p, color=KOK_HEX, sz=8):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    e = OxmlElement("w:bottom")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(sz))
    e.set(qn("w:space"), "1")
    e.set(qn("w:color"), color)
    b.append(e)
    pPr.append(b)


def _field(run, instr):
    for t, v in (("begin", None), (None, instr), ("end", None)):
        if t:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), t)
        else:
            e = OxmlElement("w:instrText")
            e.set(qn("xml:space"), "preserve")
            e.text = v
        run._r.append(e)


def yangi(til: str = "uz", landscape: bool = False, kolontitul: bool = True) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = (Cm(29.7), Cm(21)) if landscape else (Cm(21), Cm(29.7))
    sec.left_margin, sec.right_margin = Cm(2.5), Cm(1.5)
    sec.top_margin, sec.bottom_margin = Cm(2), Cm(2)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.element.rPr.rFonts.set(qn("w:cs"), "Times New Roman")
    st.paragraph_format.space_before = st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 1.15
    # yuqori kolontitul (rasmiy hujjatlarda — yo'q) va sahifa raqami
    if kolontitul:
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = hp.add_run(f"{_BOT.get(til, _BOT['uz'])} · {datetime.date.today().strftime('%d.%m.%Y')}")
        r.font.size, r.font.color.rgb, r.italic = Pt(9), KULRANG, True
        _border_bottom(hp, "BFBFBF", 4)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run()
    r.font.size = Pt(9)
    _field(r, "PAGE")
    r2 = fp.add_run(f" {_BET.get(til, 'bet')}")
    r2.font.size, r2.font.color.rgb = Pt(9), KULRANG
    return doc


def p(doc, text="", *, b=False, i=False, size=12, color=None, al="j", sb=0, sa=3, ls=1.15, left=0.0, hang=0.0,
      ind=0.0, keep=False):
    par = doc.add_paragraph()
    f = par.paragraph_format
    f.space_before, f.space_after, f.line_spacing = Pt(sb), Pt(sa), ls
    f.left_indent = Cm(left)
    f.first_line_indent = Cm(-hang) if hang else Cm(ind)
    f.keep_with_next = keep
    par.alignment = _AL[al]
    if text:
        run(par, text, b=b, i=i, size=size, color=color)
    return par


def run(par, text, *, b=False, i=False, size=12, color=None):
    r = par.add_run(text)
    r.bold, r.italic = b, i
    r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color
    return r


def sarlavha(doc, title, subtitle="", izoh=""):
    p(doc, title.upper(), b=True, size=16, color=KOK, al="c", sa=2, ls=1.0)
    if subtitle:
        p(doc, subtitle, b=True, size=13, al="c", sa=2, ls=1.1)
    if izoh:
        p(doc, izoh, i=True, size=10.5, color=KULRANG, al="c", sa=6, ls=1.0)
    line = p(doc, "", sa=8)
    _border_bottom(line, KOK_HEX, 12)


def parametrlar(doc, rows, widths=(5.0, 12.0), rasmiy=False):
    """[(nom, qiymat)] — kulrang fonli chap ustun."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _grid(t, widths)
    for k, v in rows:
        c = t.add_row().cells
        _yoz(c[0], k, b=True, size=11 if rasmiy else 10.5, w=widths[0])
        _shd(c[0]._tc, "F2F2F2" if rasmiy else OCH_HEX)
        _yoz(c[1], v, size=11 if rasmiy else 10.5, w=widths[1])
    p(doc, "", sa=4)
    return t


def bolim(doc, text, sb=12, rasmiy=False):
    if rasmiy:
        return p(doc, text, b=True, size=12, al="c", sb=sb, sa=6, keep=True, ls=1.0)
    par = p(doc, text, b=True, size=13, color=KOK, al="l", sb=sb, sa=6, keep=True, ls=1.0)
    _border_bottom(par, "8EA9C8", 6)
    return par


def izoh(doc, text, sb=2, sa=6):
    return p(doc, text, i=True, size=10.5, color=KULRANG, sb=sb, sa=sa, ls=1.05)


def _grid(t, widths):
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w * 567)))


def _yoz(cell, text, *, b=False, i=False, size=10.5, al="l", color=None, w=None):
    if w is not None:
        cell.width = Cm(w)
    par = cell.paragraphs[0]
    par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(1)
    par.paragraph_format.line_spacing = 1.0
    par.alignment = _AL[al]
    lines = str(text or "").split("\n")
    for k, ln in enumerate(lines):
        r = par.add_run(ln)
        r.bold, r.italic = b, i
        r.font.size = Pt(size)
        if color is not None:
            r.font.color.rgb = color
        if k < len(lines) - 1:
            r.add_break()


def jadval(doc, head, rows, widths, size=10, center_cols=(0,), rasmiy=False):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _grid(t, widths)
    for k, h in enumerate(head):
        c = t.rows[0].cells[k]
        _yoz(c, h, b=True, size=size, al="c", color=None if rasmiy else RGBColor(0xFF, 0xFF, 0xFF), w=widths[k])
        _shd(c._tc, "D9D9D9" if rasmiy else KOK_HEX)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)
    for n, row in enumerate(rows):
        cs = t.add_row().cells
        for k, v in enumerate(row):
            _yoz(cs[k], v, size=size, al="c" if k in center_cols else "l", w=widths[k])
            if n % 2 == 1 and not rasmiy:
                _shd(cs[k]._tc, OCH_HEX)
    p(doc, "", sa=4)
    return t


def karta(doc, n, sarlavha_, qator, belgilar, annot="", havola="", gost="", L=None):
    """Bitta adabiyot/jurnal kartochkasi."""
    par = p(doc, "", sb=6, sa=1, keep=True, ls=1.1, left=0.9, hang=0.9, al="l")
    run(par, f"{n}.\t", b=True, size=12, color=KOK)
    run(par, sarlavha_, b=True, size=12)
    if qator:
        p(doc, qator, size=10.5, left=0.9, sa=1, ls=1.05, keep=True, al="l")
    if belgilar:
        bp = p(doc, "", left=0.9, sa=1, ls=1.05, keep=bool(annot or havola or gost), al="l")
        for k, (t, rang) in enumerate(belgilar):
            run(bp, ("  ·  " if k else "") + t, size=10, color=rang or KULRANG, b=rang is not None)
    if annot:
        p(doc, annot, i=True, size=10, color=KULRANG, left=0.9, sa=1, ls=1.05, al="j")
    if havola:
        p(doc, havola, size=10, color=KOK, left=0.9, sa=1, ls=1.0, al="l")
    if gost:
        gp = p(doc, "", left=0.9, sa=4, ls=1.0, al="j")
        run(gp, (L or "GOST") + ": ", b=True, size=9.5, color=KULRANG)
        run(gp, gost, size=9.5, color=KULRANG)


def royxat(doc, items, size=12):
    """Raqamlangan bibliografik ro'yxat (osilgan chekinish bilan)."""
    for n, x in enumerate(items, 1):
        par = p(doc, "", left=1.0, hang=1.0, sa=3, ls=1.15, al="j")
        run(par, f"{n}.\t", size=size)
        run(par, x, size=size)


def sahifa_uzish(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def imzo_qatori(doc, text):
    return p(doc, text, size=12, sb=10, al="l")


def _chegarasiz(t):
    tblPr = t._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tblPr.append(b)


def imzolar(doc, rows, widths=(6.0, 4.5, 6.0)):
    """Imzo bloki: [(lavozim, ism)] → «Lavozim | ________ (imzo) | F.I.Sh.» — chegarasiz jadval."""
    t = doc.add_table(rows=0, cols=3)
    t.autofit = False
    _grid(t, widths)
    _chegarasiz(t)
    for lav, ism in rows:
        c = t.add_row().cells
        _yoz(c[0], lav, b=True, size=12, w=widths[0])
        _yoz(c[1], "______________", size=12, al="c", w=widths[1])
        _yoz(c[2], ism or "______________________", size=12, w=widths[2])
        for cc in c:
            cc.paragraphs[0].paragraph_format.space_before = Pt(10)
    return t


def tasdiq_blok(doc, qatorlar, left=10.0):
    for k, t in enumerate(qatorlar):
        p(doc, t, b=(k == 0), size=12, al="l", left=left, sa=2, ls=1.1)

"""
ragbat_dizayn.py — Tashakkurnoma, faxriy yorliq, maqtov yorlig'i, sertifikat va diplomlar dizayni (PDF, vektor)
═══════════════════════════════════════════════════════════════════════════════════════════════
• 8 ta dizayn: rasmiy (oltin, milliy naqsh, akademik, zamonaviy, premium, tabiat) va bolalar uchun
  (quvnoq, kosmos). Hammasi kod bilan chiziladi — tashqi rasm kerak emas, chop etishda sifat yo'qolmaydi.
• A4 albom (landscape). Bitta faylda bir nechta varaq (guruh uchun) — har biri alohida ism bilan.
• Ism o'rniga bo'sh chiziq (qo'lda yoziladigan shablon) ham mumkin.
• 3 til: o'zbek (lotin), rus, ingliz. Shriftda harf bo'lmasa (masalan, o'zbek kirill ҳ) — avtomatik
  zaxira shriftga o'tiladi, matn hech qachon «kvadratcha» bo'lib chiqmaydi.
• Shriftlar: assets/fonts (SIL Open Font License) — Forum, Old Standard, Marck Script, Lobster, PT Sans.
"""
from __future__ import annotations

import math
import os
import random
import tempfile
import threading

from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "fonts")
W, H = landscape(A4)            # 841.9 x 595.3 pt

_SHRIFTLAR = {
    "Forum": "Forum-Regular.ttf",
    "OldB": "OldStandard-Bold.ttf",
    "OldI": "OldStandard-Italic.ttf",
    "Marck": "MarckScript-Regular.ttf",
    "Lobster": "Lobster-Regular.ttf",
    "Sans": "PT_Sans-Web-Regular.ttf",
    "SansB": "PT_Sans-Web-Bold.ttf",
    "Serif": "LiberationSerif-Regular.ttf",
    "SerifB": "LiberationSerif-Bold.ttf",
    "SerifI": "LiberationSerif-Italic.ttf",
}
_ZAXIRA = {"Marck": "Forum", "Lobster": "SansB", "OldB": "SerifB", "OldI": "SerifI", "Forum": "Serif",
           "Sans": "Serif", "SansB": "SerifB"}
_REG_LOCK = threading.Lock()
_REG = False


def _royxat():
    global _REG
    with _REG_LOCK:
        if _REG:
            return
        for nom, fayl in _SHRIFTLAR.items():
            p = os.path.join(_DIR, fayl)
            if os.path.exists(p):
                pdfmetrics.registerFont(TTFont(nom, p))
        _REG = True


def _bor(font: str) -> bool:
    try:
        pdfmetrics.getFont(font)
        return True
    except Exception:
        return False


def _shrift(font: str, matn: str) -> str:
    """Matndagi barcha harflar shriftda bo'lmasa — zaxira shrift (o'zbek kirill, maxsus belgilar)."""
    f = font
    for _ in range(4):
        if _bor(f):
            g = pdfmetrics.getFont(f).face.charToGlyph
            if all((ord(ch) in g) or ch in " \n\t" for ch in matn):
                return f
        f = _ZAXIRA.get(f, "Serif")
    return "Serif" if _bor("Serif") else "Helvetica"


def _kenglik(matn, font, size, spc=0.0):
    return pdfmetrics.stringWidth(matn, font, size) + spc * max(0, len(matn) - 1)


def _sigdir(matn, font, maks_w, boshl, eng_kam, spc=0.0):
    s = boshl
    while s > eng_kam and _kenglik(matn, font, s, spc * s / boshl) > maks_w:
        s -= 0.5
    return s


def _markaz(c, x, y, matn, font, size, rang, spc=0.0):
    font = _shrift(font, matn)
    c.setFillColor(rang)
    if spc:
        w = _kenglik(matn, font, size, spc)
        t = c.beginText(x - w / 2, y)
        t.setFont(font, size)
        t.setCharSpace(spc)
        t.textLine(matn)
        c.drawText(t)
    else:
        c.setFont(font, size)
        c.drawCentredString(x, y, matn)


def _qatorlar(matn, font, size, maks_w):
    sozlar, qator, out = matn.split(), "", []
    for s in sozlar:
        sinov = (qator + " " + s).strip()
        if _kenglik(sinov, font, size) <= maks_w:
            qator = sinov
        else:
            if qator:
                out.append(qator)
            qator = s
    if qator:
        out.append(qator)
    return out


# ═══════════════════════════════════════════════════════════════════
# MATNLAR
# ═══════════════════════════════════════════════════════════════════
TURLAR = ("tashakkurnoma", "faxriy", "maqtov", "sertifikat", "diplom", "minnatdorchilik")
SARLAVHA = {
    "uz": {"tashakkurnoma": "TASHAKKURNOMA", "faxriy": "FAXRIY YORLIQ", "maqtov": "MAQTOV YORLIG‘I",
           "sertifikat": "SERTIFIKAT", "diplom": "DIPLOM", "minnatdorchilik": "MINNATDORCHILIK XATI"},
    "ru": {"tashakkurnoma": "БЛАГОДАРНОСТЬ", "faxriy": "ПОЧЁТНАЯ ГРАМОТА", "maqtov": "ПОХВАЛЬНЫЙ ЛИСТ",
           "sertifikat": "СЕРТИФИКАТ", "diplom": "ДИПЛОМ", "minnatdorchilik": "БЛАГОДАРСТВЕННОЕ ПИСЬМО"},
    "en": {"tashakkurnoma": "CERTIFICATE OF APPRECIATION", "faxriy": "CERTIFICATE OF HONOR",
           "maqtov": "CERTIFICATE OF MERIT", "sertifikat": "CERTIFICATE", "diplom": "DIPLOMA",
           "minnatdorchilik": "LETTER OF GRATITUDE"},
}
OSTKI = {
    "uz": {"tashakkurnoma": "taqdim etiladi", "faxriy": "bilan taqdirlanadi", "maqtov": "bilan taqdirlanadi",
           "sertifikat": "beriladi", "diplom": "bilan taqdirlanadi", "minnatdorchilik": "taqdim etiladi"},
    "ru": {"tashakkurnoma": "объявляется", "faxriy": "награждается", "maqtov": "награждается",
           "sertifikat": "выдан", "diplom": "награждается", "minnatdorchilik": "вручается"},
    "en": {"tashakkurnoma": "is proudly presented to", "faxriy": "is awarded to", "maqtov": "is awarded to",
           "sertifikat": "is awarded to", "diplom": "is awarded to", "minnatdorchilik": "is presented to"},
}
DARAJA = {"uz": "{d}-DARAJALI ", "ru": " {r} СТЕПЕНИ", "en": "{o} PLACE "}
_RIM = {1: "I", 2: "II", 3: "III"}
_ORD = {1: "FIRST", 2: "SECOND", 3: "THIRD"}
MO = {"uz": "M.O‘.", "ru": "М.П.", "en": "Seal"}
SANA_L = {"uz": "Sana", "ru": "Дата", "en": "Date"}


def sarlavha(tur: str, til: str, daraja: int | None = None) -> str:
    s = SARLAVHA[til][tur]
    if tur == "diplom" and daraja in (1, 2, 3):
        if til == "uz":
            return f"{_RIM[daraja]} DARAJALI DIPLOM"
        if til == "ru":
            return f"ДИПЛОМ {_RIM[daraja]} СТЕПЕНИ"
        return f"{_ORD[daraja]} PLACE DIPLOMA"
    return s


# ═══════════════════════════════════════════════════════════════════
# BEZAK ELEMENTLARI
# ═══════════════════════════════════════════════════════════════════
def _yulduz(c, x, y, r, n=5, ichki=0.45, rang=None, aylan=90, chiziq=None):
    p = c.beginPath()
    for i in range(2 * n):
        rr = r if i % 2 == 0 else r * ichki
        a = math.radians(aylan + i * 180 / n)
        (p.moveTo if i == 0 else p.lineTo)(x + rr * math.cos(a), y + rr * math.sin(a))
    p.close()
    if rang is not None:
        c.setFillColor(rang)
    if chiziq is not None:
        c.setStrokeColor(chiziq)
        c.drawPath(p, fill=1, stroke=1)
    else:
        c.drawPath(p, fill=1, stroke=0)


def _medal(c, x, y, r, asosiy, toq, lenta=True, belgi="star"):
    """Oltin medal va lentalar."""
    if lenta:
        for dx, rang in ((-0.45, toq), (0.45, asosiy)):
            p = c.beginPath()
            x0 = x + dx * r
            p.moveTo(x0 - r * 0.28, y - r * 0.2)
            p.lineTo(x0 + r * 0.28, y - r * 0.2)
            p.lineTo(x0 + r * 0.28 + dx * r * 0.3, y - r * 1.75)
            p.lineTo(x0 + dx * r * 0.3, y - r * 1.5)
            p.lineTo(x0 - r * 0.28 + dx * r * 0.3, y - r * 1.75)
            p.close()
            c.setFillColor(rang)
            c.drawPath(p, fill=1, stroke=0)
    # tishli chekka
    p = c.beginPath()
    n = 28
    for i in range(2 * n):
        rr = r if i % 2 == 0 else r * 0.9
        a = math.radians(i * 180 / n)
        (p.moveTo if i == 0 else p.lineTo)(x + rr * math.cos(a), y + rr * math.sin(a))
    p.close()
    c.setFillColor(asosiy)
    c.drawPath(p, fill=1, stroke=0)
    c.setFillColor(toq)
    c.circle(x, y, r * 0.76, fill=1, stroke=0)
    c.setFillColor(asosiy)
    c.circle(x, y, r * 0.68, fill=1, stroke=0)
    if belgi == "star":
        _yulduz(c, x, y, r * 0.48, rang=toq)


def _burchak_naqsh(c, x, y, sx, sy, rang, olcham=46):
    """Klassik burchak bezagi (jingalak chiziqlar). sx, sy — yo'nalish (+1/-1)."""
    c.saveState()
    c.translate(x, y)
    c.scale(sx, sy)
    c.setStrokeColor(rang)
    c.setFillColor(rang)
    c.setLineWidth(1.4)
    for k in (1.0, 0.72, 0.46):
        c.arc(-olcham * k, -olcham * k, olcham * k, olcham * k, 0, 90)
    p = c.beginPath()
    p.moveTo(0, olcham * 1.15)
    p.curveTo(olcham * 0.2, olcham * 0.75, olcham * 0.55, olcham * 0.55, olcham * 1.15, 0)
    c.drawPath(p, stroke=1, fill=0)
    # romb
    p = c.beginPath()
    d = 7
    p.moveTo(olcham * 0.32, olcham * 0.32 + d)
    p.lineTo(olcham * 0.32 + d, olcham * 0.32)
    p.lineTo(olcham * 0.32, olcham * 0.32 - d)
    p.lineTo(olcham * 0.32 - d, olcham * 0.32)
    p.close()
    c.drawPath(p, fill=1, stroke=0)
    for t in (0.62, 0.9):
        c.circle(olcham * t, 3, 2.2, fill=1, stroke=0)
        c.circle(3, olcham * t, 2.2, fill=1, stroke=0)
    c.restoreState()


def _girih(c, x, y, r, rang1, rang2):
    """8 qirrali yulduz (ikki kvadrat) — milliy naqsh motivi."""
    for k, rang in ((0, rang1), (45, rang1)):
        c.saveState()
        c.translate(x, y)
        c.rotate(k)
        c.setFillColor(rang)
        c.rect(-r * 0.71, -r * 0.71, r * 1.42, r * 1.42, fill=1, stroke=0)
        c.restoreState()
    _yulduz(c, x, y, r * 0.62, n=8, ichki=0.62, rang=rang2, aylan=22.5)
    c.setFillColor(rang1)
    c.circle(x, y, r * 0.22, fill=1, stroke=0)


def _barg(c, x, y, uzun, en, burchak, rang, tomir=None):
    c.saveState()
    c.translate(x, y)
    c.rotate(burchak)
    p = c.beginPath()
    p.moveTo(0, 0)
    p.curveTo(uzun * 0.3, en, uzun * 0.7, en, uzun, 0)
    p.curveTo(uzun * 0.7, -en, uzun * 0.3, -en, 0, 0)
    p.close()
    c.setFillColor(rang)
    c.drawPath(p, fill=1, stroke=0)
    if tomir is not None:
        c.setStrokeColor(tomir)
        c.setLineWidth(0.6)
        c.line(uzun * 0.05, 0, uzun * 0.92, 0)
    c.restoreState()


def _novda(c, x, y, burchak, n, rang1, rang2, uz=1.0):
    """Barglar novdasi (tabiat dizayni)."""
    c.saveState()
    c.translate(x, y)
    c.rotate(burchak)
    c.setStrokeColor(rang2)
    c.setLineWidth(1.6)
    L = 150 * uz
    p = c.beginPath()
    p.moveTo(0, 0)
    p.curveTo(L * 0.3, 12, L * 0.7, 8, L, 0)
    c.drawPath(p, stroke=1, fill=0)
    for i in range(n):
        t = (i + 1) / (n + 1)
        bx = L * t
        _barg(c, bx, 6, 34 * uz * (1.1 - t * 0.4), 9 * uz, 55, rang1 if i % 2 else rang2, white)
        _barg(c, bx, 4, 30 * uz * (1.1 - t * 0.4), 8 * uz, -50, rang2 if i % 2 else rang1, white)
    c.restoreState()


def _dafna(c, x, y, r, rang):
    """Dafna gulchambari (akademik)."""
    for tomon in (-1, 1):
        for i in range(9):
            a = math.radians(-100 + i * 22) if tomon == 1 else math.radians(280 - i * 22)
            bx, by = x + r * math.cos(a), y + r * math.sin(a)
            burchak = math.degrees(a) + (90 if tomon == 1 else -90) + 25 * tomon
            _barg(c, bx, by, 17, 5, burchak, rang)


# ═══════════════════════════════════════════════════════════════════
# DIZAYNLAR
# ═══════════════════════════════════════════════════════════════════
def _fon(c, rang):
    c.setFillColor(rang)
    c.rect(0, 0, W, H, fill=1, stroke=0)


def _gradient(c, yuqori, past, n=160):
    a, b = HexColor(yuqori), HexColor(past)
    for i in range(n):
        t = i / (n - 1)
        c.setFillColor(Color(a.red + (b.red - a.red) * t, a.green + (b.green - a.green) * t,
                             a.blue + (b.blue - a.blue) * t))
        c.rect(0, H - (i + 1) * H / n - 1, W, H / n + 2, fill=1, stroke=0)


def _d_oltin(c, d):
    _fon(c, HexColor("#FBF7EC"))
    g = HexColor("#B08D3C")
    c.setStrokeColor(g)
    c.setLineWidth(7)
    c.rect(18, 18, W - 36, H - 36)
    c.setLineWidth(1.2)
    c.rect(30, 30, W - 60, H - 60)
    c.rect(34, 34, W - 68, H - 68)
    for x, y, sx, sy in ((40, 40, 1, 1), (W - 40, 40, -1, 1), (40, H - 40, 1, -1), (W - 40, H - 40, -1, -1)):
        _burchak_naqsh(c, x, y, sx, sy, g)
    # sarlavha ostidagi bezak
    c.setLineWidth(1)
    c.line(W / 2 - 150, H - 162, W / 2 - 14, H - 162)
    c.line(W / 2 + 14, H - 162, W / 2 + 150, H - 162)
    _yulduz(c, W / 2, H - 162, 7, n=4, ichki=0.35, rang=g, aylan=90)


def _d_milliy(c, d):
    _fon(c, HexColor("#FFFDF7"))
    fir, toq, olt = HexColor("#0E7C86"), HexColor("#0B5563"), HexColor("#C9A227")
    b = 30
    c.setFillColor(fir)
    for x, y, w, h in ((14, 14, W - 28, b), (14, H - 14 - b, W - 28, b), (14, 14, b, H - 28), (W - 14 - b, 14, b, H - 28)):
        c.rect(x, y, w, h, fill=1, stroke=0)
    # yulduzlar zanjiri
    qadam = 34
    n = int((W - 28) // qadam)
    for i in range(n + 1):
        x = 14 + b / 2 + i * (W - 28 - b) / n
        for y in (14 + b / 2, H - 14 - b / 2):
            _yulduz(c, x, y, 9, n=8, ichki=0.55, rang=white, aylan=22.5)
            c.setFillColor(olt)
            c.circle(x, y, 2.2, fill=1, stroke=0)
    m = int((H - 28) // qadam)
    for i in range(1, m):
        y = 14 + b / 2 + i * (H - 28 - b) / m
        for x in (14 + b / 2, W - 14 - b / 2):
            _yulduz(c, x, y, 9, n=8, ichki=0.55, rang=white, aylan=22.5)
            c.setFillColor(olt)
            c.circle(x, y, 2.2, fill=1, stroke=0)
    c.setStrokeColor(olt)
    c.setLineWidth(2)
    c.rect(14 + b + 6, 14 + b + 6, W - 2 * (14 + b + 6), H - 2 * (14 + b + 6))
    for x, y in ((14 + b / 2, 14 + b / 2), (W - 14 - b / 2, 14 + b / 2), (14 + b / 2, H - 14 - b / 2),
                 (W - 14 - b / 2, H - 14 - b / 2)):
        _girih(c, x, y, 22, olt, toq)
    _girih(c, W / 2, H - 14 - b - 6, 17, olt, fir)


def _d_akademik(c, d):
    _fon(c, white)
    nav, olt = HexColor("#1B2A4A"), HexColor("#C5A253")
    c.setFillColor(nav)
    c.rect(0, 0, 92, H, fill=1, stroke=0)
    c.setStrokeColor(olt)
    c.setLineWidth(1.5)
    c.line(80, 0, 80, H)
    c.line(86, 0, 86, H)
    _dafna(c, 46, H / 2, 30, olt)
    _yulduz(c, 46, H / 2, 13, rang=olt)
    for y in (90, H - 90):
        _yulduz(c, 46, y, 6, n=4, ichki=0.35, rang=olt)
    c.setStrokeColor(nav)
    c.setLineWidth(1)
    c.rect(110, 22, W - 132, H - 44)
    c.setStrokeColor(olt)
    c.rect(116, 28, W - 144, H - 56)


def _d_zamonaviy(c, d):
    _fon(c, white)
    teal, ora, nav = HexColor("#0FA3B1"), HexColor("#F7A072"), HexColor("#2D3142")

    def uchb(pts, rang):
        p = c.beginPath()
        p.moveTo(*pts[0])
        for q in pts[1:]:
            p.lineTo(*q)
        p.close()
        c.setFillColor(rang)
        c.drawPath(p, fill=1, stroke=0)
    uchb([(0, H), (230, H), (0, H - 170)], teal)
    uchb([(0, H), (120, H), (0, H - 250)], nav)
    uchb([(W, 0), (W - 170, 0), (W, 120)], teal)
    uchb([(W, 0), (W - 80, 0), (W, 210)], ora)
    c.setFillColor(ora)
    c.rect(W / 2 - 40, H - 172, 80, 4, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#D9DCE3"))
    c.setLineWidth(1)
    c.rect(26, 26, W - 52, H - 52)


def _d_premium(c, d):
    _fon(c, HexColor("#111827"))
    g = HexColor("#D4AF37")
    c.setStrokeColor(g)
    c.setLineWidth(2.2)
    c.rect(22, 22, W - 44, H - 44)
    c.setLineWidth(0.8)
    c.rect(30, 30, W - 60, H - 60)
    L = 70
    c.setLineWidth(4)
    for x, y, sx, sy in ((40, 40, 1, 1), (W - 40, 40, -1, 1), (40, H - 40, 1, -1), (W - 40, H - 40, -1, -1)):
        c.line(x, y, x + sx * L, y)
        c.line(x, y, x, y + sy * L)
    c.setLineWidth(0.8)
    c.line(W / 2 - 160, H - 162, W / 2 + 160, H - 162)
    _yulduz(c, W / 2, H - 162, 6, n=4, ichki=0.35, rang=g)


def _d_bolalar(c, d):
    _fon(c, HexColor("#FFF8E1"))
    rnd = random.Random(7)
    ranglar = [HexColor(x) for x in ("#E8505B", "#F9C74F", "#43AA8B", "#577590", "#F3722C", "#9B5DE5")]
    # bayroqchalar gulchambari (yuqorida)
    c.setStrokeColor(HexColor("#8D99AE"))
    c.setLineWidth(1)
    n = 22
    nuq = []
    for i in range(n + 1):
        t = i / n
        x = 40 + t * (W - 80)
        y = H - 32 - 9 * math.sin(math.pi * t)
        nuq.append((x, y))
    for (x1, y1), (x2, y2) in zip(nuq, nuq[1:]):
        c.line(x1, y1, x2, y2)
    for i, ((x1, y1), (x2, y2)) in enumerate(zip(nuq, nuq[1:])):
        p = c.beginPath()
        p.moveTo(x1 + 2, y1)
        p.lineTo(x2 - 2, y2)
        p.lineTo((x1 + x2) / 2, (y1 + y2) / 2 - 17)
        p.close()
        c.setFillColor(ranglar[i % len(ranglar)])
        c.drawPath(p, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#FFB703"))
    c.setLineWidth(6)
    c.setDash(14, 8)
    c.roundRect(22, 22, W - 44, H - 44, 28)
    c.setDash()
    for _ in range(46):
        x, y = rnd.uniform(40, W - 40), rnd.uniform(40, H - 40)
        if 125 < x < W - 125 and 60 < y < H - 70:
            continue
        if y > H - 75 and 60 < x < W - 60:
            continue
        rang = rnd.choice(ranglar)
        if rnd.random() < 0.45:
            _yulduz(c, x, y, rnd.uniform(6, 13), rang=rang, aylan=rnd.uniform(0, 72))
        else:
            c.setFillColor(rang)
            c.circle(x, y, rnd.uniform(3, 7), fill=1, stroke=0)
    # shariklar
    for x, y, rang in ((70, H - 150, "#E8505B"), (105, H - 120, "#43AA8B"), (W - 75, H - 140, "#F9C74F"),
                       (W - 110, H - 175, "#9B5DE5")):
        c.setStrokeColor(HexColor("#888888"))
        c.setLineWidth(0.8)
        c.line(x, y - 26, x + 6, y - 80)
        c.setFillColor(HexColor(rang))
        c.ellipse(x - 19, y - 26, x + 19, y + 24, fill=1, stroke=0)
        c.setFillColor(white)
        c.ellipse(x - 10, y + 4, x - 3, y + 14, fill=1, stroke=0)


def _d_kosmos(c, d):
    _gradient(c, "#0B1D51", "#3A0CA3")
    rnd = random.Random(3)
    for _ in range(170):
        x, y = rnd.uniform(0, W), rnd.uniform(0, H)
        c.setFillColor(Color(1, 1, 1, alpha=rnd.uniform(0.35, 1)))
        c.circle(x, y, rnd.uniform(0.4, 1.7), fill=1, stroke=0)
    for x, y, r in ((70, 160, 9), (W - 70, 170, 7), (170, H - 60, 6), (W / 2 + 270, 60, 5)):
        _yulduz(c, x, y, r, n=4, ichki=0.3, rang=HexColor("#FFD166"))
    # sayyora
    c.setFillColor(HexColor("#F3722C"))
    c.circle(W - 95, H - 110, 44, fill=1, stroke=0)
    c.setFillColor(HexColor("#F8961E"))
    c.circle(W - 108, H - 97, 16, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#FFD166"))
    c.setLineWidth(3)
    c.saveState()
    c.translate(W - 95, H - 110)
    c.rotate(-18)
    c.ellipse(-74, -13, 74, 13, fill=0, stroke=1)
    c.restoreState()
    # oy
    c.setFillColor(HexColor("#EDEDED"))
    c.circle(88, H - 100, 28, fill=1, stroke=0)
    c.setFillColor(HexColor("#C9C9C9"))
    for dx, dy, r in ((-8, 6, 6), (9, -7, 4), (6, 12, 3)):
        c.circle(88 + dx, H - 100 + dy, r, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#FFD166"))
    c.setLineWidth(1.2)
    c.roundRect(24, 24, W - 48, H - 48, 18)


def _d_tabiat(c, d):
    _fon(c, HexColor("#F4FAF2"))
    y1, y2, y3 = HexColor("#52B788"), HexColor("#2D6A4F"), HexColor("#95D5B2")
    c.setStrokeColor(y3)
    c.setLineWidth(1.4)
    c.roundRect(26, 26, W - 52, H - 52, 16)
    _novda(c, 20, H - 30, -22, 6, y1, y2)
    _novda(c, 30, H - 20, -62, 5, y3, y1, 0.8)
    _novda(c, W - 22, 26, 112, 6, y1, y2, 0.95)
    _novda(c, W - 18, 40, 128, 4, y3, y1, 0.7)
    for x, y, r in ((W - 60, H - 60, 5), (W - 85, H - 48, 3.5), (60, 60, 5), (82, 48, 3.5), (W - 70, 45, 4)):
        c.setFillColor(HexColor("#FFB4A2"))
        c.circle(x, y, r, fill=1, stroke=0)


# kalit: (nomlar, chizuvchi, ranglar va shriftlar)
DIZAYN = {
    "oltin": dict(nom={"uz": "Rasmiy oltin", "ru": "Классика (золото)", "en": "Classic gold"}, f=_d_oltin,
                  sar_f="OldB", sar_r="#1F2A44", ism_f="Marck", ism_r="#8A6A1F", mat_f="Forum", mat_r="#2B2B2B",
                  qo_r="#6B5B3E", chiz="#B08D3C", medal=("#C9A227", "#8A6A1F"), rasmiy=True),
    "milliy": dict(nom={"uz": "Milliy naqsh", "ru": "Национальный орнамент", "en": "National ornament"},
                   f=_d_milliy, sar_f="Forum", sar_r="#0B5563", ism_f="Marck", ism_r="#0E7C86", mat_f="Forum",
                   mat_r="#24343A", qo_r="#0B5563", chiz="#C9A227", medal=("#C9A227", "#0B5563"), rasmiy=True),
    "akademik": dict(nom={"uz": "Akademik", "ru": "Академический", "en": "Academic"}, f=_d_akademik,
                     sar_f="OldB", sar_r="#1B2A4A", ism_f="Marck", ism_r="#1B2A4A", mat_f="Forum", mat_r="#2B2B2B",
                     qo_r="#5B6577", chiz="#C5A253", medal=("#C5A253", "#1B2A4A"), rasmiy=True, chap=92),
    "zamonaviy": dict(nom={"uz": "Zamonaviy", "ru": "Современный", "en": "Modern"}, f=_d_zamonaviy,
                      sar_f="SansB", sar_r="#2D3142", ism_f="SansB", ism_r="#0FA3B1", mat_f="Sans", mat_r="#2D3142",
                      qo_r="#6C7080", chiz="#2D3142", medal=("#F7A072", "#2D3142"), rasmiy=True, spc=6),
    "premium": dict(nom={"uz": "Premium (qora-oltin)", "ru": "Премиум (чёрный и золото)", "en": "Premium black & gold"},
                    f=_d_premium, sar_f="OldB", sar_r="#D4AF37", ism_f="Marck", ism_r="#F3E5AB", mat_f="Sans",
                    mat_r="#E8E8E8", qo_r="#BFA76A", chiz="#D4AF37", medal=("#D4AF37", "#7A6320"), rasmiy=True),
    "tabiat": dict(nom={"uz": "Tabiat (yashil)", "ru": "Природа (зелёный)", "en": "Nature (green)"}, f=_d_tabiat,
                   sar_f="Forum", sar_r="#2D6A4F", ism_f="Marck", ism_r="#40916C", mat_f="Sans", mat_r="#2B3A33",
                   qo_r="#52796F", chiz="#52B788", medal=("#52B788", "#2D6A4F"), rasmiy=True),
    "bolalar": dict(nom={"uz": "Bolalar — quvnoq", "ru": "Детский — яркий", "en": "Kids — cheerful"}, f=_d_bolalar,
                    sar_f="Lobster", sar_r="#E8505B", ism_f="Lobster", ism_r="#3D5A80", mat_f="Sans", mat_r="#333333",
                    qo_r="#555555", chiz="#577590", medal=("#F9C74F", "#E8505B"), rasmiy=False, sar_maks=0.66),
    "kosmos": dict(nom={"uz": "Bolalar — kosmos", "ru": "Детский — космос", "en": "Kids — space"}, f=_d_kosmos,
                   sar_f="Lobster", sar_r="#FFD166", ism_f="Lobster", ism_r="#FFFFFF", mat_f="Sans", mat_r="#EDEDED",
                   qo_r="#CFCFE8", chiz="#FFD166", medal=("#FFD166", "#F3722C"), rasmiy=False, sar_maks=0.6),
}
TARTIB = ("oltin", "milliy", "akademik", "zamonaviy", "premium", "tabiat", "bolalar", "kosmos")


# ═══════════════════════════════════════════════════════════════════
# BITTA VARAQ
# ═══════════════════════════════════════════════════════════════════
def _varaq(c, d: dict, dz: str, til: str):
    t = DIZAYN[dz]
    t["f"](c, d)
    R = HexColor
    chap = t.get("chap", 0)
    cx = chap + (W - chap) / 2
    maks = (W - chap) * 0.72
    spc = t.get("spc", 3)
    tur = d.get("tur", "tashakkurnoma")
    medal = tur in ("faxriy", "diplom") or not t["rasmiy"]

    # 1. tashkilot (yuqorida)
    y = H - 82
    if d.get("tashkilot"):
        org = d["tashkilot"].upper() if len(d["tashkilot"]) < 70 else d["tashkilot"]
        s = _sigdir(org, _shrift("SansB", org), maks, 12.5, 8.5, 1.2)
        _markaz(c, cx, y, org, "SansB", s, R(t["qo_r"]), spc=1.2 * s / 12.5)
    if d.get("raqam"):
        f = _shrift("Sans", d["raqam"])
        c.setFont(f, 9)
        c.setFillColor(R(t["qo_r"]))
        c.drawRightString(W - 60, H - 58, d["raqam"])

    # 2. sarlavha
    sar = sarlavha(tur, til, d.get("daraja"))
    sf = _shrift(t["sar_f"], sar)
    bosh = 46 if t["sar_f"] != "Lobster" else 54
    ss = _sigdir(sar, sf, (W - chap) * t.get("sar_maks", 0.72), bosh, 22, spc)
    _markaz(c, cx, H - 140, sar, sf, ss, R(t["sar_r"]), spc=spc * ss / bosh)

    # 3. «taqdim etiladi»
    ost = OSTKI[til][tur]
    _markaz(c, cx, H - 192, ost, "OldI" if t["rasmiy"] and t["mat_f"] != "Sans" else "Sans", 15,
            R(t["qo_r"]), spc=0.8)

    # 4. ism (yoki bo'sh chiziq)
    ism = (d.get("ism") or "").strip()
    yi = H - 262
    if ism:
        f = _shrift(t["ism_f"], ism)
        si = _sigdir(ism, f, maks * 0.92, 46, 20)
        _markaz(c, cx, yi, ism, f, si, R(t["ism_r"]))
    c.setStrokeColor(R(t["chiz"]))
    c.setLineWidth(0.9)
    c.line(cx - maks * 0.36, yi - 16, cx + maks * 0.36, yi - 16)

    # 5. sabab matni
    mat = (d.get("matn") or "").strip()
    if mat:
        f = _shrift(t["mat_f"], mat)
        kenglik = maks * 0.9
        for s in (21, 20, 19, 18, 17, 16, 15, 14, 13, 12):
            qat = _qatorlar(mat, f, s, kenglik)
            if len(qat) <= 4 and (yi - 56 - (len(qat) - 1) * s * 1.38) > 150:
                break
        qat = qat[:5]
        yy = yi - 56
        for q in qat:
            _markaz(c, cx, yy, q, f, s, R(t["mat_r"]))
            yy -= s * 1.38

    # 6. pastki qism: sana (chap), medal (o'rta), imzo (o'ng)
    yp = 96
    xl, xr = chap + 150 if chap else 150, W - 165
    f = _shrift("Sans", d.get("sana", "") + (d.get("shahar") or ""))
    c.setFillColor(R(t["qo_r"]))
    if d.get("sana") or d.get("shahar"):
        c.setFont(f, 12)
        c.drawCentredString(xl, yp + 4, d.get("sana", ""))
        if d.get("shahar"):
            c.setFont(f, 11)
            c.drawCentredString(xl, yp - 12, d["shahar"])
        c.setStrokeColor(R(t["chiz"]))
        c.setLineWidth(0.7)
        c.line(xl - 70, yp - 2 + (0 if not d.get("shahar") else -12) - 6, xl + 70,
               yp - 2 + (0 if not d.get("shahar") else -12) - 6)
    lav, bism = (d.get("lavozim") or "").strip(), (d.get("beruvchi") or "").strip()
    if lav or bism:
        if lav:
            fl = _shrift("Sans", lav)
            sl = _sigdir(lav, fl, 230, 11.5, 8)
            c.setFont(fl, sl)
            c.setFillColor(R(t["qo_r"]))
            c.drawCentredString(xr, yp + 22, lav)
        c.setStrokeColor(R(t["chiz"]))
        c.setLineWidth(0.8)
        c.line(xr - 105, yp + 4, xr + 105, yp + 4)
        if bism:
            fb = _shrift("SansB", bism)
            sb = _sigdir(bism, fb, 220, 12.5, 8)
            c.setFont(fb, sb)
            c.setFillColor(R(t["sar_r"]) if dz not in ("kosmos", "premium") else R(t["mat_r"]))
            c.drawCentredString(xr, yp - 12, bism)
        if t["rasmiy"]:
            c.setFont(_shrift("Sans", MO[til]), 10)
            c.setFillColor(R(t["qo_r"]))
            c.drawRightString(xr - 115, yp + 1, MO[til])
    if medal:
        _medal(c, cx, yp + 8, 30, R(t["medal"][0]), R(t["medal"][1]))


def chiz(fayl: str, varaqlar: list[dict], dizayn: str, til: str) -> str:
    """varaqlar: [{tur, ism, matn, daraja, tashkilot, lavozim, beruvchi, sana, shahar, raqam}]"""
    _royxat()
    if dizayn not in DIZAYN:
        dizayn = "oltin"
    til = til if til in SARLAVHA else "uz"
    c = canvas.Canvas(fayl, pagesize=(W, H))
    c.setTitle(SARLAVHA[til].get(varaqlar[0].get("tur", "tashakkurnoma"), "Certificate") if varaqlar else "")
    c.setAuthor("Academic Bot")
    for d in varaqlar:
        _varaq(c, d, dizayn, til)
        c.showPage()
    c.save()
    return fayl


# ═══════════════════════════════════════════════════════════════════
# RASM (PNG/JPG) — Telegram'da ko'rish va ota-onalarga yuborish uchun
# ═══════════════════════════════════════════════════════════════════
def rasmlar(pdf: str, papka: str, zoom: float = 2.0, fmt: str = "jpg", nomlar: list | None = None) -> list:
    try:
        import pymupdf as fitz
    except Exception:
        import fitz
    out = []
    with fitz.open(pdf) as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            nom = (nomlar[i] if nomlar and i < len(nomlar) else f"{i + 1:02d}") + f".{fmt}"
            p = os.path.join(papka, nom)
            if fmt == "jpg":
                pix.save(p, jpg_quality=90)
            else:
                pix.save(p)
            out.append(p)
    return out


_KOLAJ: dict = {}
_KOLAJ_LOCK: dict = {}
_KOLAJ_LOCK0 = threading.Lock()
NAMUNA = {
    "uz": dict(tur="faxriy", ism="Toshmatova Malika", tashkilot="Toshkent shahar 110-maktab",
               matn="«Kasrlar» mavzusidagi darsda faol ishtiroki va topshiriqlarni a’lo darajada bajargani uchun.",
               lavozim="Matematika o‘qituvchisi", beruvchi="A. Karimov", sana="25.11.2026"),
    "ru": dict(tur="faxriy", ism="Тошматова Малика", tashkilot="Школа № 110 г. Ташкента",
               matn="За активное участие в уроке по теме «Дроби» и отличное выполнение заданий.",
               lavozim="Учитель математики", beruvchi="А. Каримов", sana="25.11.2026"),
    "en": dict(tur="faxriy", ism="Malika Toshmatova", tashkilot="School No. 110, Tashkent",
               matn="For active participation in the lesson on fractions and excellent completion of all tasks.",
               lavozim="Mathematics teacher", beruvchi="A. Karimov", sana="25.11.2026"),
}


_NAMUNA_MATN = {
    "tashakkurnoma": {"uz": "«Kasrlar» mavzusidagi darsda faol ishtiroki va tashabbuskorligi uchun.",
                      "ru": "За активное участие и инициативность на уроке по теме «Дроби».",
                      "en": "For active participation and initiative in the lesson on fractions."},
    "maqtov": {"uz": "Matematika fanini a’lo o‘zlashtirgani va namunali xulqi uchun.",
               "ru": "За отличные успехи в изучении математики и примерное поведение.",
               "en": "For excellent progress in mathematics and exemplary conduct."},
    "sertifikat": {"uz": "«Yosh matematiklar» to‘garagini muvaffaqiyatli tamomlagani uchun.",
                   "ru": "За успешное окончание кружка «Юные математики».",
                   "en": "For successfully completing the «Young Mathematicians» club."},
    "diplom": {"uz": "Maktab matematika olimpiadasida yuqori natija ko‘rsatgani uchun.",
               "ru": "За высокий результат в школьной олимпиаде по математике.",
               "en": "For an outstanding result in the school mathematics olympiad."},
    "minnatdorchilik": {"uz": "Farzandini tarbiyalashdagi g‘amxo‘rligi va maktab bilan hamkorligi uchun.",
                        "ru": "За заботу о воспитании ребёнка и сотрудничество со школой.",
                        "en": "For caring support of your child and partnership with the school."},
}


def kolaj(til: str, tur: str = "faxriy", daraja: int | None = None) -> str:
    """8 ta dizaynning namunaviy ko'rinishi bitta rasmda (raqamlangan) — foydalanuvchi tanlagan hujjat TURI bilan.
    Har (til, tur, daraja) uchun bir marta yaratiladi va keshlanadi."""
    til = til if til in NAMUNA else "uz"
    tur = tur if tur in TURLAR else "faxriy"
    import datetime as _dt
    bugun = _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=5))).strftime("%d.%m.%Y")   # Toshkent vaqti
    kalit = (til, tur, daraja if tur == "diplom" else None, bugun)
    with _KOLAJ_LOCK0:
        lock = _KOLAJ_LOCK.setdefault(kalit, threading.Lock())
    with lock:
        p = _KOLAJ.get(kalit)
        if p and os.path.exists(p):
            return p
        namuna = dict(NAMUNA[til], tur=tur, daraja=kalit[2], sana=bugun)
        if tur in _NAMUNA_MATN:
            namuna["matn"] = _NAMUNA_MATN[tur][til]
        from PIL import Image, ImageDraw, ImageFont
        tmp = tempfile.mkdtemp(prefix="ragbat_kolaj_")
        rasm_l = []
        for i, dz in enumerate(TARTIB):
            pdf = os.path.join(tmp, f"{dz}.pdf")
            chiz(pdf, [namuna], dz, til)
            rasm_l.append(rasmlar(pdf, tmp, zoom=0.62, fmt="png", nomlar=[dz])[0])
        ims = [Image.open(x).convert("RGB") for x in rasm_l]
        w, h = ims[0].size
        pad, lab = 14, 34
        kol = Image.new("RGB", (2 * w + 3 * pad, 4 * (h + lab) + 5 * pad), "white")
        dr = ImageDraw.Draw(kol)
        try:
            fnt = ImageFont.truetype(os.path.join(_DIR, "PT_Sans-Web-Bold.ttf"), 22)
        except Exception:
            fnt = ImageFont.load_default()
        for i, (im, dz) in enumerate(zip(ims, TARTIB)):
            x = pad + (i % 2) * (w + pad)
            y = pad + (i // 2) * (h + lab + pad)
            dr.text((x + 4, y + 2), f"{i + 1}. {DIZAYN[dz]['nom'][til]}", fill=(30, 30, 30), font=fnt)
            kol.paste(im, (x, y + lab))
            dr.rectangle([x - 1, y + lab - 1, x + w, y + lab + h], outline=(200, 200, 200))
        p = os.path.join(tmp, f"dizaynlar_{til}_{tur}.jpg")
        kol.save(p, quality=88)
        for x in rasm_l:
            try:
                os.remove(x)
            except Exception:
                pass
        _KOLAJ[kalit] = p
        return p

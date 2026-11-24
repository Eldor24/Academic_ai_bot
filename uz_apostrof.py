"""O'zbek lotin alifbosidagi o‘, g‘ va tutuq belgisi (’) ni hujjatlarda standart ko'rinishga keltirish.

Lotin yozuvidagi o'zbek alifbosida o‘ va g‘ harflari ustidagi belgi — «teskari vergul» (ʻ, yozuvda ‘),
tutuq belgisi esa oddiy apostrof shaklida (ʼ, yozuvda ’) yoziladi:
    o'zbek → o‘zbek,  g'oya → g‘oya,  tog' → tog‘,  ta'lim → ta’lim,  san'at → san’at.
AI matni va foydalanuvchi kiritgan ma'lumotlarda ko'pincha to'g'ri apostrof (') yoki noto'g'ri tomonga
qaragan belgi (o’zbek) uchraydi. Ushbu modul har bir Word (.docx) va PowerPoint (.pptx) hujjat SAQLANAYOTGANDA
uni bir marta tekshiradi: hujjat o'zbek-lotin matni bo'lsa, belgilar standartga keltiriladi.

Belgilar: ‘ (U+2018) va ’ (U+2019) — ʻ (U+02BB) / ʼ (U+02BC) bilan bir xil ko'rinadi, lekin har qanday shriftda
(Times New Roman, Arial, Calibri) bor; Word'dagi qidiruv va imlo tekshiruvi bilan ham to'g'ri ishlaydi.
Ruscha, inglizcha va kirill yozuvidagi hujjatlarga tegilmaydi.
"""
from __future__ import annotations

import logging
import re

log = logging.getLogger(__name__)

OKIRISH = "‘"   # ‘  — o‘, g‘ uchun
TUTUQ = "’"     # ’  — tutuq belgisi uchun
_APOS = "'‘’ʻʼ`´′‛"
_A = re.escape(_APOS)
_HARF = "A-Za-zÀ-ÖØ-öø-ÿ"

_OG_BOSH = re.compile(rf"(?<![{_HARF}])([OoGg])[{_A}]{{1,2}}(?=[{_HARF}])")      # so'z boshida: O'zbek, g'oya
_OG_OXIR = re.compile(rf"(?<=[{_HARF}])([OoGg])[{_A}](?![{_HARF}{_A}])")         # so'z oxirida: tog', bog'
_OG_ICH = re.compile(rf"(?<=[{_HARF}])([OoGg])[{_A}]{{1,2}}(?=[{_HARF}])")        # so'z ichida: bo'lim, sog'liq
_TUT = re.compile(rf"(?<=[{_HARF}])[{_A}](?=[{_HARF}])")                         # ta'lim, san'at, mas'uliyat
_UZ_BELGI = re.compile(rf"[OoGg][{_A}][a-z]")


def tuzat(t: str) -> str:
    """Bitta matn bo'lagi uchun (o'zbek-lotin matni ekanligi oldindan aniqlangan bo'lishi kerak)."""
    if not t or not any(c in _APOS for c in t):
        return t
    t = _TUT.sub(TUTUQ, t)                                   # avval hammasi tutuq deb olinadi...
    t = _OG_BOSH.sub(lambda m: m.group(1) + OKIRISH, t)       # ...so'ng o/g dan keyingisi ‘ ga almashadi
    t = _OG_ICH.sub(lambda m: m.group(1) + OKIRISH, t)
    t = _OG_OXIR.sub(lambda m: m.group(1) + OKIRISH, t)
    return t


def ozbekchami(matn: str) -> bool:
    """Hujjat o'zbek-lotin matnimi: o'/g' birikmalari yetarli va kirill harflari ko'p emas."""
    n = len(_UZ_BELGI.findall(matn))
    if n == 0:
        return False
    ingliz = len(re.findall(rf"[A-Za-z][{_A}](?:s|t|re|ll|ve|m|d)\b", matn))   # Smith's, don't, we're ...
    if n < 3 and ingliz >= n:
        return False
    kirill = sum(1 for c in matn if "Ѐ" <= c <= "ӿ")
    lotin = sum(1 for c in matn if c.isascii() and c.isalpha())
    return lotin > kirill


def _kontekstli(text: str, oldin: str, keyin: str) -> str:
    """Matn Word'da bir necha bo'lakka (run) bo'lingan bo'lishi mumkin — qo'shni belgilar hisobga olinadi."""
    if not any(c in _APOS for c in text):
        return text
    r = tuzat(oldin + text + keyin)
    if len(r) != len(oldin) + len(text) + len(keyin):     # faqat bir xil uzunlikda xavfsiz kesib olinadi
        r2 = tuzat(text)
        return r2
    return r[len(oldin):len(r) - len(keyin)] if keyin else r[len(oldin):]


def _qismlar(package):
    for part in package.iter_parts():
        el = getattr(part, "_element", None)
        if el is not None and hasattr(el, "iter"):
            yield el


def _paragraflar(el):
    """Har bir abzas (w:p yoki a:p) ichidagi matn tugunlari (w:t yoki a:t) ro'yxati."""
    for p in el.iter():
        tag = p.tag if isinstance(p.tag, str) else ""
        if tag.endswith("}p"):
            ts = [t for t in p.iter() if isinstance(t.tag, str) and t.tag.endswith("}t") and t is not p]
            if ts:
                yield ts


def paket_tuzat(package) -> int:
    """Hujjat paketidagi barcha matnni tekshiradi; o'zgartirilgan tugunlar sonini qaytaradi."""
    qismlar = list(_qismlar(package))
    jami = "".join(t.text or "" for el in qismlar for ts in _paragraflar(el) for t in ts)
    if not ozbekchami(jami):
        return 0
    n = 0
    for el in qismlar:
        for ts in _paragraflar(el):
            for i, t in enumerate(ts):
                s = t.text or ""
                if not s:
                    continue
                oldin = (ts[i - 1].text or "")[-2:] if i else ""
                keyin = (ts[i + 1].text or "")[:1] if i + 1 < len(ts) else ""
                yangi = _kontekstli(s, oldin, keyin)
                if yangi != s:
                    t.text = yangi
                    n += 1
    return n


_ORNATILDI = False


def ornat() -> None:
    """python-docx va python-pptx saqlash funksiyalariga tuzatishni ulaydi (bir marta)."""
    global _ORNATILDI
    if _ORNATILDI:
        return
    _ORNATILDI = True
    try:
        import docx.document as _dd

        _asl = _dd.Document.save

        def save(self, path_or_stream):
            try:
                paket_tuzat(self.part.package)
            except Exception as e:                      # hujjat baribir saqlanadi
                log.warning(f"Apostrof tuzatilmadi (docx): {e}")
            return _asl(self, path_or_stream)

        _dd.Document.save = save
    except Exception as e:
        log.warning(f"uz_apostrof docx: {e}")
    try:
        import pptx.presentation as _pp

        _asl2 = _pp.Presentation.save

        def save2(self, file):
            try:
                paket_tuzat(self.part.package)
            except Exception as e:
                log.warning(f"Apostrof tuzatilmadi (pptx): {e}")
            return _asl2(self, file)

        _pp.Presentation.save = save2
    except Exception as e:
        log.warning(f"uz_apostrof pptx: {e}")

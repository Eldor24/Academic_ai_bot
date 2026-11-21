"""
matn_norm.py — foydalanuvchi kiritgan titul ma'lumotlarini tartibga keltirish
═══════════════════════════════════════════════════════════════════
Foydalanuvchi OTM, fakultet, kafedra, fan, mavzu, ism va guruhni istalgan ko'rinishda yozishi mumkin:
KATTA HARFDA, kichik harfda, kirillda, lotinda yoki aralash, «kafedrasi» so'zini qo'shib yoki qo'shmay.
Bu modul ularni bir xil, to'g'ri ko'rinishga keltiradi:

  • o'zbek kirill → lotin (o'zbekcha hujjat uchun); lotin so'z ichidagi kirill «a, e, o, p, c, x» → lotin
  • turli apostroflar (’ ‘ ʻ ʼ ` ´) → bitta ko'rinish: o', g', ' (tutuq belgisi)
  • HAMMASI KATTA yoki hammasi kichik yozilgan matn → odatiy yozuv (ism-familiya: har so'z bosh harf bilan;
    nomlar: birinchi harf katta, atoqli otlar — O'zbekiston, Toshkent, Alisher Navoiy ... — bosh harf bilan)
  • titulda alohida yoziladigan so'zlar olib tashlanadi: «Filologiya fakulteti» → «Filologiya»,
    «Ona tili kafedrasi» → «Ona tili», «14-guruh» → «14», «Andijon shahri – 2026» → «Andijon»
"""
from __future__ import annotations

import re

# ── Apostroflar ───────────────────────────────────────────────────
_APOS = "’‘ʻʼ`´‛′"


def apostrof(t: str) -> str:
    t = re.sub(f"([OoGg])[{_APOS}']", lambda m: m.group(1) + "'", t)        # o‘ g‘ → o' g'
    t = re.sub(f"[{_APOS}]", "'", t)                                        # tutuq belgisi
    return t


# ── Kirill → lotin (o'zbek alifbosi qoidalari) ───────────────────
_K2L = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k",
    "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f",
    "х": "x", "ш": "sh", "ч": "ch", "ъ": "'", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ё": "yo",
    "ў": "o'", "қ": "q", "ғ": "g'", "ҳ": "h", "ц": "ts", "щ": "sh", "ы": "i",
}
_VOWELS = set("аеёиоуэюяўАЕЁИОУЭЮЯЎ")


def kiril_lotin(t: str) -> str:
    out = []
    for i, ch in enumerate(t):
        low = ch.lower()
        if low == "е":
            prev = t[i - 1] if i > 0 else " "
            lat = "ye" if (not prev.isalpha() or prev in _VOWELS or prev in "ъьЪЬ") else "e"
        elif low in _K2L:
            lat = _K2L[low]
        else:
            out.append(ch)
            continue
        if ch.isupper() and lat:
            nxt = t[i + 1] if i + 1 < len(t) else ""
            lat = lat.upper() if (nxt.isupper() and len(lat) > 1) else lat[0].upper() + lat[1:]
        out.append(lat)
    return "".join(out)


_HOMO = str.maketrans({"а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "к": "k", "м": "m",
                       "т": "t", "н": "h", "в": "b", "А": "A", "Е": "E", "О": "O", "Р": "P", "С": "C",
                       "Х": "X", "К": "K", "М": "M", "Т": "T", "Н": "H", "В": "B"})


def _aralash_tuzat(t: str) -> str:
    """Lotin so'z ichiga tushib qolgan kirill harflari (klaviatura aralashgan) — lotinga."""
    def fix(m):
        w = m.group(0)
        lat = sum(1 for c in w if "a" <= c.lower() <= "z")
        kir = sum(1 for c in w if "а" <= c.lower() <= "я" or c.lower() in "ёўқғҳ")
        return w.translate(_HOMO) if lat and kir and lat >= kir else w
    return re.sub(r"\w+", fix, t)


def yozuv(t: str, til: str) -> str:
    """Hujjat tiliga mos yozuv: o'zbekcha (lotin) hujjatda kirill lotinga o'giriladi."""
    t = _aralash_tuzat(t)
    if til == "uz" and re.search(r"[а-яёўқғҳА-ЯЁЎҚҒҲ]", t):
        t = kiril_lotin(t)
    return apostrof(t)


# ── Harflar ───────────────────────────────────────────────────────
ATOQLI = {
    # davlat, shahar, viloyat
    "o'zbekiston", "uzbekistan", "toshkent", "tashkent", "samarqand", "samarkand", "buxoro", "bukhara", "andijon",
    "namangan", "farg'ona", "fergana", "qashqadaryo", "surxondaryo", "xorazm", "qoraqalpog'iston", "jizzax",
    "sirdaryo", "navoiy", "termiz", "qarshi", "nukus", "urganch", "guliston", "qo'qon", "chirchiq", "xiva",
    "denov", "angren", "olmaliq", "marg'ilon", "shahrisabz", "kattaqo'rg'on", "osiyo", "yevropa", "rossiya",
    "amerika", "turkiya", "xitoy", "koreya", "yaponiya", "germaniya", "angliya", "markaziy", "respublikasi",
    # mashhur shaxslar (OTM nomlarida)
    "alisher", "navoiy", "mirzo", "ulug'bek", "amir", "temur", "beruniy", "abu", "rayhon", "ibn", "sino",
    "al-xorazmiy", "xorazmiy", "muhammad", "nizomiy", "abdulla", "qodiriy", "cho'lpon", "fitrat", "behbudiy",
    "zahiriddin", "bobur", "islom", "karimov", "imom", "buxoriy", "termiziy", "mahmud", "koshg'ariy",
    "yusuf", "xos", "hojib", "ahmad", "yassaviy", "jaloliddin", "manguberdi", "berdaq", "ajiniyoz", "furqat",
    "muqimiy", "ogahiy", "hamza", "g'afur", "g'ulom", "oybek", "zulfiya", "erkin", "vohidov", "abdulla",
    "oripov", "said", "ahmad", "sharof", "rashidov", "tolstoy", "pushkin", "shakespeare",
}
_KICHIK_QOLADI = {"va", "bilan", "uchun", "hamda", "yoki", "nomidagi", "и", "в", "на", "по", "для", "с", "of",
                  "and", "the", "in", "for", "at", "on"}


def _soz_katta(w: str) -> str:
    """So'z bosh harfi (apostrof va defisni hisobga olib): «abdug'affarova» → «Abdug'affarova»."""
    return w[:1].upper() + w[1:] if w else w


def _harf_holati(t: str) -> str:
    letters = [c for c in t if c.isalpha()]
    if not letters:
        return "bosh"
    if all(c.isupper() for c in letters):
        return "katta"
    if all(c.islower() for c in letters):
        return "kichik"
    return "aralash"


def ism_harf(t: str) -> str:
    """Ism-familiya: har so'z bosh harf bilan (faqat hammasi katta/kichik yozilgan bo'lsa)."""
    if _harf_holati(t) in ("katta", "kichik"):
        t = " ".join("-".join(_soz_katta(p.lower()) for p in w.split("-")) for w in t.split())
    return t


def _en_title(t: str) -> str:
    """Inglizcha nomlar: Title Case («Samarkand State University», «Synonyms in English»)."""
    out = []
    for i, w in enumerate(t.split()):
        lw = w.lower()
        out.append(lw if (i > 0 and lw in _KICHIK_QOLADI) else "-".join(_soz_katta(x) for x in lw.split("-")))
    return " ".join(out)


def nom_harf(t: str, til: str = "uz") -> str:
    """Nom (OTM, fakultet, kafedra, fan, mavzu): birinchi harf katta, atoqli otlar bosh harf bilan."""
    holat = _harf_holati(t)
    if til == "en" and holat in ("katta", "kichik"):
        return _en_title(t)
    if holat in ("katta", "kichik"):
        words = t.lower().split()
        oldi_nomidagi = words.index("nomidagi") if "nomidagi" in words else -1
        out = []
        for i, w in enumerate(words):
            core = re.sub(r"[^\w'-]", "", w)
            if i < oldi_nomidagi or core in ATOQLI or core.split("-")[0] in ATOQLI:
                w = _soz_katta(w) if not w[:1] in "«\"“(" else w[:1] + _soz_katta(w[1:])
            out.append(w)
        t = " ".join(out)
    # birinchi harf katta
    m = re.search(r"[^\W\d_]", t)
    if m:
        t = t[:m.start()] + t[m.start()].upper() + t[m.start() + 1:]
    return t


def _toza(t: str, nuqta: bool = True) -> str:
    t = re.sub(r"\s+", " ", (t or "").replace("\u00a0", " ")).strip()
    return t.strip(" ,;:«»\"“”" + ("." if nuqta else "")).strip()


_SUFFIX = {
    "fakultet": r"(fakulteti|fakultet|fak\.|факультети|факультет|факультета|faculty)",
    "kafedra": r"(kafedrasi|kafedra|каф\.|кафедраси|кафедра|кафедры|department)",
    "fan": r"(fanidan|fani|фанидан|фани|дисциплина|предмет|subject)",
    "shahar": r"(shahri|shahar|sh\.|шаҳри|город|г\.|city)",
    "guruh": r"(guruhi|guruh|gr\.|гуруҳи|гуруҳ|группа|группы|гр\.|group)",
}
_PREFIX = {
    "fakultet": r"(faculty of|факультет)",
    "kafedra": r"(department of|кафедра)",
    "guruh": r"(guruh|группа|group)",
}


def soz_olib_tashla(t: str, tur: str) -> str:
    """Titulda o'zi yozilgan so'zni foydalanuvchi matnidan olib tashlaydi («Ona tili kafedrasi» → «Ona tili»)."""
    if tur in _SUFFIX:
        t = re.sub(rf"[\s\-–]*\b{_SUFFIX[tur]}(?=\W*$)", "", t, flags=re.I)
    if tur in _PREFIX:
        t = re.sub(rf"^\s*{_PREFIX[tur]}\b[\s:\-–]*", "", t, flags=re.I)
    return _toza(t)


def normallashtir(t: str, tur: str, til: str = "uz", strip: bool = True) -> str:
    """tur: otm | fakultet | kafedra | fan | mavzu | ism | guruh | rahbar | shahar
    strip=False — «kafedrasi», «fakulteti» kabi so'zlar qoldiriladi (rus/ingliz titulida tartib boshqacha)."""
    t = _toza(yozuv(t or "", til), nuqta=tur not in ("rahbar", "ism"))
    if tur == "shahar":
        t = re.sub(r"[\s,–\-]*\d{4}\s*(y\.?|yil|г\.?|год)?\s*$", "", t, flags=re.I)
    if strip and (tur in _SUFFIX or tur in _PREFIX):
        t = soz_olib_tashla(t, tur)
    if tur == "ism":
        return ism_harf(t)
    if tur == "guruh":
        t = re.sub(r"[-\s]*(guruh\w*|группа|group)\b", "", t, flags=re.I).strip(" -–")
        return t.upper() if re.search(r"[A-Za-zА-Яа-я]", t) else t
    if tur == "rahbar":
        if _harf_holati(t) in ("katta", "kichik"):
            parts = []
            for w in t.split():
                if "." in w and len(re.sub(r"[^\w]", "", w)) <= 6 and not re.fullmatch(r"[A-Za-zА-Яа-я]\.[A-Za-zА-Яа-я]?\.?", w):
                    parts.append(w.lower())                      # i.f.n., dots., prof.
                elif re.fullmatch(r"([A-Za-zА-Яа-яЎўҚқҒғҲҳ]\.){1,2}[A-Za-zА-Яа-я]?\.?,?", w):
                    parts.append(w.upper())                      # A.B.
                else:
                    parts.append("-".join(_soz_katta(x.lower()) for x in w.split("-")))
            t = " ".join(parts)
        t = re.sub(r"\bphd\b", "PhD", t, flags=re.I)
        t = re.sub(r"\bdsc\b", "DSc", t, flags=re.I)
        return t
    if tur == "shahar":
        return ism_harf(t)
    return nom_harf(t, til)

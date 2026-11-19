"""
kurs_ishi.py — Professional kurs ishi generatori
═══════════════════════════════════════════════════════════════════
Ikki hajm:
  • standart     — ~25-30 bet  (2 bob × 3 paragraf)
  • kengaytirilgan — ~65-70 bet (3 bob × 3 paragraf)

O'zbekiston OTMlari kurs ishi talablariga mos:
  • A4, Times New Roman 14, interval 1,5, abzas 1,25 sm,
    hoshiyalar: chap 3 sm, o'ng 1,5 sm, yuqori/past 2 sm
  • Titul varag'i (raqamsiz), MUNDARIJA (Word'ning haqiqiy mundarija maydoni —
    Word ochilganda sahifa raqamlari avtomatik aniqlanadi), sahifa raqamlari pastda markazda
  • KIRISH: dolzarbligi, o'rganilganlik darajasi, maqsad, vazifalar, obyekt, predmet,
    metodlar, amaliy ahamiyati, tuzilishi
  • Boblar va paragraflar (Heading 1/2 — mundarijaga avtomatik tushadi)
  • Haqiqiy Word snoskalari — har bir snoska adabiyotlar ro'yxatidagi manbaga ishora qiladi
  • Jadvallar: «1-jadval» raqami, nomi, manbasi; sarlavha qatori har sahifada takrorlanadi
  • Grafiklar: faqat manbasi ko'rsatilgan raqamlar bilan, «1-rasm» raqami va manba bilan
  • XULOSA VA TAKLIFLAR, FOYDALANILGAN ADABIYOTLAR (normativ / asosiy / xorijiy / internet)

Bo'limlar parallel yoziladi (3 ta oqim), har bo'lim alohida AI so'rovi — hajm kafolatlanadi:
bo'lim kutilgandan qisqa chiqsa, avtomatik davom ettiriladi.
"""
from __future__ import annotations

import concurrent.futures as cf
import datetime
import io
import json
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading

from docx import Document
from docx.enum.section import WD_SECTION_START  # noqa: F401  (tashqi foydalanish uchun)
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════
# HAJMLAR
# ═══════════════════════════════════════════════
# O'lchov (LibreOffice'da tekshirilgan): TNR 14, 1,5 interval, A4, 3/1,5/2/2 sm → o'zbekcha ~215 so'z/bet
HAJMLAR = {
    "standart": dict(boblar=2, rejalar=3, sec_words=530, kirish_words=750, xulosa_par=4,
                     takliflar=5, adab=(15, 20), bet="25–30", bet_min=25, workers=3),
    "katta":    dict(boblar=3, rejalar=3, sec_words=1100, kirish_words=1150, xulosa_par=7,
                     takliflar=7, adab=(25, 35), bet="65–70", bet_min=62, workers=3),
}

WPP = 215          # so'z / bet (oddiy matn)
CPL = 68           # belgi / qator
LPP = 34           # qator / bet (LibreOffice bilan kalibrlangan)

_API: dict = {}


def setup(call_ai_json, clean_text=None, ai_ctx=None):
    """bot.py dan chaqiriladi."""
    _API.update(call_ai_json=call_ai_json, clean_text=clean_text, ai_ctx=ai_ctx)


# ═══════════════════════════════════════════════
# TIL YORLIQLARI
# ═══════════════════════════════════════════════
LBL = {
    "uz": {
        "vazirlik": "O'ZBEKISTON RESPUBLIKASI OLIY TA'LIM, FAN VA INNOVATSIYALAR VAZIRLIGI",
        "kurs_ishi": "KURS ISHI", "fan": "Fan", "mavzu": "Mavzu", "bajardi": "Bajardi",
        "guruh": "Guruh", "rahbar": "Ilmiy rahbar", "mundarija": "MUNDARIJA", "kirish": "KIRISH",
        "xulosa": "XULOSA VA TAKLIFLAR", "adab": "FOYDALANILGAN ADABIYOTLAR RO'YXATI",
        "takliflar": "Takliflar:", "bob": "{r} BOB.", "jadval": "{n}-jadval", "rasm": "{n}-rasm.",
        "jref": "{n}-jadval", "rref": "{n}-rasm",
        "manba": "Manba", "muallif_ishlanmasi": "Muallif ishlanmasi",
        "guruhlar": {"normativ": "Normativ-huquqiy hujjatlar", "asosiy": "Asosiy adabiyotlar",
                     "xorijiy": "Xorijiy adabiyotlar", "internet": "Internet manbalari"},
        "kirish_lbl": {
            "dolzarbligi": "Mavzuning dolzarbligi.", "organilganlik_darajasi": "Muammoning o'rganilganlik darajasi.",
            "maqsadi": "Tadqiqotning maqsadi.", "vazifalari": "Tadqiqotning vazifalari:",
            "obyekti": "Tadqiqot obyekti.", "predmeti": "Tadqiqot predmeti.",
            "metodologik_asosi": "Tadqiqot metodlari.", "amaliy_ahamiyati": "Ishning amaliy ahamiyati.",
            "tuzilishi": "Kurs ishining tuzilishi.",
        },
        "tuzilish": ("Kurs ishi kirish, {b} ta bob ({p} ta paragraf), xulosa va takliflar hamda "
                     "foydalanilgan adabiyotlar ro'yxatidan iborat.{extra} Ishning umumiy hajmi {bet} betni tashkil etadi."),
        "tuz_extra": " Ishda {j} ta jadval{r} keltirilgan.", "tuz_rasm": " va {r} ta rasm",
        "shaharlar": ["Toshkent", "Samarqand", "Buxoro", "Andijon", "Namangan", "Farg'ona", "Qarshi", "Termiz",
                      "Nukus", "Urganch", "Jizzax", "Guliston", "Navoiy", "Xiva", "Qo'qon", "Chirchiq", "Denov"],
        "lang": "Butun matnni FAQAT o'zbek tilida (lotin yozuvida) yoz.",
    },
    "ru": {
        "vazirlik": "МИНИСТЕРСТВО ВЫСШЕГО ОБРАЗОВАНИЯ, НАУКИ И ИННОВАЦИЙ РЕСПУБЛИКИ УЗБЕКИСТАН",
        "kurs_ishi": "КУРСОВАЯ РАБОТА", "fan": "Дисциплина", "mavzu": "Тема", "bajardi": "Выполнил(а)",
        "guruh": "Группа", "rahbar": "Научный руководитель", "mundarija": "СОДЕРЖАНИЕ", "kirish": "ВВЕДЕНИЕ",
        "xulosa": "ЗАКЛЮЧЕНИЕ И РЕКОМЕНДАЦИИ", "adab": "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ",
        "takliflar": "Рекомендации:", "bob": "ГЛАВА {r}.", "jadval": "Таблица {n}", "rasm": "Рис. {n}.",
        "jref": "табл. {n}", "rref": "рис. {n}",
        "manba": "Источник", "muallif_ishlanmasi": "Составлено автором",
        "guruhlar": {"normativ": "Нормативно-правовые акты", "asosiy": "Основная литература",
                     "xorijiy": "Зарубежная литература", "internet": "Интернет-ресурсы"},
        "kirish_lbl": {
            "dolzarbligi": "Актуальность темы.", "organilganlik_darajasi": "Степень изученности проблемы.",
            "maqsadi": "Цель исследования.", "vazifalari": "Задачи исследования:",
            "obyekti": "Объект исследования.", "predmeti": "Предмет исследования.",
            "metodologik_asosi": "Методы исследования.", "amaliy_ahamiyati": "Практическая значимость работы.",
            "tuzilishi": "Структура работы.",
        },
        "tuzilish": ("Курсовая работа состоит из введения, {b} глав ({p} параграфов), заключения и рекомендаций, "
                     "списка использованной литературы.{extra} Общий объём работы — {bet} страниц."),
        "tuz_extra": " В работе приведены {j} таблиц{r}.", "tuz_rasm": " и {r} рисунков",
        "shaharlar": ["Ташкент", "Самарканд", "Бухара", "Андижан", "Наманган", "Фергана", "Карши", "Термез",
                      "Нукус", "Ургенч", "Джизак", "Гулистан", "Навои", "Хива", "Коканд", "Чирчик"],
        "lang": "Пиши весь текст ТОЛЬКО на русском языке.",
    },
    "en": {
        "vazirlik": "MINISTRY OF HIGHER EDUCATION, SCIENCE AND INNOVATIONS OF THE REPUBLIC OF UZBEKISTAN",
        "kurs_ishi": "COURSE WORK", "fan": "Subject", "mavzu": "Topic", "bajardi": "Prepared by",
        "guruh": "Group", "rahbar": "Scientific supervisor", "mundarija": "TABLE OF CONTENTS",
        "kirish": "INTRODUCTION", "xulosa": "CONCLUSION AND RECOMMENDATIONS", "adab": "LIST OF REFERENCES",
        "takliflar": "Recommendations:", "bob": "CHAPTER {r}.", "jadval": "Table {n}", "rasm": "Figure {n}.",
        "jref": "Table {n}", "rref": "Figure {n}",
        "manba": "Source", "muallif_ishlanmasi": "Compiled by the author",
        "guruhlar": {"normativ": "Legal and regulatory documents", "asosiy": "Core literature",
                     "xorijiy": "Foreign literature", "internet": "Internet resources"},
        "kirish_lbl": {
            "dolzarbligi": "Relevance of the topic.", "organilganlik_darajasi": "Degree of study of the problem.",
            "maqsadi": "Aim of the research.", "vazifalari": "Objectives of the research:",
            "obyekti": "Object of the research.", "predmeti": "Subject of the research.",
            "metodologik_asosi": "Research methods.", "amaliy_ahamiyati": "Practical significance.",
            "tuzilishi": "Structure of the work.",
        },
        "tuzilish": ("The course work consists of an introduction, {b} chapters ({p} sections), a conclusion with "
                     "recommendations and a list of references.{extra} The total volume of the work is {bet} pages."),
        "tuz_extra": " The work contains {j} tables{r}.", "tuz_rasm": " and {r} figures",
        "shaharlar": ["Tashkent", "Samarkand", "Bukhara", "Andijan", "Namangan", "Fergana", "Karshi", "Termez",
                      "Nukus", "Urgench", "Jizzakh", "Gulistan", "Navoi", "Khiva", "Kokand", "Chirchik"],
        "lang": "Write all text ONLY in English.",
    },
}
KIRISH_TARTIB = ["dolzarbligi", "organilganlik_darajasi", "maqsadi", "vazifalari", "obyekti", "predmeti",
                 "metodologik_asosi", "amaliy_ahamiyati", "tuzilishi"]
ROMAN = ["I", "II", "III", "IV", "V", "VI"]
GURUH_TARTIB = ["normativ", "asosiy", "xorijiy", "internet"]

# Shahar nomlari OTM nomidan (lotin) — titul pastidagi «Samarqand – 2026» uchun
_SHAHAR_KALIT = {
    "toshkent": 0, "tashkent": 0, "ташкент": 0, "samarqand": 1, "samarkand": 1, "самарканд": 1,
    "buxoro": 2, "bukhara": 2, "бухар": 2, "andijon": 3, "andijan": 3, "андижан": 3,
    "namangan": 4, "наманган": 4, "farg'ona": 5, "fargʻona": 5, "fergana": 5, "фергана": 5, "ферган": 5,
    "qarshi": 6, "karshi": 6, "карши": 6, "termiz": 7, "termez": 7, "термез": 7,
    "nukus": 8, "нукус": 8, "qoraqalpo": 8, "urganch": 9, "urgench": 9, "ургенч": 9, "xorazm": 9,
    "jizzax": 10, "jizzakh": 10, "джизак": 10, "guliston": 11, "gulistan": 11, "гулистан": 11, "sirdaryo": 11,
    "navoiy": 12, "navoi": 12, "навои": 12, "xiva": 13, "khiva": 13, "хива": 13,
    "qo'qon": 14, "qoʻqon": 14, "kokand": 14, "коканд": 14, "chirchiq": 15, "chirchik": 15, "чирчик": 15,
}


def L(til: str) -> dict:
    return LBL.get(til, LBL["uz"])


def shahar_aniqla(otm: str, til: str) -> str:
    s = (otm or "").lower()
    for k, i in _SHAHAR_KALIT.items():
        if k in s:
            sh = L(til)["shaharlar"]
            return sh[i] if i < len(sh) else sh[0]
    return L(til)["shaharlar"][0]


# ═══════════════════════════════════════════════
# YORDAMCHI: matnni tozalash
# ═══════════════════════════════════════════════
def _clean(t) -> str:
    t = str(t or "")
    fn = _API.get("clean_text")
    if fn:
        try:
            t = fn(t)
        except Exception:
            pass
    t = re.sub(r"\*\*|__|##+|`", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _strip_num(t: str) -> str:
    """AI qo'ygan raqam/prefikslarni olib tashlaydi: '1.1. ', 'I BOB. ', 'ГЛАВА 2.' va h.k."""
    t = _clean(t)
    t = re.sub(r"^(?:[IVX]+\s*(?:BOB|bob|БОБ)|(?:ГЛАВА|Глава|CHAPTER|Chapter)\s*[IVX\d]+|BOB\s*[IVX\d]+)[\s\.\:\-–—]*",
               "", t)
    t = re.sub(r"^\d+(?:\.\d+)*\.?\s+", "", t)
    t = re.sub(r"^\d+-(?:abzas|paragraf|bo'lim)[\s\.:]*", "", t, flags=re.I)
    return t.strip(" .:-–—")


def _sentence_case(t: str) -> str:
    t = t.strip()
    if t and t.isupper() and len(t) > 6:
        t = t.capitalize()
    return t[:1].upper() + t[1:] if t else t


def _words(t: str) -> int:
    return len(re.findall(r"\w+", t or ""))


def _num(v):
    try:
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).replace(" ", "").replace(" ", "").replace(",", ".").replace("%", "")
        return float(s)
    except Exception:
        return None


# ═══════════════════════════════════════════════
# AI SO'ROVLARI
# ═══════════════════════════════════════════════
def _rollar(hajm: str) -> list[str]:
    if HAJMLAR[hajm]["boblar"] == 2:
        return [
            "NAZARIY-METODOLOGIK ASOSLAR: asosiy tushunchalar va ta'riflar, olimlarning ilmiy qarashlari, "
            "rivojlanish tarixi, tasnif, omillar va baholash mezonlari",
            "AMALIY TAHLIL VA TAKOMILLASHTIRISH: O'zbekistondagi hozirgi holat, xorij tajribasi, qiyosiy tahlil, "
            "mavjud muammolar, ularni hal etish yo'llari va aniq takliflar",
        ]
    return [
        "NAZARIY ASOSLAR: asosiy tushunchalar va ta'riflar, olimlarning ilmiy qarashlari (O'zbekiston va xorij), "
        "rivojlanish tarixi, tasnif, omillar, baholash mezonlari",
        "HOZIRGI HOLAT TAHLILI: O'zbekiston amaliyoti, xorij tajribasi, qiyosiy tahlil, tendensiyalar, "
        "mavjud muammolar va ularning sabablari",
        "TAKOMILLASHTIRISH YO'LLARI VA ISTIQBOLLAR: muammolarni hal etish yo'nalishlari, xorij tajribasini "
        "qo'llash imkoniyatlari, aniq takliflar va kutilayotgan natijalar",
    ]


def _plan_prompt(p: dict) -> str:
    h = HAJMLAR[p["hajm"]]
    rollar = "\n".join(f"  {ROMAN[i]} bob — {r}" for i, r in enumerate(_rollar(p["hajm"])))
    lo, hi = h["adab"]
    return f"""Sen «{p['fan']}» fani bo'yicha 20 yillik tajribaga ega ilmiy rahbar va professorsan. {L(p['til'])['lang']}
«{p['mavzu']}» mavzusidagi KURS ISHI uchun REJA (mundarija) va FOYDALANILGAN ADABIYOTLAR ro'yxatini tuz.

TUZILMA: aynan {h['boblar']} ta bob, har bobda aynan {h['rejalar']} ta paragraf.
Boblarning vazifasi:
{rollar}
Agar mavzu gumanitar fan (adabiyot, tilshunoslik, tarix, pedagogika, psixologiya, san'at) bo'yicha bo'lsa —
boblar vazifasini fanga moslashtir (masalan: matn/asar tahlili, metodika, tarixiy jarayonlar tahlili).
Bob va paragraf nomlari mavzuga ANIQ bog'langan, ilmiy, takrorlanmaydigan bo'lsin (umumiy «Asosiy tushunchalar» emas).
Har paragraf uchun 3-4 ta "bandlar" — shu paragrafda yoritiladigan aniq masalalar.
"nomi" va "sarlavha" ga raqam yoki «I BOB», «1.1.» qo'shma — raqamlar avtomatik qo'yiladi.

ADABIYOTLAR: {lo}-{hi} ta manba. FAQAT haqiqatda mavjud, ishonching komil bo'lgan manbalar — kam bo'lsa ham mayli,
sonini to'ldirish uchun manba O'YLAB TOPMA. Turlari:
  "normativ" — O'zbekiston Respublikasi Konstitutsiyasi, kodekslar, qonunlar, Prezident farmon va qarorlari
               (faqat mavzuga aloqador, raqami va sanasini aniq bilganlaringni);
  "asosiy"   — darsliklar, monografiyalar, ilmiy maqolalar (o'zbek va rus tilida);
  "xorijiy"  — ingliz tilidagi taniqli kitob va maqolalar;
  "internet" — faqat rasmiy saytlarning bosh manzili (masalan https://lex.uz, https://stat.uz, https://www.worldbank.org).
Yozilish formati (GOST): Familiya I.O. Asar nomi. – Shahar: Nashriyot, yil. – bet soni b.
Nashriyot, yil yoki bet sonini bilmasang — o'sha qismni tashlab yoz, o'ylab topma.

FAQAT JSON qaytar (Markdown yo'q):
{{"boblar":[{{"nomi":"bob nomi","rejalar":[{{"sarlavha":"paragraf nomi","bandlar":["...","...","..."]}}]}}],
 "adabiyotlar":[{{"tur":"normativ","matn":"..."}},{{"tur":"asosiy","matn":"..."}}]}}"""


def _outline_text(plan: dict) -> str:
    rows = []
    for bi, b in enumerate(plan["boblar"]):
        rows.append(f"{ROMAN[bi]} bob. {b['nomi']}")
        for ri, r in enumerate(b["rejalar"]):
            rows.append(f"   {bi + 1}.{ri + 1}. {r['sarlavha']}")
    return "\n".join(rows)


def _adab_text(adab: list) -> str:
    return "\n".join(f"{i + 1}. {a['matn']}" for i, a in enumerate(adab))


def _kirish_prompt(p: dict, plan: dict) -> str:
    katta = p["hajm"] == "katta"
    w = {"d": 300 if katta else 190, "o": 240 if katta else 120, "m": 80 if katta else 60, "v": 6 if katta else 5,
         "ob": 70 if katta else 50, "pr": 70 if katta else 50, "me": 110 if katta else 70,
         "a": 110 if katta else 70}
    return f"""Sen «{p['fan']}» fani bo'yicha tajribali ilmiy rahbarsan. {L(p['til'])['lang']}
«{p['mavzu']}» mavzusidagi kurs ishining KIRISH qismini yoz.

Kurs ishi rejasi:
{_outline_text(plan)}

Talablar (so'z soni — taxminiy):
- "dolzarbligi": {w['d']} so'z, 2-3 abzas (abzaslarni " [ABZAS] " bilan ajrat). Mavzuning nazariy va amaliy ahamiyati,
  O'zbekistonda va jahonda mavzu bo'yicha olib borilayotgan ishlar. Faqat haqiqatda mavjud hujjat va faktlarni tilga ol.
- "organilganlik_darajasi": {w['o']} so'z. Mavzuni o'rgangan HAQIQIY olimlar (o'zbek, rus va xorijiy) va ularning
  qaysi jihatni o'rganganligi. Olim nomini bilmasang — umumiy ta'rifla, ism o'ylab topma.
- "maqsadi": {w['m']} so'z, bitta aniq maqsad.
- "vazifalari": {w['v']} ta vazifa ro'yxati (har biri — fe'l bilan boshlanadigan to'liq gap; paragraflarga mos).
- "obyekti": {w['ob']} so'z. "predmeti": {w['pr']} so'z (obyekt va predmet aniq farqlansin).
- "metodologik_asosi": {w['me']} so'z — qo'llanilgan ilmiy metodlar (tahlil, sintez, qiyosiy, tizimli, statistik,
  mantiqiy umumlashtirish va h.k.) va nazariy-uslubiy asos.
- "amaliy_ahamiyati": {w['a']} so'z — natijalardan kim, qayerda foydalanishi mumkin.
Kirishda [1] kabi havolalar va raqamli statistik ma'lumotlar YOZILMASIN.

FAQAT JSON:
{{"dolzarbligi":"...","organilganlik_darajasi":"...","maqsadi":"...","vazifalari":["...","..."],
 "obyekti":"...","predmeti":"...","metodologik_asosi":"...","amaliy_ahamiyati":"..."}}"""


def _section_rules(hajm: str, bi: int, ri: int, nbob: int) -> tuple[bool, bool, bool]:
    """(jadval_ruxsat, jadval_majburiy, grafik_ruxsat)"""
    oxirgi = bi == nbob - 1
    if bi == 0:                                   # nazariy bob
        return (ri == 1, ri == 1, False)
    if nbob == 3 and bi == 1:                     # tahlil bobi (katta)
        return (True, True, ri in (0, 1))
    if oxirgi and nbob == 3:                      # takomillashtirish bobi
        return (ri in (0, 2), ri == 0, ri == 1)
    # standart: II bob (tahlil + takliflar)
    return (True, ri in (0, 2), ri == 0)


def _section_prompt(p: dict, plan: dict, adab: list, bi: int, ri: int) -> str:
    h = HAJMLAR[p["hajm"]]
    b = plan["boblar"][bi]
    r = b["rejalar"][ri]
    words = h["sec_words"]
    npar = max(5, round(words / 115))
    lo, hi = int(words / npar * 0.85), int(words / npar * 1.15)
    rol = _rollar(p["hajm"])[bi]
    jr, jm, gr = _section_rules(p["hajm"], bi, ri, len(plan["boblar"]))
    if jr:
        jadval = (
            ("JADVAL (MAJBURIY): " if jm else "JADVAL (agar mazmunni boyitsa): ") +
            "1 ta jadval. Raqamli (statistik) jadval faqat taniqli rasmiy manba (O'zstat, Markaziy bank, vazirliklar, "
            "Jahon banki, BMT va h.k.) ma'lumotlari bilan, \"manba\" maydonida manba nomi va yili. Ishonchli raqam "
            "bilmasang — RAQAMSIZ jadval tuz (tasnif, qiyosiy tahlil, afzallik va kamchiliklar, bosqichlar, "
            "muammo – yechim – kutilayotgan natija) va \"manba\" ni bo'sh qoldir. 3-5 ustun, 4-8 qator, katakchalar qisqa. "
            "Jadval tahlil qilingan abzasda {JADVAL} belgisini yoz (u «1-jadval» / «табл. 1» / «Table 1» ga almashtiriladi) (masalan: \"{JADVAL} ma'lumotlaridan ko'rinadiki, ...\"; qo'shimchani ajratib yoz: \"{JADVAL} ma'lumotlari\"). "
            "\"keyin\" — jadval qaysi abzasdan keyin turishi (raqam).\n"
            '"jadval": {"nomi":"...","manba":"","ustunlar":["...","..."],"qatorlar":[["...","..."]],"keyin":2}'
        )
    else:
        jadval = '"jadval": null bo\'lsin.'
    if gr:
        grafik = (
            "GRAFIK (ixtiyoriy): mavzu bo'yicha taniqli rasmiy manbadan ANIQ bilgan raqamli ma'lumot bo'lsa "
            "(yillar bo'yicha dinamika yoki tarkib) — 1 ta grafik: \"tur\": \"chiziq\" (yillar dinamikasi), \"ustun\" "
            "(taqqoslash) yoki \"doira\" (tarkib, jami ≈100%). 4-10 ta qiymat. \"manba\" MAJBURIY (nomi va yili). "
            "Ishonchli raqam bilmasang — \"grafik\": null (to'qilgan raqamli grafik qo'yishdan ko'ra grafiksiz yaxshi). "
            "Grafik tahlil qilingan abzasda {RASM} belgisini yoz (masalan: \"{RASM} ma'lumotlari shuni ko'rsatadiki\").\n"
            '"grafik": {"nomi":"...","tur":"chiziq","birlik":"mlrd so\'m","manba":"O\'zstat, 2024",'
            '"qiymatlar":[["2020",10.5],["2021",12.1]],"keyin":5}'
        )
    else:
        grafik = '"grafik": null bo\'lsin.'
    bandlar = "\n".join(f"  - {x}" for x in r.get("bandlar", [])) or "  - (paragraf nomidan kelib chiqib)"
    return f"""Sen «{p['fan']}» fani bo'yicha tajribali ilmiy rahbar va akademik muallifsan. {L(p['til'])['lang']}
Kurs ishi mavzusi: «{p['mavzu']}».

Kurs ishi rejasi:
{_outline_text(plan)}

HOZIR YOZILADIGAN QISM: {ROMAN[bi]} bob «{b['nomi']}» → {bi + 1}.{ri + 1}. «{r['sarlavha']}»
Bobning vazifasi: {rol}
Shu paragrafda yoritiladigan masalalar:
{bandlar}
Boshqa paragraflardagi masalalarni takrorlama.

HAJM: aynan {npar} ta abzas, har biri {lo}-{hi} so'z (jami ≈{words} so'z). Har abzas to'liq va mazmunli.
USLUB: ilmiy-akademik; ta'riflar, dalillar, olimlarning qarashlari (faqat HAQIQIY olimlar va ularning haqiqiy fikrlari),
sabab-oqibat tahlili, misollar. Abzaslar mantiqan bog'lansin. «Ushbu paragrafda...» bilan boshlama, sarlavhani takrorlama.
Oxirgi abzas — paragraf bo'yicha qisqa xulosa.
ILMIY HAVOLALAR: quyidagi ro'yxatdagi manbalarga murojaat qilganda (iqtibos, ta'rif, muhim fikr) gapdan keyin
[raqam] qo'y — raqam ro'yxatdagi tartib raqami. Paragrafda 2-4 ta havola, bitta abzasda ko'pi bilan bitta.
Ro'yxatda yo'q manbaga havola qilma, sahifa raqami yozma.
ADABIYOTLAR:
{_adab_text(adab)}

{jadval}
{grafik}

ANIQLIK: sana, raqam, foiz va statistikani faqat aniq bilsang va manbasi bo'lsa yoz; aks holda sifat tavsifi ber.
O'tkazilmagan so'rovnoma, tajriba yoki respondentlar haqida yozma. Qonun va farmonlarni faqat aniq bilgan holda tilga ol.

FAQAT JSON (Markdown yo'q, matn ichida yangi qator yo'q):
{{"paragraflar":["1-abzas matni","2-abzas matni"],"jadval":null,"grafik":null}}"""


def _davom_prompt(p: dict, plan: dict, bi: int, ri: int, mavjud: list, kerak: int) -> str:
    b = plan["boblar"][bi]
    r = b["rejalar"][ri]
    oxirgi = " ".join(mavjud[-2:])[-1200:]
    return f"""{L(p['til'])['lang']}
«{p['mavzu']}» kurs ishining {bi + 1}.{ri + 1}. «{r['sarlavha']}» paragrafi qisqa chiqdi. Uni DAVOM ETTIR:
yana {kerak} ta yangi abzas yoz (har biri 100-130 so'z), oldingi fikrlarni takrorlamay, quyidagi masalalarni chuqurroq yorit:
{chr(10).join('  - ' + x for x in r.get('bandlar', []))}
Paragrafning oxirgi qismi (davomini shu fikrdan boshla):
«{oxirgi}»
Havola, jadval va raqamli statistika qo'shma. Oxirgi abzas paragraf bo'yicha xulosa bo'lsin.
FAQAT JSON: {{"paragraflar":["...","..."]}}"""


def _xulosa_prompt(p: dict, plan: dict) -> str:
    h = HAJMLAR[p["hajm"]]
    bandlar = []
    for bi, b in enumerate(plan["boblar"]):
        for ri, r in enumerate(b["rejalar"]):
            bandlar.append(f"{bi + 1}.{ri + 1}. {r['sarlavha']}: " + "; ".join(r.get("bandlar", [])))
    return f"""Sen tajribali ilmiy rahbarsan. {L(p['til'])['lang']}
«{p['mavzu']}» mavzusidagi kurs ishining XULOSA VA TAKLIFLAR qismini yoz.
Ishning tarkibi:
{chr(10).join(bandlar)}

"xulosalar": aynan {h['xulosa_par']} ta abzas (har biri 100-130 so'z): har bir bob bo'yicha asosiy ilmiy xulosalar
(nazariy xulosalar, tahlil natijalari, aniqlangan muammolar), umumiy xulosa.
"takliflar": aynan {h['takliflar']} ta aniq, amaliy taklif (har biri 1-2 gap; kim va nima qilishi kerakligi aniq).
Ishda bo'lmagan yangi raqam yoki faktlarni qo'shma, [1] kabi havolalar yozma, «Xulosa qilib aytganda» kabi shablonlardan qoch.
FAQAT JSON: {{"xulosalar":["...","..."],"takliflar":["...","..."]}}"""


def _ask(prompt: str, name: str, max_tokens: int, temperature: float = 0.6) -> dict:
    return _API["call_ai_json"](prompt, name, max_tokens=max_tokens, temperature=temperature)


# ═══════════════════════════════════════════════
# NATIJALARNI TEKSHIRISH
# ═══════════════════════════════════════════════
def _norm_plan(raw: dict, hajm: str) -> dict:
    h = HAJMLAR[hajm]
    boblar = []
    for b in (raw.get("boblar") or [])[:h["boblar"]]:
        if not isinstance(b, dict):
            continue
        rejalar = []
        for r in (b.get("rejalar") or [])[:h["rejalar"]]:
            if isinstance(r, str):
                r = {"sarlavha": r, "bandlar": []}
            s = _sentence_case(_strip_num(r.get("sarlavha", "")))
            if len(s) < 4:
                continue
            band = [_clean(x) for x in (r.get("bandlar") or []) if len(_clean(x)) > 3][:5]
            rejalar.append({"sarlavha": s, "bandlar": band})
        nomi = _strip_num(b.get("nomi", ""))
        if len(rejalar) == h["rejalar"] and len(nomi) >= 4:
            boblar.append({"nomi": nomi.upper(), "rejalar": rejalar})
    if len(boblar) != h["boblar"]:
        raise ValueError(f"Reja to'liq emas: {len(boblar)} bob")
    return {"boblar": boblar}


def _norm_adab(raw: list, hajm: str) -> list:
    seen, out = set(), []
    for a in raw or []:
        if isinstance(a, str):
            a = {"tur": "asosiy", "matn": a}
        if not isinstance(a, dict):
            continue
        m = _clean(a.get("matn", ""))
        m = re.sub(r"^\d+[\.\)]\s*", "", m).strip()
        m = re.sub(r"\[\d+\]", "", m).strip()
        if len(m) < 12 or m.lower() in seen:
            continue
        tur = str(a.get("tur", "asosiy")).lower().strip()
        if tur not in GURUH_TARTIB:
            tur = "internet" if re.search(r"https?://|www\.", m) else "asosiy"
        if tur == "internet" and not re.search(r"https?://|www\.|\.[a-z]{2,3}\b", m):
            tur = "asosiy"
        seen.add(m.lower())
        out.append({"tur": tur, "matn": m})
    out.sort(key=lambda a: GURUH_TARTIB.index(a["tur"]))       # GOST tartibi: normativ → asosiy → xorijiy → internet
    return out[:HAJMLAR[hajm]["adab"][1]]


def _fix_refs(text: str, nadab: int, used: set) -> str:
    """[1, 3] → [1][3]; ro'yxatda yo'q raqamlarni olib tashlaydi."""
    text = re.sub(r"\[(\d+(?:\s*[,;]\s*\d+)+)\]",
                  lambda m: "".join(f"[{x.strip()}]" for x in re.split(r"[,;]", m.group(1))), text)

    def rep(m):
        n = int(m.group(1))
        if 1 <= n <= nadab:
            used.add(n)
            return f"[{n}]"
        return ""
    return re.sub(r"\[(\d+)\]", rep, text)


def _strip_refs(text: str) -> str:
    return re.sub(r"\s*\[\d+(?:\s*[,;]\s*\d+)*\]", "", text)


_NUM_RE = re.compile(r"^[\s\-–+]*[\d\s.,]+\s*(%|‰|mln|mlrd|ming|тыс|млн|млрд|million|billion|thousand)?\.?\s*$", re.I)


def _norm_table(t, til: str):
    if not isinstance(t, dict):
        return None
    cols = [_clean(c) for c in (t.get("ustunlar") or [])][:6]
    rows = []
    for r in (t.get("qatorlar") or [])[:14]:
        if isinstance(r, list) and r:
            rr = [_clean(c) for c in r][:len(cols) or 6]
            rr += [""] * (len(cols) - len(rr))
            rows.append(rr)
    if len(cols) < 2 or len(rows) < 2:
        return None
    nomi = _strip_refs(_clean(t.get("nomi", "")))
    manba = _clean(t.get("manba", ""))
    if re.fullmatch(r"(?i)(manba|source|источник)?[:\s\-–]*", manba or ""):
        manba = ""
    cells = [c for r in rows for c in r[1:]]
    numeric = sum(1 for c in cells if c and _NUM_RE.match(c)) / max(1, len(cells))
    if numeric > 0.4 and not manba:
        log.info("Kurs ishi: manbasiz raqamli jadval olib tashlandi")
        return None
    if not manba:
        manba = L(til)["muallif_ishlanmasi"]
    return {"nomi": nomi or "-", "manba": manba, "ustunlar": cols, "qatorlar": rows,
            "keyin": int(_num(t.get("keyin")) or 2)}


def _norm_chart(g):
    if not isinstance(g, dict):
        return None
    manba = _clean(g.get("manba", ""))
    if len(manba) < 3:
        return None                     # manbasiz grafik — uydirma bo'lishi mumkin
    tur = str(g.get("tur", "ustun")).lower()
    tur = tur if tur in ("ustun", "chiziq", "doira") else "ustun"
    vals = []
    for it in (g.get("qiymatlar") or [])[:12]:
        if isinstance(it, (list, tuple)) and len(it) >= 2:
            lab, v = _clean(it[0])[:24], _num(it[1])
        elif isinstance(it, dict):
            lab, v = _clean(it.get("nom") or it.get("label") or "")[:24], _num(it.get("qiymat") or it.get("value"))
        else:
            continue
        if lab and v is not None and math.isfinite(v):
            vals.append((lab, v))
    if len(vals) < 3:
        return None
    if tur == "doira" and any(v < 0 for _, v in vals):
        tur = "ustun"
    return {"nomi": _strip_refs(_clean(g.get("nomi", ""))) or "-", "tur": tur,
            "birlik": _clean(g.get("birlik", ""))[:30], "manba": manba, "qiymatlar": vals,
            "keyin": int(_num(g.get("keyin")) or 4)}


# Matnda saqlash formati (eski tahrirlash tizimi bilan mos):
#   abzaslar " [ABZAS] " bilan; jadval [JADVAL:nomi\nA|B\n1|2\nManba: ...]; grafik [GRAFIK:{json, ] belgisisiz}]
def _enc_table(t: dict) -> str:
    def c(x):
        return str(x).replace("|", "/").replace("]", ")").replace("[", "(").replace("\n", " ")
    lines = [c(t["nomi"]), "|".join(c(x) for x in t["ustunlar"])]
    lines += ["|".join(c(x) for x in r) for r in t["qatorlar"]]
    lines.append("Manba: " + c(t["manba"]))
    return "[JADVAL:" + "\n".join(lines) + "]"


def _enc_chart(g: dict) -> str:
    d = {"nomi": g["nomi"], "tur": g["tur"], "birlik": g["birlik"], "manba": g["manba"],
         "q": [[a, b] for a, b in g["qiymatlar"]]}
    s = json.dumps(d, ensure_ascii=False).replace("[", "⟦").replace("]", "⟧")
    return "[GRAFIK:" + s + "]"


def _dec_chart(s: str):
    try:
        d = json.loads(s.replace("⟦", "[").replace("⟧", "]"))
        return {"nomi": d.get("nomi", ""), "tur": d.get("tur", "ustun"), "birlik": d.get("birlik", ""),
                "manba": d.get("manba", ""), "qiymatlar": [(a, float(b)) for a, b in d.get("q", [])]}
    except Exception:
        return None


# ═══════════════════════════════════════════════
# GENERATSIYA
# ═══════════════════════════════════════════════
def _write_section(p, plan, adab, bi, ri) -> dict:
    h = HAJMLAR[p["hajm"]]
    words = h["sec_words"]
    data = None
    for attempt in range(2):
        try:
            data = _ask(_section_prompt(p, plan, adab, bi, ri), f"{bi + 1}.{ri + 1}-paragraf",
                        max_tokens=7500 if p["hajm"] == "katta" else 4500)
            if data.get("paragraflar"):
                break
        except Exception as e:
            log.warning(f"Kurs ishi {bi + 1}.{ri + 1} xato (urinish {attempt + 1}): {e}")
            if attempt == 1:
                raise
    paras = [_clean(x) for x in (data.get("paragraflar") or []) if isinstance(x, str)]
    paras = [_strip_num(x) if re.match(r"^\d+-(abzas|paragraf)", x, re.I) else x for x in paras]
    paras = [x for x in paras if _words(x) >= 12]

    # Hajm kafolati: qisqa chiqsa — davom ettiramiz (ko'pi bilan 2 marta)
    for _ in range(2):
        jami = sum(_words(x) for x in paras)
        if jami >= words * 0.88 or not paras:
            break
        kerak = max(2, min(8, round((words - jami) / 115)))
        try:
            dv = _ask(_davom_prompt(p, plan, bi, ri, paras, kerak), f"{bi + 1}.{ri + 1} davomi", max_tokens=5000)
            add = [_strip_refs(_clean(x)) for x in (dv.get("paragraflar") or []) if isinstance(x, str)]
            add = [x for x in add if _words(x) >= 12]
            if not add:
                break
            paras += add
        except Exception as e:
            log.warning(f"Davom ettirish xatosi: {e}")
            break
    if not paras:
        raise ValueError(f"{bi + 1}.{ri + 1}-paragraf bo'sh qaytdi")

    used: set = set()
    per_par = []
    for x in paras:                                  # bitta abzasda ko'pi bilan 2 ta havola
        x = _fix_refs(x, len(adab), used)
        refs = re.findall(r"\[\d+\]", x)
        for extra in refs[2:]:
            x = x.replace(extra, "", 1)
        per_par.append(x)
    paras = per_par

    jr, jm, gr = _section_rules(p["hajm"], bi, ri, len(plan["boblar"]))
    tbl = _norm_table(data.get("jadval"), p["til"]) if jr else None
    gfk = _norm_chart(data.get("grafik")) if gr else None
    blocks = list(paras)
    inserts = []
    if tbl:
        inserts.append((min(max(1, tbl["keyin"]), len(paras)), _enc_table(tbl)))
    if gfk:
        inserts.append((min(max(1, gfk["keyin"]), len(paras)), _enc_chart(gfk)))
    for pos, block in sorted(inserts, key=lambda x: -x[0]):
        blocks.insert(pos, block)
    snoskalar = {str(n): adab[n - 1]["matn"] for n in sorted(used)}
    return {"sarlavha": f"{bi + 1}.{ri + 1}. {plan['boblar'][bi]['rejalar'][ri]['sarlavha']}",
            "matn": " [ABZAS] ".join(blocks), "snoskalar": snoskalar}


def generate(cid: int, p: dict, progress=None) -> tuple[str, dict]:
    """
    p: otm, fakultet, fan, mavzu, ism, guruh, rahbar, hajm ("standart"|"katta"), til ("uz"|"ru"|"en")
    Qaytaradi: (fayl_yo'li, data) — data tahrirlash tizimi bilan mos (bob1_data, bob2_data, ...).
    """
    p = dict(p)
    p["hajm"] = p.get("hajm") if p.get("hajm") in HAJMLAR else "standart"
    p["til"] = p.get("til") if p.get("til") in LBL else "uz"
    h = HAJMLAR[p["hajm"]]
    plock = threading.Lock()

    def prog(txt):
        if progress:
            with plock:
                try:
                    progress.update(txt)
                except Exception:
                    pass

    T = PROG.get(p["til"], PROG["uz"])
    jami_qadam = h["boblar"] * h["rejalar"]
    prog(T["reja"].format(bet=h["bet"]))

    # 1) Reja + adabiyotlar
    plan_raw, plan = None, None
    for attempt in range(3):
        plan_raw = _ask(_plan_prompt(p), "Reja", max_tokens=5000, temperature=0.5)
        try:
            plan = _norm_plan(plan_raw, p["hajm"])
            break
        except ValueError as e:
            log.warning(f"Kurs ishi rejasi: {e} (urinish {attempt + 1})")
    if not plan:
        raise ValueError("Kurs ishi rejasini tuzib bo'lmadi")
    adab = _norm_adab(plan_raw.get("adabiyotlar"), p["hajm"])
    log.info(f"[{cid}] Kurs ishi rejasi: {len(plan['boblar'])} bob, {len(adab)} ta adabiyot")

    # 2) Kirish
    prog(T["kirish"].format(bet=h["bet"]))
    kirish_raw = _ask(_kirish_prompt(p, plan), "Kirish", max_tokens=5000)
    kirish = {}
    for k in KIRISH_TARTIB[:-1]:
        v = kirish_raw.get(k)
        if isinstance(v, list):
            items = [_strip_refs(_clean(x)).rstrip(";.") for x in v if _clean(x)]
            v = " [ABZAS] ".join(f"– {x};" if i < len(items) - 1 else f"– {x}." for i, x in enumerate(items))
        else:
            v = _strip_refs(_clean(v))
        if v:
            kirish[k] = v

    # 3) Paragraflar (parallel)
    tasks = [(bi, ri) for bi in range(h["boblar"]) for ri in range(h["rejalar"])]
    natija: dict = {}
    done = [0]
    ctx = _API.get("ai_ctx")

    def run(bi, ri):
        if ctx:
            with ctx(cid):
                return _write_section(p, plan, adab, bi, ri)
        return _write_section(p, plan, adab, bi, ri)

    prog(T["bob"].format(k=0, n=jami_qadam, bet=h["bet"]))
    with cf.ThreadPoolExecutor(max_workers=h["workers"]) as ex:
        futs = {ex.submit(run, bi, ri): (bi, ri) for bi, ri in tasks}
        for f in cf.as_completed(futs):
            natija[futs[f]] = f.result()            # xato bo'lsa — butun ish to'xtaydi (tanga qaytariladi)
            done[0] += 1
            prog(T["bob"].format(k=done[0], n=jami_qadam, bet=h["bet"]))

    # 4) Xulosa va takliflar
    prog(T["xulosa"].format(bet=h["bet"]))
    xr = _ask(_xulosa_prompt(p, plan), "Xulosa", max_tokens=4500)
    xul = [_strip_refs(_clean(x)) for x in (xr.get("xulosalar") or []) if _words(_clean(x)) >= 10]
    tak = [_strip_refs(_clean(x)).rstrip(".") for x in (xr.get("takliflar") or []) if _words(_clean(x)) >= 4]
    xparts = list(xul)
    if tak:
        xparts.append(L(p["til"])["takliflar"])
        tak = [re.sub(r"^\d+[\.\)]\s*", "", x) for x in tak]
        xparts += [f"{i + 1}) {x}." for i, x in enumerate(tak)]
    xulosa = " [ABZAS] ".join(xparts)

    data = {
        "v": 2, "hajm": p["hajm"], "til": p["til"], "bob_soni": h["boblar"],
        "titul": {k: p.get(k, "") for k in ("otm", "fakultet", "fan", "mavzu", "ism", "guruh", "rahbar")},
        "plan": {"kirish": kirish, "adabiyotlar": adab},
        "xulosa_data": {"xulosa": xulosa},
    }
    for bi in range(h["boblar"]):
        data[f"bob{bi + 1}_nomi"] = L(p["til"])["bob"].format(r=ROMAN[bi]) + " " + plan["boblar"][bi]["nomi"]
        data[f"bob{bi + 1}_data"] = {"rejalar": [natija[(bi, ri)] for ri in range(h["rejalar"])]}

    prog(T["format"].format(bet=h["bet"]))
    fname = _out_name(p["mavzu"])
    info = build_docx(data, fname)
    data["bet"] = info["bet"]
    return fname, data


PROG = {
    "uz": {
        "reja": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n🗂 Reja va adabiyotlar ro'yxati tuzilmoqda...",
        "kirish": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n✍️ Kirish yozilmoqda...",
        "bob": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n📖 Paragraflar yozilmoqda: <b>{k}/{n}</b>\n"
               "<i>Jadvallar, grafiklar va snoskalar bilan</i>",
        "xulosa": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n🧾 Xulosa va takliflar yozilmoqda...",
        "format": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n📎 Mundarija, snoskalar va formatlash...",
    },
    "ru": {
        "reja": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n🗂 Составляются план и список литературы...",
        "kirish": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n✍️ Пишется введение...",
        "bob": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n📖 Пишутся параграфы: <b>{k}/{n}</b>\n"
               "<i>С таблицами, графиками и сносками</i>",
        "xulosa": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n🧾 Пишутся заключение и рекомендации...",
        "format": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n📎 Содержание, сноски и оформление...",
    },
    "en": {
        "reja": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n🗂 Building the outline and references...",
        "kirish": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n✍️ Writing the introduction...",
        "bob": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n📖 Writing sections: <b>{k}/{n}</b>\n"
               "<i>With tables, charts and footnotes</i>",
        "xulosa": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n🧾 Writing the conclusion...",
        "format": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n📎 Contents, footnotes and formatting...",
    },
}


def _out_name(mavzu: str) -> str:
    safe = re.sub(r"[^\w\s-]", "", mavzu or "kurs")[:40].strip().replace(" ", "_") or "kurs"
    d = os.path.join(tempfile.gettempdir(), "academic_bot_docs")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"Kurs_ishi_{safe}_{int(datetime.datetime.now().timestamp() * 1000) % 10**9}.docx")


# ═══════════════════════════════════════════════
# WORD HUJJATI
# ═══════════════════════════════════════════════
FONT = "Times New Roman"


def _set_font(run, size=14, bold=None, italic=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    return run


def _style_font(style, size=14, bold=False):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in list(rf.attrib):
        if "Theme" in a or "theme" in a:
            del rf.attrib[a]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def _setup_doc(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(3.0), Cm(1.5)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.0)
    sec.footer_distance = Cm(1.0)
    sec.different_first_page_header_footer = True          # titulda sahifa raqami yo'q

    st = doc.styles
    n = st["Normal"]
    _style_font(n, 14)
    pf = n.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True

    for name, upper in (("Heading 1", True), ("Heading 2", False)):
        s = st[name]
        _style_font(s, 14, bold=True)
        f = s.paragraph_format
        f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f.space_before = Pt(0)
        f.space_after = Pt(12 if upper else 6)
        f.line_spacing = 1.5
        f.keep_with_next = True
        f.first_line_indent = Cm(0)
        f.page_break_before = upper

    for name, ind, bold in (("TOC 1", 0.0, True), ("TOC 2", 0.6, False)):
        try:
            s = st[name]
        except KeyError:
            s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = n
        _style_font(s, 14, bold=bold)
        f = s.paragraph_format
        f.alignment = WD_ALIGN_PARAGRAPH.LEFT
        f.left_indent = Cm(ind)
        f.first_line_indent = Cm(0)
        f.line_spacing = 1.15
        f.space_after = Pt(4 if bold else 2)
        f.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    # Word ochilganda mundarija sahifa raqamlarini yangilash
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    doc.settings.element.append(upd)

    # Sahifa raqami (pastda markazda)
    ft = sec.footer
    fp = ft.paragraphs[0] if ft.paragraphs else ft.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.first_line_indent = Cm(0)
    _field(fp, " PAGE ", "1", size=12)


def _field(paragraph, instr: str, cached: str = "", size=14, bold=False):
    def r(kind=None, text=None, instr_text=None):
        run = paragraph.add_run()
        _set_font(run, size, bold=bold)
        if kind:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            run._r.append(fc)
        if instr_text:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = instr_text
            run._r.append(it)
        if text is not None:
            run.text = text
        return run
    r("begin")
    r(instr_text=instr)
    r("separate")
    r(text=cached)
    r("end")


def _para(doc, text="", size=14, bold=False, italic=False, align=None, indent=True, sb=0, sa=0, ls=None,
          keep=False, left=None):
    p = doc.add_paragraph()
    f = p.paragraph_format
    f.first_line_indent = Cm(1.25) if indent else Cm(0)
    f.space_before, f.space_after = Pt(sb), Pt(sa)
    if ls is not None:
        f.line_spacing = ls
    if left is not None:
        f.left_indent = Cm(left)
    if align is not None:
        p.alignment = align
    if keep:
        f.keep_with_next = True
    if text:
        _set_font(p.add_run(text), size, bold=bold, italic=italic)
    return p


# ── Snoskalar (haqiqiy Word footnote) ──
_FN_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_FN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _fn_root(doc):
    for rel in doc.part.rels.values():
        if rel.reltype == _FN_REL:
            return rel._target._element
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="{_W}">'
           '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r>'
           '</w:p></w:footnote></w:footnotes>').encode()
    part = XmlPart(PackURI("/word/footnotes.xml"), _FN_CT, etree.fromstring(xml), doc.part.package)
    doc.part.relate_to(part, _FN_REL)
    return part._element


def _sup_rpr(size_half=None):
    rpr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    rpr.append(rf)
    va = OxmlElement("w:vertAlign")
    va.set(qn("w:val"), "superscript")
    rpr.append(va)
    if size_half:
        for t in ("w:sz", "w:szCs"):
            e = OxmlElement(t)
            e.set(qn("w:val"), str(size_half))
            rpr.append(e)
    return rpr


def _add_footnote(doc, para, text: str, fid: int):
    ref = OxmlElement("w:r")
    ref.append(_sup_rpr(28))
    fr = OxmlElement("w:footnoteReference")
    fr.set(qn("w:id"), str(fid))
    ref.append(fr)
    para._p.append(ref)

    fn = OxmlElement("w:footnote")
    fn.set(qn("w:id"), str(fid))
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "both")
    ppr.append(jc)
    p.append(ppr)
    r1 = OxmlElement("w:r")
    r1.append(_sup_rpr(20))
    r1.append(OxmlElement("w:footnoteRef"))
    p.append(r1)
    r2 = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    rpr.append(rf)
    for t in ("w:sz", "w:szCs"):
        e = OxmlElement(t)
        e.set(qn("w:val"), "20")
        rpr.append(e)
    r2.append(rpr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = " " + text
    r2.append(t)
    p.append(r2)
    fn.append(p)
    _fn_root(doc).append(fn)


def _text_with_refs(doc, p, text: str, snoskalar: dict, adab: list, fid: list, size=14):
    for i, seg in enumerate(re.split(r"\[(\d+)\]", text)):
        if i % 2 == 0:
            if seg:
                _set_font(p.add_run(seg), size)
        else:
            src = snoskalar.get(seg) or snoskalar.get(str(seg))
            if not src:
                try:
                    src = adab[int(seg) - 1]["matn"]
                except Exception:
                    src = ""
            src = _clean(src)
            if len(src) >= 8:
                _add_footnote(doc, p, src.rstrip(".") + ".", fid[0])
                fid[0] += 1


# ── Jadval ──
def _parse_table_block(block: str, til: str):
    body = block[len("[JADVAL:"):-1] if block.endswith("]") else block[len("[JADVAL:"):]
    lines = [x.strip() for x in body.replace("\\n", "\n").split("\n") if x.strip()]
    if len(lines) < 3:
        return None
    nomi = lines[0]
    manba = ""
    if re.match(r"(?i)^(manba|источник|source)\s*:", lines[-1]):
        manba = lines[-1].split(":", 1)[1].strip()
        lines = lines[:-1]
    m = re.search(r"\((?:manba|источник|source)\s*:\s*([^)]+)\)", nomi, re.I)
    if m:
        manba = manba or m.group(1).strip()
        nomi = nomi.replace(m.group(0), "").strip()
    nomi = re.sub(r"^\s*(\d+\s*-\s*jadval|таблица\s*\d+|table\s*\d+)[\s\.:–-]*", "", nomi, flags=re.I).strip()
    rows = [[c.strip() for c in x.split("|")] for x in lines[1:]]
    ncol = max(len(r) for r in rows)
    if ncol < 2:
        return None
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    return {"nomi": nomi or "-", "manba": manba or L(til)["muallif_ishlanmasi"], "ustunlar": rows[0],
            "qatorlar": rows[1:]}


def _cell_border(tbl):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        x.set(qn("w:val"), "single")
        x.set(qn("w:sz"), "6")
        x.set(qn("w:color"), "000000")
        b.append(x)
    tblPr.append(b)


def _insert_table(doc, t: dict, num: int, til: str):
    lb = L(til)
    _para(doc, lb["jadval"].format(n=num), size=14, italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT,
          indent=False, sb=6, keep=True, ls=1.0)
    _para(doc, t["nomi"], size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, sa=4,
          keep=True, ls=1.0)
    ncol = len(t["ustunlar"])
    tbl = doc.add_table(rows=1 + len(t["qatorlar"]), cols=ncol)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _cell_border(tbl)
    for ci, val in enumerate(t["ustunlar"]):
        c = tbl.rows[0].cells[ci]
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].paragraph_format.first_line_indent = Cm(0)
        c.paragraphs[0].paragraph_format.line_spacing = 1.0
        _set_font(c.paragraphs[0].add_run(str(val)), 12, bold=True)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trPr.append(hdr)                                   # sarlavha qatori har sahifada takrorlanadi
    for ri, row in enumerate(t["qatorlar"]):
        for ci in range(ncol):
            c = tbl.rows[ri + 1].cells[ci]
            pp = c.paragraphs[0]
            pp.paragraph_format.first_line_indent = Cm(0)
            pp.paragraph_format.line_spacing = 1.0
            val = str(row[ci]) if ci < len(row) else ""
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER if (ci > 0 and len(val) < 18) else WD_ALIGN_PARAGRAPH.LEFT
            _set_font(pp.add_run(val), 12)
    _para(doc, f"{lb['manba']}: {t['manba']}", size=12, italic=True, indent=False, sb=2, sa=8, ls=1.0,
          align=WD_ALIGN_PARAGRAPH.LEFT)
    rows_lines = sum(max(1, math.ceil(max(len(str(c)) for c in r) / max(8, 90 // ncol))) for r in
                     [t["ustunlar"]] + t["qatorlar"])
    return 3.0 + rows_lines * 0.75


# ── Grafik ──
def _chart_png(g: dict) -> bytes | None:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
    except Exception as e:
        log.warning(f"matplotlib yo'q: {e}")
        return None
    fam = "DejaVu Serif"
    for cand in ("Times New Roman", "Liberation Serif", "DejaVu Serif"):
        if any(cand == f.name for f in font_manager.fontManager.ttflist):
            fam = cand
            break
    plt.rcParams.update({"font.family": fam, "font.size": 11, "axes.spines.top": False,
                         "axes.spines.right": False})
    labels = [str(a) for a, _ in g["qiymatlar"]]
    vals = [float(b) for _, b in g["qiymatlar"]]
    fig, ax = plt.subplots(figsize=(6.3, 3.5), dpi=200)
    color = "#1f3b63"

    def fmt(v):
        return f"{v:,.1f}".replace(",", " ").rstrip("0").rstrip(".") if abs(v) < 1e6 else f"{v:,.0f}".replace(",", " ")
    if g["tur"] == "doira":
        cols = ["#1f3b63", "#4f6d93", "#8aa1c1", "#c3cfdf", "#6b6b6b", "#a8a8a8", "#d6d6d6", "#2f2f2f",
                "#5a7fa8", "#b0bfd4", "#3d5a80", "#98a9c4"]
        ax.pie(vals, labels=None, autopct="%1.1f%%", startangle=90, colors=cols[:len(vals)],
               wedgeprops={"edgecolor": "white"}, textprops={"color": "white", "fontsize": 9})
        ax.legend(labels, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=9)
        ax.axis("equal")
    elif g["tur"] == "chiziq":
        ax.plot(labels, vals, marker="o", color=color, linewidth=2)
        for x, v in zip(labels, vals):
            ax.annotate(fmt(v), (x, v), textcoords="offset points", xytext=(0, 6), ha="center", fontsize=9)
        ax.grid(axis="y", alpha=0.3)
    else:
        bars = ax.bar(labels, vals, color=color, width=0.6)
        for bb, v in zip(bars, vals):
            ax.annotate(fmt(v), (bb.get_x() + bb.get_width() / 2, bb.get_height()), textcoords="offset points",
                        xytext=(0, 3), ha="center", fontsize=9)
        ax.grid(axis="y", alpha=0.3)
    if g["tur"] != "doira":
        if g.get("birlik"):
            ax.set_ylabel(g["birlik"])
        if max(len(x) for x in labels) > 6:
            plt.setp(ax.get_xticklabels(), rotation=25, ha="right")
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200)
    plt.close(fig)
    return buf.getvalue()


def _insert_chart(doc, g: dict, num: int, til: str) -> float:
    png = _chart_png(g)
    if not png:
        return 0.0
    lb = L(til)
    p = _para(doc, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, sb=6, keep=True, ls=1.0)
    p.add_run().add_picture(io.BytesIO(png), width=Cm(15.0))
    cap = _para(doc, "", indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, ls=1.0, keep=True)
    _set_font(cap.add_run(lb["rasm"].format(n=num) + " "), 14, bold=True)
    _set_font(cap.add_run(g["nomi"]), 14, bold=False)
    _para(doc, f"{lb['manba']}: {g['manba']}", size=12, italic=True, indent=False, sa=8, ls=1.0,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    return 14.0


def _lines(text: str) -> float:
    return max(1, math.ceil((len(text) + 8) / CPL))


def build_docx(data: dict, out_path: str, page_map: dict | None = None) -> dict:
    """data (v2) dan Word hujjatini quradi. page_map — sarlavha → aniq sahifa (bo'lsa)."""
    til = data.get("til", "uz")
    lb = L(til)
    t = data.get("titul", {})
    doc = Document()
    _setup_doc(doc)
    nbob = int(data.get("bob_soni") or sum(1 for k in data if re.fullmatch(r"bob\d+_data", k)))
    adab = data.get("plan", {}).get("adabiyotlar", []) or []
    if adab and isinstance(adab[0], str):
        adab = [{"tur": "asosiy", "matn": re.sub(r"^\d+[\.\)]\s*", "", a)} for a in adab]

    # ── 1. TITUL ──
    yil = datetime.datetime.now().year
    shahar = shahar_aniqla(t.get("otm", ""), til)
    _para(doc, lb["vazirlik"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.15, sa=12)
    _para(doc, (t.get("otm") or "").upper(), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.15,
          sa=4)
    if t.get("fakultet"):
        _para(doc, t["fakultet"], align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.15)
    p = _para(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, sb=150, sa=18, ls=1.0)
    r = p.add_run(lb["kurs_ishi"])
    _set_font(r, 26, bold=True)
    r.font.spacing = Pt(3)
    p = _para(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.15, sa=6)
    _set_font(p.add_run(f"{lb['fan']}: "), 14, bold=True)
    _set_font(p.add_run(t.get("fan", "")), 14)
    p = _para(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.15)
    _set_font(p.add_run(f"{lb['mavzu']}: "), 14, bold=True)
    _set_font(p.add_run(f"«{(t.get('mavzu') or '').strip()}»"), 14, bold=True)
    sb = 120
    for lbl, val in ((lb["bajardi"], t.get("ism")), (lb["guruh"], t.get("guruh")),
                     (lb["rahbar"], t.get("rahbar") or "____________________")):
        if not val:
            continue
        p = _para(doc, "", indent=False, left=8.0, sb=sb, ls=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
        _set_font(p.add_run(f"{lbl}: "), 14, bold=True)
        _set_font(p.add_run(val), 14)
        sb = 4
    # Uzun OTM/mavzu nomi titulni 2-betga surib yubormasligi uchun pastki bo'shliqni kamaytiramiz
    extra = max(0, math.ceil(len(t.get("otm") or "") / 60) - 1) + max(0, math.ceil(len(t.get("mavzu") or "") / 60) - 2) \
        + max(0, math.ceil(len(t.get("fakultet") or "") / 70) - 1) + max(0, math.ceil(len(t.get("rahbar") or "") / 28) - 1)
    p = _para(doc, f"{shahar} – {yil}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False,
              sb=max(20, 150 - 22 * extra), ls=1.0)

    # ── 2. MUNDARIJA (Word maydoni + tayyor ko'rinish) ──
    p = _para(doc, lb["mundarija"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, sa=12)
    p.paragraph_format.page_break_before = True
    toc_entries = [(1, lb["kirish"])]
    for bi in range(nbob):
        toc_entries.append((1, data.get(f"bob{bi + 1}_nomi", f"{ROMAN[bi]}")))
        for rj in data.get(f"bob{bi + 1}_data", {}).get("rejalar", []):
            toc_entries.append((2, _clean(rj.get("sarlavha", ""))))
    toc_entries += [(1, lb["xulosa"]), (1, lb["adab"])]
    toc_paras = []
    for i, (lvl, txt) in enumerate(toc_entries):
        tp = doc.add_paragraph(style="TOC 1" if lvl == 1 else "TOC 2")
        if i == 0:                                    # maydon boshlanishi
            for kind, instr in (("begin", None), (None, ' TOC \\o "1-2" \\h \\z \\u '), ("separate", None)):
                rr = tp.add_run()
                if kind:
                    fc = OxmlElement("w:fldChar")
                    fc.set(qn("w:fldCharType"), kind)
                    if kind == "begin":
                        fc.set(qn("w:dirty"), "true")
                    rr._r.append(fc)
                else:
                    it = OxmlElement("w:instrText")
                    it.set(qn("xml:space"), "preserve")
                    it.text = instr
                    rr._r.append(it)
        _set_font(tp.add_run(txt + "\t"), 14, bold=(lvl == 1))
        toc_paras.append(tp)
    end = toc_paras[-1].add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    end._r.append(fc)

    # Sahifa hisobi (mundarija uchun taxminiy; LibreOffice bo'lsa — aniq)
    cur = [3.0]                                       # 1 — titul, 2 — mundarija
    toc_pages = []

    def new_page():
        cur[0] = float(math.ceil(cur[0] - 1e-6)) if cur[0] % 1 > 0.02 else cur[0]

    def add_lines(n):
        cur[0] += n / LPP

    def heading(text, level):
        if level == 1:
            new_page()
        toc_pages.append(int(cur[0]))
        hp = doc.add_paragraph(style="Heading 1" if level == 1 else "Heading 2")
        _set_font(hp.add_run(text), 14, bold=True)
        add_lines(_lines(text) + (0.6 if level == 1 else 0.4))
        return hp

    fid = [1]
    counters = {"jadval": 0, "rasm": 0}

    # ── 3. KIRISH ──
    heading(lb["kirish"], 1)
    kir = dict(data.get("plan", {}).get("kirish", {}))
    if not kir.get("tuzilishi") or data.get("v") == 2:
        kir["tuzilishi"] = "{TUZILISH}"
    tuz_para = None
    for key in KIRISH_TARTIB + [k for k in kir if k not in KIRISH_TARTIB]:
        val = kir.get(key)
        if not val:
            continue
        label = lb["kirish_lbl"].get(key, "")
        chunks = [c.strip() for c in str(val).split("[ABZAS]") if c.strip()]
        list_mode = key == "vazifalari"
        for ci, ch in enumerate(chunks):
            ch = _strip_refs(_clean(ch))
            if list_mode and ci == 0 and label:
                _para(doc, label, bold=True, keep=True)
                add_lines(1)
            pp = _para(doc)
            if ch == "{TUZILISH}":
                tuz_para = pp
                add_lines(3)
                continue
            if ci == 0 and label and not list_mode:
                _set_font(pp.add_run(label + " "), 14, bold=True)
            _set_font(pp.add_run(ch), 14)
            add_lines(_lines(ch))

    # ── 4. BOBLAR ──
    for bi in range(nbob):
        heading(data.get(f"bob{bi + 1}_nomi", ROMAN[bi]), 1)
        for rj in data.get(f"bob{bi + 1}_data", {}).get("rejalar", []):
            heading(_clean(rj.get("sarlavha", "")), 2)
            matn = str(rj.get("matn", ""))
            sn = {str(k).strip("[] "): v for k, v in (rj.get("snoskalar") or {}).items()}
            parts = [x for x in re.split(r"(\[JADVAL:[^\]]*\]|\[GRAFIK:[^\]]*\])", matn, flags=re.S) if x.strip()]
            # jadval/rasm raqamlarini oldindan belgilaymiz ({JADVAL}/{RASM} havolalari uchun)
            tnums, gnums = [], []
            for x in parts:
                if x.strip().startswith("[JADVAL:"):
                    tnums.append(counters["jadval"] + len(tnums) + 1)
                elif x.strip().startswith("[GRAFIK:"):
                    gnums.append(counters["rasm"] + len(gnums) + 1)
            for x in parts:
                x = x.strip()
                if x.startswith("[JADVAL:"):
                    tb = _parse_table_block(x, til)
                    if tb:
                        counters["jadval"] += 1
                        add_lines(_insert_table(doc, tb, counters["jadval"], til))
                    continue
                if x.startswith("[GRAFIK:"):
                    g = _dec_chart(x[len("[GRAFIK:"):-1])
                    if g and len(g["qiymatlar"]) >= 2:
                        counters["rasm"] += 1
                        h_ = _insert_chart(doc, g, counters["rasm"], til)
                        if not h_:
                            counters["rasm"] -= 1
                        add_lines(h_)
                    continue
                for ab in x.split("[ABZAS]"):
                    ab = _clean(ab)
                    if _words(ab) < 3:
                        continue
                    ab = ab.replace("{JADVAL}", lb["jref"].format(n=tnums[0]) if tnums else "")
                    ab = ab.replace("{RASM}", lb["rref"].format(n=gnums[0]) if gnums else "")
                    ab = re.sub(r"\(\s*\)", "", ab).replace("  ", " ")
                    pp = _para(doc)
                    _text_with_refs(doc, pp, ab, sn, adab, fid)
                    nref = len(re.findall(r"\[\d+\]", ab))
                    add_lines(_lines(ab) + nref * 0.9)

    # ── 5. XULOSA ──
    heading(lb["xulosa"], 1)
    for ch in str(data.get("xulosa_data", {}).get("xulosa", "")).split("[ABZAS]"):
        ch = _strip_refs(_clean(ch))
        if not ch:
            continue
        if ch == lb["takliflar"] or ch.rstrip(":") == lb["takliflar"].rstrip(":"):
            _para(doc, lb["takliflar"], bold=True, keep=True, sb=6)
            add_lines(1.2)
            continue
        _para(doc, ch)
        add_lines(_lines(ch))

    # ── 6. ADABIYOTLAR ──
    heading(lb["adab"], 1)
    n = 1
    gi = 0
    by_group = [g for g in GURUH_TARTIB if any(a.get("tur") == g for a in adab)]
    show_groups = len(by_group) > 1
    for g in GURUH_TARTIB:
        items = [a for a in adab if a.get("tur", "asosiy") == g]
        if not items:
            continue
        if show_groups:
            _para(doc, f"{ROMAN[gi]}. {lb['guruhlar'][g]}", bold=True, indent=False, sb=6, sa=2, keep=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER)
            add_lines(1.3)
            gi += 1
        for a in items:
            pp = _para(doc, f"{n}. {_clean(a['matn'])}", indent=False, left=0.0, ls=1.5)
            pp.paragraph_format.first_line_indent = Cm(0)
            add_lines(_lines(a["matn"]))
            n += 1
    jami_bet = int(math.ceil(cur[0]))

    # ── Mundarija sahifa raqamlari va tuzilish matni ──
    if page_map:
        pages = [page_map.get(i, toc_pages[i] if i < len(toc_pages) else 0) for i in range(len(toc_entries))]
        jami_bet = page_map.get("jami", jami_bet)
    else:
        pages = toc_pages
    for tp, pg in zip(toc_paras, pages):
        _set_font(tp.add_run(str(pg)), 14, bold=False)
    if tuz_para is not None:
        nrej = sum(len(data.get(f"bob{bi + 1}_data", {}).get("rejalar", [])) for bi in range(nbob))
        extra = ""
        if counters["jadval"] or counters["rasm"]:
            extra = lb["tuz_extra"].format(
                j=counters["jadval"], r=lb["tuz_rasm"].format(r=counters["rasm"]) if counters["rasm"] else "")
        _set_font(tuz_para.add_run(lb["kirish_lbl"]["tuzilishi"] + " "), 14, bold=True)
        _set_font(tuz_para.add_run(lb["tuzilish"].format(b=nbob, p=nrej, extra=extra, bet=jami_bet)), 14)

    doc.core_properties.title = t.get("mavzu", "")
    doc.core_properties.author = t.get("ism", "")
    doc.save(out_path)

    info = {"bet": jami_bet, "jadval": counters["jadval"], "rasm": counters["rasm"], "snoska": fid[0] - 1,
            "toc": toc_entries}
    if page_map is None:
        exact = _exact_pages(out_path, toc_entries)
        if exact:
            return build_docx(data, out_path, page_map=exact)
    return info


def _exact_pages(path: str, entries: list) -> dict | None:
    """LibreOffice mavjud bo'lsa — PDFga aylantirib sarlavhalarning aniq sahifasini topadi."""
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if not office or os.getenv("KURS_EXACT_TOC", "1") == "0":
        return None
    try:
        import fitz
    except Exception:
        return None
    tmp = tempfile.mkdtemp(prefix="kurs_toc_")
    try:
        subprocess.run([office, "--headless", "--convert-to", "pdf", "--outdir", tmp, path],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=150, check=False)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        if not os.path.exists(pdf):
            return None
        d = fitz.open(pdf)
        texts = [re.sub(r"\s+", " ", d[i].get_text()) for i in range(d.page_count)]
        res, start = {}, 2                            # mundarijadan keyin qidiramiz
        for i, (_, txt) in enumerate(entries):
            key = re.sub(r"\s+", " ", txt).strip()[:45]
            for pg in range(start, len(texts)):
                if key and key in texts[pg]:
                    res[i] = pg + 1
                    start = pg
                    break
        res["jami"] = d.page_count
        d.close()
        return res if len(res) > len(entries) * 0.7 else None
    except Exception as e:
        log.warning(f"Aniq mundarija hisobi bo'lmadi: {e}")
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

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
    "standart": dict(boblar=2, rejalar=3, sec_words=600, kirish_words=750, xulosa_par=4,
                     takliflar=5, adab=(15, 20), bet="25–30", bet_oraliq=(27, 30), workers=3),
    "katta":    dict(boblar=3, rejalar=3, sec_words=1300, kirish_words=1150, xulosa_par=7,
                     takliflar=7, adab=(25, 35), bet="65–70", bet_oraliq=(66, 70), workers=3),
}

WPP = 215          # o'zbekcha so'z / bet (TNR 14, 1,5 interval, A4 — o'lchangan)

_API: dict = {}


def setup(call_ai_json, clean_text=None, ai_ctx=None):
    """bot.py dan chaqiriladi."""
    _API.update(call_ai_json=call_ai_json, clean_text=clean_text, ai_ctx=ai_ctx)


# ═══════════════════════════════════════════════
# TIL YORLIQLARI
# ═══════════════════════════════════════════════
LBL = {
    "uz": {
        "vazirlik": "O'ZBEKISTON RESPUBLIKASI TA'LIM VA ILM-FAN VAZIRLIGI",
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
        "vazirlik": "МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ РЕСПУБЛИКИ УЗБЕКИСТАН",
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
        "vazirlik": "MINISTRY OF EDUCATION AND SCIENCE OF THE REPUBLIC OF UZBEKISTAN",
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
Havola, jadval va raqamli statistika qo'shma. Yangi abzaslar paragrafning yakuniy xulosasidan OLDIN qo'yiladi —
shuning uchun ular xulosa bilan tugamasin, yangi jihatlar, misollar va tahlil bilan boyitsin.
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
            paras = paras[:-1] + add + paras[-1:]   # yakuniy xulosa abzasi oxirida qoladi
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
        "titul": {k: p.get(k, "") for k in ("otm", "fakultet", "kafedra", "fan", "mavzu", "ism", "guruh", "rahbar",
                                            "shahar")},
        "plan": {"kirish": kirish, "adabiyotlar": adab},
        "xulosa_data": {"xulosa": xulosa},
    }
    for bi in range(h["boblar"]):
        data[f"bob{bi + 1}_nomi"] = L(p["til"])["bob"].format(r=ROMAN[bi]) + " " + plan["boblar"][bi]["nomi"]
        data[f"bob{bi + 1}_data"] = {"rejalar": [natija[(bi, ri)] for ri in range(h["rejalar"])]}

    # 5) Hajm kafolati: sahifalar aniq hisoblanadi; kam bo'lsa — paragraflar to'ldiriladi, ko'p bo'lsa — qisqartiriladi
    _hajmga_moslash(cid, p, plan, data, lambda k, n: prog(T["hajm"].format(bet=h["bet"], k=k, n=n)))

    prog(T["format"].format(bet=h["bet"]))
    fname = _out_name(p["mavzu"])
    info = build_docx(data, fname)
    data["bet"] = info["bet"]
    log.info(f"[{cid}] Kurs ishi tayyor: {info['bet']} bet, {info['jadval']} jadval, {info['rasm']} rasm, "
             f"{info['snoska']} snoska")
    return fname, data


def _sahifalar(data: dict) -> int:
    B, _, _ = _flow(data, "99")
    return SH.paginate(B)[1]


def _hajmga_moslash(cid, p, plan, data, prog=None):
    lo, hi = HAJMLAR[p["hajm"]]["bet_oraliq"]
    maqsad = (lo + hi) // 2
    nbob = data["bob_soni"]
    keys = [(bi, ri) for bi in range(nbob) for ri in range(len(data[f"bob{bi + 1}_data"]["rejalar"]))]

    def bloklar(bi, ri):
        return data[f"bob{bi + 1}_data"]["rejalar"][ri]["matn"].split(" [ABZAS] ")

    def saqla(bi, ri, bl):
        data[f"bob{bi + 1}_data"]["rejalar"][ri]["matn"] = " [ABZAS] ".join(bl)

    def matnli(x):
        return not x.startswith("[JADVAL:") and not x.startswith("[GRAFIK:")

    ctx = _API.get("ai_ctx")
    hammasi = list(keys)
    for urinish in range(4):
        keys = list(hammasi)
        bet = _sahifalar(data)
        log.info(f"[{cid}] Kurs ishi hajmi: {bet} bet (kerak {lo}-{hi}), urinish {urinish + 1}")
        if lo <= bet <= hi:
            return
        if bet > hi:                                  # ortiqcha — o'rta abzaslarni olib tashlaymiz
            ortiq = (bet - hi + 0.5) * WPP
            while ortiq > 0:
                eng = max(keys, key=lambda k: sum(_words(x) for x in bloklar(*k) if matnli(x)))
                bl = bloklar(*eng)
                cand = [i for i in range(1, len(bl) - 1) if matnli(bl[i]) and "{" not in bl[i]]
                cand = [i for i in cand if not re.search(r"\[\d+\]", bl[i])] or cand   # havolalilar — oxirgi navbatda
                if len([x for x in bl if matnli(x)]) <= 4 or not cand:
                    keys = [k for k in keys if k != eng]
                    if not keys:
                        break
                    continue
                i = cand[len(cand) // 2]
                ortiq -= _words(bl[i])
                del bl[i]
                saqla(*eng, bl)
            continue
        # kam — eng qisqa paragraflarni davom ettiramiz (parallel)
        kerak_soz = (maqsad - bet) * WPP
        tartib = sorted(keys, key=lambda k: sum(_words(x) for x in bloklar(*k) if matnli(x)))
        soni = min(len(tartib), max(1, math.ceil(kerak_soz / 450)))
        har = max(2, min(8, math.ceil(kerak_soz / soni / 115)))
        if prog:
            prog(urinish + 1, 4)

        def ish(k):
            bi, ri = k
            mavjud = [x for x in bloklar(bi, ri) if matnli(x)]
            run = lambda: _ask(_davom_prompt(p, plan, bi, ri, mavjud, har), f"{bi + 1}.{ri + 1} to'ldirish",
                               max_tokens=5000)
            try:
                if ctx:
                    with ctx(cid):
                        dv = run()
                else:
                    dv = run()
            except Exception as e:
                log.warning(f"Hajm to'ldirish xatosi {k}: {e}")
                return k, []
            add = [_strip_refs(_clean(x)) for x in (dv.get("paragraflar") or []) if isinstance(x, str)]
            return k, [x for x in add if _words(x) >= 12]

        with cf.ThreadPoolExecutor(max_workers=3) as ex:
            for k, add in ex.map(ish, tartib[:soni]):
                if add:
                    bl = bloklar(*k)
                    saqla(*k, bl[:-1] + add + bl[-1:])     # xulosa abzasi oxirida qoladi


PROG = {
    "uz": {
        "reja": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n🗂 Reja va adabiyotlar ro'yxati tuzilmoqda...",
        "kirish": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n✍️ Kirish yozilmoqda...",
        "bob": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n📖 Paragraflar yozilmoqda: <b>{k}/{n}</b>\n"
               "<i>Jadvallar, grafiklar va snoskalar bilan</i>",
        "xulosa": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n🧾 Xulosa va takliflar yozilmoqda...",
        "hajm": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n📏 Hajm tekshirilmoqda va to'ldirilmoqda ({k}/{n})...",
        "format": "⏳ <b>Kurs ishi yaratilmoqda</b> ({bet} bet)\n\n📎 Mundarija, snoskalar va formatlash...",
    },
    "ru": {
        "reja": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n🗂 Составляются план и список литературы...",
        "kirish": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n✍️ Пишется введение...",
        "bob": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n📖 Пишутся параграфы: <b>{k}/{n}</b>\n"
               "<i>С таблицами, графиками и сносками</i>",
        "xulosa": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n🧾 Пишутся заключение и рекомендации...",
        "hajm": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n📏 Проверка и доведение объёма ({k}/{n})...",
        "format": "⏳ <b>Курсовая работа создаётся</b> ({bet} стр.)\n\n📎 Содержание, сноски и оформление...",
    },
    "en": {
        "reja": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n🗂 Building the outline and references...",
        "kirish": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n✍️ Writing the introduction...",
        "bob": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n📖 Writing sections: <b>{k}/{n}</b>\n"
               "<i>With tables, charts and footnotes</i>",
        "xulosa": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n🧾 Writing the conclusion...",
        "hajm": "⏳ <b>Creating the course work</b> ({bet} pages)\n\n📏 Checking and completing the volume ({k}/{n})...",
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

    # Snoska uslublari (Word va LibreOffice bir xil ko'rsatishi uchun)
    try:
        fs = st["Footnote Text"]
    except KeyError:
        fs = st.add_style("Footnote Text", WD_STYLE_TYPE.PARAGRAPH)
        fs.base_style = n
    _style_font(fs, 10)
    fs.paragraph_format.line_spacing = 1.0
    fs.paragraph_format.space_before = Pt(0)
    fs.paragraph_format.space_after = Pt(0)
    fs.paragraph_format.first_line_indent = Cm(0)
    try:
        fr = st["Footnote Reference"]
    except KeyError:
        fr = st.add_style("Footnote Reference", WD_STYLE_TYPE.CHARACTER)
    fr.font.superscript = True

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
        f.right_indent = Cm(1.0)                     # uzun nom raqam ustuniga kirmasdan ko'chadi
        f.first_line_indent = Cm(0)
        f.line_spacing = 1.15
        f.space_after = Pt(4 if bold else 2)
        f.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


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
           '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="100" '
           'w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:before="0" w:after="0" '
           'w:line="100" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:continuationSeparator/>'
           '</w:r></w:p></w:footnote></w:footnotes>').encode()
    part = XmlPart(PackURI("/word/footnotes.xml"), _FN_CT, etree.fromstring(xml), doc.part.package)
    doc.part.relate_to(part, _FN_REL)
    return part._element


def _sup_rpr(size_half=None):
    rpr = OxmlElement("w:rPr")
    rs = OxmlElement("w:rStyle")
    rs.set(qn("w:val"), "FootnoteReference")
    rpr.append(rs)
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
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), "FootnoteText")
    ppr.append(ps)
    sp = OxmlElement("w:spacing")                     # snoska — yakka interval, 10 pt
    for a, v in (("w:before", "0"), ("w:after", "0"), ("w:line", "240"), ("w:lineRule", "auto")):
        sp.set(qn(a), v)
    ppr.append(sp)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    ppr.append(ind)
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


# ── Grafik ──
_PNG_CACHE: dict = {}


def _chart_png(g: dict) -> bytes | None:
    key = json.dumps(g, ensure_ascii=False, sort_keys=True, default=str)
    if key in _PNG_CACHE:
        return _PNG_CACHE[key]
    png = _chart_png_draw(g)
    if png and len(_PNG_CACHE) < 64:
        _PNG_CACHE[key] = png
    return png


def _chart_png_draw(g: dict) -> bytes | None:
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



# ═══════════════════════════════════════════════
# HUJJAT TUZILMASI → SAHIFALASH → WORD
# ═══════════════════════════════════════════════
import sahifalash as SH

_AL = {"j": WD_ALIGN_PARAGRAPH.JUSTIFY, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT,
       "l": WD_ALIGN_PARAGRAPH.LEFT}


def _P(text="", size=14, bold=False, italic=False, ind=1.25, left=0.0, ls=1.5, sb=0, sa=0, align="j",
       keep=False, pbb=False, segs=None, style=None, **kw):
    b = {"k": "p", "segs": segs if segs is not None else ([("t", text, bold, italic)] if text else []),
         "size": size, "ind": ind, "left": left, "ls": ls, "sb": sb, "sa": sa, "align": align, "keep": keep,
         "pbb": pbb, "style": style}
    b.update(kw)
    return b


def _ref_segs(text: str, snoskalar: dict, adab: list) -> list:
    segs = []
    for i, seg in enumerate(re.split(r"\[(\d+)\]", text)):
        if i % 2 == 0:
            if seg:
                segs.append(("t", seg, False, False))
        else:
            src = snoskalar.get(seg) or snoskalar.get(str(seg))
            if not src:
                try:
                    src = adab[int(seg) - 1]["matn"]
                except Exception:
                    src = ""
            src = _clean(src)
            if len(src) >= 8:
                segs.append(("fn", src.rstrip(".") + "."))
    return segs


def _table_blocks(t: dict, num: int, til: str) -> list:
    lb = L(til)
    return [
        _P(lb["jadval"].format(n=num), italic=True, ind=0, ls=1.0, sb=6, align="r", keep=True),
        _P(t["nomi"], bold=True, ind=0, ls=1.0, sa=4, align="c", keep=True),
        {"k": "tbl", "cols": [str(c) for c in t["ustunlar"]], "rows": [[str(c) for c in r] for r in t["qatorlar"]],
         "size": 12, "hdr_bold": True, "sa": 0},
        _P(f"{lb['manba']}: {t['manba']}", size=12, italic=True, ind=0, ls=1.0, sb=2, sa=8, align="l"),
    ]


def _chart_blocks(g: dict, num: int, til: str) -> list:
    png = _chart_png(g)
    if not png:
        return []
    try:
        from PIL import Image
        w, h = Image.open(io.BytesIO(png)).size
        h_cm = 15.0 * h / w
    except Exception:
        h_cm = 8.33
    lb = L(til)
    return [
        _P("", ind=0, ls=1.0, sb=6, align="c", keep=True, png=png, img_h=h_cm * SH.PT_CM),
        _P(segs=[("t", lb["rasm"].format(n=num) + " ", True, False), ("t", g["nomi"], False, False)],
           ind=0, ls=1.0, sb=4, align="c", keep=True),
        _P(f"{lb['manba']}: {g['manba']}", size=12, italic=True, ind=0, ls=1.0, sa=8, align="c"),
    ]


# ═══════════════════════════════════════════════
# TITUL VARAG'I — OTM namunasi bo'yicha (ro'yxatga olish, taqriz va komissiya qismlari bo'sh qoladi)
# ═══════════════════════════════════════════════
SHAHARLAR = [
    {"uz": "Toshkent", "ru": "Ташкент", "en": "Tashkent"}, {"uz": "Samarqand", "ru": "Самарканд", "en": "Samarkand"},
    {"uz": "Buxoro", "ru": "Бухара", "en": "Bukhara"}, {"uz": "Andijon", "ru": "Андижан", "en": "Andijan"},
    {"uz": "Namangan", "ru": "Наманган", "en": "Namangan"}, {"uz": "Farg'ona", "ru": "Фергана", "en": "Fergana"},
    {"uz": "Qarshi", "ru": "Карши", "en": "Karshi"}, {"uz": "Termiz", "ru": "Термез", "en": "Termez"},
    {"uz": "Nukus", "ru": "Нукус", "en": "Nukus"}, {"uz": "Urganch", "ru": "Ургенч", "en": "Urgench"},
    {"uz": "Jizzax", "ru": "Джизак", "en": "Jizzakh"}, {"uz": "Guliston", "ru": "Гулистан", "en": "Gulistan"},
    {"uz": "Navoiy", "ru": "Навои", "en": "Navoi"}, {"uz": "Qo'qon", "ru": "Коканд", "en": "Kokand"},
    {"uz": "Chirchiq", "ru": "Чирчик", "en": "Chirchik"}, {"uz": "Xiva", "ru": "Хива", "en": "Khiva"},
]

TITUL = {
    "uz": dict(fak="fakulteti", kaf="kafedrasi", royxat="Ro'yxatga olindi №__________",
               sana="“_____” ____________ 20___ y.", kaf_line="“{x}” KAFEDRASI", fan_line="“{x}” FANIDAN",
               kurs="KURS ISHI", mavzu="MAVZU:", bajardi="Bajardi:", fak2="fakulteti", talaba="guruh talabasi",
               tekshirdi="Tekshirdi:", t_top="Kurs ishi taqrizga\ntopshirilgan sana\n“____” ________ 20___ y.",
               t_qayt="Kurs ishi taqrizdan\nqaytarilgan sana\n“____” ________ 20___ y.",
               t_himoya="Kurs ishi himoya\nqilingan sana\n“____” ________ 20___ y.\n\nBaho “____” ________",
               imzo="(imzo)", komissiya="Komissiya a'zolari:"),
    "ru": dict(fak="факультет", kaf="кафедра", royxat="Зарегистрировано №__________",
               sana="«_____» ____________ 20___ г.", kaf_line="КАФЕДРА «{x}»", fan_line="ПО ДИСЦИПЛИНЕ «{x}»",
               kurs="КУРСОВАЯ РАБОТА", mavzu="ТЕМА:", bajardi="Выполнил(а):", fak2="факультет",
               talaba="студент(ка) группы", tekshirdi="Проверил(а):",
               t_top="Сдана на рецензию\n«____» ________ 20___ г.", t_qayt="Возвращена с рецензии\n«____» ________ 20___ г.",
               t_himoya="Дата защиты\n«____» ________ 20___ г.\n\nОценка «____» ________", imzo="(подпись)",
               komissiya="Члены комиссии:"),
    "en": dict(fak="faculty", kaf="department", royxat="Registered №__________",
               sana="“_____” ____________ 20___", kaf_line="DEPARTMENT OF {x}", fan_line="SUBJECT: {x}",
               kurs="COURSE WORK", mavzu="TOPIC:", bajardi="Prepared by:", fak2="faculty", talaba="group student",
               tekshirdi="Checked by:", t_top="Submitted for review\n“____” ________ 20___",
               t_qayt="Returned from review\n“____” ________ 20___",
               t_himoya="Defended on\n“____” ________ 20___\n\nGrade “____” ________", imzo="(signature)",
               komissiya="Committee members:"),
}
_BLANK = "____________________"


def _nom_tozala(x: str, sozlar: tuple) -> str:
    """«Iqtisodiyot fakulteti» → «Iqtisodiyot» (titulda «... fakulteti» so'zi alohida turadi)."""
    x = _clean(x or "")
    for s in sozlar:
        x = re.sub(rf"\s*\b{s}\b\.?$", "", x, flags=re.I)
    return x.strip(" ,.«»\"“”")


def _tp(cell_or_doc, text="", size=14, bold=True, align="c", sb=0, sa=0, ls=1.0, left=None, italic=False, new=True):
    p = cell_or_doc.add_paragraph() if new else cell_or_doc.paragraphs[0]
    f = p.paragraph_format
    f.first_line_indent = Cm(0)
    f.space_before, f.space_after = Pt(sb), Pt(sa)
    f.line_spacing = ls
    f.keep_together = True
    if left is not None:
        f.left_indent = Cm(left)
    p.alignment = _AL[align]
    if text:
        _set_font(p.add_run(text), size, bold=bold, italic=italic)
    return p


def _no_borders(tbl):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        x.set(qn("w:val"), "nil")
        b.append(x)
    tblPr.append(b)


def _render_titul(doc, t: dict, til: str):
    lb, T = L(til), TITUL.get(til, TITUL["uz"])
    fak = _nom_tozala(t.get("fakultet"), ("fakulteti", "fakultet", "факультет", "faculty", "faculty of"))
    kaf = _nom_tozala(t.get("kafedra"), ("kafedrasi", "kafedra", "кафедра", "department", "department of"))
    otm = _clean(t.get("otm") or "").upper() or "______________________________ UNIVERSITETI"
    fan = _clean(t.get("fan") or "")
    shahar = _clean(t.get("shahar") or "") or shahar_aniqla(t.get("otm", ""), til)
    yil = datetime.datetime.now().year
    uzun = len(otm) > 70 or len(t.get("mavzu") or "") > 110

    _tp(doc, lb["vazirlik"], 16 if not uzun else 15, new=False if doc.paragraphs else True, sa=6)
    _tp(doc, otm, 16 if not uzun else 15, sa=10)

    # Fakultet | kafedra, ro'yxatga olish (chap va o'ng ustun)
    tb = doc.add_table(rows=1, cols=2)
    _no_borders(tb)
    tb.autofit = False
    for ci, (nom, soz) in enumerate(((fak, T["fak"]), (kaf, T["kaf"]))):
        c = tb.rows[0].cells[ci]
        c.width = Cm(8.25)
        al = "l" if ci == 0 else "r"
        _tp(c, (nom if nom else _BLANK) + f" {soz}", 14, align=al, sa=6, new=False)
        _tp(c, T["royxat"], 14, align=al, sa=2)
        _tp(c, T["sana"], 14, align=al)

    _tp(doc, T["kaf_line"].format(x=(kaf.upper() if kaf else "_" * 28)), 16, sb=30 if not uzun else 16)
    _tp(doc, T["fan_line"].format(x=(fan.upper() if fan else "_" * 28)), 16, sb=4)
    _tp(doc, T["kurs"], 28, sb=40 if not uzun else 18, sa=12)
    p = _tp(doc, "", 14, ls=1.15, sa=4)
    _set_font(p.add_run(T["mavzu"] + " "), 14, bold=True)
    _set_font(p.add_run((t.get("mavzu") or "").strip().upper()), 14, bold=True)

    # Talaba va rahbar (o'ng tomonda)
    guruh = _clean(t.get("guruh") or "")
    rows = [(T["bajardi"], _clean(t.get("ism") or "") or _BLANK),
            ("", f"{fak if fak else _BLANK} {T['fak2']}"),
            ("", f"{guruh if guruh else '__________'} {T['talaba']}"),
            (T["tekshirdi"], _clean(t.get("rahbar") or "") or _BLANK)]
    for i, (lbl, val) in enumerate(rows):
        p = _tp(doc, "", 14, align="l", left=8.0, sb=(34 if not uzun else 14) if i == 0 else 3)
        if lbl:
            _set_font(p.add_run(lbl + " "), 14, bold=True)
        _set_font(p.add_run(val), 14, bold=False if lbl else True)

    # Taqriz va himoya jadvali (komissiya to'ldiradi — bo'sh)
    _tp(doc, "", 6, sb=28 if not uzun else 8)
    tb = doc.add_table(rows=2, cols=3)
    tb.autofit = False
    _cell_border(tb)
    w = (Cm(5.6), Cm(4.9), Cm(6.0))
    cells = [[T["t_top"], "", T["t_qayt"]],
             [T["t_himoya"], "\n".join([f"__________\n{T['imzo']}"] * 3), T["komissiya"] + "\n\n" + "\n\n".join(["__________________"] * 3)]]
    for ri in range(2):
        for ci in range(3):
            c = tb.rows[ri].cells[ci]
            c.width = w[ci]
            lines = cells[ri][ci].split("\n")
            for li, ln in enumerate(lines):
                _tp(c, ln, 11, bold=False, align="c" if ci == 1 else "l", new=(li > 0))

    _tp(doc, f"{shahar} – {yil}", 14, sb=40 if not uzun else 12)


def titul_sigadimi(t: dict, til: str = "uz") -> bool:
    """Titul bir betga sig'ishini taxminiy tekshirish (juda uzun nomlar uchun)."""
    return len(t.get("otm") or "") < 160 and len(t.get("mavzu") or "") < 260


def _flow(data: dict, bet_matn: str = "99") -> tuple[list, list, dict]:
    """Hujjatni abstrakt bloklarga aylantiradi. Qaytaradi: bloklar, mundarija yozuvlari, hisoblagichlar."""
    til = data.get("til", "uz")
    lb = L(til)
    t = data.get("titul", {})
    nbob = int(data.get("bob_soni") or sum(1 for k in data if re.fullmatch(r"bob\d+_data", k)))
    adab = data.get("plan", {}).get("adabiyotlar", []) or []
    if adab and isinstance(adab[0], str):
        adab = [{"tur": "asosiy", "matn": re.sub(r"^\d+[\.\)]\s*", "", a)} for a in adab]
    B: list = []

    # ── 1. Titul (alohida chiziladi, 1-bet) ──
    B.append({"k": "titul", "t": t, "til": til})

    # ── 2. Mundarija ──
    B.append(_P(lb["mundarija"], bold=True, ind=0, sa=12, align="c", pbb=True))
    toc = [(1, lb["kirish"])]
    for bi in range(nbob):
        toc.append((1, data.get(f"bob{bi + 1}_nomi", ROMAN[bi])))
        for rj in data.get(f"bob{bi + 1}_data", {}).get("rejalar", []):
            toc.append((2, _clean(rj.get("sarlavha", ""))))
    toc += [(1, lb["xulosa"]), (1, lb["adab"])]
    toc_idx = []
    for i, (lvl, txt) in enumerate(toc):
        toc_idx.append(len(B))
        B.append(_P(segs=[("t", txt, lvl == 1, False)], ind=0, left=0.0 if lvl == 1 else 0.6, right=1.0, ls=1.15,
                    sa=4 if lvl == 1 else 2, align="l", style="TOC 1" if lvl == 1 else "TOC 2",
                    toc_text=txt, toc_first=(i == 0), toc_last=(i == len(toc) - 1)))
    heads = []                                     # mundarija tartibida sarlavha bloklari

    def H(text, level):
        heads.append(len(B))
        B.append(_P(text, bold=True, ind=0, ls=1.5, sa=12 if level == 1 else 6, align="c", keep=True,
                    pbb=(level == 1), style="Heading 1" if level == 1 else "Heading 2"))

    cnt = {"jadval": 0, "rasm": 0}

    # ── 3. Kirish ──
    H(lb["kirish"], 1)
    kir = dict(data.get("plan", {}).get("kirish", {}))
    kir["tuzilishi"] = "{TUZILISH}"
    tuz_i = None
    for key in KIRISH_TARTIB + [k for k in kir if k not in KIRISH_TARTIB]:
        val = kir.get(key)
        if not val:
            continue
        label = lb["kirish_lbl"].get(key, "")
        chunks = [c.strip() for c in str(val).split("[ABZAS]") if c.strip()]
        for ci, ch in enumerate(chunks):
            ch = _strip_refs(_clean(ch))
            if ch == "{TUZILISH}":
                tuz_i = len(B)
                B.append(_P(segs=[("t", label + " ", True, False), ("t", "{TUZILISH}", False, False)]))
                continue
            if key == "vazifalari":
                if ci == 0 and label:
                    B.append(_P(label, bold=True, keep=True))
                B.append(_P(ch))
                continue
            segs = ([("t", label + " ", True, False)] if (ci == 0 and label) else []) + [("t", ch, False, False)]
            B.append(_P(segs=segs))

    # ── 4. Boblar ──
    for bi in range(nbob):
        H(data.get(f"bob{bi + 1}_nomi", ROMAN[bi]), 1)
        for rj in data.get(f"bob{bi + 1}_data", {}).get("rejalar", []):
            H(_clean(rj.get("sarlavha", "")), 2)
            matn = str(rj.get("matn", ""))
            sn = {str(k).strip("[] "): v for k, v in (rj.get("snoskalar") or {}).items()}
            parts = [x for x in re.split(r"(\[JADVAL:[^\]]*\]|\[GRAFIK:[^\]]*\])", matn, flags=re.S) if x.strip()]
            tnums, gnums = [], []
            for x in parts:
                if x.strip().startswith("[JADVAL:"):
                    tnums.append(cnt["jadval"] + len(tnums) + 1)
                elif x.strip().startswith("[GRAFIK:"):
                    gnums.append(cnt["rasm"] + len(gnums) + 1)
            for x in parts:
                x = x.strip()
                if x.startswith("[JADVAL:"):
                    tb = _parse_table_block(x, til)
                    if tb:
                        cnt["jadval"] += 1
                        B.extend(_table_blocks(tb, cnt["jadval"], til))
                    continue
                if x.startswith("[GRAFIK:"):
                    g = _dec_chart(x[len("[GRAFIK:"):-1])
                    if g and len(g["qiymatlar"]) >= 2:
                        cb = _chart_blocks(g, cnt["rasm"] + 1, til)
                        if cb:
                            cnt["rasm"] += 1
                            B.extend(cb)
                    continue
                for ab in x.split("[ABZAS]"):
                    ab = _clean(ab)
                    if _words(ab) < 3:
                        continue
                    ab = ab.replace("{JADVAL}", lb["jref"].format(n=tnums[0]) if tnums else "")
                    ab = ab.replace("{RASM}", lb["rref"].format(n=gnums[0]) if gnums else "")
                    ab = re.sub(r"\(\s*\)", "", ab).replace("  ", " ").strip()
                    B.append(_P(segs=_ref_segs(ab, sn, adab)))

    # ── 5. Xulosa ──
    H(lb["xulosa"], 1)
    for ch in str(data.get("xulosa_data", {}).get("xulosa", "")).split("[ABZAS]"):
        ch = _strip_refs(_clean(ch))
        if not ch:
            continue
        if ch.rstrip(":") == lb["takliflar"].rstrip(":"):
            B.append(_P(lb["takliflar"], bold=True, keep=True, sb=6))
        else:
            B.append(_P(ch))

    # ── 6. Adabiyotlar ──
    H(lb["adab"], 1)
    groups = [g for g in GURUH_TARTIB if any(a.get("tur", "asosiy") == g for a in adab)]
    n = 1
    for gi, g in enumerate(groups):
        if len(groups) > 1:
            B.append(_P(f"{ROMAN[gi]}. {lb['guruhlar'][g]}", bold=True, ind=0, sb=6, sa=2, align="c", keep=True))
        for a in [a for a in adab if a.get("tur", "asosiy") == g]:
            B.append(_P(f"{n}. {_clean(a['matn'])}", ind=0, align="j"))
            n += 1

    # tuzilish matni (sahifa soni keyin qo'yiladi)
    if tuz_i is not None:
        nrej = sum(len(data.get(f"bob{bi + 1}_data", {}).get("rejalar", [])) for bi in range(nbob))
        ex = ""
        if cnt["jadval"] or cnt["rasm"]:
            ex = lb["tuz_extra"].format(j=cnt["jadval"], r=lb["tuz_rasm"].format(r=cnt["rasm"]) if cnt["rasm"] else "")
        B[tuz_i]["segs"][1] = ("t", lb["tuzilish"].format(b=nbob, p=nrej, extra=ex, bet=bet_matn), False, False)
        B[tuz_i]["tuz_tmpl"] = (lb["tuzilish"], nbob, nrej, ex)
    meta = {"toc": toc, "toc_idx": toc_idx, "heads": heads, "cnt": cnt, "tuz_i": tuz_i}
    return B, toc, meta


def _render(B: list, meta: dict, page_of: list, out_path: str, title: str = "", author: str = ""):
    doc = Document()
    _setup_doc(doc)
    fid = [1]
    toc_pages = [page_of[h] for h in meta["heads"]]
    ti = 0
    for b in B:
        if b["k"] == "titul":
            _render_titul(doc, b["t"], b["til"])
            continue
        if b["k"] == "tbl":
            _render_table(doc, b)
            continue
        if b.get("style"):
            p = doc.add_paragraph(style=b["style"])
        else:
            p = doc.add_paragraph()
        f = p.paragraph_format
        if not b.get("style") or b["style"].startswith("TOC") is False and not b["style"].startswith("Heading"):
            pass
        if not (b.get("style") or "").startswith(("Heading", "TOC")):
            f.first_line_indent = Cm(b.get("ind", 0))
            f.left_indent = Cm(b.get("left", 0))
            f.line_spacing = b.get("ls", 1.5)
            f.space_before = Pt(b.get("sb", 0))
            f.space_after = Pt(b.get("sa", 0))
            p.alignment = _AL[b.get("align", "j")]
            f.keep_with_next = bool(b.get("keep"))
            f.page_break_before = bool(b.get("pbb"))
        if b.get("png"):
            p.add_run().add_picture(io.BytesIO(b["png"]), width=Cm(15.0))
            continue
        if b.get("toc_text") is not None:
            lvl1 = b["style"] == "TOC 1"
            if b.get("toc_first"):
                for kind, instr in (("begin", None), (None, ' TOC \\o "1-2" \\h \\z \\u '), ("separate", None)):
                    rr = p.add_run()
                    if kind:
                        fc = OxmlElement("w:fldChar")
                        fc.set(qn("w:fldCharType"), kind)
                        rr._r.append(fc)
                    else:
                        it = OxmlElement("w:instrText")
                        it.set(qn("xml:space"), "preserve")
                        it.text = instr
                        rr._r.append(it)
            _set_font(p.add_run(b["toc_text"] + "\t"), 14, bold=lvl1)
            _set_font(p.add_run(str(toc_pages[ti])), 14, bold=False)
            ti += 1
            if b.get("toc_last"):
                rr = p.add_run()
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), "end")
                rr._r.append(fc)
            continue
        size = b.get("size", 14)
        for s in b["segs"]:
            if s[0] == "fn":
                _add_footnote(doc, p, s[1], fid[0])
                fid[0] += 1
            else:
                r = _set_font(p.add_run(s[1]), size, bold=s[2], italic=s[3])
                if b.get("spacing"):
                    r.font.spacing = Pt(b["spacing"])
    doc.core_properties.title = title
    doc.core_properties.author = author
    doc.save(out_path)
    return fid[0] - 1


def _render_table(doc, b: dict):
    ncol = len(b["cols"])
    tbl = doc.add_table(rows=1 + len(b["rows"]), cols=ncol)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    _cell_border(tbl)
    cw = Cm(16.5 / ncol)
    for row in tbl.rows:
        for c in row.cells:
            c.width = cw
    for ci, val in enumerate(b["cols"]):
        c = tbl.rows[0].cells[ci]
        pp = c.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.first_line_indent = Cm(0)
        pp.paragraph_format.line_spacing = 1.0
        _set_font(pp.add_run(val), b.get("size", 12), bold=True)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trPr.append(hdr)
    for ri, row in enumerate(b["rows"]):
        cs = OxmlElement("w:cantSplit")                     # qator sahifalar orasida bo'linmaydi
        tbl.rows[ri + 1]._tr.get_or_add_trPr().append(cs)
        for ci in range(ncol):
            pp = tbl.rows[ri + 1].cells[ci].paragraphs[0]
            pp.paragraph_format.first_line_indent = Cm(0)
            pp.paragraph_format.line_spacing = 1.0
            val = row[ci] if ci < len(row) else ""
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER if (ci > 0 and len(val) < 18) else WD_ALIGN_PARAGRAPH.LEFT
            _set_font(pp.add_run(val), b.get("size", 12))


def build_docx(data: dict, out_path: str, page_map: dict | None = None) -> dict:
    """data (v2) dan Word hujjati: avval sahifalar aniq hisoblanadi, keyin mundarija raqamlari bilan yoziladi."""
    t = data.get("titul", {})
    B, toc, meta = _flow(data, "99")
    page_of, jami = SH.paginate(B)
    if meta["tuz_i"] is not None:                    # «umumiy hajmi N bet» — haqiqiy son bilan qayta hisob
        tmpl, nb, nr, ex = B[meta["tuz_i"]]["tuz_tmpl"]
        for _ in range(2):
            B[meta["tuz_i"]]["segs"][1] = ("t", tmpl.format(b=nb, p=nr, extra=ex, bet=jami), False, False)
            page_of, jami2 = SH.paginate(B)
            if jami2 == jami:
                break
            jami = jami2
    fn = _render(B, meta, page_of, out_path, title=t.get("mavzu", ""), author=t.get("ism", ""))
    return {"bet": jami, "jadval": meta["cnt"]["jadval"], "rasm": meta["cnt"]["rasm"], "snoska": fn,
            "toc": toc, "toc_pages": [page_of[h] for h in meta["heads"]]}

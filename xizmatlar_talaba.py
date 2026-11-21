"""
xizmatlar_talaba.py — bitiruvchi va amaliyotchi talabalar uchun xizmatlar
═══════════════════════════════════════════════════════════════════
  🎓 BMI (bitiruv malakaviy ishi) — bot.py dagi kurs ishi oqimi «bmi» rejimida (kurs_ishi.py), bu yerda
                                     faqat bosh menyu tugmasi ulanadi
  🛡 Himoyaga tayyorlov            — kurs ishi / BMI (botdagi oxirgi ish, Word/PDF fayl yoki matn) asosida:
                                     himoya taqdimoti (PPTX, 9-10 slayd), 7-9 daqiqalik himoya nutqi,
                                     15 ta ehtimoliy komissiya savoli va javoblari, himoya bo'yicha maslahatlar (Word)
  📋 Amaliyot hisoboti             — o'quv / ishlab chiqarish / pedagogik / malakaviy / bitiruv oldi amaliyoti:
                                     titul, amaliyot kundaligi (ish kunlari bo'yicha jadval), hisobot (kirish,
                                     tashkilot tavsifi, bajarilgan ishlar, ko'nikmalar, muammo va takliflar,
                                     xulosa), tavsifnoma loyihasi — imzo, baho va muhr joylari bo'sh

Namunalar: VM 13.12.2024 dagi 836-son qarori (yakuniy attestatsiya), OTMlarning BMI va amaliyot bo'yicha
uslubiy ko'rsatmalari (tuzilma, himoya nutqi 7-10 daqiqa, baholash mezonlari), tavsifnoma namunalari.
"""
from __future__ import annotations

import datetime
import logging
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup

import kurs_ishi as K
import pptx_themes
import security
import xizmatlar_yangi as X

log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════
# NARX VA TUGMALAR
# ═══════════════════════════════════════════════════════════════════
X.NARX.update({"bmi": 35, "himoya": 8, "amaliyot": 8})
for _til, _b in {
    "uz": {"bmi": "🎓 BMI (bitiruv ishi)", "himoya": "🛡 Himoyaga tayyorlov", "amaliyot": "📋 Amaliyot hisoboti"},
    "ru": {"bmi": "🎓 ВКР (дипломная)", "himoya": "🛡 Подготовка к защите", "amaliyot": "📋 Отчёт по практике"},
    "en": {"bmi": "🎓 Graduation thesis", "himoya": "🛡 Defense prep", "amaliyot": "📋 Internship report"},
}.items():
    X.BTN[_til].update(_b)

_TX = {
    "uz": {
        "hm_start": ("🛡 <b>Himoyaga tayyorlov</b> — {narx} 🪙\n\n"
                     "Kurs ishi yoki bitiruv malakaviy ishingiz asosida tayyorlab beraman:\n"
                     "• 📊 <b>Himoya taqdimoti</b> (PPTX, 9-10 slayd): dolzarblik → maqsad va vazifalar → obyekt, "
                     "predmet, metodlar → natijalar (ishingizdagi jadval/grafik bilan) → takliflar → xulosa\n"
                     "• 🎤 <b>Himoya nutqi</b> 7-9 daqiqaga — har slayd uchun alohida\n"
                     "• ❓ <b>15 ta ehtimoliy komissiya savoli</b> va tayyor javoblar\n"
                     "• ✅ Himoyada nimalar baholanishi va maslahatlar\n\n"
                     "Ishingizni yuboring: 📄 <b>Word</b> yoki 📕 <b>PDF</b> fayl (yoki matnini).{last}"),
        "hm_last": "\n\nYoki botda yaratilgan oxirgi ishingizni tanlang 👇",
        "hm_last_btn": "📄 Oxirgi ishim: {m}",
        "hm_ism": ("👤 Taqdimotning birinchi slaydi uchun <b>familiya va ismingizni</b> yozing:\n"
                   "<i>Masalan: Karimova Dilnoza</i>"),
        "hm_short": "⚠️ Matn juda qisqa. Ishingizni Word/PDF fayl ko'rinishida yuboring (kamida 1-2 bet).",
        "hm_wait": "⏳ <b>Himoya materiallari tayyorlanmoqda...</b>\n\n📖 Ish o'qilmoqda va tahlil qilinmoqda...",
        "hm_prog2": "⏳ <b>Himoya materiallari tayyorlanmoqda...</b>\n\n📊 Taqdimot va nutq yozilmoqda...",
        "hm_done_ppt": "📊 <b>Himoya taqdimoti</b> — «{m}»",
        "hm_done_doc": ("🎤 <b>Himoya nutqi, komissiya savollari va maslahatlar</b>\n\n"
                        "💡 Nutqni 2-3 marta ovoz chiqarib, soat bilan mashq qiling. Slayddagi raqamlar ishingizdagi "
                        "bilan bir xil ekanini tekshiring."),
        "am_start": ("📋 <b>Amaliyot hisoboti</b> — {narx} 🪙\n\n"
                     "OTM talablari bo'yicha to'liq hujjatlar to'plami (Word):\n"
                     "• 🏛 Titul varag'i\n• 📅 <b>Amaliyot kundaligi</b> — har ish kuni uchun sana va bajarilgan ishlar "
                     "(rahbar imzosi ustuni bilan)\n• 📝 <b>Hisobot</b>: kirish, tashkilot tavsifi, bajarilgan ishlar, "
                     "egallangan ko'nikmalar, muammolar va takliflar, xulosa\n• 🧾 <b>Tavsifnoma</b> loyihasi "
                     "(tashkilot rahbari imzolaydi)\n\n<b>Amaliyot turini tanlang:</b>"),
        "am_otm": ("✅ {v}\n\n🏛 <b>OTM va fakultet</b>ni yozing (2 qatorda):\n"
                   "<code>Samarqand davlat universiteti\nFilologiya fakulteti</code>"),
        "am_talaba": ("👤 <b>Talaba ma'lumotlari</b> (3 qatorda): F.I.Sh., ta'lim yo'nalishi, kurs va guruh\n"
                      "<code>Karimova Dilnoza Akmalovna\n60111800 – Xorijiy til va adabiyoti (ingliz tili)\n3-kurs, 301-guruh</code>"),
        "am_joy": ("🏢 <b>Amaliyot joyi</b>: tashkilot nomi (va bo'lim, shahar):\n"
                   "<code>Samarqand shahar 12-umumiy o'rta ta'lim maktabi</code>\n"
                   "<code>«Hamkorbank» ATB Samarqand filiali, kredit bo'limi</code>"),
        "am_muddat": ("📅 <b>Amaliyot muddati</b> (boshlanish va tugash sanasi):\n"
                      "<code>02.02.2026 – 28.02.2026</code>"),
        "am_muddat_bad": ("⚠️ Sanani tushunmadim. Shunday yozing: <code>02.02.2026 – 28.02.2026</code>\n"
                          "(muddat 3 kundan 120 kungacha)"),
        "am_rahbar": ("👨‍🏫 <b>Amaliyot rahbarlari</b> (2 qatorda): OTMdan va tashkilotdan\n"
                      "<code>dots. Rahimov A.B.\nmaktab direktori o'rinbosari Toshev S.</code>\n"
                      "Bilmasangiz — «O'tkazib yuborish» (bo'sh chiziq qoladi)."),
        "am_qosh": ("✍️ <b>Amaliyotda aynan nimalar qildingiz?</b> (ixtiyoriy, lekin hisobot ancha aniq chiqadi)\n"
                    "<i>Masalan: 5-«A» sinfda 6 ta dars o'tdim, 2 ta tarbiyaviy soat, ota-onalar majlisida qatnashdim, "
                    "hujjatlar bilan ishladim...</i>\n\nYoki «O'tkazib yuborish»."),
        "am_wait": "⏳ <b>Amaliyot hisoboti tayyorlanmoqda...</b>\n\n📅 Kundalik va hisobot yozilmoqda...",
        "am_done": ("✅ <b>Amaliyot hujjatlari tayyor!</b>\n\n📅 Kundalik: <b>{n} ish kuni</b>\n"
                    "✍️ Imzo, baho va muhr joylari bo'sh qoldirildi. Kundalikdagi ishlarni haqiqatda bajarganlaringizga "
                    "moslab, kerak bo'lsa tahrirlang."),
        "skip": "⏭ O'tkazib yuborish",
        "turlar": {"oquv": "📘 O'quv amaliyoti", "ishlab": "🏭 Ishlab chiqarish amaliyoti",
                   "pedagogik": "🏫 Pedagogik amaliyot", "malakaviy": "🎯 Malakaviy amaliyot",
                   "bitiruv": "🎓 Bitiruv oldi amaliyoti"},
    },
    "ru": {
        "hm_start": ("🛡 <b>Подготовка к защите</b> — {narx} 🪙\n\n"
                     "На основе вашей курсовой или ВКР подготовлю:\n"
                     "• 📊 <b>Презентацию для защиты</b> (PPTX, 9-10 слайдов): актуальность → цель и задачи → объект, "
                     "предмет, методы → результаты (с таблицами/графиками из работы) → рекомендации → заключение\n"
                     "• 🎤 <b>Речь на защиту</b> на 7-9 минут — к каждому слайду\n"
                     "• ❓ <b>15 вероятных вопросов комиссии</b> с ответами\n"
                     "• ✅ Критерии оценки и советы\n\n"
                     "Отправьте работу: 📄 <b>Word</b> или 📕 <b>PDF</b> (или текст).{last}"),
        "hm_last": "\n\nИли выберите последнюю работу, созданную в боте 👇",
        "hm_last_btn": "📄 Последняя работа: {m}",
        "hm_ism": "👤 <b>Фамилия и имя</b> для первого слайда:\n<i>Например: Иванова Анна</i>",
        "hm_short": "⚠️ Слишком короткий текст. Отправьте работу файлом Word/PDF.",
        "hm_wait": "⏳ <b>Готовим материалы к защите...</b>\n\n📖 Читаем и анализируем работу...",
        "hm_prog2": "⏳ <b>Готовим материалы к защите...</b>\n\n📊 Пишутся презентация и речь...",
        "hm_done_ppt": "📊 <b>Презентация для защиты</b> — «{m}»",
        "hm_done_doc": ("🎤 <b>Речь на защиту, вопросы комиссии и советы</b>\n\n"
                        "💡 Прорепетируйте речь вслух 2-3 раза с таймером. Проверьте, что цифры на слайдах совпадают с работой."),
        "am_start": ("📋 <b>Отчёт по практике</b> — {narx} 🪙\n\n"
                     "Полный комплект документов (Word):\n• 🏛 Титульный лист\n• 📅 <b>Дневник практики</b> — дата и "
                     "выполненные работы по каждому рабочему дню (с графой для подписи руководителя)\n"
                     "• 📝 <b>Отчёт</b>: введение, характеристика организации, выполненные работы, навыки, проблемы и "
                     "предложения, заключение\n• 🧾 Проект <b>характеристики</b> (подписывает руководитель)\n\n"
                     "<b>Выберите вид практики:</b>"),
        "am_otm": "✅ {v}\n\n🏛 <b>Вуз и факультет</b> (2 строки):\n<code>Самаркандский государственный университет\nФилологический факультет</code>",
        "am_talaba": ("👤 <b>Данные студента</b> (3 строки): Ф.И.О., направление, курс и группа\n"
                      "<code>Иванова Анна Сергеевна\n60111800 – Иностранный язык и литература\n3 курс, группа 301</code>"),
        "am_joy": "🏢 <b>Место практики</b>: организация (отдел, город):\n<code>АКБ «Хамкорбанк», Самаркандский филиал, кредитный отдел</code>",
        "am_muddat": "📅 <b>Сроки практики</b>:\n<code>02.02.2026 – 28.02.2026</code>",
        "am_muddat_bad": "⚠️ Не понял даты. Пример: <code>02.02.2026 – 28.02.2026</code> (от 3 до 120 дней)",
        "am_rahbar": ("👨‍🏫 <b>Руководители практики</b> (2 строки): от вуза и от организации\n"
                      "<code>доц. Рахимов А.Б.\nначальник отдела Тошев С.</code>\nНе знаете — «Пропустить»."),
        "am_qosh": ("✍️ <b>Что именно вы делали на практике?</b> (необязательно, но отчёт будет точнее)\n"
                    "Или «Пропустить»."),
        "am_wait": "⏳ <b>Готовим отчёт по практике...</b>\n\n📅 Пишутся дневник и отчёт...",
        "am_done": ("✅ <b>Документы по практике готовы!</b>\n\n📅 Дневник: рабочих дней — <b>{n}</b>\n"
                    "✍️ Места для подписей, оценки и печати оставлены пустыми. При необходимости отредактируйте дневник "
                    "под фактически выполненные работы."),
        "skip": "⏭ Пропустить",
        "turlar": {"oquv": "📘 Учебная практика", "ishlab": "🏭 Производственная практика",
                   "pedagogik": "🏫 Педагогическая практика", "malakaviy": "🎯 Квалификационная практика",
                   "bitiruv": "🎓 Преддипломная практика"},
    },
    "en": {
        "hm_start": ("🛡 <b>Defense preparation</b> — {narx} 🪙\n\n"
                     "Based on your course work or graduation thesis I will prepare:\n"
                     "• 📊 <b>Defense slides</b> (PPTX, 9-10 slides): relevance → aim and objectives → object, subject, "
                     "methods → results (with tables/charts from your work) → recommendations → conclusion\n"
                     "• 🎤 <b>Defense speech</b> for 7-9 minutes — for every slide\n"
                     "• ❓ <b>15 likely committee questions</b> with answers\n"
                     "• ✅ Assessment criteria and tips\n\n"
                     "Send your work: 📄 <b>Word</b> or 📕 <b>PDF</b> (or its text).{last}"),
        "hm_last": "\n\nOr choose the last work created in the bot 👇",
        "hm_last_btn": "📄 My last work: {m}",
        "hm_ism": "👤 Your <b>full name</b> for the first slide:\n<i>e.g. Karimova Dilnoza</i>",
        "hm_short": "⚠️ The text is too short. Please send your work as a Word/PDF file.",
        "hm_wait": "⏳ <b>Preparing defense materials...</b>\n\n📖 Reading and analysing your work...",
        "hm_prog2": "⏳ <b>Preparing defense materials...</b>\n\n📊 Writing slides and speech...",
        "hm_done_ppt": "📊 <b>Defense presentation</b> — «{m}»",
        "hm_done_doc": ("🎤 <b>Defense speech, committee questions and tips</b>\n\n"
                        "💡 Rehearse the speech aloud 2-3 times with a timer. Check that the figures on the slides match your work."),
        "am_start": ("📋 <b>Internship report</b> — {narx} 🪙\n\n"
                     "A complete document set (Word):\n• 🏛 Title page\n• 📅 <b>Internship diary</b> — date and tasks for "
                     "every working day (with a supervisor signature column)\n• 📝 <b>Report</b>: introduction, organisation "
                     "profile, work done, skills, problems and suggestions, conclusion\n• 🧾 Draft <b>reference letter</b> "
                     "(signed by the supervisor)\n\n<b>Choose the type of internship:</b>"),
        "am_otm": "✅ {v}\n\n🏛 <b>University and faculty</b> (2 lines):\n<code>Samarkand State University\nFaculty of Philology</code>",
        "am_talaba": ("👤 <b>Student details</b> (3 lines): full name, field of study, year and group\n"
                      "<code>Karimova Dilnoza\n60111800 – Foreign Language and Literature\n3rd year, group 301</code>"),
        "am_joy": "🏢 <b>Place of internship</b>: organisation (department, city):\n<code>Hamkorbank, Samarkand branch, credit department</code>",
        "am_muddat": "📅 <b>Internship period</b>:\n<code>02.02.2026 – 28.02.2026</code>",
        "am_muddat_bad": "⚠️ I couldn't read the dates. Example: <code>02.02.2026 – 28.02.2026</code> (3 to 120 days)",
        "am_rahbar": ("👨‍🏫 <b>Internship supervisors</b> (2 lines): from the university and from the organisation\n"
                      "<code>Assoc. Prof. Rahimov A.B.\nHead of department Toshev S.</code>\nNot sure — «Skip»."),
        "am_qosh": "✍️ <b>What exactly did you do during the internship?</b> (optional, makes the report more accurate)\nOr «Skip».",
        "am_wait": "⏳ <b>Preparing the internship report...</b>\n\n📅 Writing the diary and report...",
        "am_done": ("✅ <b>Internship documents are ready!</b>\n\n📅 Diary: <b>{n} working days</b>\n"
                    "✍️ Signature, grade and stamp fields are left blank. Adjust the diary to the work you actually did if needed."),
        "skip": "⏭ Skip",
        "turlar": {"oquv": "📘 Educational internship", "ishlab": "🏭 Industrial internship",
                   "pedagogik": "🏫 Teaching practice", "malakaviy": "🎯 Professional internship",
                   "bitiruv": "🎓 Pre-graduation internship"},
    },
}


def _til(cid: int) -> str:
    t = X._til(cid)
    return t if t in _TX else "uz"


def _tt(cid: int, key: str, **kw) -> str:
    s = _TX[_til(cid)].get(key) or _TX["uz"][key]
    return s.format(**kw) if kw else s


def _api():
    return X._api


def _bot():
    return X._bot


def _send(cid, text, **kw):
    return X._send(cid, text, **kw)


def _kb1(*rows):
    kb = InlineKeyboardMarkup(row_width=1)
    kb.add(*[InlineKeyboardButton(t, callback_data=d) for t, d in rows])
    return kb


# ═══════════════════════════════════════════════════════════════════
# 1. HIMOYAGA TAYYORLOV
# ═══════════════════════════════════════════════════════════════════
def _last_v2(cid: int) -> dict | None:
    ld = (_api().holat.get(cid, {}) or {}).get("last_doc") or {}
    d = ld.get("data") if isinstance(ld, dict) else None
    return d if isinstance(d, dict) and d.get("v") == 2 and d.get("bob1_data") else None


def himoya_boshlash(msg):
    cid = msg.chat.id
    if not X._coins_ok(cid, "himoya"):
        return
    X._new_session(cid, "himoya", "hm_src")
    d = _last_v2(cid)
    kb = None
    last = ""
    if d:
        m = (d.get("titul", {}).get("mavzu") or "")[:38]
        kb = _kb1((_tt(cid, "hm_last_btn", m=m), "hm_last"))
        last = _tt(cid, "hm_last")
    _send(cid, _tt(cid, "hm_start", narx=X.NARX["himoya"], last=last), reply_markup=kb or _api().back_kb(cid))


def _hm_last(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if X._qadam(cid) != "hm_src":
        return
    d = _last_v2(cid)
    if not d:
        _send(cid, _tt(cid, "hm_short"))
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _api().update_holat(cid, "hm_src", {"type": "last"})
    ism = (d.get("titul") or {}).get("ism", "")
    if ism:
        _api().update_holat(cid, "hm_ism", ism)
        _hm_start_job(cid)
    else:
        _api().update_holat(cid, "qadam", "hm_ism")
        _send(cid, _tt(cid, "hm_ism"), reply_markup=_api().back_kb(cid))


def _hm_src(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    src = X._capture(msg)
    if not src:
        _send(cid, X._t(cid, "file_bad"))
        return
    if src["type"] == "text" and len(src["text"].strip()) < 600:
        _send(cid, _tt(cid, "hm_short"))
        return
    if src["type"] == "text":
        src["text"] = security.sanitize_long_text(src["text"], 60000)
    _api().update_holat(cid, "hm_src", src)
    _api().update_holat(cid, "qadam", "hm_ism")
    _send(cid, _tt(cid, "hm_ism"), reply_markup=_api().back_kb(cid))


def _hm_ism(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    ok, err = _api().validate_input(cid, msg.text, max_len=120, min_len=2)
    if not ok:
        _send(cid, err)
        return
    try:
        import matn_norm
        ism = matn_norm.normallashtir(msg.text, "ism", _til(cid))
    except Exception:
        ism = (msg.text or "").strip()
    _api().update_holat(cid, "hm_ism", ism)
    _hm_start_job(cid)


def _clean(t) -> str:
    t = re.sub(r"\[(?:\d+(?::\d+)?)(?:,\s*\d+(?::\d+)?)*\]", "", str(t or ""))
    return re.sub(r"\s+", " ", t).strip()


def _v2_matn(d: dict) -> tuple[str, list, list]:
    """Botdagi kurs ishi/BMI ma'lumotidan himoya uchun qisqa matn, jadvallar va grafiklar."""
    til = d.get("til", "uz")
    lb = K.LB(til, d.get("hajm", ""))
    out, jadvallar, grafiklar = [], [], []
    t = d.get("titul", {})
    out.append(f"MAVZU: {t.get('mavzu', '')}")
    kir = (d.get("plan") or {}).get("kirish") or {}
    for k in K.KIRISH_TARTIB:
        if kir.get(k):
            out.append(f"{lb['kirish_lbl'].get(k, k)} {_clean(str(kir[k]).replace('[ABZAS]', ' '))}")
    nb = int(d.get("bob_soni") or 2)
    for bi in range(nb):
        out.append(f"\n{d.get(f'bob{bi + 1}_nomi', '')}")
        for rj in (d.get(f"bob{bi + 1}_data") or {}).get("rejalar", []):
            matn = str(rj.get("matn", ""))
            for blk in re.findall(r"\[JADVAL:[^\]]*\]", matn, flags=re.S):
                tb = K._parse_table_block(blk, til)
                if tb:
                    jadvallar.append(tb)
            for blk in re.findall(r"\[GRAFIK:[^\]]*\]", matn, flags=re.S):
                g = K._dec_chart(blk[len("[GRAFIK:"):-1])
                if g and len(g.get("qiymatlar") or []) >= 2:
                    grafiklar.append(g)
            matn = re.sub(r"\[JADVAL:[^\]]*\]|\[GRAFIK:[^\]]*\]", "", matn, flags=re.S)
            paras = [_clean(x) for x in matn.split("[ABZAS]") if _clean(x)]
            keep = paras[:2] + (paras[-1:] if len(paras) > 2 else [])
            out.append(f"{rj.get('sarlavha', '')}: " + " ".join(keep)[:1800])
    out.append(f"\n{lb['xulosa']}: " + _clean(str((d.get("xulosa_data") or {}).get("xulosa", "")).replace("[ABZAS]", " ")))
    return "\n".join(out), jadvallar, grafiklar


_HM_NOM = {"uz": ("kurs ishi", "bitiruv malakaviy ishi"), "ru": ("курсовая работа", "выпускная квалификационная работа"),
           "en": ("course work", "graduation thesis")}


def _hm_prompt(material: str, til: str, bmi: bool) -> str:
    ish = "bitiruv malakaviy ishi (BMI)" if bmi else "kurs ishi"
    yangilik = ("  4 «Ilmiy yangilik» (tipo \"royxat\", 3-4 band)\n" if bmi else "")
    n = 10 if bmi else 9
    return f"""Sen ko'plab talabalarni himoyaga tayyorlagan tajribali ilmiy rahbar va davlat attestatsiya komissiyasi a'zosisan.
{X.LANG_CMD[til]}
Quyidagi {ish} asosida HIMOYAGA TAYYORLOV materiallarini tuz. FAQAT ishdagi ma'lumotlardan foydalan: ishda yo'q raqam,
natija, so'rovnoma, tajriba yoki fakt O'YLAB TOPMA. Ish matni qisqartirilgan holda berilgan.

1) "slaydlar" — aynan {n} ta slayd (titul va «E'tiboringiz uchun rahmat» slaydlari avtomatik qo'shiladi):
  1 «Mavzuning dolzarbligi» (tipo "royxat", 3-4 band)
  2 «Tadqiqot maqsadi va vazifalari» (tipo "royxat": 1-band — «Maqsad: ...», qolganlari vazifalar)
  3 «Tadqiqot obyekti, predmeti va metodlari» (tipo "royxat": «Obyekt: ...», «Predmet: ...», «Metodlar: ...»)
{yangilik}  {5 if bmi else 4} «Ishning tuzilishi» (tipo "bosqichlar": har qator «Bob nomi | 1 gapdagi mazmuni», boblar soni bo'yicha)
  so'ng «Nazariy tahlil xulosalari», «Amaliy tahlil natijalari», «Aniqlangan muammolar», «Takliflar va tavsiyalar»
  (tipo "royxat", har biri 3-5 band) va oxirida «Xulosa» (tipo "xulosa", 3-5 band).
  Har band 6-16 so'z, aniq va ishning o'ziga xos mazmuni bilan (umumiy iboralar emas). "matn" ichida bandlar \\n bilan ajratiladi.
  Slayd sarlavhalari ish tilida.
2) "nutq" — har slayd uchun himoya nutqi: [{{"slayd": 1, "matn": "..."}}, ...] — birinchi element (slayd 0) —
  salomlashish va mavzuni e'lon qilish («Hurmatli rais, hurmatli komissiya a'zolari! ...»), oxirgisi minnatdorchilik.
  Jami 950-1150 so'z (7-9 daqiqa), birinchi shaxsda, jonli og'zaki nutq uslubida, slayddagi bandlarni so'zma-so'z
  takrorlamay — izohlab.
3) "savollar" — komissiya berishi mumkin bo'lgan 15 ta savol va qisqa, ishonchli javob (har javob 40-70 so'z, ish
  mazmuniga asoslangan): mavzuni tanlash sababi, dolzarblik, obyekt va predmet farqi, metodlar, {"ilmiy yangilik, " if bmi else ""}
  amaliy ahamiyat, asosiy atamalar ta'rifi, natijalar, takliflarni joriy etish, xorij tajribasi, ishning cheklovlari va
  keyingi tadqiqotlar, rahbar/taqrizchi berishi mumkin bo'lgan tanqidiy savollar.

ISH MATNI:
{security.wrap_user_block(material, "ISH")}

FAQAT JSON (Markdown yo'q):
{{"slaydlar":[{{"tipo":"royxat","sarlavha":"...","matn":"...\\n..."}}],
 "nutq":[{{"slayd":0,"matn":"..."}}],
 "savollar":[{{"savol":"...","javob":"..."}}]}}"""


_HM_MASLAHAT = {
    "uz": ("HIMOYADA NIMALAR BAHOLANADI VA MASLAHATLAR", [
        "Nutq 7–10 daqiqa: salomlashish va mavzu (0,5 daq.), dolzarblik (1–1,5), maqsad va vazifalar (1), obyekt, "
        "predmet va metodlar (1), natijalar (3–4), takliflar va xulosa (1–1,5).",
        "Ko'plab OTMlarda ish quyidagi mezonlar bo'yicha baholanadi: tayyorgarlik, mavzuning dolzarbligi, mustaqillik "
        "darajasi, maqsad va natijalarning mosligi, rasmiylashtirish, AKTdan foydalanish, xulosalarning asoslanganligi, "
        "taqdimot va savollarga javoblar (eng katta ulush).",
        "Slaydni o'qimang: slaydda tayanch so'zlar, nutqda izoh. Har slaydga taxminan 1 daqiqa.",
        "Slayddagi raqam, jadval va grafiklar ishingizdagi bilan bir xil bo'lsin.",
        "Savolni oxirigacha tinglang, qisqa va aniq javob bering; bilmasangiz — halol ayting va bu masala keyingi "
        "tadqiqotlarda o'rganilishini qo'shing.",
        "Ilmiy rahbar xulosasi va taqrizdagi tanqidiy fikrlarga oldindan javob tayyorlang.",
        "Ish himoyadan oldin antiplagiat tekshiruvidan o'tkaziladi — matnni o'z tahlilingiz bilan boyiting.",
        "Himoyadan bir kun oldin taqdimotni himoya xonasidagi kompyuterda (USB orqali) tekshirib ko'ring.",
    ]),
    "ru": ("ЧТО ОЦЕНИВАЕТСЯ НА ЗАЩИТЕ И СОВЕТЫ", [
        "Речь 7–10 минут: приветствие и тема (0,5 мин), актуальность (1–1,5), цель и задачи (1), объект, предмет и "
        "методы (1), результаты (3–4), рекомендации и заключение (1–1,5).",
        "Во многих вузах работу оценивают по критериям: подготовка, актуальность, самостоятельность, соответствие "
        "цели и результатов, оформление, использование ИКТ, обоснованность выводов, презентация и ответы на вопросы "
        "(наибольшая доля).",
        "Не читайте слайды: на слайде ключевые слова, в речи — пояснение. Примерно 1 минута на слайд.",
        "Цифры, таблицы и графики на слайдах должны совпадать с работой.",
        "Дослушайте вопрос до конца, отвечайте кратко и точно; если не знаете — скажите честно и отметьте это как "
        "направление дальнейших исследований.",
        "Заранее подготовьте ответы на замечания научного руководителя и рецензента.",
        "Перед защитой работа проходит проверку на антиплагиат — дополните текст собственным анализом.",
        "Накануне проверьте презентацию на компьютере в аудитории защиты.",
    ]),
    "en": ("WHAT IS ASSESSED AT THE DEFENSE AND TIPS", [
        "Speech of 7–10 minutes: greeting and topic (0.5 min), relevance (1–1.5), aim and objectives (1), object, "
        "subject and methods (1), results (3–4), recommendations and conclusion (1–1.5).",
        "Many universities assess: preparation, relevance, independence, match between aims and results, formatting, "
        "use of ICT, soundness of conclusions, presentation and answers to questions (the largest share).",
        "Do not read the slides: key words on the slide, explanation in your speech. About 1 minute per slide.",
        "Figures, tables and charts on the slides must match your work.",
        "Listen to the whole question, answer briefly and precisely; if you do not know, say so honestly and mention it "
        "as a direction for further research.",
        "Prepare answers to the comments of your supervisor and reviewer in advance.",
        "The thesis is checked for plagiarism before the defense — enrich the text with your own analysis.",
        "The day before, test the presentation on the computer in the defense room.",
    ]),
}
_HM_LBL = {
    "uz": dict(nutq="HIMOYA NUTQI", vaqt="Taxminiy davomiyligi: {m} daqiqa ({w} so'z)", kirish="Kirish so'zi", xotima="Yakuniy so'z",
               slayd="{n}-slayd", savol="KOMISSIYA SAVOLLARI VA JAVOBLAR", s="Savol", j="Javob",
               taqdimot_k="Kurs ishi himoyasi", taqdimot_b="Bitiruv malakaviy ishi himoyasi", manba="Manba"),
    "ru": dict(nutq="РЕЧЬ НА ЗАЩИТУ", vaqt="Примерная длительность: {m} мин ({w} слов)", kirish="Вступление", xotima="Заключительное слово",
               slayd="Слайд {n}", savol="ВОПРОСЫ КОМИССИИ И ОТВЕТЫ", s="Вопрос", j="Ответ",
               taqdimot_k="Защита курсовой работы", taqdimot_b="Защита выпускной квалификационной работы", manba="Источник"),
    "en": dict(nutq="DEFENSE SPEECH", vaqt="Approximate duration: {m} min ({w} words)", kirish="Opening", xotima="Closing words",
               slayd="Slide {n}", savol="COMMITTEE QUESTIONS AND ANSWERS", s="Question", j="Answer",
               taqdimot_k="Course work defense", taqdimot_b="Graduation thesis defense", manba="Source"),
}


def _hm_slides(raw: list, jadvallar: list, grafiklar: list, til: str) -> tuple[list, list]:
    """Qaytaradi: (slaydlar, xarita) — xarita[k-1] = AI ning k-slaydi yakuniy ro'yxatdagi o'rni (0 dan)."""
    sl = []
    for s in raw or []:
        if not isinstance(s, dict) or not str(s.get("sarlavha", "")).strip():
            continue
        tipo = s.get("tipo") if s.get("tipo") in ("royxat", "bosqichlar", "xulosa", "matn") else "royxat"
        matn = s.get("matn")
        if isinstance(matn, list):
            matn = "\n".join(str(x) for x in matn)
        matn = "\n".join(_clean(x) for x in str(matn or "").split("\n") if _clean(x))
        sl.append({"tipo": tipo, "sarlavha": _clean(s["sarlavha"])[:90], "matn": matn or "—"})
    sl = sl[:11]
    xarita = list(range(len(sl)))
    # ishdagi HAQIQIY jadval/grafik — «natijalar» slaydidan keyin qo'yiladi
    extra = []
    if grafiklar:
        g = grafiklar[0]
        extra.append({"tipo": "ustun_diag" if g.get("tur") != "doira" else "doira_diag", "sarlavha": g["nomi"][:90],
                      "birligi": g.get("birlik", ""), "manba": g.get("manba", ""),
                      "matn": "\n".join(f"{k}: {v}" for k, v in g["qiymatlar"][:10])})
    if jadvallar:
        t = jadvallar[0]
        cols = [str(c) for c in t["ustunlar"]][:5]
        rows = [" | ".join(cols)] + [" | ".join(str(c) for c in r[:5]) for r in t["qatorlar"][:6]]
        extra.append({"tipo": "jadval_slayd", "sarlavha": t["nomi"][:90], "manba": t.get("manba", ""),
                      "matn": "\n".join(rows)})
    if extra:
        pos = next((i + 1 for i, s in enumerate(sl) if i >= 3 and re.search(r"natija|результат|result", s["sarlavha"], re.I)),
                   None)
        if pos is None:
            pos = next((i + 1 for i, s in enumerate(sl) if i >= 3 and re.search(r"tahlil|анализ|analys", s["sarlavha"], re.I)
                        and not re.search(r"nazariy|теорет|theoret", s["sarlavha"], re.I)), max(1, len(sl) - 3))
        sl[pos:pos] = extra
        xarita = [i if i < pos else i + len(extra) for i in xarita]
    return sl, xarita


def _hm_docx(cid, mavzu, ism, data, slides, til, bmi, xarita=None) -> str:
    lb = _HM_LBL.get(til, _HM_LBL["uz"])
    doc = _doc14()
    _h(doc, lb["taqdimot_b" if bmi else "taqdimot_k"].upper(), 14, sa=2)
    _h(doc, f"«{mavzu}»", 14, sa=2)
    if ism:
        _pp(doc, ism, align="c", ind=0, sa=10)
    nutq = [n for n in (data.get("nutq") or []) if isinstance(n, dict) and _clean(n.get("matn"))]
    words = sum(len(_clean(n["matn"]).split()) for n in nutq)
    _h(doc, lb["nutq"], 14, sb=6)
    _pp(doc, lb["vaqt"].format(m=max(1, round(words / 125)), w=words), italic=True, align="c", ind=0, sa=6)
    for n in nutq:
        try:
            k = int(n.get("slayd", 0))
        except (TypeError, ValueError):
            k = 0
        xr = xarita if xarita is not None else list(range(len(slides)))
        if 1 <= k <= len(xr):
            j = xr[k - 1]                              # PPTXda 1-slayd — titul, kontent 2-dan boshlanadi
            sar = f"{lb['slayd'].format(n=j + 2)}. {slides[j]['sarlavha']}"
        else:
            sar = lb["kirish"] if k <= 0 else lb["xotima"]
        _pp(doc, sar, bold=True, ind=0, sb=6, keep=True)
        for par in re.split(r"\n+", str(n["matn"])):
            if _clean(par):
                _pp(doc, _clean(par))
    _brk(doc)
    _h(doc, lb["savol"], 14)
    for i, q in enumerate([q for q in (data.get("savollar") or []) if isinstance(q, dict) and _clean(q.get("savol"))][:20], 1):
        _pp(doc, f"{i}. {_clean(q['savol'])}", bold=True, ind=0, sb=6, keep=True)
        _pp(doc, _clean(q.get("javob")))
    sar, items = _HM_MASLAHAT.get(til, _HM_MASLAHAT["uz"])
    _brk(doc)
    _h(doc, sar, 14)
    for it in items:
        _pp(doc, "• " + it, ind=0, left=0.5, hang=0.5)
    return X._save(doc, mavzu, "_Himoya_nutqi.docx")


def _hm_start_job(cid: int):
    h = _api().holat.get(cid, {})
    src = h.get("hm_src") or {}
    ism = h.get("hm_ism", "")
    _api().update_holat(cid, "qadam", "hm_jarayon")
    til_user = _til(cid)

    def worker(cid, prog):
        jadvallar, grafiklar = [], []
        if src.get("type") == "last":
            d = _last_v2(cid) or {}
            material, jadvallar, grafiklar = _v2_matn(d)
            til = d.get("til", til_user)
            bmi = d.get("hajm") == "bmi"
            mavzu = (d.get("titul") or {}).get("mavzu", "")
        else:
            material = X._extract(src, til_user)
            if len((material or "").strip()) < 500:
                raise X.UserFacingError("file_bad")
            til = til_user
            bmi = bool(re.search(r"bitiruv\s+malakaviy|выпускн\w*\s+квалификац|дипломн|graduation|bachelor",
                                 material[:6000], re.I))
            mavzu = _guess_topic(material)
            import xizmatlar_qoshimcha as XQ
            material = XQ.sample_text(material, 22000, 10)
        prog.update(_tt(cid, "hm_prog2"))
        data = X._ai_json(_hm_prompt(material, til, bmi), 9000)
        if not mavzu:
            mavzu = _clean(data.get("mavzu") or "") or "Himoya"
        slides, xarita = _hm_slides(data.get("slaydlar"), jadvallar, grafiklar, til)
        if len(slides) < 5:
            raise ValueError("Himoya slaydlari yetarli emas")
        lb = _HM_LBL.get(til, _HM_LBL["uz"])
        ppt = _api().make_fname(mavzu, "_Himoya.pptx")
        pptx_themes.build_presentation(slides, mavzu, ism, "akademik", til, ppt,
                                       labels={"taqdimot": lb["taqdimot_b" if bmi else "taqdimot_k"],
                                               "manba": lb["manba"]})
        dfile = _hm_docx(cid, mavzu, ism, data, slides, til, bmi, xarita)
        safe = X._safe_name(mavzu, "Himoya").replace(".docx", "")
        try:
            with open(ppt, "rb") as f:
                sent = _bot().send_document(cid, f, caption=_tt(cid, "hm_done_ppt", m=security.esc(mavzu[:80])),
                                            parse_mode="HTML", visible_file_name=f"{safe}_taqdimot.pptx")
            try:
                _api().log_javob(cid=cid, javob_turi="🛡 Himoya taqdimoti", javob_matn=mavzu,
                                 file_id=sent.document.file_id, file_type="document",
                                 visible_name=f"{safe}_taqdimot.pptx")
            except Exception:
                pass
            X._send_doc(cid, dfile, f"{safe}_nutq_va_savollar.docx",
                        f"{_tt(cid, 'hm_done_doc')}\n\n{X._t(cid, 'coins', c=_api().get_coins(cid))}"[:1024],
                        "🛡 Himoya nutqi", mavzu)
            _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        finally:
            X._cleanup(ppt)
            X._cleanup(dfile)
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "himoya", "hm_wait_key", worker, "Himoyaga tayyorlov")


def _guess_topic(text: str) -> str:
    m = re.search(r"(?:MAVZU|Mavzu|ТЕМА|Тема|TOPIC|Topic)\s*[:：]\s*[«\"“]?([^\n«»\"“”]{8,220})", text)
    if m:
        t = m.group(1).strip(" .«»\"“”")
        return t[:1] + t[1:].lower() if t.isupper() else t
    for line in text.split("\n")[:60]:
        s = line.strip()
        if s.startswith(("«", "\"", "“")) and 15 <= len(s) <= 220:
            return s.strip(" .«»\"“”")
    return ""


# ═══════════════════════════════════════════════════════════════════
# 2. AMALIYOT HISOBOTI
# ═══════════════════════════════════════════════════════════════════
_AM_NOM = {   # titul va matnda: «{tur} amaliyoti»
    "uz": {"oquv": "O'quv amaliyoti", "ishlab": "Ishlab chiqarish amaliyoti", "pedagogik": "Pedagogik amaliyot",
           "malakaviy": "Malakaviy amaliyot", "bitiruv": "Bitiruv oldi amaliyoti"},
    "ru": {"oquv": "Учебная", "ishlab": "Производственная", "pedagogik": "Педагогическая",
           "malakaviy": "Квалификационная", "bitiruv": "Преддипломная"},
    "en": {"oquv": "Educational", "ishlab": "Industrial", "pedagogik": "Teaching", "malakaviy": "Professional",
           "bitiruv": "Pre-graduation"},
}
_RU_ACC = {"Учебная": "учебную", "Производственная": "производственную", "Педагогическая": "педагогическую",
           "Квалификационная": "квалификационную", "Преддипломная": "преддипломную"}
_AM_LBL = {
    "uz": dict(amaliyot="{t}", hisobot="BO'YICHA HISOBOT", talaba="Talaba:", yon="Ta'lim yo'nalishi:",
               kurs="Kurs, guruh:", joy="Amaliyot joyi:", muddat="Amaliyot muddati:", r1="OTMdan amaliyot rahbari:",
               r2="Tashkilotdan amaliyot rahbari:", baho="Baho: ______________", imzo="(imzo)",
               kundalik="AMALIYOT KUNDALIGI", ustun=["№", "Sana", "Bajarilgan ishlar mazmuni", "Rahbar imzosi"],
               hafta="{n}-hafta", sh="AMALIYOT HISOBOTI", kirish="KIRISH",
               bolim=["1. AMALIYOT O'TKAZILGAN TASHKILOT TAVSIFI", "2. AMALIYOT DAVOMIDA BAJARILGAN ISHLAR",
                      "3. EGALLANGAN BILIM, KO'NIKMA VA MALAKALAR", "4. ANIQLANGAN MUAMMOLAR VA TAKLIFLAR"],
               xulosa="XULOSA", tavsif="TAVSIFNOMA",
               tavsif_sub="{ism}ning ({yon}, {kurs}) {joy}da {d1} – {d2} kunlari o'tagan {t} yuzasidan",
               t_baho="Amaliyot natijalari bo'yicha tavsiya etilgan baho: ______________",
               t_rahbar="Tashkilotdan amaliyot rahbari: ____________________  ____________", mo="M.O'.",
               sana="“____” ______________ 20___ y.", kun_tasdiq="Kundalik tashkilot rahbari tomonidan tasdiqlandi: ____________",
               tuzdi="Hisobotni tayyorladi:", tekshirdi="Tekshirdi:"),
    "ru": dict(amaliyot="{t} ПРАКТИКА", hisobot="ОТЧЁТ", talaba="Студент(ка):", yon="Направление образования:",
               kurs="Курс, группа:", joy="Место практики:", muddat="Сроки практики:", r1="Руководитель практики от вуза:",
               r2="Руководитель практики от организации:", baho="Оценка: ______________", imzo="(подпись)",
               kundalik="ДНЕВНИК ПРАКТИКИ", ustun=["№", "Дата", "Содержание выполненных работ", "Подпись руководителя"],
               hafta="Неделя {n}", sh="ОТЧЁТ ПО ПРАКТИКЕ", kirish="ВВЕДЕНИЕ",
               bolim=["1. ХАРАКТЕРИСТИКА ОРГАНИЗАЦИИ — МЕСТА ПРАКТИКИ", "2. РАБОТЫ, ВЫПОЛНЕННЫЕ В ХОДЕ ПРАКТИКИ",
                      "3. ПРИОБРЕТЁННЫЕ ЗНАНИЯ, УМЕНИЯ И НАВЫКИ", "4. ВЫЯВЛЕННЫЕ ПРОБЛЕМЫ И ПРЕДЛОЖЕНИЯ"],
               xulosa="ЗАКЛЮЧЕНИЕ", tavsif="ХАРАКТЕРИСТИКА",
               tavsif_sub="на студента(ку) {ism} ({yon}, {kurs}), проходившего(ую) {t2} практику ({joy}) с {d1} по {d2}",
               t_baho="Рекомендуемая оценка по итогам практики: ______________",
               t_rahbar="Руководитель практики от организации: ____________________  ____________", mo="М.П.",
               sana="«____» ______________ 20___ г.", kun_tasdiq="Дневник подтверждён руководителем от организации: ____________",
               tuzdi="Отчёт подготовил(а):", tekshirdi="Проверил(а):"),
    "en": dict(amaliyot="{t} INTERNSHIP", hisobot="REPORT", talaba="Student:", yon="Field of study:",
               kurs="Year, group:", joy="Place of internship:", muddat="Period:", r1="University supervisor:",
               r2="Organisation supervisor:", baho="Grade: ______________", imzo="(signature)",
               kundalik="INTERNSHIP DIARY", ustun=["No.", "Date", "Work performed", "Supervisor's signature"],
               hafta="Week {n}", sh="INTERNSHIP REPORT", kirish="INTRODUCTION",
               bolim=["1. PROFILE OF THE HOST ORGANISATION", "2. WORK PERFORMED DURING THE INTERNSHIP",
                      "3. KNOWLEDGE AND SKILLS GAINED", "4. PROBLEMS IDENTIFIED AND SUGGESTIONS"],
               xulosa="CONCLUSION", tavsif="REFERENCE LETTER",
               tavsif_sub="for {ism} ({yon}, {kurs}), who completed the {t2} internship at {joy} from {d1} to {d2}",
               t_baho="Recommended grade for the internship: ______________",
               t_rahbar="Organisation supervisor: ____________________  ____________", mo="Stamp",
               sana="“____” ______________ 20___", kun_tasdiq="Diary confirmed by the organisation supervisor: ____________",
               tuzdi="Prepared by:", tekshirdi="Checked by:"),
}


def amaliyot_boshlash(msg):
    cid = msg.chat.id
    if not X._coins_ok(cid, "amaliyot"):
        return
    X._new_session(cid, "amaliyot", "am_tur")
    T = _TX[_til(cid)]["turlar"]
    _send(cid, _tt(cid, "am_start", narx=X.NARX["amaliyot"]), reply_markup=_kb1(*[(v, f"amt_{k}") for k, v in T.items()]))


def _am_tur(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if X._qadam(cid) != "am_tur":
        return
    k = call.data.split("_", 1)[1]
    T = _TX[_til(cid)]["turlar"]
    if k not in T:
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _api().update_holat(cid, "am_tur", k)
    _api().update_holat(cid, "qadam", "am_otm")
    _send(cid, _tt(cid, "am_otm", v=T[k]), reply_markup=_api().back_kb(cid))


def _lines(text: str) -> list[str]:
    q = [x.strip() for x in (text or "").strip().split("\n") if x.strip()]
    if len(q) == 1 and q[0].count(",") >= 1:
        q = [x.strip() for x in q[0].split(",") if x.strip()]
    return q


def _nrm(cid, text, tur, strip=True):
    try:
        import matn_norm
        return matn_norm.normallashtir(text or "", tur, _til(cid), strip=strip)
    except Exception:
        return (text or "").strip()


def _am_step(field: str, nxt: str, ask: str, mx: int, mn: int = 2, kb_skip: str | None = None):
    def h(msg):
        cid = msg.chat.id
        if _api().is_back(msg):
            return
        ok, err = _api().validate_input(cid, msg.text, max_len=mx, min_len=mn)
        if not ok:
            _send(cid, err)
            return
        q = _lines(msg.text)
        if field == "am_otm":
            _api().update_holat(cid, "am_otm", _nrm(cid, q[0], "otm"))
            _api().update_holat(cid, "am_fak", _nrm(cid, " ".join(q[1:]), "fakultet", strip=False) if len(q) > 1 else "")
        elif field == "am_talaba":
            _api().update_holat(cid, "am_ism", _nrm(cid, q[0], "ism"))
            _api().update_holat(cid, "am_yon", _nrm(cid, q[1], "fan") if len(q) > 1 else "")
            _api().update_holat(cid, "am_kurs", " ".join(q[2:])[:60] if len(q) > 2 else "")
        elif field == "am_joy":
            _api().update_holat(cid, "am_joy", security.sanitize_long_text(msg.text, 250).replace("\n", ", "))
        elif field == "am_muddat":
            d = _parse_dates(msg.text)
            if not d:
                _send(cid, _tt(cid, "am_muddat_bad"))
                return
            _api().update_holat(cid, "am_d1", d[0].isoformat())
            _api().update_holat(cid, "am_d2", d[1].isoformat())
        _api().update_holat(cid, "qadam", nxt)
        kb = _kb1((_tt(cid, "skip"), kb_skip)) if kb_skip else _api().back_kb(cid)
        _send(cid, _tt(cid, ask), reply_markup=kb)
    h.__name__ = f"_am_{field}"
    return h


_DATE = re.compile(r"(\d{1,2})[./\-](\d{1,2})(?:[./\-](\d{2,4}))?")


def _parse_dates(text: str):
    ds = _DATE.findall(text or "")
    if len(ds) < 2:
        return None
    yil = datetime.date.today().year
    out = []
    for d, m, y in ds[:2]:
        y = int(y) if y else None
        if y and y < 100:
            y += 2000
        try:
            out.append([int(d), int(m), y])
        except ValueError:
            return None
    if out[1][2] is None:
        out[1][2] = out[0][2] or yil
    if out[0][2] is None:
        out[0][2] = out[1][2] if (out[0][1], out[0][0]) <= (out[1][1], out[1][0]) else out[1][2] - 1
    try:
        d1, d2 = (datetime.date(y, m, d) for d, m, y in out)
    except ValueError:
        return None
    if not (3 <= (d2 - d1).days + 1 <= 120):
        return None
    return d1, d2


def _ish_kunlari(d1, d2, olti: bool) -> list:
    days, d = [], d1
    while d <= d2:
        if d.weekday() < (6 if olti else 5):
            days.append(d)
        d += datetime.timedelta(days=1)
    return days


def _am_rahbar_text(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    q = _lines(msg.text)
    _api().update_holat(cid, "am_r1", _nrm(cid, q[0], "rahbar")[:120] if q else "")
    _api().update_holat(cid, "am_r2", _nrm(cid, " ".join(q[1:]), "rahbar")[:120] if len(q) > 1 else "")
    _am_qosh_sora(cid)


def _am_rahbar_skip(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if X._qadam(cid) != "am_rahbar":
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _api().update_holat(cid, "am_r1", "")
    _api().update_holat(cid, "am_r2", "")
    _am_qosh_sora(cid)


def _am_qosh_sora(cid):
    _api().update_holat(cid, "qadam", "am_qosh")
    _send(cid, _tt(cid, "am_qosh"), reply_markup=_kb1((_tt(cid, "skip"), "amq_skip")))


def _am_qosh_text(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    _api().update_holat(cid, "am_info", security.sanitize_long_text(msg.text or "", 2500))
    _am_start_job(cid, msg)


def _am_qosh_skip(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if X._qadam(cid) != "am_qosh":
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _api().update_holat(cid, "am_info", "")
    _am_start_job(cid, call.message)


def _fmt_d(d, til: str) -> str:
    return d.strftime("%d.%m.%Y")


def _am_prompt(p: dict, kunlar: list, til: str) -> str:
    tur_nom = _AM_NOM["uz"][p["tur"]].replace(" amaliyoti", "").replace(" amaliyot", "")
    yon = p["yon"] or "ko'rsatilmagan"
    kurs = p["kurs"] or "ko'rsatilmagan"
    ped = (", pedagogik amaliyotda: darslarni kuzatish, dars ishlanmasi tuzish, sinov va mustaqil darslar, tarbiyaviy "
           "tadbir, sinf rahbari ishi, o'quvchilarning psixologik-pedagogik tavsifi") if p["tur"] == "pedagogik" else ""
    haftalar = {}
    for d in kunlar:
        haftalar.setdefault(d.isocalendar()[1], []).append(d)
    hafta_soni = len(haftalar)
    sanalar = ", ".join(d.strftime("%d.%m") for d in kunlar)
    info = (f"Talaba o'zi aytgan ishlar (ASOSIY manba — kundalik va hisobotda shular aks etsin):\n"
            f"{security.wrap_user_block(p['info'], 'TALABA')}\n") if p.get("info") else ""
    return f"""Sen OTM amaliyot bo'limida ko'p yillik tajribaga ega amaliyot rahbarisan. {X.LANG_CMD[til]}
Talaba uchun «{tur_nom} amaliyoti» bo'yicha KUNDALIK, HISOBOT va TAVSIFNOMA loyihasini tayyorla.
Ma'lumotlar: ta'lim yo'nalishi — {yon}; kurs — {kurs};
amaliyot joyi — {p['joy']}; muddati — {_fmt_d(kunlar[0], til)} – {_fmt_d(kunlar[-1], til)} ({len(kunlar)} ish kuni, {hafta_soni} hafta).
{info}
1) "kundalik": aynan {len(kunlar)} ta yozuv — har ish kuni uchun (sanalar: {sanalar}), har biri 12-28 so'z: o'sha kuni
   bajarilgan aniq ish (o'tgan zamonda, «...bilan tanishdim», «...ni o'rgandim», «...da ishtirok etdim»). Mantiqiy ketma-ketlik:
   1-kun — tashkilot bilan tanishish, xavfsizlik texnikasi va ichki tartib; keyin bo'limlar va hujjatlar bilan tanishish,
   kuzatish, rahbar ko'rsatmasi bilan mustaqil vazifalar, oxirgi kunlar — materiallarni umumlashtirish, hisobot tayyorlash
   va topshirish. Ishlar yo'nalishga va amaliyot turiga mos bo'lsin{ped}.
   Bir xil yozuvlarni takrorlama.
2) "kirish": 2 abzas — amaliyotning maqsadi va vazifalari (vazifalar ro'yxat ko'rinishida emas, matnda).
3) "tashkilot": 2-3 abzas — tashkilotning faoliyat yo'nalishi, tuzilmasi va bo'limlari, talaba ishlagan bo'lim vazifalari.
   Tashkilot haqida ANIQ bilmagan fakt (tashkil etilgan yili, xodimlar soni, raqamlar, rahbar ismi)ni O'YLAB TOPMA —
   bunday joyda «______» qoldir; tashkilot turiga xos umumiy tavsif ber.
4) "ishlar": {hafta_soni} ta abzas — har hafta bo'yicha bajarilgan ishlar tahlili (kundalikka mos).
5) "konikmalar": 6-8 ta band — egallangan bilim, ko'nikma va malakalar.
6) "takliflar": 4-5 ta band — amaliyot davomida kuzatilgan muammolar va ularni hal etish bo'yicha takliflar (tashkilot va
   amaliyotni tashkil etish bo'yicha; xushmuomala, konstruktiv).
7) "xulosa": 2 abzas.
8) "tavsifnoma": 3 abzas, uchinchi shaxsda (tashkilot rahbari nomidan): talabaning amaliyotdagi faolligi, intizomi va
   jamoa bilan munosabati; bajargan ishlari va nazariy bilimlarini amaliyotda qo'llashi; kasbiy sifatlari va xulosa.
   Haddan tashqari maqtov va o'ylab topilgan yutuqlarsiz. Baho yozilmasin.

FAQAT JSON (Markdown yo'q):
{{"kundalik":["...","..."],"kirish":["...","..."],"tashkilot":["..."],"ishlar":["..."],"konikmalar":["..."],
 "takliflar":["..."],"xulosa":["...","..."],"tavsifnoma":["...","...","..."]}}"""


def _am_start_job(cid: int, msg=None):
    h = dict(_api().holat.get(cid, {}))
    til = _til(cid)
    p = {"tur": h.get("am_tur", "malakaviy"), "otm": h.get("am_otm", ""), "fak": h.get("am_fak", ""),
         "ism": h.get("am_ism", ""), "yon": h.get("am_yon", ""), "kurs": h.get("am_kurs", ""),
         "joy": h.get("am_joy", ""), "r1": h.get("am_r1", ""), "r2": h.get("am_r2", ""), "info": h.get("am_info", "")}
    try:
        d1 = datetime.date.fromisoformat(h["am_d1"])
        d2 = datetime.date.fromisoformat(h["am_d2"])
    except Exception:
        _send(cid, _tt(cid, "am_muddat_bad"))
        _api().update_holat(cid, "qadam", "am_muddat")
        return
    kunlar = _ish_kunlari(d1, d2, olti=p["tur"] == "pedagogik")
    if len(kunlar) < 2:
        _send(cid, _tt(cid, "am_muddat_bad"))
        _api().update_holat(cid, "qadam", "am_muddat")
        return
    _api().update_holat(cid, "qadam", "am_jarayon")
    if msg is not None:
        X._log_start(msg, f"📋 Amaliyot: {p['tur']}", f"{p['ism']} | {p['joy']} | {d1} – {d2}")

    def worker(cid, prog):
        data = X._ai_json(_am_prompt(p, kunlar, til), 9000)
        kun = [_clean(x) for x in (data.get("kundalik") or []) if _clean(x)]
        if len(kun) < max(2, len(kunlar) // 2):
            raise ValueError("Kundalik to'liq emas")
        while len(kun) < len(kunlar):                 # kam qaytsa — oxirgi umumiy ishlar bilan to'ldiramiz
            kun.insert(-1, kun[-2] if len(kun) >= 2 else kun[-1])
        fname = _am_docx(p, kunlar, kun[:len(kunlar)], data, til)
        try:
            X._send_doc(cid, fname, X._safe_name(p["ism"] or p["joy"], "Amaliyot_hisoboti"),
                        f"{_tt(cid, 'am_done', n=len(kunlar))}\n\n{X._t(cid, 'coins', c=_api().get_coins(cid))}"[:1024],
                        "📋 Amaliyot hisoboti", f"{p['tur']} | {p['joy']}")
            _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        finally:
            X._cleanup(fname)
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "amaliyot", "am_wait_key", worker, "Amaliyot hisoboti")


def _lst(v) -> list:
    if isinstance(v, str):
        v = [x for x in re.split(r"\n+|\[ABZAS\]", v)]
    return [_clean(x) for x in (v or []) if _clean(x)]


def _am_docx(p: dict, kunlar: list, kun: list, data: dict, til: str) -> str:
    lb = _AM_LBL.get(til, _AM_LBL["uz"])
    tur = _AM_NOM.get(til, _AM_NOM["uz"])[p["tur"]]
    doc = _doc14()
    shahar = K.shahar_aniqla(p["otm"], til)
    yil = kunlar[-1].year

    # ── Titul ──
    _tp(doc, K.L(til)["vazirlik"], 14, sa=4, first=True)
    _tp(doc, (p["otm"] or "______________________________").upper(), 14, sa=2)
    if p["fak"]:
        fk = p["fak"]
        if til == "uz" and not re.search(r"fakultet", fk, re.I):
            fk += " fakulteti"
        _tp(doc, fk, 14, bold=False, sa=0)
    _tp(doc, lb["amaliyot"].format(t=tur.upper()), 20, sb=70, sa=2)
    _tp(doc, lb["hisobot"], 20, sa=40)
    rows = [(lb["talaba"], p["ism"]), (lb["yon"], p["yon"]), (lb["kurs"], p["kurs"]), (lb["joy"], p["joy"]),
            (lb["muddat"], f"{_fmt_d(kunlar[0], til)} – {_fmt_d(kunlar[-1], til)}"),
            (lb["r1"], p["r1"]), (lb["r2"], p["r2"])]
    for k, v in rows:
        pp = _pp(doc, "", ind=0, left=5.0, ls=1.15, sa=4, align="l")
        r = pp.add_run(k + " "); r.bold = True
        pp.add_run(v or "______________")
    _pp(doc, lb["baho"], ind=0, left=5.0, sb=16, align="l")
    _tp(doc, f"{shahar} – {yil}", 14, sb=90)

    # ── Kundalik ──
    _brk(doc)
    _h(doc, lb["kundalik"], 14, sa=8)
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    widths = (Cm(1.1), Cm(2.7), Cm(9.8), Cm(2.9))
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w.cm * 567)))
    for i, c in enumerate(lb["ustun"]):
        _cell(t.rows[0].cells[i], c, bold=True, align="c", w=widths[i])
    _hdr_repeat(t.rows[0])
    hafta_oldingi = None
    for i, (d, ish) in enumerate(zip(kunlar, kun), 1):
        hf = d.isocalendar()[1]
        if hf != hafta_oldingi:
            hafta_oldingi = hf
            row = t.add_row()
            m = row.cells[0].merge(row.cells[3])
            nomer = len({x.isocalendar()[1] for x in kunlar if x <= d})
            _cell(m, lb["hafta"].format(n=nomer), bold=True, align="c", italic=True)
        row = t.add_row().cells
        _cell(row[0], str(i), align="c", w=widths[0])
        _cell(row[1], _fmt_d(d, til), align="c", w=widths[1])
        _cell(row[2], ish, w=widths[2])
        _cell(row[3], "", w=widths[3])
    _pp(doc, lb["kun_tasdiq"], ind=0, sb=10, align="l")
    _pp(doc, lb["mo"], ind=0, sb=2, align="l")

    # ── Hisobot ──
    _brk(doc)
    _h(doc, lb["sh"], 14, sa=10)
    _h(doc, lb["kirish"], 14, sa=6)
    for x in _lst(data.get("kirish")):
        _pp(doc, x)
    bolimlar = [_lst(data.get("tashkilot")), _lst(data.get("ishlar")), _lst(data.get("konikmalar")),
                _lst(data.get("takliflar"))]
    for bi, (sar, items) in enumerate(zip(lb["bolim"], bolimlar)):
        _h(doc, sar, 14, sb=10, sa=6)
        for x in items:
            if bi >= 2:
                _pp(doc, "– " + x.rstrip(".;") + ";", ind=0, left=0.6, hang=0.6)
            else:
                _pp(doc, x)
    _h(doc, lb["xulosa"], 14, sb=10, sa=6)
    for x in _lst(data.get("xulosa")):
        _pp(doc, x)
    _pp(doc, f"{lb['tuzdi']} {p['ism'] or '____________'}  ____________ {lb['imzo']}", ind=0, sb=18, align="l")
    _pp(doc, f"{lb['tekshirdi']} {p['r1'] or '____________'}  ____________ {lb['imzo']}", ind=0, sb=6, align="l")

    # ── Tavsifnoma ──
    _brk(doc)
    _h(doc, lb["tavsif"], 16, sa=4)
    _pp(doc, lb["tavsif_sub"].format(ism=p["ism"] or "____________", yon=p["yon"] or "____________",
                                     kurs=p["kurs"] or "____", joy=p["joy"], d1=_fmt_d(kunlar[0], til),
                                     d2=_fmt_d(kunlar[-1], til), t=(tur.lower() + "i") if tur.lower().endswith("amaliyot") else tur.lower(),
                                     t2=_RU_ACC.get(tur, tur.lower())),
        italic=True, ind=0, align="c", ls=1.15, sa=10)
    for x in _lst(data.get("tavsifnoma")):
        _pp(doc, x)
    _pp(doc, lb["t_baho"], ind=0, sb=14, align="l")
    _pp(doc, lb["t_rahbar"], ind=0, sb=14, align="l")
    _pp(doc, lb["mo"], ind=0, sb=6, align="l")
    _pp(doc, lb["sana"], ind=0, sb=6, align="l")
    try:
        doc.sections[0].different_first_page_header_footer = True     # titulda bet raqami yo'q
        _api().add_page_numbers(doc.sections[0])
    except Exception:
        pass
    return X._save(doc, p["ism"] or "amaliyot", "_Amaliyot_hisoboti.docx")


# ═══════════════════════════════════════════════════════════════════
# DOCX YORDAMCHILARI (TNR 14, 1,5 interval, OTM talablari)
# ═══════════════════════════════════════════════════════════════════
def _doc14() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(3), Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.element.rPr.rFonts.set(qn("w:cs"), "Times New Roman")
    pf = st.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = pf.space_after = Pt(0)
    return doc


_ALIGN = {"c": WD_ALIGN_PARAGRAPH.CENTER, "l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT,
          "j": WD_ALIGN_PARAGRAPH.JUSTIFY}


def _pp(doc, text="", bold=False, italic=False, align="j", ind=1.25, left=0.0, hang=0.0, sb=0, sa=0, ls=1.5,
        keep=False):
    p = doc.add_paragraph()
    f = p.paragraph_format
    f.first_line_indent = Cm(-hang if hang else ind)
    f.left_indent = Cm(left)
    f.space_before, f.space_after = Pt(sb), Pt(sa)
    f.line_spacing = ls
    f.keep_with_next = keep
    p.alignment = _ALIGN[align]
    if text:
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
    return p


def _h(doc, text, size=14, sb=0, sa=6):
    p = _pp(doc, "", align="c", ind=0, sb=sb, sa=sa, keep=True)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    return p


def _tp(doc, text, size=14, bold=True, sb=0, sa=0, first=False):
    p = doc.paragraphs[0] if (first and doc.paragraphs) else doc.add_paragraph()
    f = p.paragraph_format
    f.first_line_indent = Cm(0)
    f.space_before, f.space_after = Pt(sb), Pt(sa)
    f.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    return p


def _brk(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _cell(cell, text, bold=False, italic=False, align="l", w=None):
    if w is not None:
        cell.width = w
    p = cell.paragraphs[0]
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = _ALIGN[align]
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.size = Pt(12)


def _hdr_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)


# ═══════════════════════════════════════════════════════════════════
# RO'YXATDAN O'TKAZISH
# ═══════════════════════════════════════════════════════════════════
def register(bot, bmi_start):
    """xizmatlar_yangi.register() dan KEYIN chaqiriladi. bmi_start — bot.py dagi BMI boshlovchi."""
    for _t, _d in {"uz": {"hm_wait_key": _TX["uz"]["hm_wait"], "am_wait_key": _TX["uz"]["am_wait"]},
                   "ru": {"hm_wait_key": _TX["ru"]["hm_wait"], "am_wait_key": _TX["ru"]["am_wait"]},
                   "en": {"hm_wait_key": _TX["en"]["hm_wait"], "am_wait_key": _TX["en"]["am_wait"]}}.items():
        X.T[_t].update(_d)
    X.STARTERS.update({"bmi": bmi_start, "himoya": himoya_boshlash, "amaliyot": amaliyot_boshlash})
    ct = ["text", "document", "photo"]

    def guard(fn):
        def wrapped(x):
            m = x.message if hasattr(x, "data") else x
            if _api().is_banned(m.chat.id):
                return
            try:
                fn(x)
            except Exception as e:
                log.error(f"[{m.chat.id}] {fn.__name__}: {e}", exc_info=True)
                try:
                    _send(m.chat.id, X._t(m.chat.id, "err"), reply_markup=_api().main_kb(m.chat.id))
                except Exception:
                    pass
        wrapped.__name__ = fn.__name__
        return wrapped

    q = X._q
    bot.register_callback_query_handler(guard(_hm_last), func=lambda c: c.data == "hm_last")
    bot.register_message_handler(guard(_hm_src), func=q("hm_src"), content_types=ct)
    bot.register_message_handler(guard(_hm_ism), func=q("hm_ism"))

    bot.register_callback_query_handler(guard(_am_tur), func=lambda c: c.data.startswith("amt_"))
    bot.register_message_handler(guard(_am_step("am_otm", "am_talaba", "am_talaba", 300)), func=q("am_otm"))
    bot.register_message_handler(guard(_am_step("am_talaba", "am_joy", "am_joy", 300)), func=q("am_talaba"))
    bot.register_message_handler(guard(_am_step("am_joy", "am_muddat", "am_muddat", 250, 3)), func=q("am_joy"))
    bot.register_message_handler(guard(_am_step("am_muddat", "am_rahbar", "am_rahbar", 80, 6, "amr_skip")),
                                 func=q("am_muddat"))
    bot.register_message_handler(guard(_am_rahbar_text), func=q("am_rahbar"))
    bot.register_callback_query_handler(guard(_am_rahbar_skip), func=lambda c: c.data == "amr_skip")
    bot.register_message_handler(guard(_am_qosh_text), func=q("am_qosh"))
    bot.register_callback_query_handler(guard(_am_qosh_skip), func=lambda c: c.data == "amq_skip")
    log.info("✅ Talaba xizmatlari ulandi: BMI, Himoyaga tayyorlov, Amaliyot hisoboti")

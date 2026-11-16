"""
xizmatlar_qoshimcha.py — foydalanuvchilar takliflari asosida qo'shilgan xizmatlar
═══════════════════════════════════════════════════════════════════
  🧠 Fayldan test     — PDF / Word / rasm / matndan test (bob yoki betlar bo'yicha),
                        10-30 savol, daraja, oq toza Word + Telegram ichida interaktiv test
  🎧 Ovozli matn      — matn, fayl yoki rasmni MP3 audioga aylantirish (o'zbek, rus, ingliz,
                        qozoq, turk ovozlari). Konspekt va mustaqil ish ostida ham tugma bor
  🧑‍🏫 Sinf sardori     — yig'ilish bayonnomasi, tadbir ssenariysi, tarbiyaviy soat rejasi,
                        e'lon/taklifnoma, faoliyat hisoboti (rasmiy Word)

xizmatlar_yangi.py dagi umumiy yordamchilar (tanga, navbat, fayl o'qish) qayta ishlatiladi.
"""
from __future__ import annotations

import asyncio
import io
import logging
import re
import time

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup

import security
import xizmatlar_yangi as X

log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════
# NARX, TUGMA, MATNLAR (xizmatlar_yangi lug'atlariga qo'shiladi)
# ═══════════════════════════════════════════════════════════════════
X.NARX.update({"ftest": 4, "quiz": 2, "tts": 2, "sardor": 4})

for _til, _b in {
    "uz": {"ftest": "🧠 Test markazi", "tts": "🎧 Ovozli matn", "sardor": "🧑‍🏫 Sinf sardori"},
    "ru": {"ftest": "🧠 Тест-центр", "tts": "🎧 Текст в аудио", "sardor": "🧑‍🏫 Староста / классный час"},
    "en": {"ftest": "🧠 Quiz center", "tts": "🎧 Text to audio", "sardor": "🧑‍🏫 Class leader"},
}.items():
    X.BTN[_til].update(_b)

_TXT = {
    "uz": {
        "ft_start": ("🧠 <b>Fayldan test tuzish</b> — {narx} 🪙\n\n"
                     "Test tuziladigan materialni yuboring:\n"
                     "• 📕 <b>PDF</b> (darslik, qo'llanma — 300 betgacha)\n"
                     "• 📄 <b>Word</b> fayl\n"
                     "• 📷 <b>Rasm</b> (kitob sahifasi, konspekt, doska)\n"
                     "• ✍️ yoki shunchaki <b>matn</b>\n\n"
                     "<i>CEFR, IELTS, milliy sertifikat va imtihonlarga tayyorlanish uchun juda qulay.</i>"),
        "ft_qism": ("📖 <b>Qaysi qismidan test tuzay?</b>\n\n"
                    "Yozing, masalan:\n"
                    "• <code>20-35</code> — shu betlardan\n"
                    "• <code>3-bob</code> yoki bob/mavzu nomi\n\n"
                    "Yoki butun materialdan:"),
        "ft_all": "📚 Hammasidan",
        "ft_soni": "🔢 Nechta savol bo'lsin?",
        "ft_daraja": "📊 Qiyinlik darajasi:",
        "d_oson": "🟢 Oson", "d_orta": "🟡 O'rta", "d_qiyin": "🔴 Qiyin", "d_aralash": "🎲 Aralash",
        "ft_wait": "🧠 Test tuzilmoqda... (30-90 soniya)",
        "ft_done": ("✅ <b>Test tayyor!</b>\n\n"
                    "Word faylda savollar va oxirida javoblar kaliti bor.\n"
                    "Testni hoziroq shu yerda yechib ko'rasizmi?"),
        "ft_bad_src": "⚠️ PDF, Word, rasm yoki matn yuboring.",
        "ft_short": "⚠️ Materialdan yetarli matn olinmadi. Boshqa fayl yoki aniqroq rasm yuboring.",
        "tts_start": ("🎧 <b>Matnni ovozga aylantirish</b> — {narx} 🪙\n\n"
                      "Tinglamoqchi bo'lgan matnni yuboring:\n"
                      "• ✍️ matn, 📄 Word, 📕 PDF yoki 📷 rasm\n"
                      "• {limit} belgigacha (~{daq} daqiqa audio)\n\n"
                      "<i>Ko'zingiz charchaganda ma'ruza, referat yoki konspektni tinglab o'rganing.</i>"),
        "tts_voice": "🗣 Ovozni tanlang:",
        "tts_wait": "🎧 Audio tayyorlanmoqda... (20-90 soniya)",
        "tts_done": "✅ Audio tayyor! 🎧",
        "tts_cut": "ℹ️ Matn uzun edi — dastlabki {limit} belgisi ovozga aylantirildi.",
        "tts_none": "⚠️ Ovozga aylantiriladigan matn topilmadi. Avval hujjat yarating yoki 🎧 Ovozli matn bo'limidan foydalaning.",
        "sd_start": ("🧑‍🏫 <b>Sinf sardori va yetakchilar uchun</b> — {narx} 🪙\n\n"
                     "Qaysi hujjat kerak?"),
        "sd_wait": "🧑‍🏫 Hujjat tayyorlanmoqda... (30-60 soniya)",
        "sd_done": "✅ <b>Hujjat tayyor!</b> Kerakli joylarini to'ldirib, chop etishingiz mumkin.",
        "sd_short": "⚠️ Biroz batafsilroq yozing (kamida 10 belgi).",
    },
    "ru": {
        "ft_start": ("🧠 <b>Тест из файла</b> — {narx} 🪙\n\n"
                     "Отправьте материал для теста:\n"
                     "• 📕 <b>PDF</b> (учебник, пособие — до 300 стр.)\n"
                     "• 📄 файл <b>Word</b>\n"
                     "• 📷 <b>фото</b> (страница книги, конспект, доска)\n"
                     "• ✍️ или просто <b>текст</b>\n\n"
                     "<i>Удобно для подготовки к CEFR, IELTS, нацсертификату и экзаменам.</i>"),
        "ft_qism": ("📖 <b>По какой части составить тест?</b>\n\n"
                    "Напишите, например:\n"
                    "• <code>20-35</code> — по этим страницам\n"
                    "• <code>глава 3</code> или название темы\n\n"
                    "Или по всему материалу:"),
        "ft_all": "📚 По всему",
        "ft_soni": "🔢 Сколько вопросов?",
        "ft_daraja": "📊 Уровень сложности:",
        "d_oson": "🟢 Лёгкий", "d_orta": "🟡 Средний", "d_qiyin": "🔴 Сложный", "d_aralash": "🎲 Смешанный",
        "ft_wait": "🧠 Составляю тест... (30-90 секунд)",
        "ft_done": ("✅ <b>Тест готов!</b>\n\n"
                    "В файле Word — вопросы и ключ ответов в конце.\n"
                    "Пройти тест прямо здесь?"),
        "ft_bad_src": "⚠️ Отправьте PDF, Word, фото или текст.",
        "ft_short": "⚠️ Из материала не удалось получить достаточно текста. Отправьте другой файл или чёткое фото.",
        "tts_start": ("🎧 <b>Текст в аудио</b> — {narx} 🪙\n\n"
                      "Отправьте текст, который хотите прослушать:\n"
                      "• ✍️ текст, 📄 Word, 📕 PDF или 📷 фото\n"
                      "• до {limit} символов (~{daq} мин аудио)\n\n"
                      "<i>Слушайте лекции и конспекты, когда устали глаза.</i>"),
        "tts_voice": "🗣 Выберите голос:",
        "tts_wait": "🎧 Готовлю аудио... (20-90 секунд)",
        "tts_done": "✅ Аудио готово! 🎧",
        "tts_cut": "ℹ️ Текст длинный — озвучены первые {limit} символов.",
        "tts_none": "⚠️ Нет текста для озвучки. Сначала создайте документ или используйте раздел 🎧 Текст в аудио.",
        "sd_start": "🧑‍🏫 <b>Для старост и лидеров</b> — {narx} 🪙\n\nКакой документ нужен?",
        "sd_wait": "🧑‍🏫 Готовлю документ... (30-60 секунд)",
        "sd_done": "✅ <b>Документ готов!</b> Заполните нужные поля и распечатайте.",
        "sd_short": "⚠️ Напишите чуть подробнее (минимум 10 символов).",
    },
    "en": {
        "ft_start": ("🧠 <b>Quiz from a file</b> — {narx} 🪙\n\n"
                     "Send the material for the quiz:\n"
                     "• 📕 <b>PDF</b> (textbook, manual — up to 300 pages)\n"
                     "• 📄 <b>Word</b> file\n"
                     "• 📷 <b>photo</b> (book page, notes, whiteboard)\n"
                     "• ✍️ or just <b>text</b>\n\n"
                     "<i>Great for CEFR, IELTS and exam preparation.</i>"),
        "ft_qism": ("📖 <b>Which part should the quiz cover?</b>\n\n"
                    "Type, for example:\n"
                    "• <code>20-35</code> — these pages\n"
                    "• <code>chapter 3</code> or a topic name\n\n"
                    "Or the whole material:"),
        "ft_all": "📚 Everything",
        "ft_soni": "🔢 How many questions?",
        "ft_daraja": "📊 Difficulty:",
        "d_oson": "🟢 Easy", "d_orta": "🟡 Medium", "d_qiyin": "🔴 Hard", "d_aralash": "🎲 Mixed",
        "ft_wait": "🧠 Building the quiz... (30-90 seconds)",
        "ft_done": ("✅ <b>Quiz ready!</b>\n\n"
                    "The Word file has the questions and an answer key at the end.\n"
                    "Take the quiz right here?"),
        "ft_bad_src": "⚠️ Send a PDF, Word file, photo or text.",
        "ft_short": "⚠️ Not enough text could be read from the material. Send another file or a clearer photo.",
        "tts_start": ("🎧 <b>Text to audio</b> — {narx} 🪙\n\n"
                      "Send the text you want to listen to:\n"
                      "• ✍️ text, 📄 Word, 📕 PDF or 📷 photo\n"
                      "• up to {limit} characters (~{daq} min of audio)\n\n"
                      "<i>Listen to lectures and notes when your eyes are tired.</i>"),
        "tts_voice": "🗣 Choose a voice:",
        "tts_wait": "🎧 Preparing audio... (20-90 seconds)",
        "tts_done": "✅ Audio ready! 🎧",
        "tts_cut": "ℹ️ The text was long — the first {limit} characters were voiced.",
        "tts_none": "⚠️ No text to voice. Create a document first or use 🎧 Text to audio.",
        "sd_start": "🧑‍🏫 <b>For class leaders</b> — {narx} 🪙\n\nWhich document do you need?",
        "sd_wait": "🧑‍🏫 Preparing the document... (30-60 seconds)",
        "sd_done": "✅ <b>Document ready!</b> Fill in the blanks and print it.",
        "sd_short": "⚠️ Please write a bit more detail (at least 10 characters).",
    },
}
_TXT2 = {
    "uz": {
        "tm_menu": ("🧠 <b>Test markazi</b>\n\n"
                    "📖 <b>Materialdan test tuzish</b> — {n1} 🪙\n"
                    "Darslik (PDF), Word, rasm yoki matn yuborasiz → bot shu material asosida yangi test tuzadi. "
                    "Kerakli betlar yoki bobni ko'rsatish, o'z <b>namuna testingizni</b> berish mumkin.\n\n"
                    "🎯 <b>Tayyor testni quizga aylantirish</b> — {n2} 🪙\n"
                    "Qo'lingizdagi tayyor testni (Word, PDF, rasm, matn) yuborasiz → bot uni Telegram ichida "
                    "yechiladigan interaktiv quizga aylantiradi. Guruhga forward qilsangiz, talabalar ham yechadi."),
        "tm_b1": "📖 Materialdan test tuzish", "tm_b2": "🎯 Tayyor testni quizga aylantirish",
        "ft_namuna": ("📋 <b>Namuna testingiz bormi?</b>\n\n"
                      "Namuna yuborsangiz (Word, PDF, rasm yoki matn), savollar aynan shu ko'rinishda tuziladi: "
                      "variantlar soni, savol uslubi va uzunligi. Namuna <b>Word</b> bo'lsa, natija o'sha faylning "
                      "sahifa sozlamalari, shrifti va kolontitullari bilan chiqadi.\n\n"
                      "Namuna bo'lmasa — standart formatni tanlang:"),
        "ft_std": "➡️ Standart format (A, B, C, D)",
        "fq_start": ("🎯 <b>Tayyor testni quizga aylantirish</b> — {narx} 🪙\n\n"
                     "Tayyor testni yuboring: 📄 Word, 📕 PDF, 📷 rasm yoki ✍️ matn.\n"
                     "• To'g'ri javoblar belgilangan yoki oxirida kalit bo'lsa — o'sha olinadi\n"
                     "• Belgilanmagan bo'lsa — AI aniqlaydi va buni alohida aytadi\n"
                     "• 50 tagacha savol"),
        "fq_wait": "🎯 Test o'qilmoqda va quizga aylantirilmoqda... (20-60 soniya)",
        "fq_done": ("✅ <b>Quiz tayyor: {n} ta savol</b>{taxmin}\n\n"
                    "«Boshlash» tugmasini bosing — savollar shu yerda ketma-ket chiqadi.\n"
                    "💡 Savollarni guruhingizga <b>forward</b> qilsangiz, talabalar ham yechishi mumkin."),
        "fq_taxmin": "\n⚠️ {k} ta savolda to'g'ri javob belgilanmagan edi — AI aniqladi, tekshirib oling.",
        "fq_start_btn": "▶️ Quizni boshlash",
        "fq_none": "⚠️ Matndan test savollari topilmadi. Savol va A/B/C/D variantlari bor faylni yuboring.",
    },
    "ru": {
        "tm_menu": ("🧠 <b>Тест-центр</b>\n\n"
                    "📖 <b>Тест по материалу</b> — {n1} 🪙\n"
                    "Отправьте учебник (PDF), Word, фото или текст → бот составит новый тест по материалу. "
                    "Можно указать страницы/главу и дать свой <b>образец теста</b>.\n\n"
                    "🎯 <b>Готовый тест → квиз</b> — {n2} 🪙\n"
                    "Отправьте готовый тест (Word, PDF, фото, текст) → бот превратит его в интерактивный квиз "
                    "в Telegram. Перешлите в группу — студенты тоже смогут решать."),
        "tm_b1": "📖 Тест по материалу", "tm_b2": "🎯 Готовый тест → квиз",
        "ft_namuna": ("📋 <b>Есть образец теста?</b>\n\n"
                      "Если отправите образец (Word, PDF, фото или текст), вопросы будут составлены точно в таком же виде: "
                      "число вариантов, стиль и длина вопросов. Если образец в <b>Word</b>, результат сохранит его "
                      "поля, шрифт и колонтитулы.\n\nНет образца — выберите стандартный формат:"),
        "ft_std": "➡️ Стандартный формат (A, B, C, D)",
        "fq_start": ("🎯 <b>Готовый тест → квиз</b> — {narx} 🪙\n\n"
                     "Отправьте готовый тест: 📄 Word, 📕 PDF, 📷 фото или ✍️ текст.\n"
                     "• Если правильные ответы отмечены или есть ключ — берём их\n"
                     "• Если нет — ИИ определит и предупредит об этом\n"
                     "• до 50 вопросов"),
        "fq_wait": "🎯 Читаю тест и делаю квиз... (20-60 секунд)",
        "fq_done": ("✅ <b>Квиз готов: {n} вопросов</b>{taxmin}\n\n"
                    "Нажмите «Начать» — вопросы появятся здесь по очереди.\n"
                    "💡 <b>Перешлите</b> вопросы в группу — студенты тоже смогут решать."),
        "fq_taxmin": "\n⚠️ В {k} вопросах ответ не был отмечен — его определил ИИ, проверьте.",
        "fq_start_btn": "▶️ Начать квиз",
        "fq_none": "⚠️ В тексте не найдено тестовых вопросов. Отправьте файл с вопросами и вариантами A/B/C/D.",
    },
    "en": {
        "tm_menu": ("🧠 <b>Quiz center</b>\n\n"
                    "📖 <b>Quiz from material</b> — {n1} 🪙\n"
                    "Send a textbook (PDF), Word file, photo or text → the bot writes a new quiz from it. "
                    "You can choose pages/chapter and give your own <b>sample test</b>.\n\n"
                    "🎯 <b>Ready test → interactive quiz</b> — {n2} 🪙\n"
                    "Send an existing test (Word, PDF, photo, text) → the bot turns it into an interactive Telegram quiz. "
                    "Forward it to your group so students can take it too."),
        "tm_b1": "📖 Quiz from material", "tm_b2": "🎯 Ready test → quiz",
        "ft_namuna": ("📋 <b>Do you have a sample test?</b>\n\n"
                      "Send a sample (Word, PDF, photo or text) and the questions will follow it exactly: number of "
                      "options, question style and length. If the sample is a <b>Word</b> file, the result keeps its "
                      "margins, font and headers.\n\nNo sample — choose the standard format:"),
        "ft_std": "➡️ Standard format (A, B, C, D)",
        "fq_start": ("🎯 <b>Ready test → quiz</b> — {narx} 🪙\n\n"
                     "Send your test: 📄 Word, 📕 PDF, 📷 photo or ✍️ text.\n"
                     "• If correct answers are marked or there is a key — they are used\n"
                     "• If not — AI determines them and tells you\n"
                     "• up to 50 questions"),
        "fq_wait": "🎯 Reading the test and building the quiz... (20-60 seconds)",
        "fq_done": ("✅ <b>Quiz ready: {n} questions</b>{taxmin}\n\n"
                    "Press «Start» — the questions appear here one by one.\n"
                    "💡 <b>Forward</b> them to your group so students can take it too."),
        "fq_taxmin": "\n⚠️ {k} questions had no marked answer — AI determined it, please check.",
        "fq_start_btn": "▶️ Start quiz",
        "fq_none": "⚠️ No test questions found. Send a file with questions and A/B/C/D options.",
    },
}
for _til, _d in _TXT.items():
    X.T[_til].update(_d)
    X.T[_til].update(_TXT2[_til])

_t, _send, _til_of, _qadam, _q = X._t, X._send, X._til, X._qadam, X._q


def _api():
    return X._api


def _bot():
    return X._bot


# ═══════════════════════════════════════════════════════════════════
# UMUMIY: Telegram ichida interaktiv test
# ═══════════════════════════════════════════════════════════════════
def _valid_quiz(raw, cap: int) -> list:
    out = []
    for q in raw or []:
        if not isinstance(q, dict):
            continue
        vs = [re.sub(r"^[A-Ea-e][).]\s+", "", str(v).strip())[:100] for v in (q.get("variantlar") or [])
              if str(v).strip()][:5]
        try:
            tg = int(q.get("togri", -1))
        except (TypeError, ValueError):
            continue
        if len(vs) < 2 or not (0 <= tg < len(vs)) or len(set(vs)) != len(vs):
            continue
        out.append({"savol": str(q.get("savol", "")).strip()[:300] or "?", "variantlar": vs, "togri": tg,
                    "izoh": str(q.get("izoh", "")).strip()[:200], "taxminiy": bool(q.get("taxminiy"))})
    return out[:cap]


def send_quiz(cid: int, test: list):
    _send(cid, _t(cid, "imtihon_quiz_start"))
    for i, q in enumerate(test, 1):
        try:
            _bot().send_poll(cid, f"{i}/{len(test)}. {q['savol']}"[:300], q["variantlar"], type="quiz",
                             correct_option_id=q["togri"], is_anonymous=False, explanation=q["izoh"] or None)
            time.sleep(0.4)
        except Exception as e:
            log.warning(f"[{cid}] quiz poll xato: {e}")


# ═══════════════════════════════════════════════════════════════════
# 1. FAYLDAN TEST
# ═══════════════════════════════════════════════════════════════════
_PAGES_RE = re.compile(r"(\d{1,4})\s*[-–—]\s*(\d{1,4})")


# AI ga yuboriladigan material hajmi (token tejash): ~24 000 belgi ≈ 7-9 ming token
MATERIAL_LIMIT = 24000


def sample_text(text: str, limit: int = MATERIAL_LIMIT, parts: int = 8) -> str:
    """
    Katta matndan (butun kitob) AI ga butunini emas, teng oraliqlardan olingan bo'laklarni yuboradi.
    Shunda 300 betlik kitob ham bitta oddiy so'rov narxida qayta ishlanadi.
    """
    text = text or ""
    if len(text) <= limit:
        return text
    win = limit // parts
    step = (len(text) - win) / (parts - 1)
    chunks = []
    for i in range(parts):
        a = int(i * step)
        a = text.rfind("\n", max(0, a - 200), a) + 1 or a
        chunks.append(text[a:a + win].strip())
    return "\n…\n".join(chunks)


def ftest_boshlash(msg):
    cid = msg.chat.id
    X._new_session(cid, "ftest", "tm_menu")
    kb = InlineKeyboardMarkup(row_width=1)
    kb.add(InlineKeyboardButton(_t(cid, "tm_b1"), callback_data="tm_material"),
           InlineKeyboardButton(_t(cid, "tm_b2"), callback_data="tm_quiz"))
    _send(cid, _t(cid, "tm_menu", n1=X.NARX["ftest"], n2=X.NARX["quiz"]), reply_markup=kb)


def _tm_mode(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "tm_menu":
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    if call.data == "tm_material":
        if not X._coins_ok(cid, "ftest"):
            return
        _api().update_holat(cid, "qadam", "ft_src")
        _send(cid, _t(cid, "ft_start", narx=X.NARX["ftest"]), reply_markup=_api().back_kb(cid))
    else:
        if not X._coins_ok(cid, "quiz"):
            return
        _api().update_holat(cid, "qadam", "fq_src")
        _send(cid, _t(cid, "fq_start", narx=X.NARX["quiz"]), reply_markup=_api().back_kb(cid))


def _ft_src(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    src = X._capture(msg)
    if not src:
        _send(cid, _t(cid, "ft_bad_src")); return
    if src["type"] == "text":
        src["text"] = security.sanitize_long_text(src["text"], 30000)
        if len(src["text"]) < 40:
            _send(cid, _t(cid, "ft_short")); return
    _api().update_holat(cid, "ft_src", src)
    _api().update_holat(cid, "ft_qism", (msg.caption or "").strip()[:120])
    _api().update_holat(cid, "qadam", "ft_qism")
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton(_t(cid, "ft_all"), callback_data="ftq_all"))
    _send(cid, _t(cid, "ft_qism"), reply_markup=kb)
    X._log_start(msg, "🧠 Fayldan test", src.get("text", f"[{src['type']}] {src.get('name', '')}")[:300])


def _ask_count(cid):
    _api().update_holat(cid, "qadam", "ft_soni")
    kb = InlineKeyboardMarkup(row_width=3)
    kb.add(*[InlineKeyboardButton(f"{n}", callback_data=f"ftn_{n}") for n in (10, 20, 30)])
    _send(cid, _t(cid, "ft_soni"), reply_markup=kb)


def _ft_qism_text(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    _api().update_holat(cid, "ft_qism", security.sanitize_long_text(msg.text or "", 120))
    _ask_count(cid)


def _ft_qism_all(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "ft_qism":
        return
    _api().update_holat(cid, "ft_qism", "")
    _ask_count(cid)


def _ft_soni(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "ft_soni":
        return
    _api().update_holat(cid, "ft_soni", max(5, min(30, int(call.data.split("_")[1]))))
    _api().update_holat(cid, "qadam", "ft_daraja")
    kb = InlineKeyboardMarkup(row_width=2)
    kb.add(*[InlineKeyboardButton(_t(cid, f"d_{d}"), callback_data=f"ftd_{d}") for d in ("oson", "orta", "qiyin", "aralash")])
    _send(cid, _t(cid, "ft_daraja"), reply_markup=kb)


def _pdf_text(file_id: str, pages: tuple | None, max_pages: int = 300) -> str:
    import fitz
    data = X._download(file_id)
    with fitz.open(stream=data, filetype="pdf") as pdf:
        n = pdf.page_count
        if pages:
            a, b = max(1, pages[0]), min(n, pages[1])
            idx = range(a - 1, b) if a <= b else range(0, min(n, 40))
        else:
            idx = range(0, min(n, max_pages))
        text = "\n".join(pdf[i].get_text() for i in idx)
        if len(text.strip()) < 200:      # skanerlangan PDF — rasm sifatida o'qiymiz
            imgs = [pdf[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).tobytes("jpeg") for i in list(idx)[:8]]
            text = X._ocr_images(imgs, "uz")
    return text


def _ft_daraja(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "ft_daraja":
        return
    _api().update_holat(cid, "ft_daraja", call.data.split("_", 1)[1])
    _api().update_holat(cid, "qadam", "ft_namuna")
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton(_t(cid, "ft_std"), callback_data="ftnm_skip"))
    _send(cid, _t(cid, "ft_namuna"), reply_markup=kb)


def _ft_namuna_skip(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "ft_namuna":
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _api().update_holat(cid, "ft_namuna_src", None)
    _ft_start_job(cid)


def _ft_namuna_input(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    src = X._capture(msg)
    if not src:
        _send(cid, _t(cid, "ft_bad_src")); return
    _api().update_holat(cid, "ft_namuna_src", src)
    _ft_start_job(cid)


def _ft_start_job(cid: int):
    _api().update_holat(cid, "qadam", "ft_jarayon")
    h = dict(_api().holat.get(cid, {}))
    daraja = h.get("ft_daraja", "aralash")
    namuna_src = h.get("ft_namuna_src")
    src, qism, soni, til = h.get("ft_src") or {}, h.get("ft_qism", ""), h.get("ft_soni", 10), _til_of(cid)
    daraja_nom = {"oson": "oson (asosiy tushunchalar)", "orta": "o'rta",
                  "qiyin": "qiyin (tahlil va qo'llash)", "aralash": "aralash: oson, o'rta va qiyin"}[daraja]

    def worker(cid, prog):
        m = _PAGES_RE.search(qism or "")
        if src.get("type") == "document" and src.get("name", "").endswith(".pdf"):
            matn = _pdf_text(src["file_id"], (int(m.group(1)), int(m.group(2))) if m else None)
        else:
            matn = X._extract(src, til)
        matn = sample_text(security.sanitize_long_text(matn, 400000))     # token tejash
        if len(matn) < 80:
            raise X.UserFacingError("ft_short")
        namuna_txt, namuna_docx = "", None
        if namuna_src:
            namuna_txt = security.sanitize_long_text(X._extract(namuna_src, til), 4000)[:4000]
            if namuna_src.get("type") == "document" and namuna_src.get("name", "").endswith(".docx"):
                namuna_docx = X._download(namuna_src["file_id"])
        prog.update("🧠 …")
        fokus = (f"Faqat materialning shu qismidan savol tuz: «{qism}»." if qism and not m else
                 "Butun materialdan teng taqsimlab savol tuz.")
        namuna_qoida = (f"""
NAMUNA TEST (format uchun, MAZMUNINI olma!): savollarni AYNAN shu namunadagidek tuz — variantlar soni (2-5),
savol uslubi va uzunligi, variantlar uslubi, til. Faqat bitta to'g'ri javobli savollar bo'lsin.
{security.wrap_user_block(namuna_txt, "NAMUNA")}
""" if namuna_txt else "Har bir savolda 4 ta variant (A, B, C, D) bo'lsin.")
        prompt = f"""Sen tajribali test tuzuvchi o'qituvchisan (DTM, CEFR, IELTS, milliy sertifikat uslubida).
Quyidagi o'quv materiali asosida {soni} ta test savoli tuz. {fokus}
Qiyinlik: {daraja_nom}.
Savollar va variantlar MATERIAL QAYSI TILDA BO'LSA O'SHA TILDA bo'lsin (masalan ingliz tilidagi matn uchun inglizcha).
Faqat materialdagi ma'lumotga asoslan, o'zingdan fakt to'qima. Har bir savolning faqat bitta to'g'ri javobi bo'lsin,
noto'g'ri variantlar ishonarli bo'lsin, to'g'ri javob o'rni aralash bo'lsin.
{namuna_qoida}
{security.wrap_user_block(matn, "O'QUV MATERIALI")}

FAQAT JSON qaytar:
{{"mavzu": "material mavzusi (3-8 so'z)",
 "test": [{{"savol": "savol (250 belgigacha)", "variantlar": ["...", "...", "...", "..."], "togri": 0,
           "izoh": "nega to'g'ri — 1 gap (150 belgigacha)"}}]}}
"variantlar" ichida harf (A, B...) yozma. "test" ichida aynan {soni} ta savol bo'lsin. Markdown ishlatma."""
        data = X._ai_json(prompt, 8000)
        test = _valid_quiz(data.get("test"), soni)
        if len(test) < max(3, soni // 3):
            raise RuntimeError(f"test juda kam: {len(test)}")
        prog.update("📄 Word fayl tayyorlanmoqda...")
        mavzu = str(data.get("mavzu") or "Test")
        fname = _ft_docx(cid, mavzu, test, daraja, til, base_docx=namuna_docx)
        _api().update_holat(cid, "ftest_test", test)
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton(_t(cid, "imtihon_quiz_btn"), callback_data="ft_quiz"))
        try:
            prog.delete()
            cap = f"{_t(cid, 'ft_done')}\n\n{_t(cid, 'coins', c=_api().get_coins(cid))}"
            X._send_doc(cid, fname, X._safe_name(mavzu, "Test"), cap[:1024], "🧠 Fayldan test",
                        f"{mavzu} | {len(test)} savol | {daraja}", kb=kb)
            _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        finally:
            X._cleanup(fname)
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "ftest", "ft_wait", worker, "Fayldan test")


def _ft_docx(cid, mavzu, test, daraja, til, base_docx: bytes | None = None) -> str:
    """Oq, toza, chop etishga tayyor test (rangsiz). base_docx — foydalanuvchi namunasi (sahifa, shrift, kolontitul)."""
    L = {"uz": ("TEST", "Savollar soni", "Daraja", "F.I.Sh.", "Guruh/sinf", "Sana", "JAVOBLAR KALITI", "Izohlar",
                {"oson": "oson", "orta": "o'rta", "qiyin": "qiyin", "aralash": "aralash"}),
         "ru": ("ТЕСТ", "Количество вопросов", "Уровень", "Ф.И.О.", "Группа/класс", "Дата", "КЛЮЧ ОТВЕТОВ", "Пояснения",
                {"oson": "лёгкий", "orta": "средний", "qiyin": "сложный", "aralash": "смешанный"}),
         "en": ("QUIZ", "Questions", "Level", "Name", "Group/class", "Date", "ANSWER KEY", "Explanations",
                {"oson": "easy", "orta": "medium", "qiyin": "hard", "aralash": "mixed"})}[til]
    doc = None
    if base_docx:
        try:
            doc = Document(io.BytesIO(base_docx))
            body = doc.element.body
            for el in list(body):
                if el.tag != qn("w:sectPr"):          # namuna mazmunini olib tashlaymiz, sozlamalar qoladi
                    body.remove(el)
        except Exception as e:
            log.warning(f"Namuna docx ochilmadi, standart format: {e}")
            doc = None
    if doc is None:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(2)
        sec.top_margin = sec.bottom_margin = Cm(1.6)
        st = doc.styles["Normal"]
        st.font.name = "Times New Roman"
        st.font.size = Pt(13)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        st.paragraph_format.space_after = Pt(0)
    sec = doc.sections[0]

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{L[0]}: {mavzu.upper()}"); r.bold = True; r.font.size = Pt(15)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"{L[1]}: {len(test)}   •   {L[2]}: {L[8].get(daraja, daraja)}").italic = True
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run(f"{L[3]}: ______________________   {L[4]}: ________   {L[5]}: ________")
    doc.add_paragraph()
    for i, q in enumerate(test, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.keep_with_next = True
        p.add_run(f"{i}. ").bold = True
        p.add_run(q["savol"]).bold = True
        for j, v in enumerate(q["variantlar"]):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.8)
            p.paragraph_format.keep_with_next = j < len(q["variantlar"]) - 1
            p.add_run(f"{'ABCDE'[j]})  {v}")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(L[6]); r.bold = True; r.font.size = Pt(14)
    cols = 5
    rows = (len(test) + cols - 1) // cols
    t = doc.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, q in enumerate(test):
        r_, c_ = i % rows, i // rows
        cell = t.cell(r_, c_)
        cell.text = ""
        run = cell.paragraphs[0].add_run(f"{i + 1} — {'ABCDE'[q['togri']]}")
        run.font.size = Pt(12)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if any(q["izoh"] for q in test):
        doc.add_paragraph()
        p = doc.add_paragraph(); p.add_run(L[7]).bold = True
        for i, q in enumerate(test, 1):
            if q["izoh"]:
                p = doc.add_paragraph()
                p.add_run(f"{i}. ({'ABCDE'[q['togri']]}) ").bold = True
                r = p.add_run(q["izoh"]); r.font.size = Pt(11)
    if not base_docx:
        try:
            _api().add_page_numbers(sec)
        except Exception:
            pass
    return X._save(doc, mavzu, "_Test.docx")


def _ft_quiz(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    test = _api().holat.get(cid, {}).get("ftest_test") or []
    if test:
        try:
            _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
        except Exception:
            pass
        send_quiz(cid, test)


# ── Tayyor testni quizga aylantirish ──────────────────────────────
def _fq_src(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    src = X._capture(msg)
    if not src:
        _send(cid, _t(cid, "ft_bad_src")); return
    _api().update_holat(cid, "qadam", "fq_jarayon")
    til = _til_of(cid)
    X._log_start(msg, "🎯 Testdan quiz", src.get("text", f"[{src['type']}] {src.get('name', '')}")[:300])

    def worker(cid, prog):
        matn = security.sanitize_long_text(X._extract(src, til), 30000)[:30000]
        if len(matn) < 30:
            raise X.UserFacingError("fq_none")
        prompt = f"""Quyidagi matn — TAYYOR TEST. Uni JSON ga o'tkaz. Savol va variantlar matnini O'ZGARTIRMA
(faqat ortiqcha raqam/harf belgilarini olib tashla), tilini ham o'zgartirma.
To'g'ri javob: belgilangan bo'lsa (*, +, qalin, "javob:", oxiridagi kalit jadvali) — shuni ol, "taxminiy": false.
Hech qayerda belgilanmagan bo'lsa — o'zing aniqla va "taxminiy": true qo'y.
Variantsiz (ochiq) savollarni tashlab ket. Ko'pi bilan 50 ta savol.

{security.wrap_user_block(matn, "TAYYOR TEST")}

FAQAT JSON qaytar:
{{"mavzu": "test mavzusi (3-8 so'z)",
 "test": [{{"savol": "...", "variantlar": ["...", "..."], "togri": 0, "taxminiy": false, "izoh": ""}}]}}
Markdown ishlatma."""
        data = X._ai_json(prompt, 8000)
        test = _valid_quiz(data.get("test"), 50)
        if not test:
            raise X.UserFacingError("fq_none")
        mavzu = str(data.get("mavzu") or "Quiz")
        _api().update_holat(cid, "ftest_test", test)
        k = sum(1 for q in test if q.get("taxminiy"))
        fname = _ft_docx(cid, mavzu, test, "aralash", til)
        kb = InlineKeyboardMarkup()
        kb.add(InlineKeyboardButton(_t(cid, "fq_start_btn"), callback_data="ft_quiz"))
        try:
            prog.delete()
            cap = _t(cid, "fq_done", n=len(test), taxmin=_t(cid, "fq_taxmin", k=k) if k else "")
            X._send_doc(cid, fname, X._safe_name(mavzu, "Quiz"),
                        (cap + "\n\n" + _t(cid, "coins", c=_api().get_coins(cid)))[:1024],
                        "🎯 Testdan quiz", f"{mavzu} | {len(test)} savol | taxminiy {k}", kb=kb)
            _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        finally:
            X._cleanup(fname)
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "quiz", "fq_wait", worker, "Testdan quiz")


# ═══════════════════════════════════════════════════════════════════
# 2. MATNNI OVOZGA (edge-tts)
# ═══════════════════════════════════════════════════════════════════
TTS_LIMIT = 10000
VOICES = [
    ("uz-UZ-MadinaNeural", "🇺🇿 Madina"), ("uz-UZ-SardorNeural", "🇺🇿 Sardor"),
    ("ru-RU-SvetlanaNeural", "🇷🇺 Светлана"), ("ru-RU-DmitryNeural", "🇷🇺 Дмитрий"),
    ("en-US-AriaNeural", "🇺🇸 Aria"), ("en-GB-RyanNeural", "🇬🇧 Ryan"),
    ("kk-KZ-AigulNeural", "🇰🇿 Айгүл"), ("tr-TR-EmelNeural", "🇹🇷 Emel"),
]
_VOICE_IDS = {v for v, _ in VOICES}


def _voice_kb(prefix: str):
    kb = InlineKeyboardMarkup(row_width=2)
    kb.add(*[InlineKeyboardButton(name, callback_data=f"{prefix}{i}") for i, (_, name) in enumerate(VOICES)])
    return kb


def tts_boshlash(msg):
    cid = msg.chat.id
    if not X._coins_ok(cid, "tts"):
        return
    X._new_session(cid, "tts", "tts_src")
    _send(cid, _t(cid, "tts_start", narx=X.NARX["tts"], limit=f"{TTS_LIMIT:,}".replace(",", " "),
                  daq=TTS_LIMIT // 900), reply_markup=_api().back_kb(cid))


def _tts_src(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    src = X._capture(msg)
    if not src:
        _send(cid, _t(cid, "ft_bad_src")); return
    _api().update_holat(cid, "tts_src", src)
    _api().update_holat(cid, "qadam", "tts_voice")
    _send(cid, _t(cid, "tts_voice"), reply_markup=_voice_kb("ttsv_"))


def _clean_for_speech(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"[*#_`|•►▪■□●◆✓✔✅❌]+", " ", text)
    text = re.sub(r"[\U0001F000-\U0001FAFF☀-➿]", " ", text)   # emoji
    text = re.sub(r"[ \t]{2,}", " ", text)
    return re.sub(r"\n{2,}", "\n", text).strip()


def _chunks(text: str, size: int = 2500) -> list:
    out, cur = [], ""
    for sent in re.split(r"(?<=[.!?;:\n])\s+", text):
        if len(cur) + len(sent) + 1 > size and cur:
            out.append(cur); cur = ""
        cur += sent + " "
    if cur.strip():
        out.append(cur)
    return out


def synthesize(text: str, voice: str) -> bytes:
    import edge_tts

    async def _one(part: str) -> bytes:
        buf = bytearray()
        async for ch in edge_tts.Communicate(part, voice).stream():
            if ch.get("type") == "audio":
                buf.extend(ch["data"])
        return bytes(buf)

    audio = bytearray()
    for part in _chunks(text):
        for attempt in range(3):
            try:
                audio.extend(asyncio.run(_one(part)))
                break
            except Exception as e:
                if attempt == 2:
                    raise
                log.warning(f"TTS qayta urinish ({attempt + 1}): {e}")
                time.sleep(2)
    if not audio:
        raise RuntimeError("TTS bo'sh audio qaytardi")
    return bytes(audio)


def _tts_run(cid: int, voice: str, text_or_src):
    til = _til_of(cid)

    def worker(cid, prog):
        text = text_or_src if isinstance(text_or_src, str) else X._extract(text_or_src, til)
        text = _clean_for_speech(security.sanitize_long_text(text, TTS_LIMIT * 2))
        if len(text) < 5:
            raise X.UserFacingError("tts_none")
        cut = len(text) > TTS_LIMIT
        text = text[:TTS_LIMIT]
        audio = synthesize(text, voice)
        prog.delete()
        title = (text.split("\n", 1)[0][:60] or "Audio").strip()
        f = io.BytesIO(audio)
        f.name = "academic_bot_audio.mp3"
        sent = _bot().send_audio(cid, f, title=title, performer="Academic Bot",
                                 caption=f"{_t(cid, 'tts_done')}\n{_t(cid, 'coins', c=_api().get_coins(cid))}",
                                 parse_mode="HTML")
        if cut:
            _send(cid, _t(cid, "tts_cut", limit=f"{TTS_LIMIT:,}".replace(",", " ")))
        _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        try:
            _api().log_javob(cid=cid, javob_turi="🎧 Ovozli matn", javob_matn=text[:500],
                             file_id=sent.audio.file_id, file_type="audio", visible_name=f.name)
        except Exception:
            pass
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "tts", "tts_wait", worker, "Ovozli matn")


def _tts_voice(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "tts_voice":
        return
    try:
        voice = VOICES[int(call.data.split("_")[1])][0]
    except (ValueError, IndexError):
        return
    _api().update_holat(cid, "qadam", "tts_jarayon")
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _tts_run(cid, voice, _api().holat.get(cid, {}).get("tts_src") or {})


def _tts_last(call):
    """Konspekt / mustaqil ish ostidagi '🎧 Ovozli tinglash' tugmasi."""
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if not _api().holat.get(cid, {}).get("tts_text"):
        _send(cid, _t(cid, "tts_none")); return
    if not X._coins_ok(cid, "tts"):
        return
    _send(cid, _t(cid, "tts_voice"), reply_markup=_voice_kb("ttsl_"))


def _tts_last_voice(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    text = _api().holat.get(cid, {}).get("tts_text")
    if not text:
        _send(cid, _t(cid, "tts_none")); return
    try:
        voice = VOICES[int(call.data.split("_")[1])][0]
    except (ValueError, IndexError):
        return
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    _tts_run(cid, voice, text)


# ═══════════════════════════════════════════════════════════════════
# 3. SINF SARDORI / YETAKCHILAR
# ═══════════════════════════════════════════════════════════════════
SD_TURLAR = {
    "bayonnoma": {
        "nom": {"uz": "📋 Yig'ilish bayonnomasi", "ru": "📋 Протокол собрания", "en": "📋 Meeting minutes"},
        "sor": {"uz": "Yig'ilish haqida yozing: <b>sana, sinf/guruh, nechta kishi qatnashdi, kun tartibidagi masalalar, "
                      "nima gaplashildi va qanday qaror qabul qilindi</b> (bilganlaringizni).",
                "ru": "Опишите собрание: <b>дата, класс/группа, сколько участников, повестка, что обсуждали и какие решения "
                      "приняли</b> (что знаете).",
                "en": "Describe the meeting: <b>date, class/group, attendance, agenda, what was discussed and decided</b>."},
        "tuzilma": ("Rasmiy YIG'ILISH BAYONNOMASI: sarlavha (№__), sana va joy, raislik qiluvchi va kotib, "
                    "qatnashdilar soni, KUN TARTIBI (raqamlangan), har bir masala bo'yicha ESHITILDI / SO'ZGA CHIQDILAR / "
                    "QAROR QILINDI bo'limlari, oxirida rais va kotib imzosi uchun joy."),
    },
    "ssenariy": {
        "nom": {"uz": "🎭 Tadbir ssenariysi", "ru": "🎭 Сценарий мероприятия", "en": "🎭 Event script"},
        "sor": {"uz": "Tadbir haqida yozing: <b>nomi/bayram, kimlar uchun, qayerda, davomiyligi, nimalar bo'lishi kerak</b> "
                      "(she'r, sahna ko'rinishi, o'yinlar...).",
                "ru": "Опишите мероприятие: <b>название/праздник, для кого, где, длительность, что должно быть</b>.",
                "en": "Describe the event: <b>name/occasion, audience, place, duration, what it should include</b>."},
        "tuzilma": ("TADBIR SSENARIYSI: maqsad, vaqt va joy, ishtirokchilar va mas'ullar, kerakli jihozlar, "
                    "dastur jadvali (vaqt | qism | mas'ul), tadbirning batafsil borishi — BOSHLOVCHI MATNLARI so'zma-so'z, "
                    "she'r/qo'shiq/o'yin uchun joylar, yakunlovchi so'z."),
    },
    "tarbiyaviy": {
        "nom": {"uz": "🌱 Tarbiyaviy soat rejasi", "ru": "🌱 План классного часа", "en": "🌱 Class hour plan"},
        "sor": {"uz": "Yozing: <b>mavzu, sinf (nechanchi), davomiyligi</b> va xohishlaringiz.",
                "ru": "Напишите: <b>тема, класс, длительность</b> и пожелания.",
                "en": "Write: <b>topic, grade, duration</b> and any wishes."},
        "tuzilma": ("TARBIYAVIY SOAT ISHLANMASI: mavzu, maqsadlar (ta'limiy, tarbiyaviy, rivojlantiruvchi), "
                    "kerakli jihozlar, mashg'ulot bosqichlari jadvali (bosqich | vaqt | o'qituvchi/sardor faoliyati | "
                    "o'quvchilar faoliyati), suhbat savollari, interaktiv mashq yoki o'yin, xulosa, uyga vazifa."),
    },
    "elon": {
        "nom": {"uz": "📣 E'lon / taklifnoma", "ru": "📣 Объявление / приглашение", "en": "📣 Notice / invitation"},
        "sor": {"uz": "Yozing: <b>nima haqida, qachon, qayerda, kimlar uchun</b>.",
                "ru": "Напишите: <b>о чём, когда, где, для кого</b>.",
                "en": "Write: <b>what, when, where, for whom</b>."},
        "tuzilma": ("Chiroyli va qisqa E'LON yoki TAKLIFNOMA: jalb qiluvchi sarlavha, asosiy ma'lumot "
                    "(nima, qachon, qayerda), qisqa tavsif, aloqa uchun joy. 1 betdan oshmasin."),
    },
    "hisobot": {
        "nom": {"uz": "📊 Faoliyat hisoboti", "ru": "📊 Отчёт о работе", "en": "📊 Activity report"},
        "sor": {"uz": "Yozing: <b>qaysi davr, sinf/guruh, qanday tadbirlar o'tkazildi, natijalar, muammolar</b>.",
                "ru": "Напишите: <b>период, класс/группа, проведённые мероприятия, результаты, проблемы</b>.",
                "en": "Write: <b>period, class/group, activities held, results, problems</b>."},
        "tuzilma": ("SINF SARDORI FAOLIYATI HISOBOTI: kirish, o'tkazilgan tadbirlar jadvali (sana | tadbir | qatnashchilar | natija), "
                    "davomat va o'zlashtirish bo'yicha qisqa tahlil, yutuqlar, muammolar va takliflar, imzo uchun joy."),
    },
}


def sardor_boshlash(msg):
    cid = msg.chat.id
    if not X._coins_ok(cid, "sardor"):
        return
    X._new_session(cid, "sardor", "sd_tur")
    til = _til_of(cid)
    kb = InlineKeyboardMarkup(row_width=1)
    kb.add(*[InlineKeyboardButton(v["nom"].get(til, v["nom"]["uz"]), callback_data=f"sdt_{k}") for k, v in SD_TURLAR.items()])
    _send(cid, _t(cid, "sd_start", narx=X.NARX["sardor"]), reply_markup=kb)


def _sd_tur(call):
    cid = call.message.chat.id
    _bot().answer_callback_query(call.id)
    if _qadam(cid) != "sd_tur":
        return
    tur = call.data.split("_", 1)[1]
    if tur not in SD_TURLAR:
        return
    til = _til_of(cid)
    _api().update_holat(cid, "sd_tur", tur)
    _api().update_holat(cid, "qadam", "sd_info")
    try:
        _bot().edit_message_reply_markup(cid, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    v = SD_TURLAR[tur]
    _send(cid, f"<b>{v['nom'].get(til)}</b>\n\n{v['sor'].get(til, v['sor']['uz'])}", reply_markup=_api().back_kb(cid))


def _sd_info(msg):
    cid = msg.chat.id
    if _api().is_back(msg):
        return
    info = security.sanitize_long_text(msg.text or "", 3000)
    if len(info) < 10:
        _send(cid, _t(cid, "sd_short")); return
    tur = _api().holat.get(cid, {}).get("sd_tur", "bayonnoma")
    til = _til_of(cid)
    _api().update_holat(cid, "qadam", "sd_jarayon")
    X._log_start(msg, f"🧑‍🏫 Sardor: {tur}", info[:300])

    def worker(cid, prog):
        prompt = f"""Sen maktab va OTM uchun rasmiy hujjatlar tayyorlovchi tajribali pedagog-metodistsan. {X.LANG_CMD[til]}
Quyidagi ma'lumot asosida hujjat tayyorla. Hujjat turi va tuzilmasi:
{SD_TURLAR[tur]['tuzilma']}

Foydalanuvchi bergan ma'lumot (bo'lmagan joylarni ______ bilan to'ldirish uchun bo'sh qoldir, ism-familiya to'qima):
{security.wrap_user_block(info, "MA'LUMOT")}

FAQAT JSON qaytar:
{{"sarlavha": "hujjat sarlavhasi",
 "bolimlar": [{{"nomi": "bo'lim nomi (bo'lmasa bo'sh)", "matn": "abzaslar, \\n bilan", "bandlar": ["ro'yxat bandi"],
               "jadval": {{"ustunlar": ["..."], "qatorlar": [["...", "..."]]}}}}],
 "imzolar": ["Rais: ____________", "Kotib: ____________"]}}
Keraksiz maydonlarni bo'sh qoldir ("bandlar": [], "jadval": null). Markdown ishlatma."""
        data = X._ai_json(prompt, 6000)
        prog.update("📄 Word fayl tayyorlanmoqda...")
        sar = str(data.get("sarlavha") or SD_TURLAR[tur]["nom"]["uz"])
        fname = _sd_docx(data, sar)
        try:
            prog.delete()
            X._send_doc(cid, fname, X._safe_name(sar, tur.capitalize()),
                        f"{_t(cid, 'sd_done')}\n\n{_t(cid, 'coins', c=_api().get_coins(cid))}"[:1024],
                        f"🧑‍🏫 {SD_TURLAR[tur]['nom']['uz']}", sar)
            _send(cid, _api().tx(cid, "yangi"), reply_markup=_api().main_kb(cid))
        finally:
            X._cleanup(fname)
        _api().update_holat(cid, "qadam", None)

    X._run_job(cid, "sardor", "sd_wait", worker, "Sinf sardori")


def _sd_docx(data: dict, sar: str) -> str:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(2.5), Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.paragraph_format.line_spacing = 1.15
    st.paragraph_format.space_after = Pt(4)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(sar.upper()); r.bold = True; r.font.size = Pt(16)
    doc.add_paragraph()
    for b in data.get("bolimlar") or []:
        if not isinstance(b, dict):
            continue
        if b.get("nomi"):
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.keep_with_next = True
            p.add_run(str(b["nomi"])).bold = True
        for par in str(b.get("matn") or "").split("\n"):
            if par.strip():
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(1.0)
                p.add_run(par.strip())
        for it in b.get("bandlar") or []:
            if str(it).strip():
                doc.add_paragraph(str(it).strip(), style="List Bullet")
        jad = b.get("jadval")
        if isinstance(jad, dict) and jad.get("ustunlar") and jad.get("qatorlar"):
            cols = [str(c) for c in jad["ustunlar"]][:6]
            t = doc.add_table(rows=1, cols=len(cols))
            t.style = "Table Grid"
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for i, c in enumerate(cols):
                cell = t.rows[0].cells[i]
                cell.text = ""
                rr = cell.paragraphs[0].add_run(c); rr.bold = True; rr.font.size = Pt(12)
            for row in jad["qatorlar"]:
                cells = t.add_row().cells
                for i in range(len(cols)):
                    val = str(row[i]) if isinstance(row, list) and i < len(row) else ""
                    cells[i].text = ""
                    cells[i].paragraphs[0].add_run(val).font.size = Pt(12)
            doc.add_paragraph()
    imz = [str(x) for x in (data.get("imzolar") or []) if str(x).strip()]
    if imz:
        doc.add_paragraph()
        for x in imz:
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(10)
            p.add_run(x)
    return X._save(doc, sar, "_Hujjat.docx")


# ═══════════════════════════════════════════════════════════════════
# RO'YXATDAN O'TKAZISH
# ═══════════════════════════════════════════════════════════════════
X.STARTERS.update({"ftest": ftest_boshlash, "tts": tts_boshlash, "sardor": sardor_boshlash})


def register(bot):
    """xizmatlar_yangi.register() dan KEYIN chaqiriladi (X._api, X._bot tayyor bo'ladi)."""
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
                    _send(m.chat.id, _t(m.chat.id, "err"), reply_markup=_api().main_kb(m.chat.id))
                except Exception:
                    pass
        wrapped.__name__ = fn.__name__
        return wrapped

    bot.register_callback_query_handler(guard(_tm_mode), func=lambda c: c.data in ("tm_material", "tm_quiz"))
    bot.register_message_handler(guard(_ft_src), func=_q("ft_src"), content_types=ct)
    bot.register_message_handler(guard(_ft_namuna_input), func=_q("ft_namuna"), content_types=ct)
    bot.register_callback_query_handler(guard(_ft_namuna_skip), func=lambda c: c.data == "ftnm_skip")
    bot.register_message_handler(guard(_fq_src), func=_q("fq_src"), content_types=ct)
    bot.register_message_handler(guard(_ft_qism_text), func=_q("ft_qism"))
    bot.register_callback_query_handler(guard(_ft_qism_all), func=lambda c: c.data == "ftq_all")
    bot.register_callback_query_handler(guard(_ft_soni), func=lambda c: c.data.startswith("ftn_"))
    bot.register_callback_query_handler(guard(_ft_daraja), func=lambda c: c.data.startswith("ftd_"))
    bot.register_callback_query_handler(guard(_ft_quiz), func=lambda c: c.data == "ft_quiz")

    bot.register_message_handler(guard(_tts_src), func=_q("tts_src"), content_types=ct)
    bot.register_callback_query_handler(guard(_tts_voice), func=lambda c: c.data.startswith("ttsv_"))
    bot.register_callback_query_handler(guard(_tts_last), func=lambda c: c.data == "tts_last")
    bot.register_callback_query_handler(guard(_tts_last_voice), func=lambda c: c.data.startswith("ttsl_"))

    bot.register_callback_query_handler(guard(_sd_tur), func=lambda c: c.data.startswith("sdt_"))
    bot.register_message_handler(guard(_sd_info), func=_q("sd_info"))
    log.info("✅ Qo'shimcha xizmatlar ulandi: Fayldan test, Ovozli matn, Sinf sardori")

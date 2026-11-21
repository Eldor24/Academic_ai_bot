"""
tarjima_shim.py — eski kodda qattiq yozilgan o'zbekcha xabarlarni rus/ingliz tilidagi foydalanuvchilar uchun tarjima qiladi
═══════════════════════════════════════════════════════════════════
Botning ba'zi joylarida (validatsiya xatolari, «Qolgan tangalar», esse tekshiruvi va h.k.) xabar to'g'ridan-to'g'ri
o'zbekcha yozilgan edi — rus va ingliz tilini tanlagan foydalanuvchi ham o'zbekcha ko'rardi.
Bu modul bot.send_message / reply_to / answer_callback_query / send_document (caption) / edit_message_text ni
o'raydi: foydalanuvchi tili ru/en bo'lsa, ma'lum o'zbekcha bo'laklar tarjimasi bilan almashtiriladi.
Faqat bot xabari matniga tegadi (hujjatlar va foydalanuvchi matni o'zgarmaydi).
"""
from __future__ import annotations

import re

# (o'zbekcha bo'lak, ruscha, inglizcha) — uzunroqlari oldin
QATORLAR = [
    ("⚠️ Telefon raqam bo'sh bo'lishi mumkin emas.", "⚠️ Номер телефона не может быть пустым.", "⚠️ Phone number cannot be empty."),
    ("⚠️ Telefon formati noto'g'ri!", "⚠️ Неверный формат телефона!", "⚠️ Invalid phone format!"),
    ("⚠️ Email bo'sh bo'lishi mumkin emas.", "⚠️ Email не может быть пустым.", "⚠️ Email cannot be empty."),
    ("⚠️ Email 100 belgidan qisqa bo'lishi kerak.", "⚠️ Email должен быть короче 100 символов.", "⚠️ Email must be under 100 characters."),
    ("⚠️ Email formati noto'g'ri!", "⚠️ Неверный формат email!", "⚠️ Invalid email format!"),
    ("⚠️ Sana formati noto'g'ri!", "⚠️ Неверный формат даты!", "⚠️ Invalid date format!"),
    ("⚠️ Bunday sana mavjud emas.", "⚠️ Такой даты не существует.", "⚠️ This date does not exist."),
    ("⚠️ Faylni yuklab olishda xatolik. Qayta urinib ko'ring.", "⚠️ Ошибка загрузки файла. Попробуйте ещё раз.",
     "⚠️ Failed to download the file. Please try again."),
    ("⚠️ Fayl buzilgan yoki haqiqiy DOCX formatida emas.", "⚠️ Файл повреждён или не является DOCX.",
     "⚠️ The file is damaged or not a real DOCX."),
    ("⚠️ Fayl buzilgan yoki haqiqiy PDF formatida emas.", "⚠️ Файл повреждён или не является PDF.",
     "⚠️ The file is damaged or not a real PDF."),
    ("⚠️ Server xatosi. Qayta urinib ko'ring.", "⚠️ Ошибка сервера. Попробуйте ещё раз.", "⚠️ Server error. Please try again."),
    ("⚠️ Ruxsat etilgan formatlar:", "⚠️ Допустимые форматы:", "⚠️ Allowed formats:"),
    ("Yuborilgan:", "Отправлено:", "Sent:"),
    ("❌ Xato. Qayta urinib ko'ring.", "❌ Ошибка. Попробуйте ещё раз.", "❌ Error. Please try again."),
    ("(Savollar tugadi)", "(Вопросы закончились)", "(No questions left)"),
    ("(Admin — cheksiz savol ✅)", "(Админ — без ограничений ✅)", "(Admin — unlimited ✅)"),
    ("💰 Qolgan tangalar:", "💰 Осталось монет:", "💰 Coins left:"),
    ("• Matn juda qisqa", "• Текст слишком короткий", "• The text is too short"),
    ("• Matn dasturlash kodi ko'rinishida", "• Текст похож на программный код", "• The text looks like program code"),
    ("• Matn asosan URL/havolalardan iborat", "• Текст состоит в основном из ссылок", "• The text consists mostly of links"),
    ("• Matn juda ko'p raqam o'z ichiga oladi", "• В тексте слишком много чисел", "• The text contains too many numbers"),
    ("• Matn ro'yxat/jadval ko'rinishida (esse erkin nasr shaklida yoziladi)",
     "• Текст похож на список/таблицу (эссе пишется связной прозой)",
     "• The text looks like a list/table (an essay is written as continuous prose)"),
    ("• Matn juda kam gapdan iborat", "• В тексте слишком мало предложений", "• The text has too few sentences"),
    ("• Matn juda ko'p emoji o'z ichiga oladi", "• В тексте слишком много эмодзи", "• The text contains too many emoji"),
    ("Misol:", "Пример:", "Example:"),
]
NAQSHLAR = [
    (re.compile(r"\((\d+) ta savol qoldi\)"), r"(осталось вопросов: \1)", r"(\1 questions left)"),
    (re.compile(r"⚠️ Fayl hajmi (\d+) MB dan oshmasligi kerak\."), r"⚠️ Размер файла не должен превышать \1 МБ.",
     r"⚠️ File size must not exceed \1 MB."),
    (re.compile(r" tanlandi$"), " выбран", " selected"),
]


def tarjima(text, til: str):
    if not isinstance(text, str) or til not in ("ru", "en") or not text:
        return text
    i = 1 if til == "ru" else 2
    for row in QATORLAR:
        if row[0] in text:
            text = text.replace(row[0], row[i])
    for rx, ru, en in NAQSHLAR:
        text = rx.sub(ru if til == "ru" else en, text)
    return text


def install(bot, til_of):
    """til_of(chat_id) -> 'uz' | 'ru' | 'en'"""
    def _til(cid):
        try:
            return til_of(int(cid))
        except Exception:
            return "uz"

    orig_send = bot.send_message

    def send_message(chat_id, text, *a, **k):
        return orig_send(chat_id, tarjima(text, _til(chat_id)), *a, **k)
    bot.send_message = send_message

    orig_reply = bot.reply_to

    def reply_to(message, text, *a, **k):
        try:
            text = tarjima(text, _til(message.chat.id))
        except Exception:
            pass
        return orig_reply(message, text, *a, **k)
    bot.reply_to = reply_to

    orig_edit = bot.edit_message_text

    def edit_message_text(text, chat_id=None, *a, **k):
        if chat_id is not None:
            text = tarjima(text, _til(chat_id))
        return orig_edit(text, chat_id, *a, **k)
    bot.edit_message_text = edit_message_text

    orig_doc = bot.send_document

    def send_document(chat_id, document, *a, **k):
        if "caption" in k:
            k["caption"] = tarjima(k["caption"], _til(chat_id))
        return orig_doc(chat_id, document, *a, **k)
    bot.send_document = send_document

    orig_cb = bot.answer_callback_query

    def answer_callback_query(callback_query_id, text=None, *a, **k):
        cid = _CB_CHAT.get(callback_query_id)
        if text and cid:
            text = tarjima(text, _til(cid))
        return orig_cb(callback_query_id, text, *a, **k)
    bot.answer_callback_query = answer_callback_query

    # callback id → chat id (answer_callback_query da chat ma'lum emas)
    orig_proc = bot.process_new_callback_query

    def process_new_callback_query(calls):
        for c in calls:
            try:
                _CB_CHAT[c.id] = c.message.chat.id
            except Exception:
                pass
        if len(_CB_CHAT) > 5000:
            for k_ in list(_CB_CHAT)[:2500]:
                _CB_CHAT.pop(k_, None)
        return orig_proc(calls)
    bot.process_new_callback_query = process_new_callback_query


_CB_CHAT: dict = {}

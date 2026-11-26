"""
marketing.py — Avtomatik foydali postlar (marketing)
═══════════════════════════════════════════════════════════════════
Strategiya:
  • Haftada 2 marta (standart: dushanba va payshanba, 18:30 Toshkent vaqti) har bir foydalanuvchiga
    BITTA xizmat haqida qisqa, sodda post yuboriladi. Har kim postlarni o'z navbatida oladi —
    hamma xizmatlar bilan ~10 haftada tanishib chiqadi, keyin aylanma qayta boshlanadi.
  • Postda «👉 Sinab ko'rish» tugmasi — bosilsa, xizmat darhol ochiladi (menyudan qidirish shart emas).
  • «🔕 Bunday xabarlar kerak emas» tugmasi — foydalanuvchi o'zi o'chiradi (/xabarlar — qayta yoqish).
    Majburiy spam yo'q — botni bloklashlar kamayadi.
  • Yangi foydalanuvchilar avval o'z tanishuv ketma-ketligini oladi (onboarding); u tugamaguncha
    haftalik postlar yuborilmaydi — ikki xil xabar ustma-ust tushmaydi.
  • Botni bloklagan foydalanuvchilarga qayta urinilmaydi (botga qaytib kirsa — yana oladi).
  • Bot o'chib-yonsa ham bir slot ikki marta yuborilmaydi (holat users_db da, bulutga sinxronlanadi);
    server uzoq o'chiq qolib, slotdan 6 soatdan ko'p o'tgan bo'lsa — kechasi yubormaydi, keyingi slotni kutadi.

Admin: /marketing — statistika, /marketing test — keyingi postni o'zingizga ko'rish,
       /marketing yubor — hozir hammaga yuborish, /marketing on|off — avtomatik yuborishni yoqish/o'chirish.
"""
from __future__ import annotations

import copy
import datetime
import logging
import threading
import time

from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup

log = logging.getLogger(__name__)

KUNLAR = (0, 3)          # 0 — dushanba, 3 — payshanba
SOAT, DAQIQA = 18, 30    # Toshkent vaqti
KECHIKISH_SOAT = 6       # slotdan shuncha soat o'tsa — yubormaymiz (tunda bezovta qilmaslik)
TEZLIK = 0.06            # xabarlar orasidagi pauza (Telegram limiti ~30/s)

_A: dict = {}            # bot.py dan beriladigan funksiyalar
_LOCK = threading.Lock()

# ═══════════════════════════════════════════════
# POSTLAR (tartib — eng foydali xizmatlardan boshlab)
# {narx} — DOC_NARX[kalit], {narx2} — qo'shimcha narx; {ref}, {xush} — bonuslar
# ═══════════════════════════════════════════════
POSTS = [
    {"key": "pptx", "narx2": None,
     "uz": "📊 <b>Ertaga taqdimot? 3 daqiqada tayyor!</b>\n\nMavzuni yozasiz → 15 ta shablondan birini tanlaysiz → "
           "tayyor PowerPoint fayl.\n✅ Diagramma va jadvallar\n✅ Har slaydda tushuntirish matni\n"
           "✅ O'zbek, rus, ingliz, qoraqalpoq va boshqa tillarda\n✅ O'z shabloningizni ham yuklasangiz bo'ladi\n\n"
           "💰 Atigi <b>{narx} tanga</b>",
     "ru": "📊 <b>Презентация на завтра? Готово за 3 минуты!</b>\n\nПишете тему → выбираете один из 15 шаблонов → "
           "готовый PowerPoint.\n✅ Диаграммы и таблицы\n✅ Пояснение на каждом слайде\n"
           "✅ На узбекском, русском, английском и других языках\n✅ Можно загрузить свой шаблон\n\n"
           "💰 Всего <b>{narx} монет</b>",
     "en": "📊 <b>Presentation tomorrow? Ready in 3 minutes!</b>\n\nType the topic → pick one of 15 templates → "
           "a ready PowerPoint file.\n✅ Charts and tables\n✅ Explanation on every slide\n"
           "✅ Uzbek, Russian, English and more\n✅ You can upload your own template\n\n💰 Only <b>{narx} coins</b>"},
    {"key": "docx", "narx2": "docx_katta",
     "uz": "📄 <b>Kurs ishi uchun tunni bedor o'tkazish shart emas</b>\n\nOTM, fan va mavzuni yozasiz — bot o'zi tuzadi:\n"
           "✅ Titul, mundarija, kirish, boblar, xulosa va takliflar\n✅ Snoskalar va adabiyotlar ro'yxati\n"
           "✅ Jadval va grafiklar\n\n📘 25-30 bet — <b>{narx} tanga</b>\n📚 65-70 bet — <b>{narx2} tanga</b>\n\n"
           "💡 O'qituvchi talablarini yozib qo'ysangiz — aynan shunga moslab yozadi.",
     "ru": "📄 <b>Курсовая без бессонных ночей</b>\n\nУкажите вуз, предмет и тему — бот всё сделает:\n"
           "✅ Титул, содержание, введение, главы, заключение\n✅ Сноски и список литературы\n✅ Таблицы и графики\n\n"
           "📘 25-30 стр. — <b>{narx} монет</b>\n📚 65-70 стр. — <b>{narx2} монет</b>\n\n"
           "💡 Напишите требования преподавателя — работа будет сделана по ним.",
     "en": "📄 <b>No more sleepless nights over a course work</b>\n\nEnter the university, subject and topic — the bot does "
           "the rest:\n✅ Title page, contents, introduction, chapters, conclusion\n✅ Footnotes and references\n"
           "✅ Tables and charts\n\n📘 25-30 pages — <b>{narx} coins</b>\n📚 65-70 pages — <b>{narx2} coins</b>\n\n"
           "💡 Add your teacher's requirements and the work will follow them."},
    {"key": "bmi", "narx2": None,
     "uz": "🎓 <b>Bitiruv malakaviy ishi (BMI) — OTM talablari bo'yicha</b>\n\nYo'nalish va mavzuni yozasiz — bot to'liq "
           "BMI tuzadi:\n✅ BMI titul varag'i (ilmiy rahbar, taqrizchi, «Himoyaga ruxsat etildi»)\n"
           "✅ Annotatsiya 3 tilda — o'zbek, rus, ingliz\n✅ Kirish: dolzarblik, maqsad, vazifalar, obyekt, predmet, "
           "<b>ilmiy yangilik</b>\n✅ 3 bob, jadval va grafiklar, bet ko'rsatilgan snoskalar\n"
           "✅ 45-55 ta adabiyot, xulosa va tavsiyalar, ilovalar\n✅ Betlari aniq avtomatik mundarija\n\n"
           "📄 65-75 bet — <b>{narx} tanga</b>\n\n💡 Keyin «🛡 Himoyaga tayyorlov» bilan taqdimot va nutq ham tayyor.",
     "ru": "🎓 <b>Выпускная квалификационная работа (ВКР) — по требованиям вузов</b>\n\nУкажите направление и тему — "
           "бот подготовит полную ВКР:\n✅ Титул ВКР (руководитель, рецензент, «Допускается к защите»)\n"
           "✅ Аннотация на 3 языках — узбекском, русском, английском\n✅ Введение: актуальность, цель, задачи, объект, "
           "предмет, <b>научная новизна</b>\n✅ 3 главы, таблицы и графики, сноски со страницами\n"
           "✅ 45-55 источников, заключение и рекомендации, приложения\n✅ Автоматическое содержание с точными страницами\n\n"
           "📄 65-75 стр. — <b>{narx} монет</b>\n\n💡 Затем «🛡 Подготовка к защите» — презентация и речь.",
     "en": "🎓 <b>Graduation thesis — following university requirements</b>\n\nEnter your field of study and topic — the bot "
           "prepares the full thesis:\n✅ Thesis title page (supervisor, reviewer, “Admitted to defense”)\n"
           "✅ Annotation in 3 languages — Uzbek, Russian, English\n✅ Introduction: relevance, aim, objectives, object, "
           "subject, <b>scientific novelty</b>\n✅ 3 chapters, tables and charts, footnotes with page numbers\n"
           "✅ 45-55 references, conclusion, appendices\n✅ Automatic table of contents with exact pages\n\n"
           "📄 65-75 pages — <b>{narx} coins</b>\n\n💡 Then use «🛡 Defense prep» for slides and a speech."},
    {"key": "himoya", "narx2": None,
     "uz": "🛡 <b>Himoyaga 1 kun qoldimi? Hammasi tayyor bo'ladi</b>\n\nKurs ishi yoki BMI faylingizni yuboring "
           "(yoki botda yozilgan ishingizni tanlang):\n✅ Himoya taqdimoti — 9-10 slayd, ishingizdagi jadval va grafiklar "
           "bilan\n✅ 7-9 daqiqalik himoya nutqi — har slayd uchun\n✅ Komissiya berishi mumkin bo'lgan <b>15 ta savol va "
           "javob</b>\n✅ Himoyada nimalar baholanishi va maslahatlar\n\n💰 <b>{narx} tanga</b>",
     "ru": "🛡 <b>Защита завтра? Подготовим всё</b>\n\nОтправьте файл курсовой или ВКР (или выберите работу, сделанную в "
           "боте):\n✅ Презентация для защиты — 9-10 слайдов с таблицами и графиками из работы\n"
           "✅ Речь на 7-9 минут — к каждому слайду\n✅ <b>15 вероятных вопросов комиссии с ответами</b>\n"
           "✅ Критерии оценки и советы\n\n💰 <b>{narx} монет</b>",
     "en": "🛡 <b>Defense tomorrow? We'll get you ready</b>\n\nSend your course work or thesis file (or pick the work made "
           "in the bot):\n✅ Defense slides — 9-10 slides with tables and charts from your work\n"
           "✅ A 7-9 minute speech — for every slide\n✅ <b>15 likely committee questions with answers</b>\n"
           "✅ Assessment criteria and tips\n\n💰 <b>{narx} coins</b>"},
    {"key": "amaliyot", "narx2": None,
     "uz": "📋 <b>Amaliyot tugadi — hisobot qoldi?</b>\n\nTashkilot, muddat va rahbarlarni yozasiz — bot tayyorlaydi:\n"
           "✅ Titul varag'i\n✅ <b>Amaliyot kundaligi</b> — har ish kuni uchun sana va bajarilgan ishlar\n"
           "✅ Hisobot: tashkilot tavsifi, bajarilgan ishlar, ko'nikmalar, muammo va takliflar, xulosa\n"
           "✅ <b>Tavsifnoma</b> loyihasi — rahbar faqat imzolaydi\n\nO'quv, ishlab chiqarish, pedagogik, malakaviy va "
           "bitiruv oldi amaliyotlari uchun.\n\n💰 <b>{narx} tanga</b>",
     "ru": "📋 <b>Практика закончилась — остался отчёт?</b>\n\nУкажите организацию, сроки и руководителей — бот подготовит:\n"
           "✅ Титульный лист\n✅ <b>Дневник практики</b> — дата и работы за каждый рабочий день\n"
           "✅ Отчёт: характеристика организации, выполненные работы, навыки, проблемы и предложения, заключение\n"
           "✅ Проект <b>характеристики</b> — руководителю остаётся подписать\n\nДля учебной, производственной, "
           "педагогической, квалификационной и преддипломной практики.\n\n💰 <b>{narx} монет</b>",
     "en": "📋 <b>Internship over — report still to do?</b>\n\nEnter the organisation, dates and supervisors — the bot "
           "prepares:\n✅ Title page\n✅ <b>Internship diary</b> — date and tasks for each working day\n"
           "✅ Report: organisation profile, work done, skills, problems and suggestions, conclusion\n"
           "✅ Draft <b>reference letter</b> — the supervisor only signs it\n\nFor educational, industrial, teaching, "
           "professional and pre-graduation internships.\n\n💰 <b>{narx} coins</b>"},
    {"key": "ilmiy", "narx2": "iz_mos",
     "uz": "👨‍🔬 <b>Doktorant va mustaqil izlanuvchilar uchun yangi bo'lim!</b>\n\n"
           "📰 <b>OAK jurnallari ro'yxati — BEPUL</b>\nJurnal nomini yozing — u sizning fan tarmog'ingiz bo'yicha "
           "faolmi yoki ro'yxatdan chiqarilganmi, darhol ko'rasiz. Ro'yxat oak.uz dan avtomatik yangilanadi, "
           "o'zgarish bo'lsa 🔔 xabar beramiz.\n\n"
           "🌍 Scopus / Web of Science: OAK qaysi bazalarni tan oladi (2026-yil iyun qarori)\n"
           "📘 Ilmiy daraja tartibi 2027, maqola yuborish, shubhali jurnallardan saqlanish — bepul qo'llanmalar\n\n"
           "🧭 Maqolangizga mos jurnal — <b>{narx2} tanga</b>\n📚 Mavzu bo'yicha haqiqiy adabiyotlar (DOI bilan)\n"
           "📑 PhD/DSc individual reja · 🎯 Mavzu va ixtisoslik mosligi · 📝 GOST/APA bibliografiya",
     "ru": "👨‍🔬 <b>Новый раздел для докторантов и соискателей!</b>\n\n"
           "📰 <b>Перечень журналов ВАК — БЕСПЛАТНО</b>\nВведите название журнала — сразу увидите, активен ли он по "
           "вашей отрасли или исключён. Перечень автоматически обновляется с oak.uz, об изменениях сообщим 🔔.\n\n"
           "🌍 Scopus / Web of Science: какие базы признаёт ВАК (решение июня 2026)\n"
           "📘 Порядок присуждения степеней 2027, подача статей, хищнические журналы — бесплатные руководства\n\n"
           "🧭 Подбор журнала для статьи — <b>{narx2} монет</b>\n📚 Реальная литература по теме (с DOI)\n"
           "📑 Индивидуальный план PhD/DSc · 🎯 Тема и специальность · 📝 Библиография ГОСТ/APA",
     "en": "👨‍🔬 <b>New section for PhD/DSc candidates and researchers!</b>\n\n"
           "📰 <b>HAC (OAK) journal list — FREE</b>\nType a journal name and see instantly whether it is active for your "
           "field or has been removed. The list updates automatically from oak.uz, and we 🔔 notify you about changes.\n\n"
           "🌍 Scopus / Web of Science: which databases the HAC recognises (June 2026 decision)\n"
           "📘 Degree rules 2027, submitting papers, avoiding predatory journals — free guides\n\n"
           "🧭 Journal finder for your paper — <b>{narx2} coins</b>\n📚 Real literature on your topic (with DOIs)\n"
           "📑 PhD/DSc individual plan · 🎯 Topic vs specialty · 📝 GOST/APA bibliography"},
    {"key": "konspekt", "narx2": None,
     "uz": "📚 <b>Darslik qalinmi? Konspekt qilib beramiz</b>\n\nMa'ruza matni, PDF/Word fayl yoki daftar rasmini yuboring:\n"
           "✅ Qisqa mazmun va asosiy g'oyalar\n✅ Atamalar lug'ati\n✅ Takrorlash uchun savollar\n"
           "🎧 Xohlasangiz — audio qilib tinglaysiz\n\n💰 <b>{narx} tanga</b>",
     "ru": "📚 <b>Толстый учебник? Сделаем конспект</b>\n\nОтправьте текст лекции, PDF/Word или фото тетради:\n"
           "✅ Краткое содержание и главные идеи\n✅ Словарь терминов\n✅ Вопросы для повторения\n"
           "🎧 Можно прослушать в аудио\n\n💰 <b>{narx} монет</b>",
     "en": "📚 <b>Thick textbook? We'll summarise it</b>\n\nSend lecture text, a PDF/Word file or a photo of your notes:\n"
           "✅ Short summary and key ideas\n✅ Glossary\n✅ Review questions\n🎧 Listen to it as audio\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "ftest", "narx2": "quiz",
     "uz": "🧠 <b>Imtihondan oldin o'zingizni sinang</b>\n\nKitob yoki ma'ruza faylini yuboring (kerakli betlarni ko'rsatib) — "
           "bot undan test tuzadi:\n✅ 10, 20 yoki 30 savol, qiyinlik darajasi\n✅ Telegram ichida yechiladigan test\n"
           "✅ Chop etish uchun Word va javoblar kaliti\n\n💰 <b>{narx} tanga</b>\n"
           "🎯 Tayyor testingizni quizga aylantirish — <b>{narx2} tanga</b>",
     "ru": "🧠 <b>Проверьте себя перед экзаменом</b>\n\nОтправьте учебник или лекцию (укажите страницы) — бот составит тест:\n"
           "✅ 10, 20 или 30 вопросов, уровень сложности\n✅ Тест прямо в Telegram\n✅ Word для печати с ключом ответов\n\n"
           "💰 <b>{narx} монет</b>\n🎯 Готовый тест → квиз — <b>{narx2} монеты</b>",
     "en": "🧠 <b>Test yourself before the exam</b>\n\nSend a textbook or lecture (choose the pages) — the bot builds a test:\n"
           "✅ 10, 20 or 30 questions, difficulty level\n✅ Quiz right inside Telegram\n✅ Printable Word with answer key\n\n"
           "💰 <b>{narx} coins</b>\n🎯 Turn your ready test into a quiz — <b>{narx2} coins</b>"},
    {"key": "masala", "narx2": None,
     "uz": "🧮 <b>Masala chiqmayaptimi?</b>\n\nRasmini yuboring — bosqichma-bosqich yechib, tushuntirib beramiz:\n"
           "✅ Matematika va fizika\n✅ Kimyoviy reaksiyalarni tenglash\n✅ Biologik sxemalarni tushuntirish\n"
           "📸 Qo'lda yozilgan bo'lsa ham bo'ladi\n\n💰 <b>{narx} tanga</b>",
     "ru": "🧮 <b>Не получается задача?</b>\n\nОтправьте фото — решим по шагам и объясним:\n✅ Математика и физика\n"
           "✅ Уравнивание химических реакций\n✅ Разбор биологических схем\n📸 Можно рукописный текст\n\n"
           "💰 <b>{narx} монеты</b>",
     "en": "🧮 <b>Stuck on a problem?</b>\n\nSend a photo — we'll solve it step by step and explain:\n✅ Maths and physics\n"
           "✅ Balancing chemical equations\n✅ Biology diagrams explained\n📸 Handwriting is fine\n\n💰 <b>{narx} coins</b>"},
    {"key": "kabinet", "narx2": None,
     "uz": "🎁 <b>Tekin tanga olishni xohlaysizmi?</b>\n\nDo'stlaringizni taklif qiling:\n• sizga — har biri uchun <b>+{ref} tanga</b>\n"
           "• do'stingizga — <b>+{xush} tanga</b> sovg'a\n\n5 do'st = {ref5} tanga ≈ {taq} ta taqdimot tekin!\n\n"
           "👤 Kabinet → «📤 Do'stlarga ulashish» — bir bosishda yuboriladi.",
     "ru": "🎁 <b>Хотите бесплатные монеты?</b>\n\nПриглашайте друзей:\n• вам — <b>+{ref} монет</b> за каждого\n"
           "• другу — <b>+{xush} монет</b> в подарок\n\n5 друзей = {ref5} монет ≈ {taq} презентаций бесплатно!\n\n"
           "👤 Кабинет → «📤 Поделиться с друзьями» — в одно касание.",
     "en": "🎁 <b>Want free coins?</b>\n\nInvite friends:\n• you get <b>+{ref} coins</b> for each\n"
           "• your friend gets <b>+{xush} coins</b>\n\n5 friends = {ref5} coins ≈ {taq} free presentations!\n\n"
           "👤 Profile → «📤 Share with friends» — one tap."},
    {"key": "imtihon", "narx2": None,
     "uz": "🎯 <b>Imtihonga necha kun qoldi?</b>\n\nFan yoki mavzularni yozing — tayyorlov to'plami tayyor:\n"
           "✅ Flash-kartalar\n✅ 10 ta test (Telegram'da yechasiz)\n✅ Ehtimoliy savollar va javob rejasi\n"
           "✅ Imtihongacha kunlik reja\n\n💰 <b>{narx} tanga</b>",
     "ru": "🎯 <b>Сколько дней до экзамена?</b>\n\nНапишите предмет или темы — получите набор для подготовки:\n"
           "✅ Флеш-карточки\n✅ 10 тестов в Telegram\n✅ Вероятные вопросы и план ответа\n✅ План по дням до экзамена\n\n"
           "💰 <b>{narx} монет</b>",
     "en": "🎯 <b>How many days until the exam?</b>\n\nType the subject or topics — get a prep kit:\n✅ Flashcards\n"
           "✅ 10 quiz questions in Telegram\n✅ Likely questions with answer outlines\n✅ Day-by-day study plan\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "referat", "narx2": None,
     "uz": "📑 <b>Referat kerakmi?</b>\n\nMavzuni yozing — 8-12 betlik tayyor Word fayl:\n"
           "✅ Titul, reja, kirish, asosiy qism, xulosa\n✅ Adabiyotlar ro'yxati\n✅ Topshirishga tayyor format\n\n"
           "💰 <b>{narx} tanga</b>",
     "ru": "📑 <b>Нужен реферат?</b>\n\nНапишите тему — готовый Word на 8-12 страниц:\n"
           "✅ Титул, план, введение, основная часть, заключение\n✅ Список литературы\n✅ Готов к сдаче\n\n"
           "💰 <b>{narx} монет</b>",
     "en": "📑 <b>Need an essay paper?</b>\n\nType the topic — get a ready 8-12 page Word file:\n"
           "✅ Title page, outline, introduction, main part, conclusion\n✅ References\n✅ Ready to submit\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "tts", "narx2": None,
     "uz": "🎧 <b>O'qishga vaqt yo'qmi? Tinglang!</b>\n\nMatn, PDF yoki rasm yuboring — MP3 audio qilib beramiz:\n"
           "✅ O'zbek, rus, ingliz, qozoq, turk ovozlari\n✅ Yo'lda, transportda tinglaysiz\n\n💰 <b>{narx} tanga</b>",
     "ru": "🎧 <b>Нет времени читать? Слушайте!</b>\n\nОтправьте текст, PDF или фото — сделаем MP3:\n"
           "✅ Узбекский, русский, английский, казахский, турецкий голоса\n✅ Слушайте в дороге\n\n"
           "💰 <b>{narx} монеты</b>",
     "en": "🎧 <b>No time to read? Listen!</b>\n\nSend text, a PDF or a photo — get an MP3:\n"
           "✅ Uzbek, Russian, English, Kazakh, Turkish voices\n✅ Listen on the go\n\n💰 <b>{narx} coins</b>"},
    {"key": "mustaqil", "narx2": None,
     "uz": "🗂 <b>Mustaqil ish — talablar bo'yicha</b>\n\n✅ Titul, reja, 3 bob, xulosa\n"
           "✅ Adabiyotlar GOST, APA yoki MLA formatida\n✅ Ovozli tinglash imkoniyati\n\n💰 <b>{narx} tanga</b>",
     "ru": "🗂 <b>Самостоятельная работа по требованиям</b>\n\n✅ Титул, план, 3 главы, заключение\n"
           "✅ Литература по ГОСТ, APA или MLA\n✅ Можно прослушать\n\n💰 <b>{narx} монет</b>",
     "en": "🗂 <b>Independent work done right</b>\n\n✅ Title page, outline, 3 chapters, conclusion\n"
           "✅ References in GOST, APA or MLA\n✅ Audio version available\n\n💰 <b>{narx} coins</b>"},
    {"key": "esse", "narx2": None,
     "uz": "📝 <b>Esse yozdingizmi? Ball qancha bo'larkin?</b>\n\nEsseni yuboring — milliy sertifikat mezonlari bo'yicha "
           "baholab, xatolarni ko'rsatamiz:\n✅ Ball (75 dan) va PDF hisobot\n✅ Grammatika va uslub tahlili\n\n"
           "🎁 Birinchi tekshiruv — <b>BEPUL</b>, keyingilari {narx} tanga",
     "ru": "📝 <b>Написали эссе? Узнайте балл!</b>\n\nОтправьте эссе — оценим по критериям национального сертификата "
           "и покажем ошибки:\n✅ Балл (из 75) и PDF-отчёт\n✅ Разбор грамматики и стиля\n\n"
           "🎁 Первая проверка — <b>БЕСПЛАТНО</b>, далее {narx} монет",
     "en": "📝 <b>Wrote an essay? Check your score!</b>\n\nSend it — we grade it by national certificate criteria and "
           "show mistakes:\n✅ Score (out of 75) and PDF report\n✅ Grammar and style review\n\n"
           "🎁 First check is <b>FREE</b>, then {narx} coins"},
    {"key": "tarjima", "narx2": None,
     "uz": "🌍 <b>Ilmiy matnni tarjima qilish kerakmi?</b>\n\nMatn yoki rasm yuboring — ilmiy uslubda, atamalarni to'g'ri "
           "saqlab tarjima qilamiz.\n✅ O'zbek ⇄ rus ⇄ ingliz va boshqa tillar\n\n💰 <b>{narx} tanga</b>",
     "ru": "🌍 <b>Нужен научный перевод?</b>\n\nОтправьте текст или фото — переведём в научном стиле с верными терминами.\n"
           "✅ Узбекский ⇄ русский ⇄ английский и другие\n\n💰 <b>{narx} монеты</b>",
     "en": "🌍 <b>Need an academic translation?</b>\n\nSend text or a photo — translated in academic style with correct "
           "terms.\n✅ Uzbek ⇄ Russian ⇄ English and more\n\n💰 <b>{narx} coins</b>"},
    {"key": "maqola", "narx2": None,
     "uz": "📝 <b>Ilmiy maqola yozyapsizmi?</b>\n\nMavzu va kalit so'zlarni yozing — IMRAD formatida maqola:\n"
           "✅ Annotatsiya, kirish, metodlar, natijalar, muhokama, xulosa\n✅ Adabiyotlar ro'yxati\n"
           "💡 O'z ma'lumotlaringizni yozsangiz — natijalar aynan shu asosda yoziladi.\n\n💰 <b>{narx} tanga</b>",
     "ru": "📝 <b>Пишете научную статью?</b>\n\nУкажите тему и ключевые слова — статья в формате IMRAD:\n"
           "✅ Аннотация, введение, методы, результаты, обсуждение, выводы\n✅ Список литературы\n"
           "💡 Дайте свои данные — результаты будут описаны по ним.\n\n💰 <b>{narx} монет</b>",
     "en": "📝 <b>Writing a research article?</b>\n\nEnter the topic and keywords — get an IMRAD article:\n"
           "✅ Abstract, introduction, methods, results, discussion, conclusion\n✅ References\n"
           "💡 Provide your own data and the results will be based on it.\n\n💰 <b>{narx} coins</b>"},
    {"key": "original", "narx2": None,
     "uz": "✨ <b>Matn originalligi past chiqyaptimi?</b>\n\nMatnni yuboring — ma'nosini saqlagan holda qayta ifodalab "
           "beramiz, takrorlar kamayadi.\n\n💰 <b>{narx} tanga</b>",
     "ru": "✨ <b>Низкая уникальность текста?</b>\n\nОтправьте текст — перефразируем с сохранением смысла.\n\n"
           "💰 <b>{narx} монет</b>",
     "en": "✨ <b>Low originality score?</b>\n\nSend the text — we rephrase it while keeping the meaning.\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "cv", "narx2": None,
     "uz": "💼 <b>Ish yoki amaliyot izlayapsizmi?</b>\n\nMa'lumotlaringizni kiriting — zamonaviy, professional "
           "Resume/CV tayyor Word faylda.\n\n💰 <b>{narx} tanga</b>",
     "ru": "💼 <b>Ищете работу или практику?</b>\n\nВведите данные — современное профессиональное резюме в Word.\n\n"
           "💰 <b>{narx} монет</b>",
     "en": "💼 <b>Looking for a job or internship?</b>\n\nEnter your details — a modern professional resume in Word.\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "grammatika", "narx2": None,
     "uz": "✍️ <b>Matningizda xato yo'qmi?</b>\n\nMatnni yuboring — imlo, grammatika va uslub xatolarini tuzatib, "
           "akademik ko'rinishga keltiramiz.\n\n💰 <b>{narx} tanga</b>",
     "ru": "✍️ <b>Нет ли ошибок в тексте?</b>\n\nОтправьте текст — исправим орфографию, грамматику и стиль.\n\n"
           "💰 <b>{narx} монеты</b>",
     "en": "✍️ <b>Any mistakes in your text?</b>\n\nSend it — we fix spelling, grammar and style.\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "nutq", "narx2": None,
     "uz": "🎤 <b>Chiqish qilishingiz kerakmi?</b>\n\nMavzu, vaqt va auditoriyani yozing — ravon, tushunarli nutq matni.\n"
           "✅ Seminar, himoya, tadbir uchun\n\n💰 <b>{narx} tanga</b>",
     "ru": "🎤 <b>Нужно выступить?</b>\n\nУкажите тему, время и аудиторию — получите понятный текст речи.\n"
           "✅ Для семинара, защиты, мероприятия\n\n💰 <b>{narx} монеты</b>",
     "en": "🎤 <b>Giving a talk?</b>\n\nEnter the topic, time and audience — get a clear speech.\n"
           "✅ For seminars, defences, events\n\n💰 <b>{narx} coins</b>"},
    {"key": "spiker", "narx2": None,
     "uz": "📽 <b>Taqdimotingiz bor, lekin nima deyishni bilmayapsizmi?</b>\n\nPPTX faylni yuboring — har slayd uchun "
           "gapiriladigan matn tayyor.\n\n💰 <b>{narx} tanga</b>",
     "ru": "📽 <b>Есть презентация, но не знаете, что говорить?</b>\n\nОтправьте PPTX — подготовим текст для каждого "
           "слайда.\n\n💰 <b>{narx} монеты</b>",
     "en": "📽 <b>Have slides but not sure what to say?</b>\n\nSend the PPTX — get a script for every slide.\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "tezis", "narx2": None,
     "uz": "📋 <b>Konferensiyaga tezis kerakmi?</b>\n\n✅ Talablarga mos tuzilma\n✅ Annotatsiya va kalit so'zlar\n"
           "✅ Adabiyotlar\n\n💰 <b>{narx} tanga</b>",
     "ru": "📋 <b>Нужны тезисы на конференцию?</b>\n\n✅ Структура по требованиям\n✅ Аннотация и ключевые слова\n"
           "✅ Литература\n\n💰 <b>{narx} монет</b>",
     "en": "📋 <b>Need a conference abstract?</b>\n\n✅ Required structure\n✅ Abstract and keywords\n✅ References\n\n"
           "💰 <b>{narx} coins</b>"},
    {"key": "gipoteza", "narx2": None,
     "uz": "🔬 <b>Magistratura yoki dissertatsiya uchun</b>\n\nMavzuni yozing — ilmiy muammo, dolzarblik, gipoteza, "
           "maqsad va vazifalarni tizimli shakllantirib beramiz.\n\n💰 <b>{narx} tanga</b>",
     "ru": "🔬 <b>Для магистратуры и диссертации</b>\n\nНапишите тему — сформулируем проблему, актуальность, гипотезу, "
           "цель и задачи.\n\n💰 <b>{narx} монет</b>",
     "en": "🔬 <b>For master's and PhD students</b>\n\nEnter your topic — get the research problem, relevance, hypothesis, "
           "aim and objectives.\n\n💰 <b>{narx} coins</b>"},
    {"key": "sardor", "narx2": None,
     "uz": "🧑‍🏫 <b>Guruh sardorimisiz?</b>\n\nYig'ilish bayonnomasi, tadbir ssenariysi, tarbiyaviy soat rejasi, "
           "e'lon va hisobot — bir necha daqiqada Word'da.\n\n💰 <b>{narx} tanga</b>",
     "ru": "🧑‍🏫 <b>Вы староста?</b>\n\nПротокол собрания, сценарий мероприятия, план воспитательного часа, объявление "
           "и отчёт — за пару минут в Word.\n\n💰 <b>{narx} монеты</b>",
     "en": "🧑‍🏫 <b>Are you a class leader?</b>\n\nMeeting minutes, event scripts, class hour plans, announcements and "
           "reports — in Word in minutes.\n\n💰 <b>{narx} coins</b>"},
]

# «yangi» sanasi bor postlar — YANGI xizmat e'loni: navbatdagi slotda hali olmagan har bir foydalanuvchiga
# aylanmadan OLDIN yuboriladi (shu sanadan YANGI_KUN kungacha). Keyin odatdagi aylanmaga qo'shiladi.
POSTS.extend([
    {"key": "ragbat", "narx2": "rg_dars", "yangi": "2026-11-25",
     "uz": "🏅 <b>Yangi bo'lim: Rag'batlantirish!</b>\n\nO'qituvchi, tarbiyachi, murabbiy va sardorlar uchun:\n"
           "🏆 <b>Faxriy yorliq, tashakkurnoma, maqtov yorlig'i, sertifikat, diplom</b> — 8 xil chiroyli dizayn "
           "(rasmiy va bolalar uchun), chop etishga tayyor PDF + rasm. Butun sinfga — bir zumda, har biriga alohida ism bilan.\n"
           "🎓 <b>Dars to'plami</b> — mavzu yoki ma'ruza matningizdan tarqatma, savollar, javob kaliti va dars rejasi — "
           "<b>{narx2} tanga</b>\n"
           "📊 Natijalarni yuborsangiz — har o'quvchiga ballga mos yorliq va <b>shaxsiy maqtov matni</b>.\n\n"
           "💰 Yorliq — <b>{narx} tangadan</b>",
     "ru": "🏅 <b>Новый раздел: Поощрение!</b>\n\nДля учителей, воспитателей, наставников и старост:\n"
           "🏆 <b>Почётная грамота, благодарность, похвальный лист, сертификат, диплом</b> — 8 красивых дизайнов "
           "(официальные и детские), готовый к печати PDF + картинка. Для всего класса — сразу, каждому со своим именем.\n"
           "🎓 <b>Пакет к уроку</b> — по теме или тексту лекции: раздатка, вопросы, ключ ответов и план урока — "
           "<b>{narx2} монет</b>\n"
           "📊 Отправьте результаты — каждому ученику грамота по баллам и <b>личный текст похвалы</b>.\n\n"
           "💰 Грамота — <b>от {narx} монет</b>",
     "en": "🏅 <b>New section: Awards!</b>\n\nFor teachers, tutors, mentors and class leaders:\n"
           "🏆 <b>Certificates of honor, merit, appreciation, diplomas</b> — 8 beautiful designs (formal and for kids), "
           "print-ready PDF + image. For a whole class at once, each with their own name.\n"
           "🎓 <b>Lesson pack</b> — from a topic or your lecture: handout, questions, answer key and lesson plan — "
           "<b>{narx2} coins</b>\n"
           "📊 Send the results — every learner gets the right certificate and a <b>personal praise note</b>.\n\n"
           "💰 Certificates — <b>from {narx} coins</b>"},
    {"key": "test", "narx2": None,
     "uz": "📝 <b>Test kerakmi? Kimga mo'ljallanganini ayting — qolganini bot qiladi</b>\n\n"
           "🎒 1-4 sinf · 📘 5-9 sinf · 🎓 10-11 sinf · 🏁 Abituriyent (DTM) · 🏛 Talaba · 👔 Mutaxassis\n\n"
           "✅ Mavzu yoki ma'ruza matningiz asosida\n✅ Oddiydan murakkabga (Bloom taksonomiyasi)\n"
           "✅ Har bir javob mustaqil tekshiriladi\n✅ 1, 2 yoki 4 variant — ko'chirib bo'lmaydi\n"
           "✅ O'quvchiga — javobsiz test va javoblar varaqasi, o'qituvchiga — kalit va izohlar\n\n💰 <b>{narx} tanga</b>",
     "ru": "📝 <b>Нужен тест? Скажите, для кого — остальное сделает бот</b>\n\n"
           "🎒 1-4 класс · 📘 5-9 класс · 🎓 10-11 класс · 🏁 Абитуриент (ДТМ) · 🏛 Студент · 👔 Специалист\n\n"
           "✅ По теме или вашему тексту лекции\n✅ От простого к сложному (таксономия Блума)\n"
           "✅ Каждый ответ проверяется независимо\n✅ 1, 2 или 4 варианта — не списать\n"
           "✅ Ученикам — тест без ответов и бланк, учителю — ключи и пояснения\n\n💰 <b>{narx} монет</b>",
     "en": "📝 <b>Need a test? Tell us who it's for — the bot does the rest</b>\n\n"
           "🎒 Grades 1-4 · 📘 5-9 · 🎓 10-11 · 🏁 Applicants · 🏛 Students · 👔 Professionals\n\n"
           "✅ From a topic or your own lecture text\n✅ Simple to complex (Bloom's taxonomy)\n"
           "✅ Every answer independently checked\n✅ 1, 2 or 4 versions — no copying\n"
           "✅ Learners get a clean test and answer sheet, the teacher gets keys and explanations\n\n💰 <b>{narx} coins</b>"},
    {"key": "qayta", "narx2": None,
     "uz": "🔄 <b>Matningizni boshqa uslubda qayta yozib beramiz</b>\n\nAkademik, neytral yoki rasmiy uslubni tanlang — "
           "mazmun saqlanadi, jumlalar ravon va tushunarli bo'ladi. Ariza, xat, maqola yoki referat qismlari uchun qulay.\n\n"
           "💰 <b>{narx} tanga</b>",
     "ru": "🔄 <b>Перепишем ваш текст в другом стиле</b>\n\nВыберите академический, нейтральный или официальный стиль — "
           "смысл сохранится, фразы станут ясными и гладкими. Удобно для заявлений, писем, статей и рефератов.\n\n"
           "💰 <b>{narx} монет</b>",
     "en": "🔄 <b>We'll rewrite your text in another style</b>\n\nChoose academic, neutral or formal — the meaning stays, "
           "the sentences become clear and smooth. Handy for applications, letters, articles and essays.\n\n"
           "💰 <b>{narx} coins</b>"},
])
YANGI_KUN = 45

BTN_GO = {"uz": "👉 Sinab ko'rish", "ru": "👉 Попробовать", "en": "👉 Try it"}
BTN_KAB = {"uz": "👤 Kabinetni ochish", "ru": "👤 Открыть кабинет", "en": "👤 Open profile"}
BTN_OFF = {"uz": "🔕 Bunday xabarlar kerak emas", "ru": "🔕 Не присылать такие сообщения",
           "en": "🔕 Don't send these"}
OFF_TXT = {"uz": "🔕 Yaxshi, endi xizmatlar haqidagi xabarlarni yubormaymiz.\nQayta yoqish: /xabarlar",
           "ru": "🔕 Хорошо, больше не будем присылать такие сообщения.\nВключить снова: /xabarlar",
           "en": "🔕 Got it, we won't send these messages anymore.\nTurn back on: /xabarlar"}
ON_TXT = {"uz": "🔔 Xizmatlar haqidagi foydali xabarlar yoqildi (haftasiga 2 marta).\nO'chirish: /xabarlar",
          "ru": "🔔 Полезные сообщения о сервисах включены (2 раза в неделю).\nВыключить: /xabarlar",
          "en": "🔔 Helpful service tips are on (twice a week).\nTurn off: /xabarlar"}


def setup(**api):
    """bot.py dan: bot, get_db, db_lock, mark_dirty, narx (DOC_NARX), til_of(uid), open_service(call, key),
    admin_id, tz, ref_bonus, xush_bonus."""
    _A.update(api)


# ── Holat (users_db["marketing"]) ─────────────────────────────────
def _mk() -> dict:
    db = _A["get_db"]()
    m = db.setdefault("marketing", {})
    m.setdefault("idx", {})          # uid -> keyingi post raqami
    m.setdefault("off", [])          # o'chirganlar
    m.setdefault("blok", {})         # uid -> bloklagan vaqti
    m.setdefault("yoqilgan", True)
    m.setdefault("tarix", [])        # oxirgi yuborishlar statistikasi
    m.setdefault("yangi_ok", {})     # yangi post kaliti -> uni olgan uid'lar (str)
    return m


def matn(uid: int, post: dict, til: str) -> str:
    narx = _A["narx"]
    t = post.get(til) or post["uz"]
    pptx = max(1, narx.get("pptx", 5))
    ref, xush = _A.get("ref_bonus", 10), _A.get("xush_bonus", 5)
    return t.format(narx=narx.get(post["key"], ""), narx2=narx.get(post.get("narx2") or "", ""),
                    ref=ref, xush=xush, ref5=ref * 5, taq=(ref * 5) // pptx)


def klaviatura(post: dict, til: str) -> InlineKeyboardMarkup:
    kb = InlineKeyboardMarkup(row_width=1)
    if post["key"] == "kabinet":
        kb.add(InlineKeyboardButton(BTN_KAB.get(til, BTN_KAB["uz"]), callback_data="mk_go_kabinet"))
    else:
        kb.add(InlineKeyboardButton(BTN_GO.get(til, BTN_GO["uz"]), callback_data=f"mk_go_{post['key']}"))
    kb.add(InlineKeyboardButton(BTN_OFF.get(til, BTN_OFF["uz"]), callback_data="mk_off"))
    return kb


def _posts_for_send() -> list:
    """Narxi topilmagan xizmat postlari (masalan o'chirilgan xizmat) chiqarib tashlanadi."""
    narx = _A["narx"]
    return [p for p in POSTS if p["key"] in ("kabinet", "ilmiy") or p["key"] in narx]


def _yangi_postlar(posts: list) -> list:
    """E'lon muddati tugamagan yangi xizmat postlari."""
    bugun = datetime.date.today()
    out = []
    for p in posts:
        try:
            d = datetime.date.fromisoformat(p.get("yangi") or "")
        except Exception:
            continue
        if 0 <= (bugun - d).days <= YANGI_KUN:
            out.append(p)
    return out


def onboarding_faol(db: dict, uid) -> bool:
    """Yangi foydalanuvchi tanishuv (onboarding) xabarlarini hali olayotgan bo'lsa — True."""
    ob = db.get("onboarding", {}).get(str(uid))
    if not ob:
        return False
    kun = ob.get("kun", 1)
    if kun > 7:
        return False
    if ob.get("birinchi_doc") and kun > 4:
        return False
    return True


def _yuborishi_mumkin(uid: int, m: dict, db: dict) -> bool:
    """DIQQAT: db_lock ichida chaqiriladi — bu yerda boshqa lock oluvchi funksiya chaqirilmaydi."""
    if uid in m["off"] or str(uid) in m["off"]:
        return False
    if uid in db.get("banned_users", []):
        return False
    b = m["blok"].get(str(uid))
    if b:
        if db.get("last_seen", {}).get(str(uid), 0) <= b:
            return False                 # bloklagan va o'shandan beri qaytmagan
        m["blok"].pop(str(uid), None)
    if onboarding_faol(db, uid):
        return False                     # yangi foydalanuvchi — avval tanishuv xabarlari
    return True


def kampaniya(sabab: str = "jadval") -> dict:
    """Barcha mos foydalanuvchilarga navbatdagi postni yuboradi. Natija statistikasi qaytadi."""
    if not _LOCK.acquire(blocking=False):
        return {"xato": "boshqa yuborish davom etmoqda"}
    try:
        bot = _A["bot"]
        posts = _posts_for_send()
        with _A["db_lock"]:
            m = _mk()
            db = _A["get_db"]()
            uids = list(db.get("all_users", []))
            nomzod = [u for u in uids if _yuborishi_mumkin(u, m, db)]
            idx = dict(m["idx"])
            yangi_ok = {k: set(v) for k, v in m["yangi_ok"].items()}
        yangilar = _yangi_postlar(posts)
        ok = blok = xato = 0
        post_stat: dict = {}
        for uid in nomzod:
            # 1) yangi xizmat e'loni — hali olmagan bo'lsa, navbatdan oldin
            post = next((p for p in yangilar if str(uid) not in yangi_ok.get(p["key"], set())), None)
            i = None
            if post is None:
                i = idx.get(str(uid), 0) % len(posts)
                post = posts[i]
            til = _A["til_of"](uid)
            try:
                bot.send_message(uid, matn(uid, post, til), reply_markup=klaviatura(post, til),
                                 parse_mode="HTML", disable_web_page_preview=True)
                ok += 1
                if i is None:
                    yangi_ok.setdefault(post["key"], set()).add(str(uid))
                else:
                    idx[str(uid)] = i + 1
                post_stat[post["key"]] = post_stat.get(post["key"], 0) + 1
            except Exception as e:
                s = str(e).lower()
                if "blocked" in s or "deactivated" in s or "chat not found" in s or "403" in s:
                    blok += 1
                    with _A["db_lock"]:
                        _mk()["blok"][str(uid)] = time.time()
                else:
                    xato += 1
                    log.debug(f"[MARKETING] {uid}: {e}")
            time.sleep(TEZLIK)
        with _A["db_lock"]:
            m = _mk()
            m["idx"].update(idx)
            # eski (muddati o'tgan) e'lonlar ro'yxati tozalanadi — baza o'smasin
            amal = {p["key"] for p in yangilar}
            m["yangi_ok"] = {k: sorted(v) for k, v in yangi_ok.items() if k in amal}
            m["tarix"] = (m["tarix"] + [{"vaqt": int(time.time()), "sabab": sabab, "ok": ok, "blok": blok,
                                          "xato": xato, "nomzod": len(nomzod)}])[-30:]
        _A["mark_dirty"]()
        res = {"ok": ok, "blok": blok, "xato": xato, "nomzod": len(nomzod), "postlar": post_stat}
        log.info(f"📣 Marketing ({sabab}): {res}")
        try:
            bot.send_message(_A["admin_id"],
                             f"📣 <b>Marketing posti yuborildi</b> ({sabab})\n✅ {ok} ta yetkazildi • "
                             f"🚫 {blok} ta bloklagan • ⚠️ {xato} ta xato\n"
                             f"Postlar: {', '.join(f'{k}: {v}' for k, v in post_stat.items()) or '—'}",
                             parse_mode="HTML")
        except Exception:
            pass
        return res
    finally:
        _LOCK.release()


# ── Jadval ────────────────────────────────────────────────────────
def _slot_kalit(now: datetime.datetime):
    if now.weekday() not in KUNLAR:
        return None
    slot = now.replace(hour=SOAT, minute=DAQIQA, second=0, microsecond=0)
    if now < slot or now > slot + datetime.timedelta(hours=KECHIKISH_SOAT):
        return None
    return slot.strftime("%Y-%m-%d %H:%M")


def _loop():
    time.sleep(90)                       # bot to'liq ishga tushib, baza tiklanishini kutamiz
    while True:
        try:
            now = datetime.datetime.now(_A["tz"])
            key = _slot_kalit(now)
            if key:
                with _A["db_lock"]:
                    m = _mk()
                    yoq = m.get("yoqilgan", True)
                    oxirgi = m.get("oxirgi_slot")
                    if yoq and oxirgi != key:
                        m["oxirgi_slot"] = key       # avval belgilaymiz — qayta ishga tushsa takrorlanmaydi
                if yoq and oxirgi != key:
                    _A["mark_dirty"]()
                    kampaniya("jadval " + key)
        except Exception as e:
            log.error(f"[MARKETING THREAD] {e}")
        time.sleep(300)


def start():
    threading.Thread(target=_loop, daemon=True, name="marketing").start()
    log.info(f"📣 Marketing postlari: har {'/'.join(['Du','Se','Ch','Pa','Ju','Sh','Ya'][d] for d in KUNLAR)} "
             f"{SOAT:02d}:{DAQIQA:02d}")


# ── Telegram handlerlar ───────────────────────────────────────────
def register(bot):
    @bot.callback_query_handler(func=lambda c: c.data == "mk_off")
    def _off(call):
        uid = call.message.chat.id
        with _A["db_lock"]:
            m = _mk()
            if uid not in m["off"]:
                m["off"].append(uid)
        _A["mark_dirty"]()
        til = _A["til_of"](uid)
        try:
            bot.answer_callback_query(call.id)
            bot.edit_message_reply_markup(uid, call.message.message_id, reply_markup=None)
        except Exception:
            pass
        bot.send_message(uid, OFF_TXT.get(til, OFF_TXT["uz"]))

    @bot.callback_query_handler(func=lambda c: c.data.startswith("mk_go_"))
    def _go(call):
        key = call.data[len("mk_go_"):]
        try:
            bot.answer_callback_query(call.id)
        except Exception:
            pass
        with _A["db_lock"]:
            m = _mk()
            st = m.setdefault("bosishlar", {})
            st[key] = st.get(key, 0) + 1
        _A["mark_dirty"]()
        _A["open_service"](call, key)

    @bot.message_handler(commands=["xabarlar"])
    def _toggle(msg):
        uid = msg.chat.id
        with _A["db_lock"]:
            m = _mk()
            if uid in m["off"]:
                m["off"].remove(uid)
                txt = ON_TXT
            else:
                m["off"].append(uid)
                txt = OFF_TXT
        _A["mark_dirty"]()
        til = _A["til_of"](uid)
        bot.send_message(uid, txt.get(til, txt["uz"]))

    @bot.message_handler(commands=["marketing"])
    def _admin(msg):
        if msg.chat.id != _A["admin_id"]:
            return
        arg = (msg.text.split(maxsplit=1) + [""])[1].strip().lower()
        uid = msg.chat.id
        if arg == "test":
            posts = _posts_for_send()
            with _A["db_lock"]:
                i = _mk()["idx"].get("_admin_test", 0) % len(posts)
                _mk()["idx"]["_admin_test"] = i + 1
            til = _A["til_of"](uid)
            bot.send_message(uid, f"🧪 Post {i + 1}/{len(posts)} ({posts[i]['key']}):")
            bot.send_message(uid, matn(uid, posts[i], til), reply_markup=klaviatura(posts[i], til), parse_mode="HTML")
            return
        if arg == "yubor":
            bot.send_message(uid, "📣 Yuborish boshlandi — tugagach hisobot keladi.")
            threading.Thread(target=kampaniya, args=("admin",), daemon=True).start()
            return
        if arg in ("on", "off"):
            with _A["db_lock"]:
                _mk()["yoqilgan"] = arg == "on"
            _A["mark_dirty"]()
            bot.send_message(uid, "✅ Avtomatik postlar " + ("yoqildi" if arg == "on" else "o'chirildi"))
            return
        with _A["db_lock"]:
            m = copy.deepcopy(_mk())
            jami = len(_A["get_db"]().get("all_users", []))
        tarix = m.get("tarix", [])[-5:]
        bos = m.get("bosishlar", {})
        kun = "/".join(["Du", "Se", "Ch", "Pa", "Ju", "Sh", "Ya"][d] for d in KUNLAR)
        lines = [
            "📣 <b>Marketing postlari</b>",
            f"Holat: {'🟢 yoqilgan' if m.get('yoqilgan', True) else '🔴 o' + chr(39) + 'chirilgan'} • jadval: {kun} {SOAT:02d}:{DAQIQA:02d}",
            f"Postlar soni: {len(_posts_for_send())} (aylanma)",
            f"👥 Jami: {jami} • 🔕 o'chirgan: {len(m.get('off', []))} • 🚫 bloklagan: {len(m.get('blok', {}))}",
            f"Oxirgi slot: {m.get('oxirgi_slot', '—')}",
        ]
        if bos:
            lines.append("👉 Tugma bosishlar: " + ", ".join(f"{k}: {v}" for k, v in sorted(bos.items(), key=lambda x: -x[1])))
        for t in tarix:
            lines.append(f"• {datetime.datetime.fromtimestamp(t['vaqt'], _A['tz']).strftime('%d.%m %H:%M')} — "
                         f"✅{t['ok']} 🚫{t['blok']} ⚠️{t['xato']}")
        lines.append("\n<code>/marketing test</code> — keyingi postni ko'rish\n<code>/marketing yubor</code> — hozir "
                     "hammaga\n<code>/marketing on</code> | <code>off</code>")
        bot.send_message(uid, "\n".join(lines), parse_mode="HTML")

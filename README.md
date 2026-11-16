# 🎓 Academic AI Bot

Talabalar uchun Telegram bot: taqdimot, kurs ishi, maqola, mustaqil ish, konspekt, imtihonga tayyorlov va boshqa
o'quv hujjatlarini sun'iy intellekt yordamida bir necha daqiqada tayyorlab beradi.

- **Tillar:** o'zbekcha, ruscha, inglizcha
- **AI:** DeepSeek (matn yaratish) va Google Gemini (rasm, PDF o'qish, masala yechish)
- **Natija:** tahrirlanadigan Word (.docx) va PowerPoint (.pptx) fayllar

![Taqdimot shablonlari](assets/themes/_hammasi.jpg)

---

## 📋 Mundarija

- [Xizmatlar](#-xizmatlar)
- [Taqdimot shablonlari](#-taqdimot-shablonlari)
- [Loyiha tuzilishi](#-loyiha-tuzilishi)
- [O'rnatish va ishga tushirish](#-ornatish-va-ishga-tushirish)
- [Sozlamalar (.env)](#%EF%B8%8F-sozlamalar-env)
- [Tanga tizimi](#-tanga-tizimi)
- [Admin buyruqlari](#-admin-buyruqlari)
- [Xavfsizlik](#-xavfsizlik)
- [Serverga joylash](#-serverga-joylash)

---

## ✨ Xizmatlar

Narxlar tangada (🪙). Admin uchun barcha xizmatlar bepul.

### Hujjatlar

| Xizmat | Narx | Natija |
|---|---|---|
| 📊 Taqdimot | 6 | 15 ta shablondan biri bilan PowerPoint taqdimot: diagrammalar, jadvallar, bosqichlar |
| 📄 Kurs ishi | 12 | 28–30 betlik Word hujjat: titul, mundarija, boblar, adabiyotlar |
| 🗂 Mustaqil ish | 7 | Titul, reja, 3 bob, xulosa va **GOST / APA / MLA** formatidagi adabiyotlar ro'yxati |
| 📑 Referat | 6 | 8–12 betlik referat |
| 📰 Maqola | 9 | IMRAD formatidagi ilmiy maqola |
| 📋 Tezis | 6 | Konferensiya tezisi |

### O'qish va imtihon

| Xizmat | Narx | Natija |
|---|---|---|
| 📚 Konspekt | 4 | Matn, Word, PDF yoki rasmdan konspekt: qisqa mazmun, asosiy g'oyalar, atamalar lug'ati, savollar |
| 🎯 Imtihonga tayyorlov | 5 | Flash-kartalar, 10 ta test (Telegram ichida interaktiv quiz), ehtimoliy savollar, kunlik reja |
| 🧮 Masala yechish | 3 | Rasm yoki matndagi masalani bosqichma-bosqich yechish (matematika, fizika, iqtisod va h.k.) |
| 🧪 Test savollari | 4 | Mavzu bo'yicha test savollari |

### Matn bilan ishlash

| Xizmat | Narx | Natija |
|---|---|---|
| 🌐 Tarjima | 3 | Matn yoki rasmdagi matnni tarjima qilish |
| ✍️ Grammatika | 2 | Imlo va grammatik xatolarni tuzatish |
| 🔄 Originallashtirish | 6 | Matnni qayta ifodalash |
| 📝 Esse tekshirish | 5 | Esse bo'yicha baho va tavsiyalar (PDF hisobot) |

### Boshqa

| Xizmat | Narx | Natija |
|---|---|---|
| 🎤 Nutq | 4 | Ma'ruza/chiqish uchun nutq matni |
| 🗣 Spiker nutqi | 4 | Taqdimot faylidan spiker matni |
| 💼 CV | 9 | Rezyume (Word) |
| 🔬 Gipoteza | 5 | Ilmiy muammo va gipoteza shakllantirish |

Tayyor hujjatni bot ichida **tahrirlash** ham mumkin ("✏️ Hujjatni tahrirlash" tugmasi).

---

## 🎨 Taqdimot shablonlari

Foydalanuvchi slaydlar sonini kiritgandan keyin 15 ta shablonning rasmli galereyasini ko'radi.
Rasmlardagi katta raqam tanlash tugmalaridagi raqam bilan bir xil.

Shablonlar faqat rangi bilan emas, **maketi** bilan ham farq qiladi:

| Slayd turi | Variantlar |
|---|---|
| Titul | 15 xil |
| Sarlavha | ostiga chiziq, rangli tasma, chap ustun, raqamli, yorliqli, markaziy naqsh va boshqalar (9 xil) |
| Statistika | vertikal ustun, gorizontal ustun, chiziqli grafik, maydon grafik, progress-chiziqlar, lollipop |
| Taqsimot (foizlar) | doughnut, doira, 100% tasma, waffle 10×10, pufakchalar, halqalar |
| Ro'yxat | qatorlar, kartalar, ikonka ustunlari, tabletkalar, zigzag, katta raqamlar |
| Reja | ustunlar, tugunlar, pog'onali yorliqlar, chap panel, vertikal raqamlar |
| Bosqichlar | gorizontal, vertikal, strelkalar, zinapoya |
| Jadval, taqqoslash, statistika kartalari, xulosa | 4 xildan |

Diagrammalar PowerPoint'ning haqiqiy diagrammalari, ya'ni raqamlarni keyin o'zingiz tahrirlay olasiz.
Har bir slaydda sarlavha ostida slaydning asosiy fikrini tushuntiruvchi izoh bor.

Shablonlar ro'yxati: Akademik, Minimal, Korporativ, Texno, Tabiat, Gradient, Ijodiy, Milliy naqsh, Qora-oltin,
Ilmiy tadqiqot, Kosmos neon, Jurnal, Startap pitch, Okean, Sinf doskasi.

---

## 📁 Loyiha tuzilishi

```
Academic_ai_bot/
├── bot.py                # Asosiy bot: menyu, xizmatlar, tanga tizimi, admin panel
├── pptx_themes.py        # Taqdimot shablonlari dvigateli (15 shablon)
├── xizmatlar_yangi.py    # Konspekt, Imtihonga tayyorlov, Masala yechish, Mustaqil ish
├── security.py           # Xavfsizlik qatlami
├── requirements.txt      # Python kutubxonalari
├── .env.example          # Sozlamalar namunasi (.env shu asosda yaratiladi)
├── ishga_tushirish.bat   # Windows'da bir bosishda ishga tushirish
├── logo.png              # Esse hisoboti va sertifikat uchun logo
├── assets/themes/        # Shablonlar ko'rinish rasmlari (botdagi galereya)
├── SERVERGA_YUKLASH.md   # Linux serverga joylash yo'riqnomasi
└── OZGARISHLAR.md        # Versiya o'zgarishlari
```

Bot ishlaganda papkada `users.json` (foydalanuvchilar bazasi), `bot.log` va `admin_audit.log` paydo bo'ladi.
Ular `.gitignore` ga kiritilgan, GitHub'ga yuklanmaydi.

---

## 🚀 O'rnatish va ishga tushirish

Python **3.10 yoki undan yangi** kerak (3.11 yoki 3.12 tavsiya etiladi).

### Windows (eng oson yo'l)

1. Repozitoriyani yuklab oling: **Code → Download ZIP** yoki
   ```bat
   git clone https://github.com/Eldor24/Academic_ai_bot.git
   ```
2. `.env.example` faylidan nusxa olib, nomini `.env` qiling va ichini to'ldiring ([Sozlamalar](#%EF%B8%8F-sozlamalar-env)).
3. `ishga_tushirish.bat` faylini ikki marta bosing. U virtual muhit yaratadi, kutubxonalarni o'rnatadi va botni ishga tushiradi.

### Linux / macOS

```bash
git clone https://github.com/Eldor24/Academic_ai_bot.git
cd Academic_ai_bot
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env      # keyin .env ni tahrirlang
chmod 600 .env
python bot.py
```

> ⚠️ Bitta bot tokeni bilan bir vaqtda faqat **bitta** nusxa ishlashi mumkin. Aks holda Telegram
> `409 Conflict` xatosini beradi. Sinov uchun @BotFather'da alohida bot yarating.

---

## ⚙️ Sozlamalar (.env)

| O'zgaruvchi | Majburiy | Tavsif |
|---|---|---|
| `BOT_TOKEN` | ✅ | Telegram bot tokeni (@BotFather) |
| `DEEPSEEK_API_KEY` | ✅ | DeepSeek API kaliti — barcha matnlarni yaratadi |
| `GEMINI_API_KEY` | ✅ | Google Gemini kaliti — rasm va PDF o'qish, masala yechish |
| `ADMIN_ID` | ✅ | Admin Telegram ID raqami |
| `LOG_CHANNEL_ID` | — | Bot harakatlari yuboriladigan kanal ID si (`0` = o'chirilgan) |
| `CARD_NUMBER` | — | Hisobni to'ldirish uchun karta raqami |
| `CARD_OWNER` | — | Karta egasi |

---

## 🪙 Tanga tizimi

- Yangi foydalanuvchiga **5 tanga** bonus beriladi.
- Do'stini taklif qilgan foydalanuvchiga **10 tanga** beriladi. Faqat yangi foydalanuvchilar hisobga olinadi.
- Haftalik bonus, promokodlar va obunalar (Lite, Standart) bor.
- Hisob karta orqali to'ldiriladi, to'lovni admin `/topupconfirm` bilan tasdiqlaydi.
- AI xato bersa, sarflangan tangalar avtomatik qaytariladi.

---

## 🛠 Admin buyruqlari

Faqat `ADMIN_ID` egasi ishlata oladi. Barcha admin buyruqlari `admin_audit.log` ga yoziladi.

| Guruh | Buyruqlar |
|---|---|
| Panel | `/admin` |
| Statistika | `/stat`, `/fullstat`, `/activestat`, `/docstat`, `/topstat`, `/toprefs`, `/users`, `/userslist`, `/userinfo` |
| Tangalar | `/addcoin`, `/removecoins`, `/setcoins`, `/coins`, `/topupconfirm` |
| Foydalanuvchilar | `/ban`, `/unban`, `/banlist`, `/resetuser`, `/msg` |
| Promokodlar | `/newpromo`, `/delpromo`, `/promolist`, `/promostat`, `/promoguide` |
| Obunalar | `/delob` |
| Xabar tarqatish | `/broadcast`, `/send`, `/send_confirm`, `/cancel_broadcast` |
| Baza | `/getdb`, `/dbinfo`, `/cleanup` |

Foydalanuvchi buyruqlari: `/start`, `/kabinet`, `/topup`, `/promo`, `/obuna`, `/lang`.

---

## 🔒 Xavfsizlik

`security.py` quyidagilarni ta'minlaydi:

- Bot tokeni va API kalitlari log faylga ham, log kanaliga ham tushmaydi.
- Bot faqat shaxsiy chatda ishlaydi. Guruhga qo'shilsa, o'zi chiqib ketadi.
- Ko'p xabar yuborgan foydalanuvchi bir daqiqaga to'xtatiladi (10 soniyada 18 tadan ortiq xabar).
- Yuklangan fayllar tekshiriladi: Word ichidagi makros va zip-bomba, parolli yoki juda katta PDF, juda katta rasmlar.
- AI so'rovlarida foydalanuvchi matni ko'rsatma sifatida bajarilmaydi (prompt-injection himoyasi).
- Matnda `<` yoki `&` bo'lsa ham xabar yetib boradi.
- Serverda `.env` ruxsatlari avtomatik `600` qilinadi.

**Muhim:** `.env`, `users.json` va loglarni hech qachon GitHub'ga yuklamang. Ular `.gitignore` da turibdi.
Agar kalit tasodifan oshkor bo'lsa, uni darhol yangilang: @BotFather'da `/revoke`, DeepSeek va Gemini panelida yangi kalit oling.

---

## 🌐 Serverga joylash

Linux serverga SSH kalit bilan joylash, systemd xizmatini sozlash va yangilash tartibi:
[SERVERGA_YUKLASH.md](SERVERGA_YUKLASH.md).

---

## 📌 Eslatma

Bot yaratgan matnlar o'quv maqsadida namuna sifatida mo'ljallangan. Adabiyotlar ro'yxatidagi manbalarni
Google Scholar yoki kutubxona orqali tekshirib chiqing, chunki AI ba'zan yil yoki nashriyotni xato yozishi mumkin.

**Muallif:** [@Eldor24](https://github.com/Eldor24) · Telegram: [@endlessscience](https://t.me/endlessscience)

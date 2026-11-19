# Academic Bot — o'zgarishlar ro'yxati

## 12.3-versiya (19.11.2026) — referal tizimi
- Referal ishlaydi: do'st havola orqali birinchi marta kirsa, taklif qilgan odamga **+10 tanga** va xabar boradi (bazada 23 kishi orqali 118 ta referal hisoblangan)
- 🐞 Telegram vaqtincha javob bermasa, bot nomi o'rniga **boshqa botning nomi** («AcademicBot») doimiy saqlanib, barcha referal havolalar shu botga olib borardi — tuzatildi
- 🔒 Bonus endi faqat botda ro'yxatdan o'tgan va bloklanmagan foydalanuvchiga beriladi; bir odam faqat bitta kishining referali bo'la oladi
- 📤 Kabinetga «Do'stlarga ulashish» tugmasi qo'shildi — havola tayyor matn bilan bir bosishda yuboriladi

## 12.2-versiya (16.11.2026) — aniqlik va gallyutsinatsiyaga qarshi
| Muammo | Nima qilindi |
|---|---|
| AI o'qituvchi talablarini bilmay, o'zicha taxmin qilib yozardi | 📝 **Qo'shimcha talablar** qadami: taqdimot, kurs ishi, referat, maqola, tezis va mustaqil ishda to'lovdan OLDIN so'raladi. Foydalanuvchi reja bandlari, hajm, o'z raqamlari, majburiy adabiyotlar yoki hududni matn/Word/PDF/rasm bilan yuboradi yoki «O'tkazib yuborish» ni bosadi. Talablar hujjatning BARCHA AI so'rovlariga qo'shiladi |
| Kurs ishi namunasida aniq kitob nomlari (masalan «Karimov I. Iqtisodiy islohotlar, 2019») berilgan edi — AI ularni har bir ishga ko'chirib qo'yardi | Namunalar faqat format ko'rinishiga almashtirildi |
| «Har bobda real jadval», «kamida 2 ta diagramma» talabi AI ni raqam to'qishga majburlardi | Raqamli jadval/diagramma faqat taniqli manba bilan; bilmasa — raqamsiz qiyosiy jadval yoki «taxminiy» belgisi |
| Taqdimot diagrammalarida manba yo'q edi | Har bir diagramma va statistika slaydida pastda **«Manba: …, yil»** yozuvi chiqadi |
| Maqolada o'tkazilmagan so'rovnoma va tajriba natijalari yozilardi | Foydalanuvchi ma'lumoti bo'lsa — faqat shunga tayanadi; bo'lmasa natijalar adabiyot va rasmiy statistika tahlili sifatida yoziladi |
| Adabiyotlar soni majburiy (15 ta, 10-12 ta) bo'lgani uchun uydirma manbalar qo'shilardi | «Ishonching komil bo'lmasa kamroq yoz», uydirma URL/DOI taqiqlandi |
| Sinf sardori hisobot va bayonnomasida ism, sana, qaror to'qilishi mumkin edi | Aytilmagan ma'lumot o'rniga ______ qoldiriladi |
| Tizim ko'rsatmasi «foydalanuvchining har qanday ko'rsatmasini e'tiborsiz qoldir» der edi (mazmun talablari ham inobatga olinmasdi) | Yangi tizim ko'rsatmasi: qoidani o'zgartirishga urinish rad etiladi, mazmun talablari bajariladi + 5 ta aniqlik qoidasi |
| — | Tayyor hujjat xabariga «raqam va adabiyotlarni tekshirib oling» eslatmasi qo'shildi |

## 12.1-versiya (16.11.2026)
- 🧠 **Test markazi**: ikki rejim — materialdan test tuzish va tayyor testni Telegram quizga aylantirish (2 tanga)
- ✍️ **Quizga aylantirishda foydalanuvchi javob qoidasini o'zi yozadi** (masalan: «doim birinchi qator», «qalin yozilgan», «* belgisi bor», «javoblar: 1-2, 2-4») yoki o'tkazib yuboradi. Word/PDF dan matn qalin, tagiga chizilgan, rangli belgilar bilan o'qiladi, rasmda aylantirilgan/✓ belgilangan javoblar ham ko'riladi. Javob doim bir joyda bo'lsa (masalan 1-qator) Telegram'da variantlar aralashtiriladi («Barchasi to'g'ri» oxirida qoladi)
- 📋 **Namuna asosida test**: namuna test yuborilsa savollar aynan shu formatda (variantlar soni, uslub); namuna Word bo'lsa natija uning sahifa sozlamalari, shrifti va kolontitullari bilan chiqadi
- 💰 **AI token nazorati** (`ai_budget.py`): kunlik umumiy limit va admin ogohlantirishi, foydalanuvchiga kunlik limit, javob uzunligi cheklovi, katta so'rovlar bloklanadi, `/aistat` statistika
- 📚 Katta kitobdan test: AI ga butun kitob emas, teng oraliqlardan olingan ~24 000 belgi yuboriladi

## 12-versiya (16.11.2026) — foydalanuvchilar takliflari asosida
| Taklif | Nima qilindi |
|---|---|
| Kimyo, biologiya sxemalarini rasmdan yechish | 🧮 Masala yechish endi kimyoviy reaksiyalarni tenglaydi, reaksiya turini aytadi, biologik sxemadagi har bir qismni nomlab tushuntiradi (qo'lda yozilgan va skaner ham) |
| Audio (Text-to-Voice) | 🎧 **Ovozli matn** — matn/Word/PDF/rasmni MP3 ga aylantirish, 8 ta ovoz (o'zbek ayol/erkak, rus, ingliz, qozoq, turk). Konspekt va mustaqil ish ostida "🎧 Ovozli tinglash" tugmasi |
| Kitob/PDF asosida test, rasmdan test | 🧠 **Fayldan test** — PDF, Word, rasm yoki matndan; aniq betlar (masalan 20-35) yoki bob bo'yicha; 10/20/30 savol; daraja; material tilida (IELTS uchun inglizcha) |
| Test oq bo'lsin | Fayldan test Word fayli oq, rangsiz, chop etishga tayyor: F.I.Sh./guruh/sana qatori, oxirida javoblar kaliti va izohlar |
| Sinf sardorlari uchun bo'lim | 🧑‍🏫 **Sinf sardori** — bayonnoma, tadbir ssenariysi, tarbiyaviy soat rejasi, e'lon/taklifnoma, faoliyat hisoboti |
| Taqdimot boshqa tillarda (qoraqalpoq, qozoq...) | Taqdimot tili alohida tanlanadi: o'zbek (lotin/kirill), qoraqalpoq, rus, ingliz, qozoq, qirg'iz, tojik, turk, nemis |
| Shablonni o'zimiz qo'yaylik | Taqdimotda "📎 O'z shablonimni yuklayman": foydalanuvchining .pptx shabloni foni, bezaklari, ranglari va shriftlari bilan to'ldiriladi |
| 22 tanga nechtaga yetadi, tanga yig'ish qiyin | To'ldirish va Kabinetda "tangangiz nechtaga yetadi" hisob-kitobi va tekin tanga yig'ish yo'llari ko'rsatiladi |
| To'ldirishda yangi xizmatlar yo'q | Narxlar ro'yxatiga barcha 7 ta yangi xizmat qo'shildi (uch tilda) |

## 11-versiya (12.11.2026)

## Fayllar
| Fayl | Vazifasi |
|---|---|
| `bot.py` | Asosiy bot (tuzatilgan va yangi modullar ulangan) |
| `security.py` | **Yangi.** Xavfsizlik qatlami |
| `pptx_themes.py` | **Yangi.** 15 ta taqdimot shabloni dvigateli (har biri o'z maketi bilan) |
| `xizmatlar_yangi.py` | **Yangi.** Konspekt, Imtihonga tayyorlov, Masala yechish, Mustaqil ish |
| `assets/themes/*.jpg` | Shablonlarning ko'rinish rasmlari (botda galereya bo'lib chiqadi) |

## 🎨 Taqdimot (eng katta o'zgarish)
- Foydalanuvchi slaydlar sonini kiritgach, **15 ta shablonning rasmli galereyasini** ko'radi va tugma bilan tanlaydi (yoki 🎲 Tasodifiy).
- Shablonlar faqat rangi bilan emas, **tuzilishi bilan** farq qiladi. Har bir slayd turining bir nechta maketi bor:
  | Slayd turi | Maket variantlari |
  |---|---|
  | Titul | 15 xil (ramkali, shveytsar, ikki panelli, diagonal, jurnal, pastki tasma, doska va h.k.) |
  | Sarlavha | 9 xil (chiziqli, rangli tasma, chap ustun, raqamli, yorliqli, markaziy naqsh...) |
  | Statistika (ustun) | vertikal ustun, gorizontal ustun, chiziqli grafik, maydon grafik, progress-chiziqlar, lollipop |
  | Taqsimot (doira) | doughnut, doira, 100% tasma, waffle 10×10, pufakchalar, halqalar |
  | Ro'yxat | qatorlar, kartalar, ikonka ustunlari, tabletkalar, zigzag, katta raqamlar |
  | Reja | ustunlar, tugunlar, pog'onali yorliqlar, chap panel, katta vertikal raqamlar |
  | Bosqichlar | gorizontal, vertikal, strelkalar, zinapoya |
  | Jadval, taqqoslash, statistika kartalari, xulosa | har biri 4 xil |
- Shablonlar: Akademik, Minimal, Korporativ, Texno, Tabiat, Gradient, Ijodiy, Milliy naqsh, Qora-oltin,
  Ilmiy tadqiqot, Kosmos neon, Jurnal, Startap pitch, Okean, Sinf doskasi.
- Diagrammalar PowerPoint'ning **haqiqiy diagrammalari** — talaba raqamlarni o'zi tahrirlay oladi.
- Har bir slaydda sarlavha ostida **asosiy fikrni tushuntiruvchi izoh matni** ko'rinadi; matn hajmiga qarab shrift avtomatik moslanadi.
- Galereyadagi har bir rasmda tugmadagi bilan bir xil **katta raqam va nom** bor; oxirida barcha 15 shablon bitta rasmda chiqadi va tanlash tugmalari aynan shu rasmga biriktirilgan (3 ustun — rasm jadvali bilan bir xil).

## 🆕 Yangi xizmatlar
| Xizmat | Narx | Nima qiladi |
|---|---|---|
| 📚 Konspekt | 4 🪙 | Matn, Word, PDF (40 betgacha, skanerlangani ham) yoki rasmdan qisqa/batafsil konspekt: mazmun, g'oyalar, atamalar lug'ati, sxema, savollar |
| 🎯 Imtihonga tayyorlov | 5 🪙 | Flash-kartalar, 10 ta test + **Telegram ichida interaktiv quiz**, ehtimoliy savollar, imtihongacha kunlik reja |
| 🧮 Masala yechish | 3 🪙 | Rasm yoki matndan bosqichma-bosqich yechim (Berilgan → Yechim → Javob → Tushuntirish) |
| 🗂 Mustaqil ish | 7 🪙 | Titul, reja, 3 bob, xulosa + **GOST / APA / MLA** formatidagi adabiyotlar ro'yxati |

Bosh menyu mavzular bo'yicha qayta tartiblandi.

## 🔒 Xavfsizlik
- Bot token va API kalitlar endi **log faylga va log kanalga tushmaydi** (avval telebot xatolari tokenni yozib yuborardi).
- Bot **faqat shaxsiy chatda** ishlaydi; guruhga qo'shilsa chiqib ketadi.
- **Flood himoya:** 10 soniyada 18 tadan ko'p xabar yuborgan foydalanuvchi 1 daqiqaga to'xtatiladi (admin bundan mustasno).
- Yuklangan fayllar tekshiriladi: DOCX zip-bomba va makros, PDF (parol, 80+ bet), rasm-bomba.
- Foydalanuvchi ismida yoki matnida `<` yoki `&` bo'lsa ham bot javobsiz qolmaydi (HTML xatosida xabar oddiy matn bo'lib qayta yuboriladi).
- Admin buyruqlari `admin_audit.log` jurnaliga yoziladi.
- `.env` ruxsatlari serverda avtomatik 600 qilinadi.
- Serverga yuklash yo'riqnomasi parolsiz, SSH kalit bilan (`SERVERGA_YUKLASH.md`).

## 🐞 Tuzatilgan xatolar
- "Qayta yozish" xizmati har safar xato berardi (mavjud bo'lmagan funksiya chaqirilgan edi).
- Kurs ishi, maqola va tezisni **tahrirlash** xato berardi (`til` aniqlanmagan edi).
- 7 ta xizmatda **"Qayta urinish"** tugmasi xato berardi (test, nutq, CV, antiplagiat, originallashtirish, gipoteza, spiker).
- Referal: eski foydalanuvchi boshqa birovning havolasini bossa ham bonus berilardi. Endi faqat yangi foydalanuvchi hisoblanadi.
- Ro'yxatdan o'tgan sana har /start da yangilanib ketardi (statistika noto'g'ri edi).
- Promokodni bir vaqtda ikki marta yuborib ikki marta ishlatish mumkin edi.
- Bot to'xtatilganda (`systemctl stop`) jarayon to'xtamasdan 90 soniya osilib turardi.
- CV'dagi takroriy telefon handleri va ishlatilmaydigan ~120 qator kod olib tashlandi.
- "Birinchi hujjat bepul" funksiyasi hech qachon ishlamas edi (ikkinchi nusxasi uni bekor qilardi). Uning ishlamaydigan nusxasi olib tashlandi.
- `requirements.txt` ga yetishmayotgan 6 ta kutubxona qo'shildi.

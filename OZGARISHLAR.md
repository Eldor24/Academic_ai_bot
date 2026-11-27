# Academic Bot — o'zgarishlar ro'yxati

## 16.8-versiya (27.11.2026) — Snoskalar: yil va bet har doim, eski adabiyot kirillda; Taqdimot + savol kartochkalari; taqdimot sxemalari
- 📎 **Kurs ishi snoskalari va adabiyotlar ro'yxati:**
  - Har bir snoskada **nashr yili ham, kitob beti ham** bo'ladi. Ro'yxatga faqat nashriyoti, yili va umumiy bet
    soni (kitob) yoki jurnal nomi va betlari (maqola) aniq manbalar kiradi; to'liq bo'lmaganlari to'ldiriladi yoki
    chiqariladi (`_toliqmi`, `_adab_toliqlash`). Nazorat bosqichi 2 marta aylanadi; hisobotda «yilsiz» hisoblagichi.
  - Maqola snoskasida endi maqolaning bet oralig'i emas, **iqtibos olingan aniq bet** («– B. 47.») turadi.
  - **1993-yilgacha nashr etilgan o'zbekcha adabiyotlar asl holidagidek kirill yozuvida** yoziladi (ro'yxatda ham,
    snoskada ham): «Ҳожиев А. ... – Тошкент: Ўқитувчи, 1974. – 304 б.», snoskada «– Б. 45.» (`matn_norm.lotin_kiril`).
  - Tuzatilgan xato: kirill «б.» bilan yozilgan umumiy bet soni snoskadan olib tashlanmasdi.
- 🎴 **Rag'batlantirish → «Taqdimot + savol kartochkalari»** (`ragbat_kartochka.py`):
  - Mavzu yoziladi → odatdagi taqdimot oqimi (ism, slaydlar soni, til, **dizayn shabloni tanlash**) → taqdimot
    tayyor bo'lgach kartochka sozlamasi avtomatik ochiladi. Oxirgi taqdimotdan yoki o'z .pptx faylidan ham mumkin;
    har qanday taqdimot yuborilgandan keyin «🎴 Shu taqdimotdan savol kartochkalari» taklifi chiqadi.
  - Sozlamalar: varaqda **4/6/8/9/12** ta, varaqlar **1–6** (jami = varaqda × varaq, ko'pi bilan 72), shakl
    (to'rtburchak, yumaloq, oval, oltiburchak, chipta), **9 rang yoki kamalak**, javob joyi (kartada — buklanadigan
    pastki qism / orqa tomonda — ikki tomonlama chop, ko'zgu tartibida / faqat kalitda), qiyinlik (Bloom bo'yicha).
  - Savollar faqat slaydlar mazmunidan; har bir savol-javob alohida AI bosqichida slaydlar bilan solishtirib
    tekshiriladi (tuzatiladi yoki chiqariladi), takrorlar olib tashlanadi, taqdimot tartibida raqamlanadi.
  - Natija: chop etishga tayyor vektor PDF (kesish chiziqlari, raqam, Bloom darajasi), oxirida o'qituvchi uchun
    javoblar kaliti, birinchi varaq rasmi va darsda qo'llash usullari. Boshqa shakl/rangda qayta chizish — bepul.
  - Narx: 24 tagacha — 3 🪙, 25–72 — 5 🪙.
- 👥 **«Mavzularni taqsimlash» (o'quvchilar dars o'tadi):** o'quvchilar va mavzular ro'yxatidan guruhlar (1–4 kishi,
  ro'yxat tartibida yoki aralash), har chiqishga sana (har dars kuni / kun ora / haftada 1, yakshanba o'tkaziladi,
  bir darsda 1–2 chiqish), guruh ichida rollar, talaba darsi tuzilmasi, tayyorgarlik eslatmasi, **har bir chiqish
  uchun oldindan to'ldirilgan o'qituvchi rubrikasi** (20 ball → 5 ballik baho) va «Ikki yulduz va bitta tilak»
  o'zaro baholash varag'i (Word) — 2 🪙.
- 📊 **Taqdimot bo'limi kuchaytirildi:**
  - 3 ta yangi raqamsiz diagramma turi, 15 ta shablonning hammasida: **«sxema»** (tushunchalar xaritasi —
    rasmiy shablonlarda daraxt, boshqalarida markazli xarita), **«sikl»** (davriy jarayon, yo'nalish strelkalari bilan),
    **«piramida»** (darajalar/ierarxiya).
  - AI qoidalari: kamida 3 xil vizual slayd, «matn» slaydlari ≤ 1/3; ilmiy-akademik uslub (shaxssiz, atamaga
    ta'rif, sabab-oqibat/mexanizm/miqdor, olim va yil faqat aniq bilinganda).
  - Slaydni tahrirlash endi slayd turini (diagramma, sxema, jadval) va manbasini saqlaydi; foydalanuvchi talabi
    «faqat ma'lumot» chegarasida beriladi.
- 📣 Marketing: yangi «Taqdimot + savol kartochkalari» posti (e'lon muddatida birinchi yuboriladi).

## 16.7-versiya (26.11.2026) — Kurs ishi: o'qituvchi e'tirozlari, snoskalar nazorati; professional test generatori
- 📄 **Kurs ishi** (o'qituvchi e'tirozlari bo'yicha):
  - Kurs ishida endi faqat **«XULOSA»** bo'limi bor. Takliflar alohida raqamlanmaydi — xulosaning oxirgi
    abzaslarida uzluksiz matn sifatida beriladi. Mundarija, «Kurs ishining tuzilishi» bandi va bot xabarlari ham
    shunga moslandi. BMI'da «Xulosa va tavsiyalar» avvalgidek qoladi.
  - Kirishga **«Tadqiqotning materiali (amaliy bazasi)»** bandi qo'shildi: tahlil qaysi aniq qoidalar, lug'atlar,
    darsliklar, matnlar, statistik manbalar asosida olib borilgani adabiyotlar ro'yxatidagi nomlar bilan ochiq
    ko'rsatiladi. Tahlil boblari shu materiallarga tayanib, aniq misollar bilan yoziladi.
  - Kurs ishi matnida «mazkur bitiruv malakaviy ishi», «ВКР», «graduation thesis» kabi iboralar ishlatilmaydi:
    AI'ga qat'iy qoida beriladi, qolganlari avtomatik tuzatiladi.
- 📎 **Snoskalar va adabiyotlar ro'yxati — qat'iy moslik.** Topilgan va tuzatilgan sabablar:
  - Hajmni to'ldirish uchun qo'shiladigan abzaslarda havolalar majburan o'chirilardi (olim fikri snoskasiz
    qolardi, ayniqsa 3.2-3.3 fasllarda). Endi bu abzaslar ham ro'yxatdagi manbadan bet bilan havola oladi.
  - «Bitta manba ko'pi bilan 2 marta» va «abzasda ko'pi bilan 1 ta havola» cheklovlari snoskalarni jim o'chirib
    yuborardi. Endi snoska hech qachon o'chirilmaydi, manbalar paragraflarga oldindan taqsimlanadi.
  - Snoska kaliti topilmasa, ro'yxatdagi shu raqamli BOSHQA kitob olinib qolishi mumkin edi. Tuzatildi.
  - Kurs ishini tahrirlashda (bobni qayta yozish) AI ro'yxatda yo'q, uydirma manbalarni snoskaga qo'yardi. Endi
    tahrir `kurs_ishi.qism_qayta_yoz` orqali, faqat adabiyotlar ro'yxatidan, bet bilan qilinadi.
  - **Yangi nazorat bosqichi** (`_snoska_nazorati`): olim nomi bor-u havolasi yo'q har bir gapga mos manba va bet
    qo'yiladi yoki gap muallif nomisiz qayta yoziladi; kitobga snoskada bet yo'q bo'lsa, bet aniqlanadi.
  - **Adabiyotlar to'liqligi** (`_adab_toliqlash`): yili, shahri yoki nashriyoti yo'q manbalar to'ldiriladi,
    aniq bilinmaganlari ro'yxatdan chiqariladi. Internet manbalari GOST bo'yicha: «Nomi. – URL: ...
    (murojaat sanasi: ...)».
  - `snoska_hisobot()`: har snoska ro'yxatdagi manbaning aynan o'zimi, betsiz kitob snoskasi va havolasiz olim
    fikri qolmaganmi — logga yoziladi va testlarda tekshiriladi.
- 🏅 **Rag'batlantirish:** yorliqdagi sana — hujjat yaratilgan kundagi Toshkent sanasi. Dizaynlar namunasidagi
  sana ham bugungi kunni ko'rsatadi.
- 📝 **Test generatori qaytadan yozildi** (`test_pro.py`):
  - Auditoriya tanlanadi: 1-4 sinf (3 variantli, sodda), 5-9 sinf, 10-11 sinf / litsey, abituriyent (DTM
    uslubi), talaba (OTM), mutaxassis (attestatsiya).
  - Test mavzu bo'yicha yoki ma'ruza matni / Word / PDF / rasm asosida tuziladi.
  - Qiyinlik: oson / o'rta / qiyin / aralash — Bloom taksonomiyasi bo'yicha aniq taqsimot. Avval reja tuziladi:
    har savol mavzuning boshqa jihatini tekshiradi.
  - **Mustaqil tekshiruv:** har savolni AI kalitga qaramay qayta yechadi. Mos kelmasa yoki noaniq bo'lsa, savol
    almashtiriladi.
  - Javoblar A/B/C/D bo'yicha teng taqsimlanadi. «Barchasi to'g'ri», «hech biri» kabi variantlar taqiqlangan.
  - 1, 2 yoki 4 variant: savollar va javoblar tartibi aralashtiriladi, har variantga alohida kalit beriladi.
  - Ikki fayl:
    - o'quvchilar uchun — javobsiz test, yo'riqnoma, tavsiya etilgan vaqt, javoblar varaqasi;
    - o'qituvchi uchun — kalitlar, har savol izohi, Bloom darajasi, baholash shkalasi.
  - Narx: 30 tagacha savol — 4 🪙, 40-50 ta — 6 🪙.
  - Eski oqimdagi xato tuzatildi: «O'rta» darajada AI'ga ko'rsatma umuman berilmasdi (`orta` / `o'rta` farqi).
    Javoblar o'quvchi nusxasida ✅ bilan ko'rinib turardi — endi faqat o'qituvchi faylida.

## 16.6-versiya (25.11.2026) — 🏅 Rag'batlantirish bo'limi, foydalanuvchi navbati, marketing
- 🏅 **Yangi bo'lim «Rag'batlantirish»** (`ragbat.py`, `ragbat_dizayn.py`) — 3 tilda, bosh menyuda «Sinf sardori»
  yonida:
  - **Tashakkurnoma va yorliqlar.**
    - Turlar: tashakkurnoma, faxriy yorliq, maqtov yorlig'i, sertifikat, diplom (I-II-III daraja),
      minnatdorchilik xati.
    - Bitta kishiga (2 🪙) yoki guruhga — 2-60 ism, har biriga alohida varaq (4 🪙). Bo'sh shablon — ism qo'lda
      yoziladi (1 🪙).
    - Hujjat tili alohida tanlanadi: o'zbek, rus yoki ingliz.
    - Sabab matnini AI rasmiy uslubga keltiradi; matn boshiga «=» qo'yilsa, o'zgarishsiz qoladi. Tayyor
      namunalar ham bor.
    - Beruvchi: lavozim, F.I.Sh., tashkilot, shahar — shaxs, jamoa yoki tashkilot nomidan.
  - **8 ta dizayn**, kod bilan chiziladi (vektor PDF, chop etishda sifat yo'qolmaydi): rasmiy oltin, milliy naqsh,
    akademik, zamonaviy, premium, tabiat va bolalar uchun 2 ta (quvnoq, kosmos). Dizayn tanlashda 8 tasining
    namunasi bitta rasmda ko'rsatiladi.
  - **Natija fayllari.** Bitta kishi: rasm + PDF. Guruh: barcha varaqlar bitta PDF'da va har kimning rasmi
    alohida ZIP'da (ota-onalarga yuborish uchun).
  - **🎓 Dars to'plami (6 🪙).** Mavzu yoki ma'ruza matni / fayl / rasm asosida tuziladi.
    - Tarqatma: asosiy tushunchalar; Bloom taksonomiyasi bo'yicha savollar (test, to'g'ri/noto'g'ri, bo'sh joy
      to'ldirish, moslashtirish, qisqa javob, ijodiy); refleksiya «3-2-1»; o'z-o'zini baholash.
    - O'qituvchi uchun: javob kaliti, baholash mezoni, mukofot shkalasi, dars rejasi, maqtov iboralari.
    - Yakunda tugmalar: «Natijalarni kiritib taqdirlash» va «Shu mavzuda taqdimot» (taqdimot oqimi mavzu
      kiritilgan holda boshlanadi).
  - **🏆 Natijalar bo'yicha taqdirlash (5 🪙).** «Ism — ball» ro'yxatidan ishlaydi: 17, 4,5, 9/20 kabi
    yozuvlarni tushunadi.
    - Rejimlar:
      - ballga qarab: 86% va yuqori — faxriy yorliq, 71-85% — maqtov yorlig'i, 56-70% — tashakkurnoma,
        qolganlarga ishtirok sertifikati;
      - g'oliblar: 1-2-3 o'rin diplom, qolganlarga tashakkurnoma;
      - hammaga bir xil hujjat;
      - 56% dan pastlarga hujjat berilmaydi.
    - Har kim uchun shaxsiy maqtov matni yoziladi: natijani emas, harakat va strategiyani e'tirof etadi.
    - Natijalar qaydnomasi Word'da: o'rtacha natija, mukofotlar taqsimoti va yordam kerak bo'lganlar.
  - **💡 Bepul qo'llanma:** samarali rag'batlantirishning 7 tamoyili (Hattie & Timperley, Black & Wiliam, Dweck,
    Brophy).
- 📥 **Foydalanuvchi navbati.** Hujjat tayyorlanayotganda yana so'rov yuborilsa, u rad etilmaydi: navbatga
  qo'yiladi va oldingisi tugashi bilan avtomatik boshlanadi.
  - Har foydalanuvchida bir vaqtda faqat bitta ish bajariladi va navbatda yana 2 tasi turishi mumkin. To'la
    bo'lsa, xabar beriladi va tanga qaytariladi.
  - Navbatda kutgan ish qabul qilingan paytdagi ma'lumotlar bilan bajariladi: shu orada boshqa bo'limga
    o'tilsa ham aralashib ketmaydi.
  - Umumiy parallel ishlar soni avvalgidek 3 ta, 1 GB RAM'li server uchun. Bajarilayotgan yoki navbatdagi
    ishi bor foydalanuvchining sessiyasi RAM tozalashda o'chirilmaydi.
- 📣 **Marketing.** Yangi postlar qo'shildi: «Rag'batlantirish», «Test savollari», «Qayta yozish».
  - Yangi xizmat e'lonlari (`"yangi": "YYYY-MM-DD"`) navbatdagi slotda hali olmagan har bir foydalanuvchiga
    aylanmadan oldin yuboriladi (45 kun davomida). Keyin odatdagi aylanmaga qo'shiladi.
- 🔤 Shriftlar (OFL): Forum, Old Standard, Marck Script, Lobster, PT Sans — `assets/fonts/SHRIFTLAR.txt`.

## 16.5-versiya (24.11.2026) — O‘/G‘ va tutuq belgisi, titul tekisligi, joriy yil
- 🔤 **O'zbek alifbosi belgilari.** Barcha Word va PowerPoint hujjatlarda (kurs ishi, BMI, mustaqil ish, referat,
  maqola, tezis, taqdimot va boshqalar) saqlash paytida belgilar standartga keltiriladi:
  - o‘ va g‘ teskari vergul (‘) bilan yoziladi: o‘zbek, g‘oya, tog‘;
  - tutuq belgisi ’ bilan yoziladi: ta’lim, san’at, ma’lumot.

  Hujjat o'zbek-lotin matni ekani avtomatik aniqlanadi. Ruscha, inglizcha va kirill yozuvidagi hujjatlarga
  tegilmaydi. Yangi modul: `uz_apostrof.py`. Mundarija sahifalash hisobida ham belgi kengligi shunga mos o'lchanadi.
- 📐 **Kurs ishi tituli.** O'ng tomondagi blokda («... kafedrasi», «Ro‘yxatga olindi №», sana) uchala qator endi bir
  xil chap chiziqdan boshlanadi va blok o'ng chetga yopishib turadi. Avval qisqa qator ichkariga kirib, qiyshiq
  ko'rinardi.
- 📅 **Yil.** Titul pastidagi «Shahar – yil», referat va taqdimot titulidagi yil hujjat tayyorlangan kundagi yil
  bo'ladi: 2027-yilda avtomatik 2027 chiqadi. Yil Toshkent vaqti bo'yicha olinadi (server UTC vaqtida ishlaydi).
  Shahar so'ralganda ko'rsatiladigan misol («Toshkent – 2026») ham joriy yilni ko'rsatadi.

## 16.4-versiya (21.11.2026) — Taqdimot: slaydlarda ma'lumot ancha ko'p
- Taqdimot endi ikki bosqichda tayyorlanadi:
  1. AI avval batafsil reja tuzadi: har slayd uchun sarlavha, turi va aniq yoritiladigan fikrlar. Reja mantiqiy
     ketma-ketlikda boradi: tushuncha → nazariya → tarkib → jarayon → O'zbekiston va xorij tajribasi → muammolar →
     istiqbol → xulosa.
  2. Slaydlar 5 tadan parallel, to'liq yoziladi.

  Avval 20 tagacha slayd bitta so'rovda yozilardi va matn qisqarib qolardi. Reja olinmasa, bot avvalgi bir so'rovli
  usulga o'tadi.
- Yangi hajm talablari:
  - matn slaydi: 7-9 gap (130-170 so'z);
  - ro'yxat: 5-6 band, «Atama — izoh» shaklida;
  - xulosa: 5 ta aniq natija yoki tavsiya;
  - bosqichlar: 4-5 bosqich;
  - taqqoslash: har tomonda 4-5 band;
  - jadval: 4-5 qator;
  - har slayd ostidagi asosiy fikr: 2 gap.
- Ro'yxatlarda atama qalin harf bilan ajratiladi. Slayd ostidagi izoh uchun joy 220 belgidan 300 belgiga oshirildi.
  Kartochkalarda matn ramkadan chiqib ketmaydi.
- Mazmunli slayd matnida «update», «sanab» kabi so'zlar uchraganda u xizmat slaydi deb noto'g'ri o'chirilib
  ketardi. Bu xato tuzatildi.

## 16.3-versiya (21.11.2026) — Ilmiy bo'lim: dalillarga asoslangan javoblar, yagona Word shabloni, xavfsizlik
- 💡 **Mavzu tanlash** endi haqiqiy tadqiqotlarga tayanadi. Bot so'rovdan soha va qidiruv so'zlarini ajratadi,
  OpenAlex'dan 2019-yildan keyingi eng yaqin ishlarni oladi va yo'nalishda nechta ish borligini hisoblaydi.
  Har bir variantda: obyekt, predmet, dolzarblik, tadqiqot bo'shlig'i, ilmiy yangilik, ishchi gipoteza, metodlar,
  ma'lumot manbalari, xavf. Tayanilgan ishlar [n] bilan ko'rsatiladi, hujjat oxirida GOST bo'yicha
  «Asoslovchi adabiyotlar» ro'yxati beriladi.
- 🎯 **Mavzu va ixtisoslik**: mavzuga eng yaqin so'nggi ishlar va yo'nalish dinamikasi (so'nggi 5 yil va jami ishlar
  soni) qo'shildi. «Ilmiy yangilik va takrorlanish xavfi» bandi [n] havolalari bilan beriladi.
- 🧑‍🏫 **Ilmiy maslahatchi**:
  - Savol turini aniqlaydi: tartib-qoida yoki ilmiy mazmun. Ilmiy savolda so'nggi tadqiqotlarni dalil sifatida
    oladi.
  - Javob akademik tuzilmada beriladi: Qisqa javob → Asoslash → Amaliy qadamlar → E'tibor bering → Keyingi qadam.
  - Manbalar ro'yxatida rasmiy hujjatlar, OAK yangiliklari va ilmiy maqolalar (GOST) ko'rsatiladi.
- 📄 Bo'limdagi barcha Word hujjatlar yagona shablonda (`ilmiy_shablon.py`). Shablonda sarlavha bloki, parametrlar
  jadvali, kolontitul va sahifa raqami bor. Individual reja rasmiy shaklda: «TASDIQLAYMAN» bloki, kulrang sarlavhali
  jadvallar, imzo jadvali.
- ↩️ «Orqaga» tugmasi xizmat ichidan «Ilmiy izlanuvchi» menyusiga qaytaradi.
- 🛡 **Xavfsizlik**:
  - Foydalanuvchi matni va tashqi manbalar (OpenAlex, OAK yangiliklari) AI'ga «faqat ma'lumot» chegarasida beriladi.
    Ular ichidagi ko'rsatmalar bajarilmaydi (prompt-injection himoyasi).
  - Callback ma'lumotlari qat'iy formatda va chegaralarda tekshiriladi. Eskirgan yoki qo'lda o'zgartirilgan tugmalar
    jim e'tiborsiz qoldiriladi.
  - Tugmalarni tez-tez bosish cheklangan. Bepul Word ro'yxatini qayta olish uchun 20 soniya kutish kerak.
  - «🗑 Tarixni tozalash» tugmasi tasdiqlash bilan ishlaydi.
- 🐍 bot.py: Python 3.11 bilan mos kelmaydigan f-string tuzatildi.

## 16.2-versiya (21.11.2026) — «📚 Adabiyot va tadqiqotlar»: faqat yangi va mavzuga mos manbalar
- Faqat **2019-yildan keyin** chop etilgan ishlar olinadi. Avval asosiy ro'yxat 2012-yildan, «eng ko'p iqtibos»
  bo'limi 2000-yildan edi, dissertatsiyalar uchun esa yil chegarasi yo'q edi.
- Qaytarib olingan (retracted) maqolalar va muqova/mundarija kabi yozuvlar chiqarib tashlanadi.
- Chuqurroq qidiruv: 3 ta inglizcha (asosiy, sinonimli, torroq), ruscha va o'zbekcha so'rov, iqtiboslar bo'yicha,
  dissertatsiyalar — jami 100+ nomzod ish.
- AI har bir topilgan ishning sarlavhasi va annotatsiyasiga qarab mavzuga mosligini baholaydi (bevosita / qisman / mos emas).
  Mos kelmaganlari olib tashlanadi. AI hech qanday manba qo'shmaydi — faqat haqiqiy topilgan ishlarni saralaydi.
- Natijada 5 ta bo'lim bor: eng yaqin tadqiqotlar, eng ko'p iqtibos qilinganlar, rus tilidagi, **o'zbek tilidagi**,
  dissertatsiyalar. Har bir ishda yili, iqtiboslar soni, mosligi, DOI yoki ochiq PDF ko'rsatiladi.
- Hujjat boshida qidiruv qanday o'tkazilgani yoziladi: nechta ish ko'rib chiqildi, qaysi yillar oralig'i.

## 16.1-versiya (21.11.2026) — 2027-yildagi yangi tizim, aniqlashtiruvchi savollar, tarix
- 🧾 «Ilmiy daraja» qo'llanmasi (3 tilda) qayta yozildi:
  - Prezidentning 26.05.2026 dagi **PF-98-son Farmoni** qabul qilingan hujjat sifatida bayon qilinadi: nimalar bekor
    qilinadi, kengashlar vakolati, OAK tanlab tekshiruvi, kengashlar reytingi, «Loyiha doirasida ilmiy daraja»;
  - o'tish muddatlari — OAK Rayosatining 391/2-son qarori;
  - OAKning 05.11.2026 dagi rasmiy savol-javoblari (Q1–Q2 majburiy emas, dissertatsiyani qayta yozish shart emas);
  - yangi nizom talablari (PhD 8 ish, DSc 16 ish ...) — «loyiha» yozuvi olib tashlandi.
- 📄 Admin uchun `/nizom <hujjat nomi, raqami, sanasi>` buyrug'i: tasdiqlangan nizom rekvizitlari qo'llanmada va
  maslahatchi javoblarida ko'rsatiladi (`/nizom -` — o'chirish).
- 📢 OAK rasmiy yangiliklari oak.uz dan avtomatik olinadi (har 6 soatda) — «📰 OAK jurnallari → 📢 OAK rasmiy
  yangiliklari» va ilmiy maslahatchi bilim bazasida.
- 🧑‍🏫 Ilmiy maslahatchi endi:
  - aniq javob uchun ma'lumot yetishmasa, avval 1–3 ta **aniqlashtiruvchi savol** beradi (PhD/DSc, shakl, ixtisoslik,
    bosqich ...). Ular uchun tanga olinmaydi, ko'pi bilan 2 marta so'raydi;
  - javob oxirida **rasmiy manbalar** ro'yxati (Farmon, OAK qarorlari, oak.uz yangiliklari) — faqat bilim bazasidagi manbalar;
  - «❓ Shu mavzuda davom etish» — oldingi suhbat konteksti saqlanadi; «🆕 Yangi savol».
- 📜 **Mening tarixim**: maslahatchi savol-javoblari va barcha ilmiy xizmatlar natijalari (mos jurnallar, adabiyotlar,
  mavzu tahlili, mavzu variantlari, bibliografiya, individual reja) foydalanuvchi bo'yicha saqlanadi. Oxirgi 30 ta yozuv
  users bazasida (cloud_db orqali doimiy) turadi; ro'yxatdan istalganini to'liq ochish mumkin.
- Individual reja: nashrlar rejasi 2027-yildan amal qiladigan talablar bo'yicha.

## 16-versiya (21.11.2026) — 👨‍🔬 «Ilmiy izlanuvchi» bo'limi
Tayanch doktorant, doktorant, mustaqil izlanuvchi va magistrantlar uchun alohida bo'lim. Bosh menyuda
«👨‍🔬 Ilmiy izlanuvchi», ichida o'z menyusi va «⬅️ Orqaga» tugmasi bor.

### 📰 OAK jurnallari — BEPUL, avtomatik yangilanadi (`oak_jurnallar.py`)
- Manba — oak.uz → «Ilmiy nashrlar ro'yxati» sahifasidagi rasmiy ro'yxat ilovasi
  (journal-statistics-oak.vercel.app). Ro'yxat ~1 700 jurnal: milliy, MDH, xorijiy; har biri fan tarmog'i bo'yicha
  holati, qo'shilgan va chiqarilgan sana, qaror raqami bilan.
- Bot ro'yxatni ishga tushganda va **har 6 soatda** tekshiradi, keshda saqlaydi. OAK yangilasa — botda ham yangilanadi.
- **🔎 Nomi bo'yicha tekshirish** (lotin/kirill): jurnal qaysi fan kodlari bo'yicha ✅ faol, qaysi biri ❌ chiqarilgan
  (sana, qaror).
- **📂 Fan tarmog'i bo'yicha**: 23 ta soha → milliy / MDH / xorijiy ro'yxati sahifalab, **📥 to'liq ro'yxat Word**da.
- **🆕 So'nggi o'zgarishlar**: oxirgi qo'shilgan va chiqarilgan jurnallar.
- **🔔 Obuna**: ro'yxatda o'zgarish bo'lsa, bot qo'shilgan/chiqarilgan jurnallarni obunachilarga (har kimga o'z tilida)
  va adminga yuboradi.

### 🌍 Xalqaro bazalar — BEPUL
OAK 10.06.2026 qarori bo'yicha A/B/C toifali bazalar (Web of Science, Scopus, JCR; AGRIS, CAS, GeoRef, PubMed, Compendex;
4 ta ijtimoiy-gumanitar baza) va endi mustaqil mezon bo'lmagan platformalar (Springer, CrossRef, DOAJ, ResearchGate,
SCImago va boshq.) — ma'lumot jonli ro'yxatdan olinadi.

### 📘 Qo'llanmalar — BEPUL (3 tilda, `ilmiy_qollanma.py`)
1. Ilmiy daraja: o'tish davri (13.11 / 30.11 / 31.12.2026), PF-98 Farmoni, yangi nizom **loyihasi** (PhD 8 ish,
   DSc 16 ish, hajm, dissertatsiyasiz yo'l), amaldagi nizom talablari
2. OAK jurnaliga maqola yuborish tartibi
3. Scopus / WoS jurnaliga maqola: tanlash, tekshirish (MJL, Scopus Sources, discontinued, hijacked), taqriz, APC
4. Shubhali va o'g'irlangan jurnallar: belgilar va tekshirish tartibi
5. Doktoranturaga qabul (VM 304-son qarori): shakllar, hujjatlar, imtihonlar, muddatlar, individual reja
6. Jurnal tanlash strategiyasi

### Arzon xizmatlar (raqobatchilardan 5–10 baravar arzon)
| Xizmat | Tanga | Nima beradi |
|---|---|---|
| 🧭 Mos jurnal topish | 3 | AI fan kodini aniqlaydi → **faqat OAK ro'yxatidagi faol** jurnallardan tanlaydi + OpenAlex orqali o'xshash maqolalar chop etilgan xalqaro jurnallar (h-indeks, OA, APC, WoS/Scopus tekshirish havolalari), Word |
| 📚 Adabiyot va tadqiqotlar | 3 | OpenAlex'dan **haqiqiy** maqolalar va dissertatsiyalar: eng mos, eng ko'p iqtibos qilingan, rus tilidagi; DOI, ochiq PDF, GOST 7.1 va APA 7 yozuvlari, Word |
| 🎯 Mavzu va ixtisoslik | 3 | moslik bahosi, xavflar, tavsiyalar, 4 ta aniqlashtirilgan variant, obyekt-predmet-maqsad |
| 💡 Mavzu tanlash | 4 | 7 ta dissertatsiya mavzusi: dolzarblik, yangilik, metod, ma'lumot manbalari, xavf (Word) |
| 📑 Individual reja | PhD 12 / DSc 18 | yakka tartibdagi ish rejasi (Word): «TASDIQLAYMAN», ma'lumotlar jadvali, maqsad-vazifalar, nazariy tayyorgarlik, 6 semestrlik kalendar reja, nashrlar rejasi (2027 loyihasi talablari bo'yicha), dissertatsiya tuzilmasi, imzolar |
| 📝 Bibliografiya | 2 | ro'yxatni GOST 7.1-2003 / APA 7 bo'yicha rasmiylashtirish, yetishmayotgan elementlarni belgilash (hech narsa qo'shmaydi) |
| 🧑‍🏫 Ilmiy maslahatchi | 1 / savol | rasmiy ma'lumotlar bazasiga tayangan javob; bilmagan raqamni o'ylab topmaydi |

- Bo'lim menyusida mavjud xizmatlar ham bor: 📝 Maqola (IMRAD), 📋 Tezis, 🔬 Ilmiy muammo, 🌍 Ilmiy tarjimon.
- Yangi modullar: `ilmiy_izlanuvchi.py`, `oak_jurnallar.py`, `openalex.py`, `ilmiy_qollanma.py`.
- Marketing: «Ilmiy izlanuvchi» haqida yangi post (3 tilda) — jami 25 ta; narxlar ro'yxatiga yangi guruh.
- Ixtiyoriy muhit o'zgaruvchilari: `OPENALEX_EMAIL` (OpenAlex «polite pool»), `OPENALEX_KEY`.

## 15-versiya (21.11.2026) — bitiruvchi va amaliyotchi talabalar uchun 3 ta yangi xizmat
Namunalar internetdan o'rganildi: VM 13.12.2024 dagi 836-son qarori (yakuniy attestatsiya, BMI, antiplagiat,
komissiya), OTMlarning BMI uslubiy ko'rsatmalari (tuzilma, hajm, baholash mezonlari), BMI himoyasi bo'yicha qo'llanmalar
(nutq 7–10 daqiqa, 7–10 slayd), amaliyot hisoboti va tavsifnoma namunalari.

### 🎓 BMI — bitiruv malakaviy ishi (35 🪙)
- Kurs ishi dvigatelida yangi «bmi» rejimi: **65–75 bet**, 3 bob × 3 paragraf, har bob boshida kirish matni
- **BMI titul varag'i**: vazirlik, OTM, fakultet, kafedra, «Qo'lyozma huquqida», muallif, mavzu,
  «… ta'lim yo'nalishi bo'yicha bakalavr darajasini olish uchun», ilmiy rahbar, **taqrizchi**,
  «Himoyaga ruxsat etildi — Kafedra mudiri ___», shahar – yil (rus va ingliz tilida ham)
- **Annotatsiya 3 tilda** (o'zbek, rus, ingliz) + kalit so'zlar — titul va mundarija orasida
- Kirishda **«Tadqiqotning ilmiy yangiligi»** bandi; «Xulosa va tavsiyalar»; 45–55 adabiyot (kamida 30 o'zbekcha)
- **Ilovalar**: 1-ilova — tayanch tushunchalar lug'ati, 2-ilova — tavsiyalarni amalga oshirish rejasi (mundarijada ham)
- III bob: takomillashtirish yo'llari; pedagogik yo'nalishda tajriba-sinov ishlarining dasturi (natija raqamlari o'ylab topilmaydi)
- Qadamlar: OTM → fakultet → **ta'lim yo'nalishi (shifri bilan)** → mavzu → ism → ilmiy rahbar va taqrizchi → kafedra → shahar
- Mundarija betlari LibreOffice bilan 3 tilda tekshirildi — hammasi aniq

### 🛡 Himoyaga tayyorlov (8 🪙) — yangi `xizmatlar_talaba.py`
- Manba: botda yaratilgan oxirgi kurs ishi/BMI (bitta tugma) yoki Word/PDF fayl, yoki matn
- **Himoya taqdimoti** (PPTX, 9–10 slayd): dolzarblik → maqsad va vazifalar → obyekt, predmet, metodlar →
  (ilmiy yangilik) → ish tuzilishi → nazariy xulosalar → amaliy natijalar → **ishning o'zidagi grafik va jadval** →
  muammolar → takliflar → xulosa
- **Himoya nutqi** (Word) — har slayd uchun, 7–9 daqiqa; **15 ta komissiya savoli va javoblar**; baholash mezonlari va maslahatlar

### 📋 Amaliyot hisoboti (8 🪙)
- Turlari: o'quv, ishlab chiqarish, pedagogik, malakaviy, bitiruv oldi
- Word: titul → **amaliyot kundaligi** (har ish kuni: №, sana, bajarilgan ishlar, rahbar imzosi; haftalar bo'yicha;
  pedagogik amaliyotda 6 kunlik hafta) → hisobot (kirish, tashkilot tavsifi, bajarilgan ishlar, ko'nikmalar,
  muammolar va takliflar, xulosa) → **tavsifnoma** loyihasi (baho, imzo, M.O'. bo'sh)
- Tashkilot haqida bilinmagan faktlar o'ylab topilmaydi — «______» qoldiriladi; talaba o'zi yozgan ishlar asos qilinadi

### Boshqa
- Bosh menyuda yangi qatorlar: 🎓 BMI · 🛡 Himoyaga tayyorlov · 📋 Amaliyot hisoboti; narxlar ro'yxatiga qo'shildi
- Marketing postlari: BMI, himoya va amaliyot haqida 3 ta yangi post (3 tilda) — jami 24 ta
- 🐞 Konspekt, test markazi va boshqa yangi xizmatlardan foydalanilganda oxirgi hujjat (tahrirlash uchun) o'chib ketardi — tuzatildi
- Sahifalash: «keyingisi bilan birga» sarlavhalarda sahifa oxiridagi qator hisobi aniqlashtirildi; jadvallarda ustun kengliklari

## 14.1-versiya (21.11.2026) — kurs ishi ma'lumotlarini kiritish va 3 til
- 🐞 «Fakultet va Kafedra nomini yozing» deb so'ralib, kafedra keyin yana so'ralardi — endi avval faqat **fakultet**, kafedra keyinroq bir marta
- ✍️ Har bir qadamda **namuna**: OTM («Samarqand davlat universiteti»), fakultet («Filologiya»), fan («O'zbek tili praktikumi»),
  mavzu, familiya-ism («Toshmatova Malika»), guruh va rahbar, kafedra («O'zbek tilshunosligi»), shahar; kiritilgan qiymat
  tasdiq bilan qaytariladi («✅ Fan: O'zbek tili praktikumi»)
- 🔤 Yangi `matn_norm.py` — foydalanuvchi qanday yozsa ham titul to'g'ri chiqadi:
  - HAMMASI KATTA / hammasi kichik → odatiy yozuv (ism-familiya har so'z bosh harf bilan; nomlarda atoqli otlar — O'zbekiston,
    Toshkent, Mirzo Ulug'bek ... — bosh harf bilan; inglizchada Title Case)
  - kirillda yozilgan o'zbekcha matn o'zbekcha hujjatda lotinga o'giriladi; so'z ichida aralashib qolgan kirill harflari tuzatiladi
  - turli apostroflar (’ ‘ ʻ ʼ `) bir xil ko'rinishga keltiriladi
  - «Ona tili kafedrasi», «Filologiya fakulteti», «14-guruh», «Andijon shahri-2026» — titulda so'z ikki marta yozilmaydi
  - rahbar: «f.f.n., dots. karimov a.b.» → «f.f.n., dots. Karimov A.B.», «phd» → «PhD»
- 🌐 Rus va ingliz titulida tartib o'sha tilga mos: «Кафедра русского языка», «Филологический факультет»,
  «студент(ка) группы РЯ-21», «Department of English Philology», «student of group ENG-21»
- 🌐 Rus tilidagi kurs ishida kamida 25 ta rus/o'zbek tilidagi manba (12+ O'zbekiston mualliflari), inglizchada xalqaro manbalar
- 🌐 Yangi `tarjima_shim.py`: kodda qattiq yozilgan o'zbekcha xabarlar («💰 Qolgan tangalar», telefon/email/sana/fayl xatolari,
  esse tekshiruvi izohlari ...) rus va ingliz tilidagi foydalanuvchilarga tarjima qilib yuboriladi
- Barcha menyu tugmalari uch tilda qayta tekshirildi (muammo topilmadi)

## 14-versiya (21.11.2026) — o'qituvchi takliflari bo'yicha kurs ishi
| Taklif | Nima qilindi |
|---|---|
| O'zbek adabiyotlari kam — kamida 25 ta bo'lsin | Adabiyotlar: standart 30–36, kengaytirilgan 36–45 ta; **kamida 25 (kengaytirilganda 28) tasi o'zbek tilida**. Yetmasa bot qo'shimcha o'zbekcha manbalarni alohida so'raydi (darslik, monografiya, lug'at, dissertatsiya, jurnal maqolalari, qonun va qarorlar) |
| Bitta muallif (masalan A. Hojiyev) haddan tashqari ko'p takrorlangan | Ro'yxatda bir muallifning ko'pi bilan 2 ta asari; har paragrafga **turli manbalar** biriktiriladi; butun ish bo'ylab bitta manbaga ko'pi bilan 2 ta snoska |
| Snoskada bet ko'rsatilmagan | Snoskalar endi betli: «Hojiyev A. O'zbek tili sinonimlarining izohli lug'ati. – T.: O'qituvchi, 1974. – B. 45.» (rus manbalarida «С.», inglizchada «P.»); bet kitobning umumiy hajmidan oshmaydi; qonun va farmonlarga bet qo'yilmaydi |
| Bob sarlavhasidan keyin darhol 1.1 boshlanib ketyapti | Har bob boshida **bob kirish qismi** (3 abzas, ~350 so'z — bir betdan ko'p): bobning maqsadi, paragraflar mazmuni, yondashuv va oldingi bob bilan bog'liqligi; undan keyin 1.1, 1.2 ... |

- Hajm baribir kafolatlanadi (standart 27–30, kengaytirilgan 66–70 bet), mundarija raqamlari aniq qoladi (LibreOffice'da qayta tekshirildi)
- Adabiyotlar ro'yxati va snoskalarda «– 250 b.», «– B. 45.» qismlari qatorlar orasida uzilmaydi

## 13.3-versiya (19.11.2026) — kurs ishi titul varag'i OTM namunasi bo'yicha
- 📄 Titul OTM namunasi bo'yicha qayta tuzildi: vazirlik, universitet, chapda fakultet va o'ngda kafedra bilan
  «Ro'yxatga olindi №___» va sana qatorlari, «“…” KAFEDRASI», «“…” FANIDAN», KURS ISHI, MAVZU, Bajardi /
  fakultet / guruh talabasi / Tekshirdi, taqriz va himoya jadvali (sana, imzo, baho, komissiya a'zolari), pastda shahar – yil
- Talabaga bog'liq bo'lmagan qismlar (ro'yxatga olish raqami va sanasi, taqriz, himoya, baho, komissiya) bo'sh qoldiriladi
- Bot qo'shimcha so'raydi: **kafedra** (o'tkazib yuborsa bo'ladi) va **shahar** — OTM nomidan aniqlangan shahar birinchi
  turadi, 16 ta viloyat markazi tugmada yoki qo'lda yoziladi («Andijon-2026» yozilsa ham to'g'ri olinadi)
- Rus va ingliz tillarida ham shu tuzilma (КАФЕДРА «…», ПО ДИСЦИПЛИНЕ «…», Выполнил(а), Проверил(а), Члены комиссии ...)
- Uzun OTM yoki mavzu nomlarida titul bitta betga sig'ishi uchun oraliqlar avtomatik qisqaradi

## 13.2-versiya (19.11.2026) — kurs ishi: aniq mundarija va kafolatlangan hajm
- 📑 **Mundarija sahifa raqamlari endi to'g'ri chiqadi — Ctrl+A / F9 bosish shart emas.** Yangi `sahifalash.py`
  hujjatni Word'dagidek qatorlarga bo'lib, sahifalarga joylab hisoblaydi (Times New Roman bilan bir xil o'lchamli
  Liberation Serif shriftining harf kengliklari, 1,5 interval, sarlavhalarning keyingi abzas bilan birga turishi,
  yetim/beva qatorlar, snoskalar, jadvallar, rasmlar). LibreOffice'da 45 ta hujjatda sinaldi — barcha bo'lim
  sahifalari va umumiy bet soni mos keldi
- 📏 **Hajm kafolati:** tayyor ish sahifalab tekshiriladi; standart 27–30, kengaytirilgan 66–70 betdan kam bo'lsa —
  eng qisqa paragraflar avtomatik to'ldiriladi, ortiq bo'lsa — qisqartiriladi (avval AI qisqa yozsa 54 bet chiqardi)
- 🏛 Vazirlik nomi yangilandi: **O'zbekiston Respublikasi Ta'lim va ilm-fan vazirligi** / Министерство образования и
  науки Республики Узбекистан / Ministry of Education and Science of the Republic of Uzbekistan (kurs ishi, referat)
- Tayyor hujjat xabaridan «Ctrl+A, F9» va «tekshirib oling / taxminiy» eslatmalari, jadval va rasm soni olib tashlandi;
  aniq hajm ko'rsatiladi: «📄 Hajmi: 68 bet»
- Taqdimotda ishonchli ma'lumot bo'lmasa «taxminiy» diagramma o'rniga ro'yxat yoki bosqichlar slaydi qo'yiladi
- Snoskalar: alohida «Footnote Text» uslubi (10 pt, yakka interval), ajratuvchi chiziq ixcham
- Mundarijadagi uzun bob nomlari sahifa raqami ustuniga kirmasdan ko'chadi

## 13.1-versiya (19.11.2026) — avtomatik marketing postlari (`marketing.py`)
- 📣 **Haftada 2 marta** (dushanba va payshanba, 18:30) har bir foydalanuvchiga bitta xizmat haqida qisqa, sodda post:
  21 ta post (taqdimot, kurs ishi, konspekt, test markazi, masala, tekin tanga, imtihon, referat, audio, ...),
  har kim o'z navbatida oladi — ~10 haftada barcha xizmatlar bilan tanishadi, keyin aylanma takrorlanadi. 3 tilda
- 👉 Postdagi **«Sinab ko'rish»** tugmasi xizmatni darhol ochadi; narxlar `DOC_NARX` dan avtomatik olinadi
- 🔕 **«Bunday xabarlar kerak emas»** tugmasi va `/xabarlar` buyrug'i — foydalanuvchi o'zi o'chiradi/yoqadi
- 🆕 Yangi foydalanuvchilar avval 7 ta tanishuv xabarini oladi (1-xabar 24 soatdan keyin, keyingilari **har 2 kunda**);
  bu davrda haftalik postlar yuborilmaydi — xabarlar ustma-ust tushmaydi
- 🚫 Botni bloklaganlarga qayta yuborilmaydi; bot o'chib-yonsa ham bir slot takrorlanmaydi; server uzoq o'chiq bo'lsa tunda yubormaydi
- 👨‍💻 Admin: `/marketing` (statistika, tugma bosishlar), `/marketing test`, `/marketing yubor`, `/marketing on|off`;
  har yuborishdan keyin adminga hisobot
- 🐞 Tanishuv xabarlarida qator ko'chishi o'rniga so'zma-so'z «\n» chiqardi — tuzatildi
- 🐞 Kurs ishi tugmasi bosilganda oldingi xizmatning qadami qolib ketardi — tuzatildi

## 13-versiya (19.11.2026) — yangi kurs ishi va arzon narxlar
### 📄 Kurs ishi to'liq qayta yozildi (`kurs_ishi.py`)
| | 📘 Standart | 📚 Kengaytirilgan |
|---|---|---|
| Hajm | 25–30 bet | 65–70 bet |
| Tuzilma | Kirish, 2 bob × 3 paragraf | Kirish, 3 bob × 3 paragraf (nazariy, tahliliy, takomillashtirish) |
| Adabiyotlar | 15–20 | 25–35 |
| Narx | 12 🪙 | 25 🪙 (~20–25 ming so'm) |

- Bo'limga kirganda avval **hajm tanlanadi**, narx tugmada ko'rinadi; tanga yetmasa darhol aytiladi
- Titulga **guruh va ilmiy rahbar** qo'shildi (ixtiyoriy); shahar OTM nomidan aniqlanadi (Samarqand, Buxoro...)
- OTM talablari: A4, Times New Roman 14, 1,5 interval, hoshiyalar 3/1,5/2/2 sm, abzas 1,25 sm, titulda raqam yo'q
- **Mundarija** — Word'ning haqiqiy mundarija maydoni (Heading 1/2), sahifa raqamlari hisoblab qo'yiladi (LibreOffice bilan tekshirilgan farq ±1 bet) va Word ochilganda yangilanadi
- **Kirish**: dolzarblik, o'rganilganlik darajasi, maqsad, vazifalar ro'yxati, obyekt, predmet, metodlar, amaliy ahamiyat, tuzilish (bob, paragraf, jadval, rasm soni va bet hajmi avtomatik)
- **Snoskalar** — haqiqiy Word snoskalari, har biri adabiyotlar ro'yxatidagi manbaga ishora qiladi (ro'yxatda yo'q manbaga havola o'chiriladi)
- **Jadvallar** — «1-jadval», nomi, manbasi, sarlavha qatori har sahifada takrorlanadi; manbasiz raqamli jadval olib tashlanadi
- **Grafiklar** — faqat manbasi ko'rsatilgan raqamlar bilan (chiziq, ustun, doira), «1-rasm» va manba bilan
- **Xulosa va takliflar** (raqamlangan takliflar), **adabiyotlar** GOST tartibida: normativ hujjatlar → asosiy → xorijiy → internet
- Har paragraf alohida yoziladi (3 ta parallel oqim) — hajm kafolatlanadi: qisqa chiqsa avtomatik davom ettiriladi
- Eski kurs ishi tahrirlash tizimi yangi formatda ham ishlaydi (III bob paragraflari ham)
- Xarajat: 70 betlik ish DeepSeek'da ~0,05–0,1 $ (~1 000 so'm) — narxdan ko'p marta arzon

### 💰 Narxlar talabalar uchun arzonlashtirildi
| Paket | Avval | Endi |
|---|---|---|
| Kirish | 5 tanga — 9 900 | **10 tanga — 9 900** |
| Ommabop | 15 tanga — 24 900 | **25 tanga — 19 900** |
| O'rta | 35 tanga — 52 000 | **60 tanga — 39 900** |
| Katta | 70 tanga — 95 000 | **150 tanga — 89 900** |

- 1 tanga 1 980 so'mdan **990–599 so'm**ga tushdi (taqdimot ~4 000 so'm, kurs ishi ~9 600 so'm)
- Xizmatlar: taqdimot 6→5, referat 6→5, tezis 6→5, maqola 9→8, mustaqil ish 7→6, CV 9→6
- 🎁 Birinchi to'lovga +20% bonus tanga (avtomatik)
- 📅 Har dushanba balans 3 emas, **5** tangagacha to'ldiriladi — haftasiga 1 taqdimot tekin
- Obunalar: Lite 19 900 → 35 tanga (avval 25), Standart 34 900 → 75 tanga (avval 50)
- To'ldirish oynasi narxlarni sozlamalardan avtomatik chiqaradi (endi bir joyda o'zgartiriladi), so'mdagi narxi ham ko'rsatiladi
- Eslatma xabarida «ikkalangiz ham 10 tanga olasiz» deyilgan edi (aslida do'stga 5 tanga) — tuzatildi

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

"""
ilmiy_qollanma.py — ilmiy izlanuvchilar uchun qo'llanmalar (3 tilda)
═══════════════════════════════════════════════════════════════════
Manbalar (21.11.2026 holatiga tekshirildi):
  • oak.uz — «Ilmiy nashrlar ro'yxati» (03.08.2026), «Xalqaro nashriyotlar va ma'lumotlar bazalari ro'yxati
    optimallashmoqda» (10.06.2026), «Xorijiy jurnallar ro'yxatiga o'zgartirishlar» (13.06.2026),
    «Shubhali nashrlar to'g'risida» (19.02.2026), «Attestatsiya jarayonlarining yangi tizimga o'tishi» (20.11.2026)
  • Prezidentning 26.05.2026 dagi PF-98-son Farmoni; «Ilmiy darajalar berish tartibi to'g'risida»gi yangi nizom
    (OAK e'lon qilgan, 2027-yil 1-yanvardan)
  • VMning 22.05.2017 dagi 304-son qarori (oliy ta'limdan keyingi ta'lim; 2026-yilgacha o'zgartirishlar bilan)
  • Vazirlikning 31.01.2024 dagi 24-son buyrug'i (ilmiy etiket talablari)
Raqamlar o'zgarishi mumkin — har bir qo'llanma oxirida rasmiy manbani tekshirish eslatiladi.
"""

ROYXAT = ["daraja", "oak_maqola", "xalqaro_maqola", "shubhali", "qabul", "jurnal_tanlash"]

Q = {
    "uz": {
        "daraja": ("🧾 Ilmiy daraja: 2027-yildan yangi tizim", """🧾 <b>Ilmiy daraja olish: 2027-yil 1-yanvardan yangi tizim</b>
<i>Prezidentning 26.05.2026 dagi PF-98-son «Yangi avlod ilmiy kadrlarini tayyorlash va ularni attestatsiya qilish tizimini takomillashtirish to'g'risida»gi Farmoni</i>

<b>1. O'tish davri (OAK Rayosatining 29.10.2026 dagi 391/2-son qarori)</b>
• Amaldagi tartibda hujjat qabuli <b>13-noyabr</b> da tugadi.
• OAK saytidagi himoya e'lonlari <b>30-noyabrgacha</b> joylanadi.
• Bu himoyalar <b>31-dekabrgacha</b> o'tkaziladi.
• Ulgurmaganlar ishini davom ettiradi va <b>1-yanvar 2027</b> dan yangi tizimda himoya qiladi. Dissertatsiyani qayta yozish shart emas (OAK, 05.11.2026).

<b>2. Bekor qilinadi (Farmon)</b>
• Chet tili malaka imtihoni (mustaqil izlanuvchilar bundan mustasno)
• Mavzuni OAKda ro'yxatga olish va himoyani OAK saytida e'lon qilish
• Yetakchi tashkilotda muhokama
• Ijtimoiy-gumanitar fanlarda DSc uchun monografiya
• Avtoreferatni chop etish
• Natijalar amaliyotga joriy etilgani haqidagi xulosa
• OAK ekspert kengashi ekspertizasi va kengash qarorini OAKda tasdiqlash

<b>3. Yangi tartib</b>
• Ilmiy kengashlar daraja berish bo'yicha qarorni <b>o'zi</b> qabul qiladi va sifat uchun javob beradi.
• OAK kengashlar ma'qullagan ishlarni <b>tanlab</b> tekshiradi (xorijiy ekspertlar ham jalb etilishi mumkin) va asosli bo'lsa qarorni bekor qila oladi.
• Kengashlar <b>reytingi</b> OAKning yagona elektron portalida ochiq e'lon qilinadi; past reytingli kengashlar faoliyati cheklanishi mumkin.
• Yangi dastur — <b>«Loyiha doirasida ilmiy daraja»</b> (PhD): Innovatsion rivojlanish agentligi har yili tanlov o'tkazadi, 3 yilgacha loyiha, loyiha rahbari ilmiy rahbar bo'ladi.

<b>4. Yangi nizom bo'yicha talablar</b>
• <b>PhD:</b> kamida 8 ta ilmiy ish — kamida 3 ta milliy jurnal maqolasi, 1 ta xalqaro bazadagi jurnal maqolasi, xalqaro va respublika konferensiyalaridan 2 tadan material.
• <b>DSc:</b> kamida 16 ta ish — 8 ta milliy, 2 ta xalqaro bazadagi jurnal maqolasi, har turdagi konferensiyadan 3 tadan.
• Hajm: PhD — 120 betgacha, DSc — 200 betgacha (ijtimoiy-gumanitar fanlarda 30% gacha ko'p).
• Dissertatsiyasiz yo'l: PhD — 6 maqola (IGF — 4), DSc — 12 (8); yarmidan ko'pi oxirgi 3 yilda.
• Dastlabki ekspertiza: PhD — 2 oy, DSc — 3 oy.
• Ko'chirmachilik aniqlansa: ish himoyadan olinadi, shu mavzuda qayta himoya huquqi yo'qolishi mumkin.
{nizom}

<b>5. OAKning rasmiy tushuntirishlari (05.11.2026)</b>
• Farmonda hamma uchun Q1–Q2 jurnalida maqola majburiy emas — nashr talablari normativ hujjatlar bilan belgilanadi.
• Yangi kengashlar amaldagi kengashlar tarkibi asosida tuziladi.
• Avvalgi seminar xulosalari va taqrizlarni hisobga olish tartibi alohida tushuntiriladi.

📌 Rasmiy matn va o'tish qoidalari: <b>oak.uz</b>, t.me/attestatsiya. Call-markaz: (55) 514-11-43, ishonch telefoni: (95) 193-11-43."""),

        "oak_maqola": ("📤 OAK jurnaliga maqola yuborish tartibi", """📤 <b>OAK ro'yxatidagi jurnalga maqola yuborish — bosqichma-bosqich</b>

<b>1. Jurnalni to'g'ri tanlang</b>
• Botdagi «📰 OAK jurnallari» bo'limida jurnal nomini tekshiring: jurnal <b>sizning fan tarmog'ingiz kodi</b> bo'yicha (masalan, 13.00.00 — pedagogika) <b>«✅ faol»</b> bo'lishi kerak.
• Bir jurnal bir sohada faol, boshqasida chiqarilgan bo'lishi mumkin.
• Maqola chiqqan sanada jurnal ro'yxatda faol bo'lgani muhim. Shubha bo'lsa — OAKdan so'rang.

<b>2. Jurnal talablarini o'qing</b>
Jurnal saytidagi «Mualliflar uchun» bo'limida hajm, shablon, UDK, annotatsiya va kalit so'zlar (odatda o'zbek, rus, ingliz tillarida), adabiyotlar uslubi va to'lov haqidagi ma'lumot bo'ladi.

<b>3. Maqolani IMRAD tuzilmasida yozing</b>
Kirish → Adabiyotlar tahlili → Metodlar → Natijalar → Muhokama → Xulosa → Adabiyotlar.
💡 Botdagi «📝 Maqola (IMRAD)» xizmati qoralama tayyorlashga yordam beradi — lekin natijalar sizning tadqiqotingizdan bo'lishi shart.

<b>4. Ilmiy etiket talablari</b>
• Ko'chirmachilikka yo'l qo'ymang: o'zingiz antiplagiat tizimida tekshiring.
• Har bir iqtibosga manba ko'rsating.
• Vazirlikning 31.01.2024 dagi 24-son buyrug'i bilan ilmiy nashrlarga ilmiy etiket talablari belgilangan. Ularni buzgan jurnallar ro'yxatdan chiqarilmoqda (masalan, 10.06.2026 qarori).

<b>5. Yuborish</b>
Jurnal saytidagi tizim (OJS) yoki rasmiy e-pochta orqali yuboring. Qo'shimcha xat va muallif ma'lumotlarini (ORCID, ish joyi, lavozim) ilova qiling.

<b>6. Taqriz va nashr</b>
Odatda 2–8 hafta. Taqrizchi izohlariga bandma-band javob bering. To'lov faqat qabul qilingandan keyin va rasmiy rekvizitlarga.

<b>7. Nashrdan keyin</b>
Jurnal sonining PDF nusxasini, mundarija va muqovasini saqlab qo'ying — attestatsiya hujjatlari uchun kerak bo'ladi.

⚠️ «3 kunda chiqaramiz, OAKga 100% o'tadi» degan vositachilardan ehtiyot bo'ling — avval jurnalni ro'yxatdan tekshiring."""),

        "xalqaro_maqola": ("🌍 Scopus / Web of Science jurnaliga maqola", """🌍 <b>Scopus yoki Web of Science jurnaliga maqola chiqarish</b>

<b>1. OAK qaysi xalqaro bazalarni tan oladi (10.06.2026 dan)</b>
• <b>A toifa:</b> Web of Science, Scopus, Journal Citation Reports
• <b>B toifa:</b> AGRIS, Chemical Abstracts (CAS), GeoRef, PubMed, Compendex
• <b>C toifa (ijtimoiy-gumanitar):</b> Anthropological Index Online, Anthropological Literature, Anthropology Plus, Bibliography of Asian Studies
Springer, CrossRef (DOI), DOAJ, Index Copernicus, ResearchGate, SCImago, ResearchBib va boshqalar endi <b>mustaqil baholash mezoni emas</b>. Maqolaga DOI berilgani yoki uning taniqli nashriyotda chiqqani o'z-o'zidan yetarli emas — jurnal A/B/C bazalarida indekslanishi kerak.

<b>2. Jurnal tanlash</b>
• Botdagi «🧭 Mos jurnal topish» xizmati sizning mavzuingizga o'xshash haqiqiy maqolalar chop etilgan jurnallarni topadi.
• Rasmiy vositalar: Elsevier JournalFinder, Springer Nature Journal Suggester, Wiley Journal Finder, Clarivate Manuscript Matcher.

<b>3. Jurnalni albatta tekshiring</b>
• Web of Science: mjl.clarivate.com (Master Journal List, bepul) — ISSN bo'yicha qidiring.
• Scopus: scopus.com/sources (Source list) — jurnal <b>hozir</b> indekslanayaptimi («coverage»), qaysi yillar.
• Scopus'dan chiqarilgan jurnallar ro'yxati (Discontinued titles) — OAK 19.02.2026 da shunday ro'yxatni e'lon qilgan.
• Sayt manzilini ISSN va bazadagi havola bilan solishtiring — <b>o'g'irlangan (hijacked) jurnal</b> saytlari bor.

<b>4. Maqolani tayyorlash</b>
• «Guide for Authors» ni to'liq bajaring.
• Ingliz tilini professional darajada tahrir qiling.
• Adabiyotlar asosan oxirgi 5 yildagi xalqaro manbalardan bo'lsin.
• Cover letter yozing, ORCID oching.

<b>5. Yuborish va taqriz</b>
• Jurnalning Editorial Manager / ScholarOne tizimi orqali yuboring.
• Taqriz odatda 1–6 oy davom etadi; rad etilsa — boshqa jurnalga moslab qayta yuboring.
• Bitta maqolani bir vaqtda ikki jurnalga yuborish taqiqlanadi.

<b>6. To'lov</b>
• Obuna asosidagi (subscription) jurnallarda odatda to'lov yo'q.
• Ochiq kirishli (Gold OA) jurnallarda APC to'lanadi — summasi jurnal saytida rasmiy ko'rsatiladi.
• Taqrizdan oldin to'lov so'ralsa — bu xavf belgisi.

📌 Vaqtni rejalang: xalqaro maqola uchun yuborishdan nashrgacha odatda 3–12 oy ketadi."""),

        "shubhali": ("⚠️ Shubhali va o'g'irlangan jurnallar", """⚠️ <b>Shubhali (yirtqich) va o'g'irlangan jurnallarni qanday aniqlash mumkin</b>

<b>Xavf belgilari</b>
• «Maqolangiz 3–7 kunda chiqadi», «100% qabul» kabi va'dalar.
• Taqrizdan oldin yoki taqrizsiz to'lov so'rash.
• Global Impact Factor, SJIF, Universal IF kabi soxta «impakt-faktor»lar (OAK ularni 2024-yildayoq ro'yxatdan chiqargan).
• Bitta jurnal «hamma fanlar bo'yicha» maqola qabul qiladi.
• Tahrir hay'ati a'zolari haqiqiy emas yoki ularni topib bo'lmaydi.
• Spam xatlar va Telegram orqali «Scopus'ga maqola joylash» takliflari.
• Sayt manzili bazadagi rasmiy havola bilan mos kelmaydi (o'g'irlangan jurnal).

<b>Tekshirish tartibi</b>
1. OAK ro'yxati: botdagi «📰 OAK jurnallari → 🔎 Nomi bo'yicha» — jurnal sizning fan kodingiz bo'yicha faolmi?
2. Web of Science: mjl.clarivate.com; Scopus: scopus.com/sources — hozir ham indekslanayaptimi?
3. Scopus'dan chiqarilgan jurnallar (Discontinued titles) — OAK 19.02.2026 da ehtiyotkorlikka chaqirgan.
4. Retraction Watch Hijacked Journal Checker — o'g'irlangan jurnallar ro'yxati.
5. Jurnal saytidagi ISSN ni portal.issn.org da tekshiring.

<b>Rasmiy misollar</b>
• OAK 10.06.2026 da «Development of science» jurnalini texnika va kimyo fanlari bo'yicha ro'yxatdan chiqardi — sabab: ilmiy etiket talablari buzilgan.
• Xuddi shu kuni o'nlab xorijiy jurnallar past indekslanish va faoliyati to'xtagani sababli chiqarildi.

📌 Shubha bo'lsa — maqolani yubormang, avval ilmiy rahbaringiz va OAK bilan maslahatlashing."""),

        "qabul": ("🎓 Doktoranturaga qabul: hujjatlar va imtihonlar", """🎓 <b>Tayanch doktorantura (PhD) va doktoranturaga (DSc) qabul</b>
<i>VMning 22.05.2017 dagi 304-son qarori asosida (2026-yilgacha kiritilgan o'zgarishlar bilan)</i>

<b>Shakllar</b>
• Tayanch doktorantura — PhD uchun.
• Doktorantura — DSc uchun.
• Mustaqil izlanuvchilik — ishdan ajralmagan holda, PhD yoki DSc.
• Stajyor-tadqiqotchilik — tayyorgarlik bosqichi.

<b>Kimlar topshira oladi</b>
• Tayanch doktorantura: magistratura (tibbiyotda — klinik ordinatura) yoki mutaxassislik bo'yicha oliy ma'lumoti borlar.
• Doktorantura: fan nomzodi yoki PhD darajasi borlar.

<b>Hujjatlar</b>
• Ariza, ma'lumotnoma (obyektivka), diplom nusxasi, chop etilgan ishlar ro'yxati va nusxalari.
• DSc uchun qo'shimcha — dissertatsiya tadqiqotining batafsil rejasi.
• Mehnat va diplom ma'lumotlari «Raqamli hukumat» orqali olinadi.

<b>Imtihonlar</b>
• Mutaxassislik fani va chet tili.
• Xalqaro sertifikat (IELTS, TOEFL, CEFR) chet tili imtihonidan ozod qiladi.
• Kollegial organ bilan suhbat.
• Tayanch doktoranturada birinchi fandan umumiy ballning 50% idan kam olganlar keyingi bosqichga o'tmaydi.

<b>Muddatlar</b>
• Kvota bo'yicha takliflar — iyunda beriladi.
• Qabul — 5-dekabrgacha.
• O'qish qabuldan keyingi yilning 1-yanvaridan boshlanadi.
• Aniq sanalar har bir muassasaning e'lonida bo'ladi.

<b>Qabuldan keyin</b>
• Yakka tartibdagi (individual) reja ilmiy rahbar bilan tuziladi va <b>1 oy ichida</b> Kollegial organ qarori bilan tasdiqlanadi.
• Botdagi «📑 Individual reja» xizmati namunaviy rejani tayyorlab beradi.

<b>Tayyorgarlik bo'yicha maslahatlar</b>
• Mavzuingiz ixtisoslik pasportiga mos ekanini tekshiring («🎯 Mavzu va ixtisoslik»).
• Chet tili sertifikatini oldindan oling.
• 1–2 ta maqolangiz va tadqiqot rejangiz tayyor bo'lsin.
• Qabul e'lonlarini muassasa va vazirlik saytlarida kuzating.

📌 Qabul tartibi har yili yangilanishi mumkin — muassasa e'loni va lex.uz dagi amaldagi tahrirni tekshiring."""),

        "jurnal_tanlash": ("🧭 Jurnal tanlash strategiyasi", """🧭 <b>Dissertatsiya uchun maqolalarni qayerda chiqarish kerak — strategiya</b>

<b>1. Avval talablarni bilib oling</b>
Yangi tizimda PhD uchun kamida 3 ta milliy va 1 ta xalqaro bazadagi jurnal maqolasi, DSc uchun 8 va 2 ta kerak — «🧾 Ilmiy daraja» qo'llanmasiga qarang.

<b>2. Milliy maqolalar</b>
• Fan tarmog'ingiz kodi bo'yicha faol jurnallardan 3–4 tasini tanlang. Bot ro'yxatni OAK saytidan avtomatik yangilab boradi.
• Turli jurnallarda chiqaring — bir jurnalda ko'p maqola ekspertlar e'tiborini tortadi.

<b>3. Xalqaro maqola</b>
• Faqat A/B/C toifali bazalar hisobga olinadi: Scopus, Web of Science, PubMed va boshqalar.
• Oldindan 6–12 oy vaqt qoldiring.
• MDH va xorijiy jurnallarning OAK ro'yxatida faol ekanini ham tekshiring.

<b>4. Konferensiyalar</b>
Xalqaro va respublika konferensiyalarida tezislar bilan qatnashing. Botdagi «📋 Tezis» xizmati tezis tayyorlashga yordam beradi.

<b>5. Kalendar</b>
• 1-yil — adabiyotlar tahlili maqolasi.
• 2-yil — metodika va natijalar.
• 3-yil — xalqaro maqola va yakuniy natijalar.

<b>6. Kuzatib boring</b>
«🔔 Obuna» tugmasini bosing — OAK ro'yxatida o'zgarish bo'lsa, bot darhol xabar beradi."""),
    },
    "ru": {
        "daraja": ("🧾 Учёная степень: новая система с 2027", """🧾 <b>Учёные степени: новая система с 1 января 2027 года</b>
<i>Указ Президента № УП-98 от 26.05.2026 «О совершенствовании системы подготовки и аттестации научных кадров нового поколения»</i>

<b>1. Переходный период (решение Президиума ВАК № 391/2 от 29.10.2026)</b>
• Приём документов по действующему порядку завершён <b>13 ноября</b>.
• Объявления о защитах публикуются на oak.uz до <b>30 ноября</b>.
• Защиты проводятся до <b>31 декабря</b>.
• Остальные защищаются по новой системе с <b>1 января 2027</b>. Переписывать диссертацию не нужно (ВАК, 05.11.2026).

<b>2. Отменяется (Указ)</b>
• Квалификационный экзамен по иностранному языку (кроме соискателей)
• Регистрация темы в ВАК и объявление о защите на сайте ВАК
• Обсуждение в ведущей организации
• Монография для DSc в СГН
• Печать автореферата
• Акт о внедрении
• Экспертиза экспертного совета ВАК и утверждение решения совета в ВАК

<b>3. Новый порядок</b>
• Научные советы сами присуждают степени и отвечают за качество.
• ВАК <b>выборочно</b> проверяет работы (в том числе с зарубежными экспертами) и может отменить решение.
• <b>Рейтинг</b> советов публикуется на едином портале ВАК.
• Программа <b>«Учёная степень в рамках проекта»</b> (PhD): конкурс Агентства инновационного развития, проекты до 3 лет.

<b>4. Требования нового положения</b>
• <b>PhD:</b> не менее 8 работ — 3 статьи в национальных журналах, 1 в журнале международной базы, по 2 материала международных и республиканских конференций.
• <b>DSc:</b> не менее 16 — 8 национальных, 2 международных, по 3 материала конференций.
• Объём: PhD до 120 стр., DSc до 200 (СГН +30%).
• Без диссертации: PhD — 6 статей (СГН 4), DSc — 12 (8).
• Предварительная экспертиза: PhD 2 мес., DSc 3 мес.
{nizom}

<b>5. Разъяснения ВАК (05.11.2026)</b>
• Указ не вводит обязательных статей в Q1–Q2 для всех.
• Новые советы формируются на базе действующих.

📌 Официальные тексты: <b>oak.uz</b>, t.me/attestatsiya. Call-центр: (55) 514-11-43."""),
        "oak_maqola": ("📤 Как подать статью в журнал ВАК", """📤 <b>Статья в журнал из перечня ВАК — по шагам</b>

<b>1. Выберите журнал.</b> В разделе «📰 Журналы ВАК» проверьте, что журнал <b>«✅ активен» по вашему коду отрасли</b> (например, 13.00.00). Важно, чтобы журнал был в перечне на дату выхода статьи.

<b>2. Изучите требования журнала:</b> объём, шаблон, УДК, аннотация и ключевые слова (обычно на узбекском, русском и английском), стиль списка литературы, оплата.

<b>3. Структура IMRAD:</b> введение → обзор литературы → методы → результаты → обсуждение → выводы.

<b>4. Научная этика:</b> проверьте текст на плагиат; требования утверждены приказом министерства № 24 от 31.01.2024. Журналы-нарушители исключаются из перечня.

<b>5. Подача:</b> через систему OJS или официальную почту журнала, с ORCID и сведениями об авторе.

<b>6. Рецензия:</b> обычно 2–8 недель. Оплата — только после принятия и по официальным реквизитам.

<b>7. После публикации:</b> сохраните PDF номера, обложку и содержание — они нужны для аттестационного дела.

⚠️ Остерегайтесь посредников «публикация за 3 дня, 100% ВАК»."""),
        "xalqaro_maqola": ("🌍 Статья в Scopus / Web of Science", """🌍 <b>Публикация в журналах Scopus / Web of Science</b>

<b>1. Какие базы признаёт ВАК (с 10.06.2026)</b>
• <b>Категория A:</b> Web of Science, Scopus, Journal Citation Reports
• <b>Категория B:</b> AGRIS, CAS, GeoRef, PubMed, Compendex
• <b>Категория C (СГН):</b> Anthropological Index Online, Anthropological Literature, Anthropology Plus, Bibliography of Asian Studies
Springer, CrossRef (DOI), DOAJ, Index Copernicus, ResearchGate, SCImago и др. больше <b>не являются самостоятельным критерием</b>.

<b>2. Подбор журнала:</b> «🧭 Подбор журнала» в боте, Elsevier JournalFinder, Springer Journal Suggester, Clarivate Manuscript Matcher.

<b>3. Проверка:</b> mjl.clarivate.com (WoS), scopus.com/sources (Scopus, текущая индексация), список Discontinued titles. Сверьте сайт с официальной ссылкой — существуют <b>журналы-клоны</b>.

<b>4. Подготовка:</b> Guide for Authors, качественный английский, свежая международная литература, cover letter, ORCID.

<b>5. Подача и рецензия:</b> Editorial Manager / ScholarOne; 1–6 месяцев. Нельзя подавать одну статью в два журнала одновременно.

<b>6. Оплата:</b> в подписных журналах обычно бесплатно; в Gold OA — APC по прайсу на сайте. Оплата до рецензии — тревожный знак.

📌 Закладывайте 3–12 месяцев от подачи до публикации."""),
        "shubhali": ("⚠️ Хищнические и клонированные журналы", """⚠️ <b>Как распознать хищнический или клонированный журнал</b>

<b>Признаки:</b>
• «Публикация за 3–7 дней», «100% принятие».
• Оплата до рецензии.
• Фальшивые «импакт-факторы» (Global IF, SJIF, Universal IF — исключены ВАК ещё в 2024 г.).
• Журнал «по всем наукам».
• Несуществующая редколлегия.
• Спам и предложения «разместить в Scopus».
• Адрес сайта не совпадает со ссылкой в базе.

<b>Проверка:</b>
1. Перечень ВАК в боте — активен ли журнал по вашему коду.
2. mjl.clarivate.com и scopus.com/sources.
3. Список исключённых из Scopus журналов (ВАК предупреждал 19.02.2026).
4. Retraction Watch Hijacked Journal Checker.
5. ISSN на portal.issn.org.

📌 При сомнении — не отправляйте статью, посоветуйтесь с руководителем и ВАК."""),
        "qabul": ("🎓 Поступление в докторантуру", """🎓 <b>Поступление в базовую докторантуру (PhD) и докторантуру (DSc)</b>
<i>Постановление КМ № 304 от 22.05.2017 (с изменениями до 2026 г.)</i>

<b>Формы:</b> базовая докторантура (PhD), докторантура (DSc), самостоятельное соискательство, стажёр-исследователь (подготовительный этап).

<b>Кто может поступать:</b> PhD — с магистратурой (в медицине — клиническая ординатура) или специалитетом; DSc — с учёной степенью кандидата наук / PhD.

<b>Документы:</b> заявление, объективка, копия диплома, список и копии публикаций; для DSc — развёрнутый план диссертационного исследования.

<b>Экзамены:</b> специальность и иностранный язык (сертификаты IELTS/TOEFL/CEFR освобождают от экзамена по языку), собеседование. В базовой докторантуре менее 50% по первому предмету — не проходят дальше.

<b>Сроки:</b> приём до 5 декабря, обучение с 1 января следующего года (точные даты — в объявлении вуза/НИИ).

<b>После зачисления:</b> индивидуальный план утверждается Коллегиальным органом в течение 1 месяца — раздел «📑 Индивидуальный план».

📌 Проверяйте объявления организаций и актуальную редакцию на lex.uz."""),
        "jurnal_tanlash": ("🧭 Стратегия выбора журналов", """🧭 <b>Где публиковаться для диссертации — стратегия</b>

<b>1.</b> Узнайте требования для вашей степени («🧾 Учёная степень»).

<b>2. Национальные статьи:</b> 3–4 разных активных журнала по вашему коду (бот автоматически обновляет перечень с сайта ВАК).

<b>3. Международная статья:</b> только базы категорий A/B/C; заложите 6–12 месяцев.

<b>4. Конференции:</b> международные и республиканские — раздел «📋 Тезис».

<b>5. Календарь:</b> 1-й год — обзорная статья, 2-й — методика и результаты, 3-й — международная статья.

<b>6.</b> Включите «🔔 Подписку» — бот сообщит об изменениях в перечне ВАК."""),
    },
    "en": {
        "daraja": ("🧾 Academic degrees: new system from 2027", """🧾 <b>Academic degrees: the new system from 1 January 2027</b>
<i>Presidential Decree PF-98 of 26.05.2026 on improving the training and attestation of new-generation research staff</i>

<b>1. Transition (HAC Presidium decision No. 391/2 of 29.10.2026)</b>
• Submission of documents under the old rules closed on <b>13 November</b>.
• Defense announcements are posted on oak.uz until <b>30 November</b>.
• These defenses take place by <b>31 December</b>.
• Everyone else defends under the new system from <b>1 January 2027</b>. There is no need to rewrite the dissertation (HAC, 05.11.2026).

<b>2. Abolished (Decree)</b>
• The foreign-language qualifying exam (except for independent researchers)
• HAC topic registration and HAC defense announcements
• The leading-organisation review
• The DSc monograph in social sciences and humanities
• Printing the avtoreferat
• The implementation certificate
• The HAC expert-council review and HAC approval of council decisions

<b>3. New procedure</b>
• Academic councils award degrees themselves and are responsible for quality.
• The HAC checks a <b>selection</b> of approved works (including with foreign experts) and may annul decisions.
• Council <b>ratings</b> are published on the HAC's single portal.
• New <b>«Degree within a project»</b> PhD programme: an annual Innovation Development Agency competition, projects of up to 3 years.

<b>4. Requirements of the new regulation</b>
• <b>PhD:</b> at least 8 works — 3 national journal articles, 1 international-database journal article, 2 papers each from international and national conferences.
• <b>DSc:</b> at least 16 — 8 national, 2 international, 3 conference papers of each type.
• Length: PhD up to 120 pages, DSc up to 200 (+30% in social sciences and humanities).
• Without a dissertation: PhD — 6 articles (4 in social sciences and humanities), DSc — 12 (8).
• Preliminary review: PhD 2 months, DSc 3 months.
{nizom}

<b>5. HAC clarifications (05.11.2026)</b>
• The Decree does not make Q1–Q2 articles mandatory for everyone.
• New councils are formed from the existing ones.

📌 Official texts: <b>oak.uz</b>, t.me/attestatsiya. Call centre: (55) 514-11-43."""),
        "oak_maqola": ("📤 Submitting to an HAC-listed journal", """📤 <b>Publishing in an HAC (OAK) listed journal — step by step</b>

<b>1. Choose the journal.</b> In «📰 HAC journals», check that it is <b>«✅ active» for your field code</b> (e.g. 13.00.00). The journal should be on the list on the publication date.

<b>2. Read the author guidelines:</b> length, template, UDC, abstract and keywords (often in Uzbek, Russian and English), reference style, fees.

<b>3. Use IMRAD:</b> introduction → literature review → methods → results → discussion → conclusion.

<b>4. Research ethics:</b> check for plagiarism. Ethics requirements were set by Ministry order No. 24 of 31.01.2024, and violators are removed from the list.

<b>5. Submit</b> through the journal's OJS site or official e-mail, with your ORCID.

<b>6. Review:</b> usually 2–8 weeks. Pay only after acceptance and only to official details.

<b>7. After publication:</b> keep the PDF of the issue, its cover and contents page for your attestation file.

⚠️ Beware of brokers promising «published in 3 days, 100% HAC»."""),
        "xalqaro_maqola": ("🌍 Publishing in Scopus / Web of Science", """🌍 <b>Publishing in Scopus / Web of Science journals</b>

<b>1. Databases recognised by the HAC (since 10.06.2026)</b>
• <b>Category A:</b> Web of Science, Scopus, Journal Citation Reports
• <b>Category B:</b> AGRIS, CAS, GeoRef, PubMed, Compendex
• <b>Category C (social sciences and humanities):</b> Anthropological Index Online, Anthropological Literature, Anthropology Plus, Bibliography of Asian Studies
Springer, CrossRef (DOI), DOAJ, Index Copernicus, ResearchGate, SCImago and others are <b>no longer independent criteria</b>.

<b>2. Find a journal:</b> «🧭 Journal finder» in the bot, Elsevier JournalFinder, Springer Nature Journal Suggester, Clarivate Manuscript Matcher.

<b>3. Verify it:</b> mjl.clarivate.com (WoS), scopus.com/sources (current coverage), the Scopus discontinued titles list. Check that the website matches the database link — <b>hijacked journal</b> sites exist.

<b>4. Prepare:</b> follow the Guide for Authors, get professional English editing, cite recent international literature, write a cover letter, register an ORCID.

<b>5. Submit and review:</b> Editorial Manager or ScholarOne; review takes 1–6 months. Never submit the same paper to two journals at once.

<b>6. Fees:</b> subscription journals are usually free; Gold OA journals charge an APC listed on their site. A request to pay before review is a red flag.

📌 Plan 3–12 months from submission to publication."""),
        "shubhali": ("⚠️ Predatory and hijacked journals", """⚠️ <b>How to spot predatory and hijacked journals</b>

<b>Red flags:</b>
• «Published in 3–7 days», «100% acceptance».
• Payment before review.
• Fake «impact factors» (Global IF, SJIF, Universal IF — removed by the HAC in 2024).
• «All subjects» scope.
• An editorial board you cannot verify.
• Spam offers to «place your paper in Scopus».
• A website that does not match the database link.

<b>How to check:</b>
1. The HAC list in the bot — is it active for your field code?
2. mjl.clarivate.com and scopus.com/sources.
3. The Scopus discontinued titles list (the HAC warned about it on 19.02.2026).
4. Retraction Watch Hijacked Journal Checker.
5. The ISSN on portal.issn.org.

📌 If in doubt, do not submit — consult your supervisor and the HAC."""),
        "qabul": ("🎓 Doctoral admission", """🎓 <b>Admission to basic doctoral studies (PhD) and doctoral studies (DSc)</b>
<i>Cabinet Resolution No. 304 of 22.05.2017 (as amended up to 2026)</i>

<b>Forms:</b> basic doctorate (PhD), doctorate (DSc), independent researcher, research trainee (preparatory stage).

<b>Eligibility:</b> PhD — a master's degree (clinical residency in medicine) or a specialist degree; DSc — a candidate of sciences or PhD degree.

<b>Documents:</b> application, CV (obyektivka), diploma copy, list and copies of publications; for DSc also a detailed research plan.

<b>Exams:</b> the specialty subject and a foreign language (IELTS, TOEFL or CEFR certificates exempt you from the language exam), plus an interview. In basic doctoral studies, scoring below 50% on the first subject ends the application.

<b>Timeline:</b> admission by 5 December; studies start on 1 January of the next year (exact dates are in each institution's announcement).

<b>After enrolment:</b> the individual plan is approved by the Collegial body within 1 month — see «📑 Individual plan».

📌 Check institution announcements and the current text on lex.uz."""),
        "jurnal_tanlash": ("🧭 Journal strategy", """🧭 <b>Where to publish for your dissertation — strategy</b>

<b>1.</b> Learn the requirements for your degree («🧾 Academic degrees»).

<b>2. National articles:</b> 3–4 different journals that are active for your code (the bot updates the HAC list automatically).

<b>3. International article:</b> only category A/B/C databases count; allow 6–12 months.

<b>4. Conferences:</b> international and national — see «📋 Thesis».

<b>5. Timeline:</b> year 1 — review article, year 2 — methods and results, year 3 — international article.

<b>6.</b> Turn on «🔔 Alerts» to get notified when the HAC list changes."""),
    },
}


NIZOM_REKVIZIT = ""        # admin /nizom buyrug'i bilan o'rnatadi (tasdiqlangan hujjat nomi, raqami, sanasi)
_NIZOM_LBL = {"uz": "📄 Hujjat: {r}", "ru": "📄 Документ: {r}", "en": "📄 Document: {r}"}


def get(til: str, key: str) -> tuple[str, str]:
    sar, matn = Q.get(til, Q["uz"]).get(key) or Q["uz"][key]
    if "{nizom}" in matn:
        matn = matn.replace("{nizom}", _NIZOM_LBL.get(til, _NIZOM_LBL["uz"]).format(r=NIZOM_REKVIZIT)
                            if NIZOM_REKVIZIT else "")
    return sar, matn

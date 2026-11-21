"""
oak_jurnallar.py — OAK «Dissertatsiyalar asosiy ilmiy natijalarini chop etish tavsiya etilgan ilmiy nashrlar
ro'yxati»ni rasmiy manbadan avtomatik yuklash, qidirish va o'zgarishlarni kuzatish
═══════════════════════════════════════════════════════════════════
Manba: oak.uz → «Ilmiy nashrlar ro'yxati» (oak.uz/material/31738) sahifasiga joylashtirilgan rasmiy ro'yxat
ilovasi — https://journal-statistics-oak.vercel.app/ . Sahifa HTMLida ro'yxat to'liq JSON ko'rinishida
(Next.js RSC «self.__next_f.push» bloklari) keladi:
    {"payload": {"total", "columns": [...], "dicts": {ustun: [qiymatlar]}, "rows": [[indekslar]]},
     "publishers": [...A/B/C toifali xalqaro bazalar...], "disciplines": [{name, code}]}
Har qator = jurnal × fan tarmog'i. Bu modul qatorlarni jurnal bo'yicha guruhlaydi.

Yangilanish: bot ishga tushganda va har 6 soatda tekshiriladi; ro'yxat o'zgarsa (xesh boshqa) — qo'shilgan va
chiqarilgan jurnallar aniqlanib, tarixga yoziladi va obunachilarga xabar yuboriladi (ilmiy_izlanuvchi.py).
Tarmoq ishlamasa — oxirgi saqlangan nusxa (oak_cache.json) ishlatiladi.
"""
from __future__ import annotations

import datetime
import hashlib
import json
import logging
import os
import re
import threading
import time
import urllib.request

log = logging.getLogger(__name__)

MANBA_URL = "https://journal-statistics-oak.vercel.app/"
OAK_SAHIFA = "https://oak.uz/ilmiy-nashrlar-royxati"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "oak_cache.json")
UA = "Mozilla/5.0 (compatible; AcademicBot/1.0; +https://t.me) Python-urllib"

FAOL = {"Актив", "Қўшилди", "Ўзгартирилди"}
CHIQ = {"Чиқарилди", "Олдин чиқарилган"}
TUR = {"Миллий": "milliy", "МДҲ": "mdh", "Хорижий": "xorijiy"}

# fan tarmoqlari (OAK kodlari) — manbadagi «disciplines» bilan yangilanadi
FANLAR = {
    "01.00.00": "Fizika-matematika fanlari", "02.00.00": "Kimyo fanlari", "03.00.00": "Biologiya fanlari",
    "04.00.00": "Geologiya-mineralogiya fanlari", "05.00.00": "Texnika fanlari",
    "06.00.00": "Qishloq xo'jaligi fanlari", "07.00.00": "Tarix fanlari", "08.00.00": "Iqtisodiyot fanlari",
    "09.00.00": "Falsafa fanlari", "10.00.00": "Filologiya fanlari", "11.00.00": "Geografiya fanlari",
    "12.00.00": "Yuridik fanlar", "13.00.00": "Pedagogika fanlari", "14.00.00": "Tibbiyot fanlari",
    "15.00.00": "Farmatsevtika fanlari", "16.00.00": "Veterinariya fanlari", "17.00.00": "San'atshunoslik fanlari",
    "18.00.00": "Arxitektura", "19.00.00": "Psixologiya fanlari", "21.00.00": "Harbiy fanlar",
    "22.00.00": "Sotsiologiya fanlari", "23.00.00": "Siyosiy fanlar", "24.00.00": "Islomshunoslik fanlari",
}
FANLAR_RU = {
    "01.00.00": "Физико-математические науки", "02.00.00": "Химические науки", "03.00.00": "Биологические науки",
    "04.00.00": "Геолого-минералогические науки", "05.00.00": "Технические науки",
    "06.00.00": "Сельскохозяйственные науки", "07.00.00": "Исторические науки", "08.00.00": "Экономические науки",
    "09.00.00": "Философские науки", "10.00.00": "Филологические науки", "11.00.00": "Географические науки",
    "12.00.00": "Юридические науки", "13.00.00": "Педагогические науки", "14.00.00": "Медицинские науки",
    "15.00.00": "Фармацевтические науки", "16.00.00": "Ветеринарные науки", "17.00.00": "Искусствоведение",
    "18.00.00": "Архитектура", "19.00.00": "Психологические науки", "21.00.00": "Военные науки",
    "22.00.00": "Социологические науки", "23.00.00": "Политические науки", "24.00.00": "Исламоведение",
}
FANLAR_EN = {
    "01.00.00": "Physics and Mathematics", "02.00.00": "Chemistry", "03.00.00": "Biology",
    "04.00.00": "Geology and Mineralogy", "05.00.00": "Engineering", "06.00.00": "Agriculture",
    "07.00.00": "History", "08.00.00": "Economics", "09.00.00": "Philosophy", "10.00.00": "Philology",
    "11.00.00": "Geography", "12.00.00": "Law", "13.00.00": "Pedagogy (Education)", "14.00.00": "Medicine",
    "15.00.00": "Pharmacy", "16.00.00": "Veterinary science", "17.00.00": "Art studies", "18.00.00": "Architecture",
    "19.00.00": "Psychology", "21.00.00": "Military science", "22.00.00": "Sociology", "23.00.00": "Political science",
    "24.00.00": "Islamic studies",
}


def fan_nomi(kod: str, til: str = "uz") -> str:
    return {"ru": FANLAR_RU, "en": FANLAR_EN}.get(til, FANLAR).get(kod) or FANLAR.get(kod, kod)


# ═══════════════════════════════════════════════════════════════════
# 1. YUKLASH VA TAHLIL
# ═══════════════════════════════════════════════════════════════════
def yukla(url: str = MANBA_URL, timeout: int = 40) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="replace")


_PUSH = re.compile(r"self\.__next_f\.push\((\[.*?\])\)</script>", re.S)


def _payload_top(html: str) -> dict:
    """HTML dagi RSC bloklaridan {"payload":..., "publishers":..., "disciplines":...} obyektini topadi."""
    parts = []
    for m in _PUSH.finditer(html):
        try:
            v = json.loads(m.group(1))
        except Exception:
            continue
        if isinstance(v, list) and len(v) > 1 and isinstance(v[1], str):
            parts.append(v[1])
    blob = "".join(parts) or html
    i = blob.find('{"payload":')
    if i < 0:
        raise ValueError("OAK ro'yxati topilmadi (sahifa tuzilishi o'zgargan bo'lishi mumkin)")
    obj, _ = json.JSONDecoder().raw_decode(blob, i)
    if not isinstance(obj, dict) or "payload" not in obj:
        raise ValueError("OAK ro'yxati formati noto'g'ri")
    return obj


def _lat(t: str) -> str:
    try:
        import matn_norm
        return matn_norm.kiril_lotin(t)
    except Exception:
        return t


def _norm(t: str) -> str:
    """Qidiruv uchun: kirill → lotin, kichik harf, faqat harf va raqam."""
    t = _lat(str(t or "")).lower()
    t = t.replace("ʻ", "'").replace("‘", "'").replace("’", "'").replace("`", "'")
    t = re.sub(r"[^0-9a-zа-яёўқғҳ' ]+", " ", t)
    t = t.replace("'", "")
    return re.sub(r"\s+", " ", t).strip()


def tahlil(html: str) -> dict:
    top = _payload_top(html)
    p = top["payload"]
    cols, D, rows = p["columns"], p["dicts"], p["rows"]
    ci = {c: i for i, c in enumerate(cols)}

    def val(r, c):
        if c not in ci:
            return ""
        x = r[ci[c]]
        if c == "year":
            return x
        try:
            return D[c][x]
        except (KeyError, IndexError, TypeError):
            return ""

    jur: dict = {}
    for r in rows:
        status = val(r, "status")
        if status not in FAOL and status not in CHIQ:
            continue                                   # «Изоҳ» kabi izoh qatorlari
        nom = str(val(r, "name") or "").strip()
        nom_toza = re.sub(r"^\s*\[\s*Чиқарилган\s*\]\s*", "", nom).strip()
        if not nom_toza:
            continue
        joy = str(val(r, "place") or "").strip()
        tur = TUR.get(val(r, "kind"), "xorijiy")
        key = (_norm(nom_toza), _norm(joy))
        j = jur.get(key)
        if j is None:
            org = str(val(r, "org") or "").strip()
            j = jur[key] = {"nom": nom_toza, "joy": _lat(joy) if tur == "milliy" else joy,
                            "org": _lat(org) if tur == "milliy" else org, "tur": tur,
                            "link": str(val(r, "link") or "").strip(), "sohalar": []}
        if not j["link"] and val(r, "link"):
            j["link"] = str(val(r, "link")).strip()
        kod = str(val(r, "code") or "").strip()
        holat = "faol" if (status in FAOL and not nom.lstrip().startswith("[")) else "chiqarilgan"
        s = {"kod": kod, "holat": holat, "qoshilgan": _lat(str(val(r, "added") or "")),
             "chiqarilgan": str(val(r, "removed") or ""), "qaror": _lat(str(val(r, "decision") or "")),
             "izoh": _lat(str(val(r, "note") or ""))}
        old = next((x for x in j["sohalar"] if x["kod"] == kod), None)
        if old is None:
            j["sohalar"].append(s)
        elif holat == "faol":                          # bir soha bo'yicha faol yozuv ustun
            old.update(s)
    jurnallar = sorted(jur.values(), key=lambda x: ({"milliy": 0, "mdh": 1, "xorijiy": 2}[x["tur"]], _norm(x["nom"])))
    for i, j in enumerate(jurnallar):
        j["id"] = i
        j["sohalar"].sort(key=lambda s: s["kod"])
    for d in top.get("disciplines") or []:
        if isinstance(d, dict) and d.get("code") and d.get("name"):
            FANLAR.setdefault(d["code"], d["name"])
    pubs = [x for x in (top.get("publishers") or []) if isinstance(x, dict) and x.get("name")]
    imzo = hashlib.sha256(json.dumps([jurnallar, pubs], ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:20]
    return {"olingan": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), "xesh": imzo,
            "jami_qator": p.get("total") or len(rows), "jurnallar": jurnallar, "bazalar": pubs}


# ═══════════════════════════════════════════════════════════════════
# 2. SAQLASH VA YANGILASH
# ═══════════════════════════════════════════════════════════════════
_LOCK = threading.Lock()
_DATA: dict | None = None


def data() -> dict | None:
    return _DATA


def _saqla(d: dict):
    try:
        tmp = CACHE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, CACHE)
    except Exception as e:
        log.warning(f"OAK kesh saqlanmadi: {e}")


def kesh_yukla() -> dict | None:
    global _DATA
    try:
        with open(CACHE, encoding="utf-8") as f:
            d = json.load(f)
        if d.get("jurnallar"):
            with _LOCK:
                _DATA = d
            return d
    except Exception:
        pass
    return None


def _juftlar(d: dict) -> dict:
    out = {}
    for j in d.get("jurnallar", []):
        for s in j["sohalar"]:
            out[(_norm(j["nom"]), _norm(j["joy"]), s["kod"])] = (s["holat"], j)
    return out


def farq(eski: dict | None, yangi: dict) -> dict:
    """Qo'shilgan va chiqarilgan (jurnal, soha) juftlari."""
    if not eski:
        return {"qoshildi": [], "chiqarildi": []}
    a, b = _juftlar(eski), _juftlar(yangi)
    qosh, chiq = [], []
    for k, (h, j) in b.items():
        e = a.get(k)
        if h == "faol" and (e is None or e[0] != "faol"):
            qosh.append({"nom": j["nom"], "tur": j["tur"], "kod": k[2]})
        if h == "chiqarilgan" and e is not None and e[0] == "faol":
            chiq.append({"nom": j["nom"], "tur": j["tur"], "kod": k[2]})
    for k, (h, j) in a.items():                        # ro'yxatdan butunlay o'chirilganlar
        if k not in b and h == "faol":
            chiq.append({"nom": j["nom"], "tur": j["tur"], "kod": k[2]})
    return {"qoshildi": qosh, "chiqarildi": chiq}


def yangila(html: str | None = None) -> tuple[dict, dict]:
    """Manbadan yangilaydi. Qaytaradi: (data, farq). Xatoda istisno."""
    global _DATA
    d = tahlil(html if html is not None else yukla())
    if len(d["jurnallar"]) < 100 and html is None:
        raise ValueError(f"OAK ro'yxati shubhali darajada qisqa ({len(d['jurnallar'])})")
    with _LOCK:
        eski = _DATA
        f = farq(eski, d) if (eski and eski.get("xesh") != d["xesh"]) else {"qoshildi": [], "chiqarildi": []}
        if eski and eski.get("xesh") == d["xesh"]:
            d["olingan"] = eski.get("olingan", d["olingan"])
            d["tekshirildi"] = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        _DATA = d
    _saqla(d)
    return d, f


def fon_yangilash(on_change=None, interval: int = 6 * 3600, birinchi_kechikish: int = 20):
    """Fon oqimi: kesh → darhol foydalanish, keyin manbadan yangilash; o'zgarsa on_change(data, farq, eski_xesh)."""
    kesh_yukla()

    def loop():
        time.sleep(birinchi_kechikish)
        while True:
            try:
                eski_xesh = (_DATA or {}).get("xesh")
                d, f = yangila()
                log.info(f"📰 OAK ro'yxati tekshirildi: {len(d['jurnallar'])} jurnal, xesh {d['xesh']}")
                try:
                    yangiliklar_yangila()
                except Exception as e:
                    log.warning(f"OAK yangiliklari olinmadi: {e}")
                if on_change and eski_xesh != d["xesh"]:
                    try:
                        on_change(d, f, eski_xesh)
                    except Exception as e:
                        log.error(f"OAK on_change xato: {e}", exc_info=True)
                time.sleep(interval)
            except Exception as e:
                log.warning(f"OAK ro'yxatini yangilab bo'lmadi: {e}")
                time.sleep(30 * 60)
    threading.Thread(target=loop, daemon=True, name="oak-yangilash").start()


# ═══════════════════════════════════════════════════════════════════
# 3. QIDIRUV VA FILTRLAR
# ═══════════════════════════════════════════════════════════════════
def faol_sohalar(j: dict) -> list:
    return [s for s in j["sohalar"] if s["holat"] == "faol"]


def qidir(q: str, limit: int = 8) -> list:
    d = _DATA or {}
    nq = _norm(q)
    if len(nq) < 3:
        return []
    sozlar = [w for w in nq.split() if len(w) >= 2]
    natija = []
    for j in d.get("jurnallar", []):
        nn = _norm(j["nom"])
        if nq in nn:
            ball = 100 - min(50, len(nn) - len(nq)) / 5 + (10 if nn.startswith(nq) else 0)
        else:
            mos = sum(1 for w in sozlar if w in nn)
            if not sozlar or mos < max(1, len(sozlar) - (1 if len(sozlar) >= 3 else 0)):
                continue
            ball = 40 * mos / len(sozlar)
        if faol_sohalar(j):
            ball += 5
        natija.append((ball, j))
    natija.sort(key=lambda x: -x[0])
    return [j for _, j in natija[:limit]]


def boyicha(kod: str, tur: str | None = None, faqat_faol: bool = True) -> list:
    d = _DATA or {}
    out = []
    for j in d.get("jurnallar", []):
        if tur and j["tur"] != tur:
            continue
        for s in j["sohalar"]:
            if s["kod"] == kod and (s["holat"] == "faol" or not faqat_faol):
                out.append(j)
                break
    return out


def statistika() -> dict:
    d = _DATA or {}
    st = {"milliy": 0, "mdh": 0, "xorijiy": 0, "chiqarilgan": 0}
    for j in d.get("jurnallar", []):
        if faol_sohalar(j):
            st[j["tur"]] += 1
        else:
            st["chiqarilgan"] += 1
    return st


_OY = {"yanvar": 1, "fevral": 2, "mart": 3, "aprel": 4, "may": 5, "iyun": 6, "iyul": 7, "avgust": 8,
       "sentyabr": 9, "sentabr": 9, "oktyabr": 10, "oktabr": 10, "noyabr": 11, "dekabr": 12}


def sana(t: str) -> datetime.date | None:
    t = (t or "").strip().lower()
    m = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", t)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})\s+([a-z']+)\s+(\d{4})", t)
    if m and m.group(2) in _OY:
        try:
            return datetime.date(int(m.group(3)), _OY[m.group(2)], int(m.group(1)))
        except ValueError:
            return None
    return None


def songgi(n: int = 12) -> tuple[list, list]:
    """Eng so'nggi qo'shilgan va chiqarilgan (jurnal, soha) yozuvlari — sanasi bo'yicha."""
    d = _DATA or {}
    q, c = [], []
    for j in d.get("jurnallar", []):
        for s in j["sohalar"]:
            if s["holat"] == "faol":
                dt = sana(s["qoshilgan"])
                if dt:
                    q.append((dt, j, s))
            else:
                dt = sana(s["chiqarilgan"])
                if dt:
                    c.append((dt, j, s))
    q.sort(key=lambda x: x[0], reverse=True)
    c.sort(key=lambda x: x[0], reverse=True)
    return q[:n], c[:n]


def bazalar() -> list:
    return list((_DATA or {}).get("bazalar") or [])


# ═══════════════════════════════════════════════════════════════════
# 4. OAK RASMIY YANGILIKLARI (ilmiy maslahatchi uchun eng so'nggi rasmiy ma'lumot)
# ═══════════════════════════════════════════════════════════════════
OAK_SITE_API = "https://oak.uz/api/public/site"
OAK_MATERIAL_API = "https://oak.uz/api/materiallar/{id}"
_YANGI: list = []


def _matn_html(h: str) -> str:
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h or "", flags=re.S | re.I)
    h = re.sub(r"<br\s*/?>|</p>|</li>|</h\d>", "\n", h, flags=re.I)
    h = re.sub(r"<[^>]+>", " ", h)
    import html as _h
    h = _h.unescape(h).replace("**", "")
    h = re.sub(r"[ \t]+", " ", h)
    return re.sub(r"\n\s*\n+", "\n", h).strip()


def _blok_matn(blocks) -> str:
    out = []

    def walk(x):
        if isinstance(x, str):
            out.append(x)
        elif isinstance(x, dict):
            for k, v in x.items():
                if k not in ("type", "src", "height", "image", "url", "href", "id"):
                    walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(blocks)
    return _matn_html(" ".join(out))


def yangiliklar_yukla(n: int = 8) -> list:
    req = urllib.request.Request(OAK_SITE_API, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=40) as r:
        d = json.loads(r.read().decode("utf-8", errors="replace"))
    items = d.get("news") or []
    if isinstance(items, dict):
        items = items.get("items") or list(items.values())
    out = []
    for it in items[:n]:
        if not isinstance(it, dict) or not it.get("title"):
            continue
        matn = _matn_html(it.get("bodyHtml") or it.get("body") or "")
        if len(matn) < 200 and it.get("id"):
            try:
                rq = urllib.request.Request(OAK_MATERIAL_API.format(id=it["id"]), headers={"User-Agent": UA})
                with urllib.request.urlopen(rq, timeout=25) as r2:
                    m = json.loads(r2.read().decode("utf-8", errors="replace"))
                matn = _blok_matn(m.get("blocks")) or matn
            except Exception:
                pass
        if len(matn) < 60:
            matn = _matn_html(it.get("excerptHtml") or it.get("excerpt") or "")
        href = it.get("href") or ""
        if href.startswith("/"):
            href = "https://oak.uz" + href
        out.append({"sarlavha": _lat(it["title"]).strip(), "sana": it.get("date") or "", "havola": href,
                    "matn": _lat(matn)[:2200]})
    return out


def yangiliklar() -> list:
    return list(_YANGI or (_DATA or {}).get("yangiliklar") or [])


def yangiliklar_yangila():
    global _YANGI
    y = yangiliklar_yukla()
    if y:
        _YANGI = y
        if _DATA is not None:
            _DATA["yangiliklar"] = y
            _saqla(_DATA)
    return y

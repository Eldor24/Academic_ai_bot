"""
openalex.py — OpenAlex (https://openalex.org) ochiq ilmiy ma'lumotlar bazasi bilan ishlash
═══════════════════════════════════════════════════════════════════
OpenAlex — 250+ mln ilmiy ish, jurnal va mualliflar haqidagi ochiq (CC0) ma'lumotlar bazasi, kalitsiz ishlaydi.
Bot undan:
  • mavzu bo'yicha HAQIQIY maqolalar va dissertatsiyalarni (DOI, iqtiboslar soni bilan) topish,
  • o'xshash maqolalar qaysi jurnallarda chop etilganini aniqlash (mos xalqaro jurnallar),
  • jurnal ko'rsatkichlari (h-indeks, o'rtacha iqtibos, ochiq kirish, APC) uchun foydalanadi.
Scopus / Web of Science holati OpenAlex'da yo'q — foydalanuvchiga tekshirish havolalari beriladi.
Ixtiyoriy: OPENALEX_EMAIL (polite pool), OPENALEX_KEY muhit o'zgaruvchilari.
"""
from __future__ import annotations

import json
import logging
import os
import re
import urllib.parse
import urllib.request

log = logging.getLogger(__name__)
API = "https://api.openalex.org"
UA = "AcademicBot/1.0 (Telegram bot for students; mailto:{m})"


def _get(path: str, params: dict, timeout: int = 25) -> dict:
    params = {k: v for k, v in params.items() if v not in (None, "")}
    mail = os.getenv("OPENALEX_EMAIL", "")
    if mail:
        params["mailto"] = mail
    if os.getenv("OPENALEX_KEY"):
        params["api_key"] = os.getenv("OPENALEX_KEY")
    url = f"{API}{path}?{urllib.parse.urlencode(params, safe=':,|')}"
    req = urllib.request.Request(url, headers={"User-Agent": UA.format(m=mail or "none"), "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _abstract(inv: dict | None, limit: int = 60) -> str:
    if not isinstance(inv, dict) or not inv:
        return ""
    pos = []
    for w, ix in inv.items():
        for i in ix:
            pos.append((i, w))
    pos.sort()
    words = [w for _, w in pos[:limit]]
    return " ".join(words) + (" …" if len(pos) > limit else "")


def _work(w: dict) -> dict:
    src = ((w.get("primary_location") or {}).get("source") or {})
    b = w.get("biblio") or {}
    authors = [((a.get("author") or {}).get("display_name") or "").strip() for a in (w.get("authorships") or [])]
    doi = (w.get("doi") or "").replace("https://doi.org/", "")
    oa = w.get("open_access") or {}
    return {
        "id": (w.get("id") or "").rsplit("/", 1)[-1],
        "sarlavha": (w.get("title") or w.get("display_name") or "").strip(),
        "yil": w.get("publication_year"),
        "mualliflar": [a for a in authors if a],
        "jurnal": (src.get("display_name") or "").strip(),
        "issn": src.get("issn_l") or "",
        "jild": b.get("volume") or "", "son": b.get("issue") or "",
        "bet1": b.get("first_page") or "", "bet2": b.get("last_page") or "",
        "doi": doi, "iqtibos": w.get("cited_by_count") or 0,
        "oa_url": oa.get("oa_url") or "", "oa": bool(oa.get("is_oa")),
        "tur": w.get("type") or "", "til": w.get("language") or "",
        "annot": _abstract(w.get("abstract_inverted_index")),
        "nashriyot": (src.get("host_organization_name") or "").strip(),
        "manba_turi": src.get("type") or "",
        "doaj": bool(src.get("is_in_doaj")),
        "core": bool(src.get("is_core")),
        "indexed": [str(x).lower() for x in (w.get("indexed_in") or [])],
        "rel": float(w.get("relevance_score") or 0),
    }


_FIELDS = ("id,title,display_name,publication_year,authorships,primary_location,biblio,doi,cited_by_count,"
           "open_access,type,language,abstract_inverted_index")


def ishlar(query: str, n: int = 20, dan_yil: int | None = None, sort: str | None = None, tur: str | None = None,
           til: str | None = None, aniq: bool = True) -> list:
    """Ishlarni qidiradi. Har doim: qaytarib olingan (retracted) va paratext (muqova, mundarija) ishlar chiqarib tashlanadi."""
    f = ["is_retracted:false", "is_paratext:false"]
    if dan_yil:
        f.append(f"from_publication_date:{dan_yil}-01-01")
    if tur:
        f.append(f"type:{tur}")
    if til:
        f.append(f"language:{til}")
    q = re.sub(r"[,:|!()\"]+", " ", query or "").strip()
    if aniq and q:
        # faqat sarlavha va annotatsiya bo'yicha (to'liq matndagi tasodifiy so'z mosligi hisobga olinmaydi)
        d = _get("/works", {"filter": ",".join(f + [f"title_and_abstract.search:{q}"]), "per-page": min(50, n),
                            "sort": sort or "relevance_score:desc"})
        res = [_work(w) for w in d.get("results", []) if (w.get("title") or w.get("display_name"))]
        if len(res) >= min(5, n):
            return res
    d = _get("/works", {"search": q, "filter": ",".join(f), "per-page": min(50, n), "sort": sort})
    return [_work(w) for w in d.get("results", []) if (w.get("title") or w.get("display_name"))]


def soni(query: str, dan_yil: int | None = None) -> int | None:
    """Sarlavha va annotatsiyasida shu so'rov bo'lgan ishlar soni (mavzu qanchalik o'rganilganini baholash uchun)."""
    q = re.sub(r"[,:|!()\"]+", " ", query or "").strip()
    if not q:
        return None
    f = ["is_retracted:false", "is_paratext:false", f"title_and_abstract.search:{q}"]
    if dan_yil:
        f.append(f"from_publication_date:{dan_yil}-01-01")
    d = _get("/works", {"filter": ",".join(f), "per-page": 1})
    try:
        return int((d.get("meta") or {}).get("count"))
    except (TypeError, ValueError):
        return None


def jurnallar(query: str, n: int = 10, dan_yil: int = 2019) -> list:
    """O'xshash maqolalar eng ko'p chop etilgan jurnallar (manba bo'yicha guruhlash) + ularning ko'rsatkichlari."""
    d = _get("/works", {"search": query, "filter": f"from_publication_date:{dan_yil}-01-01,type:article",
                        "group_by": "primary_location.source.id"})
    groups = [g for g in d.get("group_by", []) if g.get("key") and "openalex.org/S" in str(g.get("key"))]
    out = []
    for g in groups[: n + 4]:
        sid = str(g["key"]).rsplit("/", 1)[-1]
        try:
            s = _get(f"/sources/{sid}", {})
        except Exception as e:
            log.debug(f"OpenAlex source {sid}: {e}")
            continue
        if s.get("type") not in (None, "journal"):
            continue
        st = s.get("summary_stats") or {}
        out.append({
            "nom": s.get("display_name") or g.get("key_display_name") or "",
            "issn": s.get("issn_l") or ((s.get("issn") or [""])[0] if s.get("issn") else ""),
            "issnlar": s.get("issn") or [],
            "nashriyot": s.get("host_organization_name") or "",
            "mos_soni": g.get("count") or 0,
            "h": st.get("h_index"), "iqt2": st.get("2yr_mean_citedness"),
            "oa": bool(s.get("is_oa")), "doaj": bool(s.get("is_in_doaj")),
            "apc": s.get("apc_usd"), "sayt": s.get("homepage_url") or "",
            "davlat": s.get("country_code") or "", "ishlar": s.get("works_count") or 0,
        })
        if len(out) >= n:
            break
    return out


def ishonch(w: dict, oak_nomlar: set | None = None, norm=None) -> list:
    """Ishning ishonchlilik belgilari (kalitlar): oak, pubmed, doaj, doi, jurnal, core."""
    out = []
    if oak_nomlar and norm and w.get("jurnal") and norm(w["jurnal"]) in oak_nomlar:
        out.append("oak")
    if "pubmed" in w.get("indexed", []):
        out.append("pubmed")
    if w.get("core"):
        out.append("core")
    if w.get("doaj"):
        out.append("doaj")
    if w.get("doi"):
        out.append("doi")
    if w.get("jurnal") and (w.get("manba_turi") in ("journal", "") or w.get("tur") in ("article", "review")):
        out.append("jurnal")
    return out


def qidiruv_havolalari(mavzu: str, en: str = "", ru: str = "") -> list:
    """Rasmiy va ishonchli bazalarda mavzu bo'yicha tayyor qidiruv havolalari."""
    from urllib.parse import quote_plus as qp
    m, e, r = mavzu[:150], (en or mavzu)[:150], (ru or mavzu)[:150]
    return [
        ("Lex.uz — O'zbekiston qonunchiligi (normativ-huquqiy hujjatlar)", f"https://lex.uz/uz/search/all?query={qp(m)}"),
        ("Google Scholar (2019-yildan)", f"https://scholar.google.com/scholar?q={qp(e)}&as_ylo=2019"),
        ("CyberLeninka (rus tilidagi ochiq maqolalar)", f"https://cyberleninka.ru/search?q={qp(r)}"),
        ("eLIBRARY.RU (RINTs)", f"https://www.elibrary.ru/querybox.asp?keywords={qp(r)}"),
        ("PubMed (OAK B toifasi — tibbiyot, biologiya)", f"https://pubmed.ncbi.nlm.nih.gov/?term={qp(e)}&filter=years.2019-2026"),
        ("DOAJ — ochiq kirishli taqrizlanadigan jurnallar", f"https://doaj.org/search/articles?source=%7B%22query%22%3A%7B%22query_string%22%3A%7B%22query%22%3A%22{qp(e)}%22%7D%7D%7D"),
        ("OpenAlex", f"https://openalex.org/works?search={qp(e)}&filter=publication_year:2019-2026"),
    ]


def tekshir_havolalar(issn: str) -> dict:
    """Jurnalni xalqaro bazalarda tekshirish havolalari (rasmiy saytlar)."""
    q = urllib.parse.quote(issn or "")
    return {
        "Web of Science (Master Journal List)": f"https://mjl.clarivate.com/search-results?issn={q}&hide_exact_match_fl=true",
        "Scopus (SCImago — Scopus ma'lumotlari)": f"https://www.scimagojr.com/journalsearch.php?q={q}",
        "Scopus Sources": "https://www.scopus.com/sources",
    }


# ── Bibliografik rasmiylashtirish ──
def _ism_gost(full: str) -> str:
    """«John A. Smith» → «Smith J. A.»"""
    p = [x for x in full.replace(",", " ").split() if x]
    if len(p) < 2:
        return full
    fam = p[-1]
    ini = " ".join(x[0].upper() + "." for x in p[:-1] if x[0].isalpha())
    return f"{fam} {ini}".strip()


def _ism_apa(full: str) -> str:
    p = [x for x in full.replace(",", " ").split() if x]
    if len(p) < 2:
        return full
    return f"{p[-1]}, " + " ".join(x[0].upper() + "." for x in p[:-1] if x[0].isalpha())


def gost(w: dict) -> str:
    """GOST 7.1-2003 / 7.0.5 uslubiga yaqin yozuv (O'zbekiston dissertatsiyalarida keng qo'llanadi)."""
    a = w["mualliflar"]
    bosh = _ism_gost(a[0]) + " " if a and len(a) <= 3 else ""
    s = f"{bosh}{w['sarlavha'].rstrip('.')}"
    if a:
        qolgan = ", ".join(_ism_gost(x) for x in a[:3])
        s += f" / {qolgan}{' [et al.]' if len(a) > 3 else ''}"
    if w["jurnal"]:
        s += f" // {w['jurnal']}"
    if w["yil"]:
        s += f". – {w['yil']}"
    if w["jild"]:
        s += f". – Vol. {w['jild']}"
    if w["son"]:
        s += f", № {w['son']}"
    if w["bet1"]:
        s += f". – P. {w['bet1']}" + (f"–{w['bet2']}" if w["bet2"] and w["bet2"] != w["bet1"] else "")
    if w["doi"]:
        s += f". – DOI: {w['doi']}"
    return s + "."


def apa(w: dict) -> str:
    a = w["mualliflar"]
    if not a:
        au = ""
    elif len(a) == 1:
        au = _ism_apa(a[0])
    elif len(a) <= 20:
        au = ", ".join(_ism_apa(x) for x in a[:-1]) + ", & " + _ism_apa(a[-1])
    else:
        au = ", ".join(_ism_apa(x) for x in a[:19]) + ", … " + _ism_apa(a[-1])
    s = f"{au} ({w['yil'] or 'n.d.'}). {w['sarlavha'].rstrip('.')}."
    if w["jurnal"]:
        s += f" {w['jurnal']}"
        if w["jild"]:
            s += f", {w['jild']}"
            if w["son"]:
                s += f"({w['son']})"
        if w["bet1"]:
            s += f", {w['bet1']}" + (f"–{w['bet2']}" if w["bet2"] and w["bet2"] != w["bet1"] else "")
        s += "."
    if w["doi"]:
        s += f" https://doi.org/{w['doi']}"
    return s.strip()

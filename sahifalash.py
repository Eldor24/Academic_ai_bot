"""
sahifalash.py — Word hujjati sahifalarini oldindan aniq hisoblash (mundarija uchun)
═══════════════════════════════════════════════════════════════════
Word mundarijasidagi sahifa raqamlarini foydalanuvchi Ctrl+A → F9 bosmasdan to'g'ri chiqarish uchun
hujjat matni shrift o'lchamlari bo'yicha qatorlarga bo'linadi va sahifalarga joylanadi:

  • Times New Roman bilan bir xil o'lchamli Liberation Serif shriftining haqiqiy harf kengliklari
    (assets/fonts, SIL OFL litsenziyasi) — har bir qator Word'dagidek so'zma-so'z ko'chiriladi
  • qator balandligi, abzas oldi/keyingi bo'shliqlar, «keyingisi bilan birga» (sarlavhalar),
    yetim/beva qatorlar nazorati (widow/orphan), yangi sahifadan boshlanadigan boblar
  • snoskalar sahifaning pastki qismidan joy oladi (havola turgan qator bilan bir sahifada)
  • jadval qatorlari (sarlavha qatori yangi sahifada takrorlanadi), rasmlar

LibreOffice bilan solishtirib tekshirilgan. Bloklar ro'yxati kurs_ishi.py da tuziladi.
"""
from __future__ import annotations

import math
import os
import re

_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "fonts")
_FILES = {(False, False): "LiberationSerif-Regular.ttf", (True, False): "LiberationSerif-Bold.ttf",
          (False, True): "LiberationSerif-Italic.ttf", (True, True): "LiberationSerif-Bold.ttf"}

PT_CM = 72 / 2.54
PAGE_H = (29.7 - 2.0 - 2.0) * PT_CM          # matn maydoni balandligi
PAGE_W = (21.0 - 3.0 - 1.5) * PT_CM          # matn maydoni kengligi
LINE_K = 2355 / 2048                          # TNR/Liberation: (ascent+descent+lineGap)/unitsPerEm
FN_SIZE = 10
FN_SEP = 5.7                                  # snoska ajratuvchi chizig'i (5 pt aniq interval) + bo'shliq
CELL_PAD = 0.19 * PT_CM                       # jadval katagining chap/o'ng ichki chekkasi
# "1,5 interval"da qo'shimcha bo'shliq qator OSTIGA qo'yiladi (Word ham, LibreOffice ham) va sahifadagi oxirgi
# qatorda u talab qilinmaydi — qator matni sig'sa yetarli. LibreOffice bilan 40 ta hujjatda tekshirilgan.
LO_MODE = True

_cache: dict = {}
_wcache: dict = {}


def _font(bold: bool, italic: bool, size: float):
    key = (bold, italic, size)
    f = _cache.get(key)
    if f is None:
        from PIL import ImageFont
        f = ImageFont.truetype(os.path.join(_FONT_DIR, _FILES[(bold, italic)]), size=int(round(size * 20)))
        _cache[key] = f
    return f


def width(text: str, size: float = 14, bold: bool = False, italic: bool = False) -> float:
    k = (text, size, bold, italic)
    w = _wcache.get(k)
    if w is None:
        w = _font(bold, italic, size).getlength(text) / 20.0
        if len(_wcache) < 200_000:
            _wcache[k] = w
    return w


def available() -> bool:
    try:
        _font(False, False, 14)
        return True
    except Exception:
        return False


def line_h(size: float, ls: float) -> float:
    return size * LINE_K * ls


def wrap(segs: list, size: float, avail_first: float, avail_rest: float) -> tuple[int, list]:
    """segs: [("t", matn, bold, italic) | ("fn", snoska_matni)] → (qatorlar soni, har snoskaning qator raqami)."""
    tokens = []                                   # (so'z, kenglik, oldidan_bo'shliq, [snoskalar])
    pending_space = False
    for s in segs:
        if s[0] == "fn":
            if tokens:
                w, sp, fns = tokens[-1][1], tokens[-1][2], tokens[-1][3]
                tokens[-1] = (tokens[-1][0], w + width("9", size * 0.65), sp, fns + [s[1]])
            continue
        _, txt, b, it = s
        parts = re.split(r"(\s+)", txt)
        for part in parts:
            if not part:
                continue
            if part.isspace():
                pending_space = True
                continue
            w = width(part, size, b, it)
            if tokens and not pending_space:      # bo'shliqsiz davomi (masalan, qalin + oddiy)
                t0 = tokens[-1]
                tokens[-1] = (t0[0] + part, t0[1] + w, t0[2], t0[3])
            else:
                tokens.append((part, w, pending_space and bool(tokens), []))
            pending_space = False
    space = width(" ", size)
    lines = 1
    cur = 0.0
    fn_lines = []
    avail = avail_first
    for word, w, sp, fns in tokens:
        add = (space if (sp and cur > 0) else 0) + w
        if cur > 0 and cur + add > avail + 0.05:
            lines += 1
            cur = 0.0
            avail = avail_rest
            add = w
        while add > avail and cur == 0 and add > 0:   # juda uzun so'z — bo'linadi
            add -= avail
            lines += 1
        cur += add
        for fn in fns:
            fn_lines.append((lines - 1, fn))
    return lines, fn_lines


def fn_height(text: str) -> float:
    n, _ = wrap([("t", text, False, False)], FN_SIZE, PAGE_W - width("99", FN_SIZE), PAGE_W)
    return n * line_h(FN_SIZE, 1.0)


class _Page:
    def __init__(self):
        self.y = 0.0          # ishlatilgan balandlik (yuqoridan)
        self.fn = 0.0         # snoskalar maydoni

    def free(self) -> float:
        return PAGE_H - self.y - self.fn


def _para_metrics(b: dict):
    size = b.get("size", 14)
    left = b.get("left", 0.0) * PT_CM
    right = b.get("right", 0.0) * PT_CM
    ind = b.get("ind", 0.0) * PT_CM
    avail_first = PAGE_W - left - right - max(0.0, ind)
    avail_rest = PAGE_W - left - right - max(0.0, -ind)
    if b.get("img_h"):
        return 1, [], b["img_h"] * 1.0 + 2
    n, fns = wrap(b.get("segs", []), size, avail_first, avail_rest)
    return n, fns, line_h(size, b.get("ls", 1.5))


def _block_first_height(b: dict) -> float:
    """«Keyingisi bilan birga» uchun keyingi blokning birinchi qismi balandligi."""
    if b["k"] == "tbl":
        return _row_h(b, b["cols"]) + _row_h(b, b["rows"][0] if b["rows"] else b["cols"])
    n, _, lh = _para_metrics(b)
    return b.get("sb", 0) + lh * min(n, 2)


def _row_h(b: dict, row: list) -> float:
    ncol = max(1, len(b["cols"]))
    cw = PAGE_W / ncol - 2 * CELL_PAD
    size = b.get("size", 12)
    lines = 1
    for c in row:
        n, _ = wrap([("t", str(c), b.get("hdr_bold", False) and row is b["cols"], False)], size, cw, cw)
        lines = max(lines, n)
    return lines * line_h(size, 1.0) + 1.0


def paginate(blocks: list, start_page: int = 1, trace: list | None = None) -> tuple[list, int]:
    """Har blok boshlangan sahifa raqamini va jami sahifalar sonini qaytaradi."""
    pages = [0] * len(blocks)
    pg = _Page()
    num = start_page

    def new_page():
        nonlocal pg, num
        pg = _Page()
        num += 1

    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b.get("pbb") and (pg.y > 0 or pg.fn > 0):
            new_page()
        if b["k"] == "titul":                        # titul varag'i — butun 1-bet
            pages[i] = num
            pg.y = PAGE_H
            i += 1
            continue
        if b["k"] == "tbl":
            pages[i] = num
            hdr = _row_h(b, b["cols"])
            if pg.free() < hdr + _row_h(b, b["rows"][0] if b["rows"] else b["cols"]) and pg.y > 0:
                new_page()
                pages[i] = num
            pg.y += hdr
            for r in b["rows"]:
                h = _row_h(b, r)
                if h > pg.free():
                    new_page()
                    pg.y += hdr                       # sarlavha qatori takrorlanadi
                pg.y += h
            pg.y += b.get("sa", 0)
            i += 1
            continue

        n, fns, lh = _para_metrics(b)
        size_ = b.get("size", 14)
        if trace is not None:
            trace.append((i, num, round(pg.y, 1), round(pg.fn, 1)))
        sb = b.get("sb", 0) if pg.y > 0 else 0
        # «keyingisi bilan birga» zanjiri (sarlavha + keyingi abzasning 2 qatori, rasm + izoh ...)
        if b.get("keep") and i + 1 < len(blocks) and pg.y > 0:
            need = sb + n * lh + b.get("sa", 0)
            j = i + 1
            while j < len(blocks) and blocks[j].get("keep") and not blocks[j].get("pbb"):
                nj, _, lhj = _para_metrics(blocks[j]) if blocks[j]["k"] != "tbl" else (0, [], 0)
                need += (blocks[j].get("sb", 0) + nj * lhj + blocks[j].get("sa", 0)) if blocks[j]["k"] != "tbl" \
                    else _block_first_height(blocks[j])
                j += 1
            if j < len(blocks) and not blocks[j].get("pbb"):
                need += _block_first_height(blocks[j])
            if need > pg.free() and need <= PAGE_H:
                new_page()
                sb = 0
        # snoskalar har qator bo'yicha
        fn_by_line: dict = {}
        for ln, txt in fns:
            fn_by_line.setdefault(ln, []).append(fn_height(txt))

        def line_cost(k):
            extra = sum(fn_by_line.get(k, []))
            return lh, extra

        # qancha qator sig'adi
        def fit_count(start):
            y, f, cnt = pg.y + (sb if start == 0 else 0), pg.fn, 0
            relief = (lh - size_ * LINE_K) if (LO_MODE and lh > size_ * LINE_K) else 0.0
            for k in range(start, n):
                h, e = line_cost(k)
                if e and f == 0:
                    e += FN_SEP
                if y + h - relief + f + e > PAGE_H + 0.01:
                    break
                y += h
                f += e
                cnt += 1
            return cnt

        start = 0
        pages[i] = None
        while start < n:
            k = fit_count(start)
            rest = n - start
            if k < rest:                              # abzas bo'linadi — yetim/beva qatorlar
                if b.get("img_h"):
                    k = 0
                elif rest >= 2 and k == 1:
                    k = 0
                elif rest - k == 1 and rest >= 3:
                    k -= 1
                if k <= 0:
                    if pg.y == 0 and pg.fn == 0:      # bo'sh sahifaga ham sig'masa — majburan
                        k = max(1, fit_count(start))
                    else:
                        new_page()
                        sb = 0
                        continue
            if pages[i] is None:
                pages[i] = num
            for kk in range(start, start + k):
                h, e = line_cost(kk)
                if e and pg.fn == 0:
                    e += FN_SEP
                pg.y += h + ((sb if kk == 0 else 0))
                pg.fn += e
            start += k
            if start < n:
                new_page()
        pg.y += b.get("sa", 0)
        if pg.y + pg.fn > PAGE_H:
            pg.y = PAGE_H - pg.fn
        i += 1
    return pages, num

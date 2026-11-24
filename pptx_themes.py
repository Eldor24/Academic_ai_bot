"""
pptx_themes.py — Taqdimot shablonlari dvigateli (Academic Bot, v2)
═══════════════════════════════════════════════════════════════════
15 ta shablon. Har bir shablon faqat rangi bilan emas, TUZILISHI bilan farq qiladi:
har bir slayd turining (titul, sarlavha, matn, ro'yxat, reja, ustun/doira diagramma,
jadval, taqqoslash, statistika, bosqichlar, xulosa) bir nechta maket varianti bor
va har bir shablon o'z kombinatsiyasini ishlatadi.

Masalan "ustun_diag" (statistika) slaydi shablonga qarab:
  vertikal ustunlar • gorizontal ustunlar • chiziqli grafik • maydon grafik •
  progress-chiziqlar • lollipop infografika
ko'rinishida chiziladi. "doira_diag" esa:
  doughnut • to'liq doira • 100% tasma • waffle (10x10) • pufakchalar • halqalar

Asosiy funksiya:
    build_presentation(slides, mavzu, ism, theme_key, til, out_path, labels=None)

Ko'rinish rasmlarini qayta yaratish (LibreOffice kerak):
    python pptx_themes.py --previews
"""
from __future__ import annotations

import datetime
import math
import os
import random
import re
from dataclasses import dataclass, field

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

SW, SH = 13.333, 7.5
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "themes")



def _joriy_yil() -> int:
    try:
        import matn_norm
        return matn_norm.joriy_yil()
    except Exception:
        return datetime.date.today().year

def _hex(h: str) -> tuple:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ═══════════════════════════════════════════════════════════════════
# SHABLON TA'RIFI
# ═══════════════════════════════════════════════════════════════════
@dataclass
class Theme:
    key: str
    names: dict
    desc: dict
    decor: str                  # kontent slayd foni bezagi
    v: dict                     # har bir slayd turi uchun maket varianti
    bg: str
    title_bg: str
    primary: str
    accent: str
    accent2: str
    text: str
    muted: str
    card: str
    on_title: str
    chart: list = field(default_factory=list)
    font_head: str = "Calibri"
    font_body: str = "Calibri"
    dark: bool = False
    grad: tuple | None = None
    line: str = "#E3E6EA"       # ajratuvchi chiziqlar rangi
    progress: bool = False      # pastda progress chizig'i

    def c(self, name: str) -> RGBColor:
        return RGBColor(*_hex(getattr(self, name)))

    def name(self, til: str) -> str:
        return self.names.get(til, self.names["uz"])


THEMES: dict[str, Theme] = {}


def _t(**kw):
    THEMES[kw["key"]] = Theme(**kw)


_t(key="akademik", decor="classic",
   names={"uz": "🎓 Akademik", "ru": "🎓 Академический", "en": "🎓 Academic"},
   desc={"uz": "Klassik ilmiy uslub: ramkali titul, ustun diagramma, kartalar",
         "ru": "Классический научный стиль: рамка, столбцы, карточки",
         "en": "Classic scholarly: framed title, column charts, cards"},
   v=dict(title="framed_center", header="underline", text="quote_card", list="rows", agenda="columns",
          bar="column", pie="doughnut", table="banded", compare="cards_vs", stats="cards",
          timeline="horizontal", concl="check_rows"),
   bg="#FFFFFF", title_bg="#FFFFFF", primary="#1B2A4A", accent="#B8892B", accent2="#3C5A8A",
   text="#2B2F36", muted="#6B7280", card="#F3F5F9", on_title="#1B2A4A",
   chart=["#1B2A4A", "#B8892B", "#3C5A8A", "#8FA3C4", "#D9B871", "#5B6B82"], font_head="Cambria")

_t(key="minimal", decor="swiss",
   names={"uz": "◻️ Minimal", "ru": "◻️ Минимализм", "en": "◻️ Minimal"},
   desc={"uz": "Shveytsar uslubi: katta raqamlar, chiziqli grafik, 100% tasma",
         "ru": "Швейцарский стиль: крупные цифры, линейный график, 100% полоса",
         "en": "Swiss style: big numbers, line chart, 100% strip"},
   v=dict(title="swiss_left", header="square_tag", text="big_statement", list="numbered_big", agenda="big_vertical",
          bar="line", pie="stacked100", table="minimal_lines", compare="split_line", stats="inline",
          timeline="vertical", concl="numbered_big"),
   bg="#F6F4EF", title_bg="#F6F4EF", primary="#111111", accent="#E63E21", accent2="#111111",
   text="#222222", muted="#7A7A7A", card="#FFFFFF", on_title="#111111", line="#D8D4CC",
   chart=["#E63E21", "#111111", "#8C8C8C", "#C9C5BC", "#F08A73", "#4A4A4A"], font_body="Calibri Light")

_t(key="korporativ", decor="plain",
   names={"uz": "💼 Korporativ", "ru": "💼 Корпоративный", "en": "💼 Corporate"},
   desc={"uz": "Biznes uslub: rangli sarlavha tasmasi, gorizontal ustunlar, doira diagramma",
         "ru": "Бизнес-стиль: цветная полоса заголовка, горизонтальные столбцы, круговая",
         "en": "Business: colour title band, horizontal bars, pie chart"},
   v=dict(title="split_panel", header="top_band", text="two_column", list="icon_columns", agenda="left_panel",
          bar="bar_h", pie="pie", table="header_col", compare="colored_halves", stats="circles",
          timeline="chevrons", concl="cards"),
   bg="#FFFFFF", title_bg="#0B4F9C", primary="#0B4F9C", accent="#1E88E5", accent2="#00B3A4",
   text="#1F2937", muted="#6B7280", card="#EEF4FB", on_title="#FFFFFF",
   chart=["#0B4F9C", "#1E88E5", "#00B3A4", "#90CAF9", "#FFB300", "#5C6BC0"])

_t(key="texno", decor="dark_geo", dark=True, progress=True,
   names={"uz": "🌌 Texno", "ru": "🌌 Техно", "en": "🌌 Tech"},
   desc={"uz": "Qorong'i IT uslub: progress-chiziqlar, halqa diagrammalar, tugunli reja",
         "ru": "Тёмный IT-стиль: прогресс-полосы, кольца, узлы",
         "en": "Dark IT style: progress bars, ring charts, node agenda"},
   v=dict(title="geo_rings", header="bracket", text="highlight_box", list="pills", agenda="nodes",
          bar="progress", pie="rings", table="banded", compare="cards_vs", stats="hero_left",
          timeline="horizontal", concl="banner"),
   bg="#0F172A", title_bg="#0B1120", primary="#F8FAFC", accent="#22D3EE", accent2="#A78BFA",
   text="#E2E8F0", muted="#94A3B8", card="#1E293B", on_title="#F8FAFC", line="#263248",
   chart=["#22D3EE", "#A78BFA", "#34D399", "#F472B6", "#FBBF24", "#60A5FA"], font_body="Calibri Light")

_t(key="tabiat", decor="organic",
   names={"uz": "🌿 Tabiat", "ru": "🌿 Природа", "en": "🌿 Nature"},
   desc={"uz": "Organik shakllar: maydon grafik, waffle diagramma, zinapoya bosqichlar",
         "ru": "Органические формы: диаграмма-область, вафля, ступени",
         "en": "Organic shapes: area chart, waffle chart, stair steps"},
   v=dict(title="organic_blobs", header="badge", text="split_panel", list="zigzag", agenda="tabs",
          bar="area", pie="waffle", table="card_rows", compare="table_compare", stats="circles",
          timeline="stairs", concl="check_rows"),
   bg="#F7F4EC", title_bg="#F7F4EC", primary="#1F4D3A", accent="#4F8A5B", accent2="#C08B3E",
   text="#2E3A33", muted="#6E7B72", card="#ECE6D6", on_title="#1F4D3A", line="#DDD5C2",
   chart=["#1F4D3A", "#4F8A5B", "#9CBF8B", "#C08B3E", "#E3C38C", "#6E7B72"], font_head="Cambria")

_t(key="gradient", decor="gradient", progress=True,
   names={"uz": "🌈 Gradient", "ru": "🌈 Градиент", "en": "🌈 Gradient"},
   desc={"uz": "Yorqin zamonaviy: gradient tasma, lollipop grafik, strelkali bosqichlar",
         "ru": "Яркий современный: градиент, леденцовый график, стрелки",
         "en": "Bright modern: gradient band, lollipop chart, chevrons"},
   v=dict(title="full_gradient", header="top_band", text="quote_card", list="cards", agenda="nodes",
          bar="lollipop", pie="doughnut", table="card_rows", compare="colored_halves", stats="cards",
          timeline="chevrons", concl="banner"),
   bg="#FFFFFF", title_bg="#5B21B6", primary="#3B0764", accent="#7C3AED", accent2="#EC4899",
   text="#1F1530", muted="#6B6880", card="#F5F0FF", on_title="#FFFFFF",
   chart=["#7C3AED", "#EC4899", "#F59E0B", "#06B6D4", "#A78BFA", "#F9A8D4"], grad=("#4C1D95", "#DB2777"))

_t(key="ijodiy", decor="pastel", progress=True,
   names={"uz": "🎨 Ijodiy", "ru": "🎨 Креатив", "en": "🎨 Creative"},
   desc={"uz": "Pastel ranglar: yumaloq kartalar, pufakcha diagramma, zinapoya",
         "ru": "Пастель: скруглённые карточки, пузырьки, ступени",
         "en": "Pastel: rounded cards, bubble chart, stairs"},
   v=dict(title="collage", header="badge", text="highlight_box", list="cards", agenda="tabs",
          bar="column", pie="bubbles", table="card_rows", compare="cards_vs", stats="circles",
          timeline="stairs", concl="cards"),
   bg="#FFF8F2", title_bg="#FFF1E6", primary="#3D2C5E", accent="#FF7A59", accent2="#2EC4B6",
   text="#3A3346", muted="#857C93", card="#FFFFFF", on_title="#3D2C5E", line="#F1E3D8",
   chart=["#FF7A59", "#2EC4B6", "#9B8AFB", "#FFC145", "#FF9EBB", "#3D2C5E"])

_t(key="milliy", decor="ornament",
   names={"uz": "🇺🇿 Milliy naqsh", "ru": "🇺🇿 Национальный", "en": "🇺🇿 Uzbek ornament"},
   desc={"uz": "Lojuvard va oltin, sharqona naqsh: markaziy sarlavha, doira diagramma",
         "ru": "Лазурит и золото, восточный орнамент",
         "en": "Lapis and gold with eastern ornament"},
   v=dict(title="ornament", header="centered", text="quote_card", list="zigzag", agenda="columns",
          bar="column", pie="pie", table="banded", compare="split_line", stats="cards",
          timeline="horizontal", concl="check_rows"),
   bg="#FBF8F1", title_bg="#0B2545", primary="#0B2545", accent="#1BA3A3", accent2="#C9A13B",
   text="#23303F", muted="#6A7380", card="#F1ECDF", on_title="#FBF8F1", line="#E4DCC8",
   chart=["#0B2545", "#1BA3A3", "#C9A13B", "#4F7CAC", "#8CD3D3", "#E0C47A"], font_head="Cambria")

_t(key="oltin", decor="luxe", dark=True,
   names={"uz": "✨ Qora-oltin", "ru": "✨ Чёрное золото", "en": "✨ Black & gold"},
   desc={"uz": "Nafis rasmiy: ramka, chiziqli grafik, katta raqamlar, vertikal bosqichlar",
         "ru": "Элегантный: рамка, линейный график, крупные цифры",
         "en": "Elegant: frame, line chart, big numbers, vertical steps"},
   v=dict(title="luxe_frame", header="centered", text="big_statement", list="rows", agenda="big_vertical",
          bar="line", pie="doughnut", table="minimal_lines", compare="split_line", stats="inline",
          timeline="vertical", concl="numbered_big"),
   bg="#121212", title_bg="#0A0A0A", primary="#E8C766", accent="#C9A227", accent2="#8C7A4B",
   text="#E9E6DF", muted="#A39E92", card="#1D1D1D", on_title="#E8C766", line="#2E2B25",
   chart=["#C9A227", "#E8C766", "#8C7A4B", "#F3E3B0", "#6B5E3A", "#BFB6A0"],
   font_head="Cambria", font_body="Calibri Light")

_t(key="ilmiy", decor="science", progress=True,
   names={"uz": "🔬 Ilmiy tadqiqot", "ru": "🔬 Научный", "en": "🔬 Research"},
   desc={"uz": "Ma'lumotga boy: raqamlangan bo'limlar, maydon grafik, 100% tasma",
         "ru": "Насыщенный данными: нумерация, диаграмма-область, 100% полоса",
         "en": "Data-rich: numbered sections, area chart, 100% strip"},
   v=dict(title="science_side", header="big_number", text="two_column", list="rows", agenda="nodes",
          bar="area", pie="stacked100", table="banded", compare="table_compare", stats="hero_left",
          timeline="horizontal", concl="check_rows"),
   bg="#FFFFFF", title_bg="#FFFFFF", primary="#0E3B43", accent="#0F9D8A", accent2="#F28C28",
   text="#1E2A2E", muted="#6A7C80", card="#EFF7F6", on_title="#0E3B43",
   chart=["#0F9D8A", "#F28C28", "#0E3B43", "#7FD1C4", "#F7C08A", "#5B8A92"])

_t(key="kosmos", decor="neon", dark=True,
   names={"uz": "🚀 Kosmos neon", "ru": "🚀 Космос неон", "en": "🚀 Neon space"},
   desc={"uz": "Neon kosmik: chap ustunli sarlavha, halqalar, progress-chiziqlar",
         "ru": "Неоновый космос: заголовок в боковой колонке, кольца",
         "en": "Neon space: side-column titles, rings, progress bars"},
   v=dict(title="diagonal", header="side_column", text="highlight_box", list="pills", agenda="left_panel",
          bar="progress", pie="rings", table="banded", compare="cards_vs", stats="circles",
          timeline="vertical", concl="cards"),
   bg="#13072E", title_bg="#0B0420", primary="#FFFFFF", accent="#FF2E97", accent2="#00E5FF",
   text="#E6E1FF", muted="#9C93C9", card="#22114A", on_title="#FFFFFF", line="#2C1A5C",
   chart=["#FF2E97", "#00E5FF", "#B388FF", "#FFD600", "#69F0AE", "#FF8A65"], grad=("#3A0CA3", "#F72585"))

_t(key="jurnal", decor="magazine",
   names={"uz": "📰 Jurnal", "ru": "📰 Журнал", "en": "📰 Editorial"},
   desc={"uz": "Gazeta-jurnal maketi: ulkan tipografiya, ustunli matn, lollipop, waffle",
         "ru": "Журнальная вёрстка: крупная типографика, колонки, вафля",
         "en": "Magazine layout: huge type, columns, lollipop, waffle"},
   v=dict(title="typographic", header="big_number", text="two_column", list="numbered_big", agenda="big_vertical",
          bar="lollipop", pie="waffle", table="banded", compare="table_compare", stats="hero_left",
          timeline="horizontal", concl="banner"),
   bg="#FBFAF7", title_bg="#FBFAF7", primary="#111111", accent="#C1121F", accent2="#1D3557",
   text="#1E1E1E", muted="#6B6B6B", card="#F0EDE6", on_title="#111111", line="#DCD8CF",
   chart=["#C1121F", "#1D3557", "#8D99AE", "#E9C46A", "#457B9D", "#2A2A2A"], font_head="Cambria")

_t(key="startap", decor="startup",
   names={"uz": "💡 Startap pitch", "ru": "💡 Стартап питч", "en": "💡 Startup pitch"},
   desc={"uz": "Pitch-deck uslubi: yorliqli sarlavha, gorizontal ustunlar, pufakchalar",
         "ru": "Стиль питч-дека: ярлык раздела, горизонтальные столбцы, пузырьки",
         "en": "Pitch deck: section pills, horizontal bars, bubbles"},
   v=dict(title="bottom_band", header="pill", text="big_statement", list="icon_columns", agenda="nodes",
          bar="bar_h", pie="bubbles", table="card_rows", compare="colored_halves", stats="circles",
          timeline="chevrons", concl="cards"),
   bg="#FFFFFF", title_bg="#FFF4EC", primary="#0F1B3D", accent="#FF6B2C", accent2="#2D5BFF",
   text="#1C2440", muted="#6B7390", card="#F4F6FB", on_title="#0F1B3D",
   chart=["#FF6B2C", "#2D5BFF", "#14C38E", "#FFB020", "#8E7CFF", "#0F1B3D"])

_t(key="okean", decor="ocean",
   names={"uz": "🌊 Okean", "ru": "🌊 Океан", "en": "🌊 Ocean"},
   desc={"uz": "Tinch moviy: chap ustunli sarlavha, maydon grafik, doira, zinapoya",
         "ru": "Спокойный синий: боковой заголовок, область, круговая, ступени",
         "en": "Calm blue: side titles, area chart, pie, stairs"},
   v=dict(title="circle_focus", header="side_column", text="split_panel", list="cards", agenda="tabs",
          bar="area", pie="pie", table="header_col", compare="table_compare", stats="cards",
          timeline="stairs", concl="check_rows"),
   bg="#F4F9FC", title_bg="#E8F4FA", primary="#0B3C5D", accent="#1D9BD1", accent2="#26C6A6",
   text="#1D3346", muted="#5E7A8C", card="#FFFFFF", on_title="#0B3C5D", line="#D6E6EF",
   chart=["#0B3C5D", "#1D9BD1", "#26C6A6", "#8FD3F0", "#F4A259", "#5E7A8C"])

_t(key="doska", decor="chalk", dark=True,
   names={"uz": "🧑‍🏫 Sinf doskasi", "ru": "🧑‍🏫 Школьная доска", "en": "🧑‍🏫 Chalkboard"},
   desc={"uz": "Bo'r va doska: yog'och ramka, progress-chiziqlar, halqalar. Maktab va pedagogika",
         "ru": "Мел и доска: деревянная рамка, прогресс-полосы, кольца",
         "en": "Chalk and board: wooden frame, progress bars, rings"},
   v=dict(title="chalk", header="underline", text="quote_card", list="zigzag", agenda="columns",
          bar="progress", pie="rings", table="minimal_lines", compare="cards_vs", stats="inline",
          timeline="vertical", concl="numbered_big"),
   bg="#24382F", title_bg="#1E2F27", primary="#F5F1E3", accent="#F2C94C", accent2="#7FD1AE",
   text="#E9E4D4", muted="#A9B8AE", card="#2E463B", on_title="#F5F1E3", line="#3B5548",
   chart=["#F2C94C", "#7FD1AE", "#F28B82", "#9CC3FF", "#F5F1E3", "#C3A6FF"])

THEME_ORDER = list(THEMES.keys())
DEFAULT_THEME = "akademik"
LEGACY_PAL_MAP = ["texno", "oltin", "ilmiy", "oltin", "tabiat", "milliy", "tabiat", "korporativ", "gradient", "oltin"]


def get_theme(key: str | None) -> Theme:
    if key == "random":
        return THEMES[random.choice(THEME_ORDER)]
    return THEMES.get(key or DEFAULT_THEME, THEMES[DEFAULT_THEME])


def preview_path(key: str) -> str:
    return os.path.join(ASSETS_DIR, f"{key}.jpg")


# ═══════════════════════════════════════════════════════════════════
# PAST DARAJALI YORDAMCHILAR
# ═══════════════════════════════════════════════════════════════════
def _clean(t) -> str:
    t = str(t or "")
    t = re.sub(r"[*#`~_]{1,3}", "", t)
    t = t.replace("\\n", "\n").replace("\r", "")
    return re.sub(r"[ \t]{2,}", " ", t).strip()


def _rgb(th: Theme, color) -> RGBColor:
    if isinstance(color, RGBColor):
        return color
    if isinstance(color, str) and color.startswith("#"):
        return RGBColor(*_hex(color))
    return th.c(color)


def _hx(th, name_or_hex: str) -> str:
    return name_or_hex if name_or_hex.startswith("#") else getattr(th, name_or_hex)


def _onc(th, color) -> str:
    """Rangli fon ustidagi matn uchun qora yoki oq."""
    r, g, b = _hex(_hx(th, color))
    return "#111111" if (0.299 * r + 0.587 * g + 0.114 * b) > 170 else "#FFFFFF"


def _alpha_fill(fill, opacity: float):
    try:
        sf = fill._xPr.find(qn("a:solidFill"))
        clr = sf[0]
        for old in clr.findall(qn("a:alpha")):
            clr.remove(old)
        etree.SubElement(clr, qn("a:alpha")).set("val", str(int(max(0, min(1, opacity)) * 100000)))
    except Exception:
        pass


def _box(sl, th, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=None, line_w=1.0,
         opacity=None, radius=None, rot=0):
    sh = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(max(w, 0.01)), Inches(max(h, 0.01)))
    if color is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(th, color)
        if opacity is not None:
            _alpha_fill(sh.fill, opacity)
    if line:
        sh.line.color.rgb = _rgb(th, line)
        sh.line.width = Pt(line_w)
    else:
        sh.line.fill.background()
    if radius is not None:
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
    if rot:
        sh.rotation = rot
    sh.shadow.inherit = False
    sh.text_frame.text = ""
    return sh


def _grad(sl, th, x, y, w, h, c1, c2, angle=0, shape=MSO_SHAPE.RECTANGLE):
    sh = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.gradient()
    sh.fill.gradient_angle = angle
    st = sh.fill.gradient_stops
    st[0].color.rgb = _rgb(th, c1); st[0].position = 0
    st[1].color.rgb = _rgb(th, c2); st[1].position = 1.0
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def _line(sl, th, x1, y1, x2, y2, color, w=1.5):
    ln = sl.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = _rgb(th, color)
    ln.line.width = Pt(w)
    return ln


def _est_lines(text: str, w_in: float, size: float) -> int:
    cpl = max(6, int(w_in * 72 / (size * 0.5)))
    return sum(max(1, math.ceil(len(p) / cpl)) for p in (text or " ").split("\n"))


def _fit(text: str, w: float, h: float, mx: float, mn: float = 11, spacing: float = 1.18) -> float:
    s = mx
    while s > mn:
        if _est_lines(text, w, s) * s * spacing / 72 <= h:
            return s
        s -= 1
    return mn


def _txt(sl, th, x, y, w, h, text, size=18, color="text", bold=False, font=None,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False, spacing=1.12,
         fit=False, min_size=11, caps=False, spc=None, lead_bold=False):
    text = _clean(text)
    if caps:
        text = text.upper()
    if fit:
        size = _fit(text, w - 0.1, h - 0.05, size, min_size, spacing + 0.06)
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(max(w, 0.1)), Inches(max(h, 0.1)))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, par in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        # «Atama — izoh» ko'rinishidagi bandlarda atama qalin ajratiladi (ro'yxatni o'qish osonlashadi)
        m = re.match(r"^(.{2,70}?)(\s[—–-]\s|:\s)(.+)$", par) if lead_bold and not bold else None
        qismlar = [(m.group(1) + m.group(2), True), (m.group(3), False)] if m else [(par, bold)]
        for matn_, qalin in qismlar:
            r = p.add_run()
            r.text = matn_
            f = r.font
            f.size = Pt(size); f.bold = qalin; f.italic = italic
            f.name = font or th.font_body
            f.color.rgb = _rgb(th, color)
            if spc is not None:
                r._r.get_or_add_rPr().set("spc", str(int(spc * 100)))
    return tb


def _bg(sl, th, color):
    f = sl.background.fill
    f.solid()
    f.fore_color.rgb = _rgb(th, color)


def _items(matn: str) -> list[str]:
    out = [x.strip(" •-–—\t") for x in _clean(matn).split("\n")]
    out = [re.sub(r"^\d{1,2}[.)]\s*", "", x) for x in out if x]
    return [x for x in out if len(x) > 1]


_NUM = re.compile(r"-?\d+(?:[.,]\d+)?")


def _pairs(matn: str) -> list[tuple[str, float]]:
    out = []
    for line in _clean(matn).split("\n"):
        if ":" not in line:
            continue
        k, v = line.rsplit(":", 1)
        m = _NUM.search(v.replace(" ", ""))
        if k.strip() and m:
            try:
                out.append((k.strip()[:40], float(m.group(0).replace(",", "."))))
            except ValueError:
                pass
    return out[:10]


def _sents(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?])\s+", _clean(text).replace("\n", " "))
    return [p.strip() for p in parts if len(p.strip()) > 2]


def _fmt(v: float) -> str:
    return (f"{v:,.0f}" if abs(v) >= 100 or v == int(v) else f"{v:,.1f}").replace(",", " ")


def _col(th, i):
    return th.chart[i % len(th.chart)]


# ═══════════════════════════════════════════════════════════════════
# FON BEZAKLARI (kontent slaydlar)
# ═══════════════════════════════════════════════════════════════════
def _star(sl, th, cx, cy, r, color):
    return _box(sl, th, cx - r, cy - r, 2 * r, 2 * r, color, MSO_SHAPE.STAR_8_POINT)


def _ornament_band(sl, th, y, h=0.42):
    _box(sl, th, 0, y, SW, h, "primary")
    step = 0.62
    for i in range(int(SW / step) + 1):
        cx = i * step + step / 2
        if i % 2 == 0:
            _star(sl, th, cx, y + h / 2, h * 0.36, "accent2")
        else:
            _box(sl, th, cx - h * 0.16, y + h / 2 - h * 0.16, h * 0.32, h * 0.32, "accent", MSO_SHAPE.DIAMOND)


def _dot_grid(sl, th, color, x0=0.25, y0=0.25, nx=27, ny=15, step=0.5, size=0.025):
    for gx in range(nx):
        for gy in range(ny):
            if (gx + gy) % 2 == 0:
                _box(sl, th, x0 + gx * step, y0 + gy * step, size, size, color, MSO_SHAPE.OVAL)


def _blank(prs):
    """Bo'sh maket: standart shablonda 6-maket, foydalanuvchi shablonida eng kam placeholderli maket."""
    lay = getattr(prs, "_ab_blank", None)
    if lay is None:
        lays = list(prs.slide_layouts)
        lay = lays[6] if len(lays) > 6 and lays[6].name.lower().startswith("blank") else \
            min(lays, key=lambda l: len(l.placeholders))
        prs._ab_blank = lay
    return lay


def _decor(sl, th: Theme):
    d = th.decor
    if d == "template":
        return          # foydalanuvchi shablonining o'z foni va bezaklari saqlanadi
    _bg(sl, th, "bg")
    if d == "classic":
        _box(sl, th, 0, 0, SW, 0.09, "primary"); _box(sl, th, 0, 0.09, SW, 0.025, "accent")
    elif d == "dark_geo":
        for i in range(1, 8):
            _box(sl, th, i * 1.9, 0, 0.006, SH, "#1A2438")
        _box(sl, th, SW - 2.2, -1.1, 3.0, 3.0, None, MSO_SHAPE.OVAL, line="accent", line_w=0.75)
    elif d == "organic":
        _box(sl, th, SW - 2.6, SH - 2.2, 3.6, 3.6, "accent", MSO_SHAPE.OVAL, opacity=0.10)
        _box(sl, th, -0.9, -0.9, 2.0, 2.0, "accent2", MSO_SHAPE.OVAL, opacity=0.12)
    elif d == "pastel":
        _box(sl, th, SW - 1.6, -0.8, 2.4, 2.4, "accent", MSO_SHAPE.OVAL, opacity=0.16)
        _box(sl, th, -0.6, SH - 1.0, 1.6, 1.6, "#9B8AFB", MSO_SHAPE.OVAL, opacity=0.16)
    elif d == "ornament":
        _ornament_band(sl, th, 0, 0.30); _box(sl, th, 0, SH - 0.08, SW, 0.08, "accent")
    elif d == "luxe":
        _box(sl, th, 0.28, 0.28, SW - 0.56, SH - 0.56, None, line="accent", line_w=0.75)
    elif d == "science":
        _dot_grid(sl, th, "#D5E6E3"); _box(sl, th, 0, 0, 0.12, SH, "accent")
    elif d == "neon":
        _dot_grid(sl, th, "#3A2470", x0=4.2, step=0.75, nx=12, ny=10, size=0.03)
        _box(sl, th, SW - 1.2, SH - 1.2, 2.4, 2.4, None, MSO_SHAPE.OVAL, line="accent2", line_w=0.75)
    elif d == "magazine":
        _box(sl, th, 0.75, 0.25, SW - 1.5, 0.02, "primary")
        _box(sl, th, 0.75, 0.31, SW - 1.5, 0.006, "primary")
    elif d == "startup":
        _box(sl, th, SW - 0.9, -0.9, 1.8, 1.8, "accent", MSO_SHAPE.OVAL)
        _box(sl, th, SW - 0.35, 0.75, 0.5, 0.5, "accent2", MSO_SHAPE.OVAL)
    elif d == "ocean":
        _box(sl, th, 2.0, SH - 0.9, 7.5, 2.2, "accent", MSO_SHAPE.OVAL, opacity=0.10)
        _box(sl, th, 6.5, SH - 0.7, 9.0, 2.2, "accent2", MSO_SHAPE.OVAL, opacity=0.10)
    elif d == "chalk":
        _box(sl, th, 0, 0, SW, SH, None, line="#7A5230", line_w=14)
        for (x, w) in ((1.2, 0.9), (10.4, 0.6)):
            _box(sl, th, x, 6.98, w, 0.08, "#F5F1E3", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5, opacity=0.8)


# ═══════════════════════════════════════════════════════════════════
# SARLAVHA VARIANTLARI → kontent maydoni (x, y, w, h)
# ═══════════════════════════════════════════════════════════════════
def _header(sl, th: Theme, title: str, num: int, izoh: str = "") -> tuple:
    v = th.v["header"]
    title = _clean(title) or "—"
    hf = th.font_head
    L = 0.75
    W = SW - 2 * L
    if v == "underline":
        _txt(sl, th, L, 0.42, W, 0.95, title, 30, "primary", True, hf, fit=True, min_size=20, anchor=MSO_ANCHOR.BOTTOM)
        _box(sl, th, L + 0.02, 1.42, 1.2, 0.05, "accent")
        return (L, 1.85, W, SH - 2.6)
    if v == "square_tag":
        _box(sl, th, L, 0.68, 0.22, 0.22, "accent")
        _txt(sl, th, L + 0.35, 0.38, W - 0.35, 0.85, title, 32, "primary", True, hf, fit=True, min_size=20,
             anchor=MSO_ANCHOR.MIDDLE)
        return (L, 1.7, W, SH - 2.45)
    if v == "top_band":
        if th.grad:
            _grad(sl, th, 0, 0, SW, 1.35, th.grad[0], th.grad[1], 0)
        else:
            _box(sl, th, 0, 0, SW, 1.35, "primary")
            _box(sl, th, 0, 1.35, SW, 0.06, "accent2")
        _txt(sl, th, L, 0.2, W - 1.2, 0.95, title, 28, "#FFFFFF", True, hf, fit=True, min_size=18,
             anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, SW - 1.7, 0.3, 1.0, 0.75, f"{num:02d}", 26, "#FFFFFF", True, hf, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.MIDDLE)
        return (L, 1.8, W, SH - 2.55)
    if v == "bracket":
        _txt(sl, th, L, 0.38, 1.5, 0.4, f"/{num:02d}", 13, "accent", True, spc=2)
        _box(sl, th, L, 0.82, 0.08, 0.6, "accent")
        _txt(sl, th, L + 0.25, 0.7, W - 0.3, 0.85, title, 30, "primary", True, hf, fit=True, min_size=20,
             anchor=MSO_ANCHOR.MIDDLE)
        return (L, 1.9, W, SH - 2.65)
    if v == "badge":
        _box(sl, th, L, 0.45, 0.72, 0.72, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
        _txt(sl, th, L, 0.45, 0.72, 0.72, str(num), 22, _onc(th, "accent"), True, hf, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, L + 0.95, 0.38, W - 0.95, 0.85, title, 29, "primary", True, hf, fit=True, min_size=20,
             anchor=MSO_ANCHOR.MIDDLE)
        return (L, 1.6, W, SH - 2.35)
    if v == "centered":
        top = 0.5
        _txt(sl, th, L, top, W, 0.85, title, 30, "primary", th.decor != "luxe", hf, fit=True, min_size=20,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spc=1 if th.decor == "luxe" else None)
        _box(sl, th, SW / 2 - 1.0, top + 0.95, 2.0, 0.025, "accent2" if th.decor == "ornament" else "accent")
        if th.decor == "ornament":
            _star(sl, th, SW / 2, top + 0.96, 0.13, "accent")
        else:
            _box(sl, th, SW / 2 - 0.08, top + 0.88, 0.16, 0.16, "accent", MSO_SHAPE.DIAMOND)
        return (L, top + 1.35, W, SH - top - 2.1)
    if v == "big_number":
        _txt(sl, th, L - 0.05, 0.42, 1.3, 0.95, f"{num:02d}", 40, "accent", True, hf, anchor=MSO_ANCHOR.MIDDLE)
        _box(sl, th, L + 1.25, 0.6, 0.025, 0.62, "accent2" if th.key == "ilmiy" else "primary")
        _txt(sl, th, L + 1.45, 0.42, W - 1.45, 0.95, title, 28, "primary", True, hf, fit=True, min_size=18,
             anchor=MSO_ANCHOR.MIDDLE)
        if th.decor == "magazine":
            _box(sl, th, L, 1.47, W, 0.012, "primary")
        return (L, 1.75, W, SH - 2.5)
    if v == "pill":
        _box(sl, th, L, 0.4, 1.25, 0.4, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        _txt(sl, th, L, 0.4, 1.25, 0.4, f"● {num:02d}", 12, _onc(th, "accent"), True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, L, 0.85, W - 1.5, 0.85, title, 30, "primary", True, hf, fit=True, min_size=20,
             anchor=MSO_ANCHOR.MIDDLE)
        return (L, 1.95, W, SH - 2.7)
    if v == "side_column":
        cw = 3.7
        if th.grad:
            _grad(sl, th, 0, 0, cw, SH, th.grad[0], th.grad[1], 90)
        else:
            _box(sl, th, 0, 0, cw, SH, "primary")
        _box(sl, th, cw, 0, 0.06, SH, "accent")
        _txt(sl, th, 0.45, 0.55, 2.0, 1.0, f"{num:02d}", 44, "#FFFFFF", True, hf)
        _box(sl, th, 0.5, 1.6, 0.8, 0.05, "accent2" if th.dark else "accent")
        if izoh:
            _txt(sl, th, 0.45, 1.85, cw - 0.8, 2.0, title, 28, "#FFFFFF", True, hf, fit=True, min_size=16, spacing=1.08)
            _box(sl, th, 0.5, 4.0, cw - 1.0, 0.012, "#FFFFFF", opacity=0.5)
            _txt(sl, th, 0.45, 4.15, cw - 0.8, SH - 5.0, izoh, 16, "#FFFFFF", fit=True, min_size=11, spacing=1.2)
        else:
            _txt(sl, th, 0.45, 1.85, cw - 0.8, 4.3, title, 28, "#FFFFFF", True, hf, fit=True, min_size=16, spacing=1.08)
        return (cw + 0.55, 0.65, SW - cw - 1.15, SH - 1.35)
    _txt(sl, th, L, 0.42, W, 0.95, title, 30, "primary", True, hf, fit=True, min_size=20)
    return (L, 1.7, W, SH - 2.45)


def _footer(sl, th: Theme, num: int, total: int, topic: str):
    if th.v["header"] == "side_column":
        _txt(sl, th, 0.45, SH - 0.65, 2.5, 0.35, f"{num} / {total}", 11, "#FFFFFF")
        return
    y = SH - 0.5
    short = _clean(topic)
    short = short if len(short) <= 70 else short[:67] + "…"
    _txt(sl, th, 0.75, y, 8.5, 0.32, short, 10, "muted")
    _txt(sl, th, SW - 1.95, y, 1.2, 0.32, f"{num} / {total}", 10, "muted", align=PP_ALIGN.RIGHT)
    if th.progress:
        ratio = min(1.0, num / max(1, total))
        _box(sl, th, 0, SH - 0.06, SW, 0.06, "card")
        if th.grad:
            _grad(sl, th, 0, SH - 0.06, SW * ratio, 0.06, th.grad[0], th.grad[1], 0)
        else:
            _box(sl, th, 0, SH - 0.06, SW * ratio, 0.06, "accent")


# Izoh (sarlavha ostidagi tushuntirish) ko'rsatilmaydigan slayd turlari: matnning o'zi tushuntirish
_NO_IZOH = {"matn", "reja"}


def _lead(sl, th, rect, izoh):
    """Sarlavha ostida slaydning asosiy fikrini ko'rsatadigan 1-2 gapli izoh. Kontent maydonini qaytaradi."""
    x, y, w, h = rect
    lh = 0.78
    _box(sl, th, x, y - 0.05, 0.07, lh - 0.1, "accent")
    _txt(sl, th, x + 0.25, y - 0.1, w - 0.25, lh, izoh, 20, "text", italic=th.decor in ("luxe", "magazine", "classic"),
         anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=12, spacing=1.12)
    return (x, y + lh + 0.15, w, h - lh - 0.15)


def _start(prs, th, sd, num):
    sl = prs.slides.add_slide(_blank(prs))
    _decor(sl, th)
    izoh = _clean(sd.get("izoh", ""))
    if sd.get("tipo") in _NO_IZOH or len(izoh) < 8:
        izoh = ""
    izoh = izoh if len(izoh) <= 300 else izoh[:297].rsplit(" ", 1)[0] + "…"
    side = th.v["header"] == "side_column"
    rect = _header(sl, th, sd.get("sarlavha"), num, izoh if side else "")
    if izoh and not side:
        rect = _lead(sl, th, rect, izoh)
    return sl, rect


def _card(sl, th, x, y, w, h, color="card", line=None):
    rounded = th.decor not in ("swiss", "luxe", "classic", "magazine")
    return _box(sl, th, x, y, w, h, color, MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                radius=0.18 if th.decor in ("pastel", "startup") else 0.07,
                line=line or ("accent2" if th.decor == "luxe" else ("line" if th.decor == "ocean" else None)),
                line_w=0.75)


def _num_circle(sl, th, x, y, d, n, color=None, shape=MSO_SHAPE.OVAL):
    color = color or th.accent
    _box(sl, th, x, y, d, d, color, shape, radius=0.3)
    _txt(sl, th, x, y, d, d, str(n), d * 30, _onc(th, color), True, th.font_head,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ═══════════════════════════════════════════════════════════════════
# MATN — 5 variant
# ═══════════════════════════════════════════════════════════════════
def slide_text(prs, th, sd, num, total, topic):
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    s = _sents(sd.get("matn", "")) or [_clean(sd.get("matn", "")) or "—"]
    lead, rest = s[0], " ".join(s[1:])
    v = th.v["text"]
    if v == "quote_card" and rest:
        cw = min(4.3, w * 0.38)
        _card(sl, th, x, y, cw, h)
        _box(sl, th, x, y, 0.09, h, "accent")
        _txt(sl, th, x + 0.3, y + 0.1, 0.9, 0.9, "“", 66, "accent", True, th.font_head)
        _txt(sl, th, x + 0.35, y + 1.0, cw - 0.6, h - 1.2, lead, 24, "primary", True, th.font_head, fit=True, min_size=14)
        _txt(sl, th, x + cw + 0.5, y, w - cw - 0.5, h, rest, 22, "text", fit=True, min_size=13, spacing=1.3,
             anchor=MSO_ANCHOR.MIDDLE)
    elif v == "big_statement" and rest:
        _txt(sl, th, x + 0.6, y, w - 1.2, h * 0.58, f"«{lead}»", 36, "primary", True, th.font_head,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=18, spacing=1.1)
        _box(sl, th, x + w / 2 - 0.5, y + h * 0.62, 1.0, 0.04, "accent")
        _txt(sl, th, x + 1.0, y + h * 0.68, w - 2.0, h * 0.32, rest, 18, "muted", align=PP_ALIGN.CENTER,
             fit=True, min_size=11, spacing=1.25)
    elif v == "two_column":
        words = " ".join(s).split(" ")
        half = len(words) // 2
        cut = half
        for i in range(half, len(words)):
            if words[i - 1].endswith((".", "!", "?")):
                cut = i
                break
        a, b = " ".join(words[:cut]), " ".join(words[cut:]) or " "
        cw = (w - 0.6) / 2
        fs = min(_fit(a, cw, h, 25, 12, 1.36), _fit(b, cw, h, 25, 12, 1.36))
        _txt(sl, th, x, y, cw, h, a, fs, "text", spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
        _box(sl, th, x + cw + 0.29, y + 0.1, 0.015, h - 0.2, "line")
        _txt(sl, th, x + cw + 0.6, y, cw, h, b, fs, "text", spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
    elif v == "split_panel" and rest:
        pw = w * 0.42
        _box(sl, th, x, y, pw, h, "card" if th.dark else "primary")
        _txt(sl, th, x + 0.4, y + 0.4, pw - 0.8, h - 0.8, lead, 24, "accent" if th.dark else "#FFFFFF", True,
             th.font_head, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=14)
        _txt(sl, th, x + pw + 0.5, y, w - pw - 0.5, h, rest, 21, "text", fit=True, min_size=12, spacing=1.32,
             anchor=MSO_ANCHOR.MIDDLE)
    elif v == "highlight_box" and rest:
        body, key = " ".join(s[:-1]), s[-1]
        bh = min(1.6, h * 0.33)
        _txt(sl, th, x, y, w, h - bh - 0.35, body, 26, "text", fit=True, min_size=12, spacing=1.3,
             anchor=MSO_ANCHOR.MIDDLE)
        _box(sl, th, x, y + h - bh, w, bh, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12, opacity=0.14)
        _box(sl, th, x, y + h - bh, 0.1, bh, "accent")
        _txt(sl, th, x + 0.35, y + h - bh, 0.7, bh, "💡", 28, "accent", anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, x + 1.1, y + h - bh + 0.1, w - 1.4, bh - 0.2, key, 20, "primary", True, anchor=MSO_ANCHOR.MIDDLE,
             fit=True, min_size=12)
    else:
        _txt(sl, th, x, y, w, h, " ".join(s), 24, "text", fit=True, min_size=13, spacing=1.32, anchor=MSO_ANCHOR.MIDDLE)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# RO'YXAT — 7 variant
# ═══════════════════════════════════════════════════════════════════
def _list_render(sl, th, rect, items, v, check=False):
    x, y, w, h = rect
    n = len(items)
    mark = "✓" if check else None
    if v == "cards" and 2 <= n <= 6:
        cols = 3 if n in (3, 5, 6) else 2
        rows = math.ceil(n / cols)
        g = 0.3
        cw, ch = (w - g * (cols - 1)) / cols, (h - g * (rows - 1)) / rows
        fs = min(_fit(t, cw - 0.6, ch - 1.15, 22, 11) for t in items)
        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            cx, cy = x + c * (cw + g), y + r * (ch + g)
            _card(sl, th, cx, cy, cw, ch)
            col = _col(th, i) if th.decor in ("pastel", "gradient", "startup", "neon") else th.accent
            _num_circle(sl, th, cx + 0.25, cy + 0.25, 0.5, mark or i + 1, col)
            _txt(sl, th, cx + 0.25, cy + 0.9, cw - 0.5, ch - 1.05, it, fs, "text", spacing=1.18, lead_bold=True)
        return
    if v == "icon_columns" and 2 <= n <= 5:
        cw = w / n
        fs = min(_fit(t, cw - 0.4, h - 1.8, 20, 12) for t in items)
        for i, it in enumerate(items):
            cx = x + i * cw
            d = 1.05
            col = _col(th, i)
            _box(sl, th, cx + cw / 2 - d / 2, y + 0.15, d, d, col, MSO_SHAPE.OVAL, opacity=0.15)
            _num_circle(sl, th, cx + cw / 2 - 0.35, y + 0.32, 0.7, mark or i + 1, col)
            _txt(sl, th, cx + 0.2, y + 1.55, cw - 0.4, h - 1.6, it, fs, "text", align=PP_ALIGN.CENTER, spacing=1.2, lead_bold=True)
            if i < n - 1:
                _box(sl, th, cx + cw - 0.01, y + 1.6, 0.012, h - 1.8, "line")
        return
    if v == "pills":
        g = 0.18
        rh = min(1.0, (h - g * (n - 1)) / n)
        fs = min(_fit(t, w - 1.6, rh - 0.1, 20, 11) for t in items)
        oy = y + (h - (rh * n + g * (n - 1))) / 2
        for i, it in enumerate(items):
            ry = oy + i * (rh + g)
            col = _col(th, i)
            _box(sl, th, x, ry, w, rh, "card", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
            _box(sl, th, x, ry, rh * 1.25, rh, col, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
            _txt(sl, th, x, ry, rh * 1.25, rh, mark or f"{i + 1:02d}", min(20, rh * 26), _onc(th, col), True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, x + rh * 1.25 + 0.3, ry, w - rh * 1.25 - 0.5, rh, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE, lead_bold=True)
        return
    if v == "zigzag" and n >= 2:
        cx = x + w / 2
        _box(sl, th, cx - 0.015, y, 0.03, h, "line")
        rh = h / n
        cw = w / 2 - 0.6
        fs = min(_fit(t, cw - 0.1, rh * 1.6, 19, 11) for t in items)
        for i, it in enumerate(items):
            ry = y + i * rh
            _num_circle(sl, th, cx - 0.25, ry + rh / 2 - 0.25, 0.5, mark or i + 1, _col(th, i))
            if i % 2 == 0:
                _txt(sl, th, x, ry - rh * 0.3, cw, rh * 1.6, it, fs, "text", align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, lead_bold=True)
            else:
                _txt(sl, th, cx + 0.6, ry - rh * 0.3, cw, rh * 1.6, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE, lead_bold=True)
        return
    if v == "numbered_big":
        cols = 2 if n >= 4 else 1
        per = math.ceil(n / cols)
        cw = (w - 0.5 * (cols - 1)) / cols
        rh = h / per
        fs = min(_fit(t, cw - 1.5, rh - 0.1, 20, 11) for t in items)
        for i, it in enumerate(items):
            c, r = divmod(i, per)
            cx, ry = x + c * (cw + 0.5), y + r * rh
            _txt(sl, th, cx, ry, 1.3, rh, mark or f"{i + 1:02d}", min(44, rh * 40), "accent", True, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, cx + 1.4, ry, cw - 1.4, rh, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE, lead_bold=True)
            _box(sl, th, cx + 1.4, ry + rh - 0.02, cw - 1.4, 0.012, "line")
        return
    # rows
    rh = h / max(n, 1)
    ms = min(0.46, rh * 0.7)
    fs = min(_fit(t, w - 1.0, rh - 0.05, 23, 12) for t in items)
    for i, it in enumerate(items):
        ry = y + i * rh
        if check:
            _num_circle(sl, th, x, ry + (rh - ms) / 2, ms, "✓", th.accent)
        elif th.decor in ("classic", "luxe", "swiss", "magazine"):
            _txt(sl, th, x - 0.05, ry, ms + 0.4, rh, f"{i + 1:02d}", ms * 40, "accent", True, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE)
        else:
            _num_circle(sl, th, x, ry + (rh - ms) / 2, ms, i + 1, th.accent)
        _txt(sl, th, x + ms + 0.35, ry, w - ms - 0.35, rh, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE, lead_bold=True)
        if i < n - 1:
            _box(sl, th, x + ms + 0.35, ry + rh - 0.01, w - ms - 0.35, 0.012, "line")


def slide_list(prs, th, sd, num, total, topic):
    sl, rect = _start(prs, th, sd, num)
    items = _items(sd.get("matn", ""))[:8] or [_clean(sd.get("matn", "")) or "—"]
    v = th.v["list"]
    if len(items) > 6 and v in ("cards", "zigzag"):
        v = "rows"
    if v == "icon_columns" and len(items) > 5:
        v = "rows"
    _list_render(sl, th, rect, items, v)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# REJA — 5 variant
# ═══════════════════════════════════════════════════════════════════
def slide_agenda(prs, th, sd, num, total, topic):
    items = _items(sd.get("matn", ""))[:6] or ["—"]
    n = len(items)
    v = th.v["agenda"]
    if v == "left_panel":
        sl = prs.slides.add_slide(_blank(prs))
        _bg(sl, th, "bg")
        pw = 4.8
        if th.grad:
            _grad(sl, th, 0, 0, pw, SH, th.grad[0], th.grad[1], 60)
        else:
            _box(sl, th, 0, 0, pw, SH, "primary")
        _box(sl, th, pw - 1.2, SH - 1.6, 2.4, 2.4, "accent", MSO_SHAPE.OVAL, opacity=0.5)
        _txt(sl, th, 0.7, 1.2, pw - 1.2, 2.5, _clean(sd.get("sarlavha")) or "Reja", 46, "#FFFFFF", True,
             th.font_head, fit=True, min_size=26, anchor=MSO_ANCHOR.BOTTOM)
        _box(sl, th, 0.75, 3.9, 1.0, 0.06, "accent2")
        x, w = pw + 0.7, SW - pw - 1.4
        rh = min(1.25, (SH - 1.6) / n)
        oy = (SH - rh * n) / 2
        fs = min(_fit(t, w - 1.4, rh - 0.1, 22, 12) for t in items)
        for i, it in enumerate(items):
            ry = oy + i * rh
            _txt(sl, th, x, ry, 1.2, rh, f"{i + 1:02d}", 34, "accent", True, th.font_head, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, x + 1.3, ry, w - 1.3, rh, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE)
            if i < n - 1:
                _box(sl, th, x, ry + rh - 0.01, w, 0.012, "line")
        _txt(sl, th, SW - 1.95, SH - 0.5, 1.2, 0.32, f"{num} / {total}", 10, "muted", align=PP_ALIGN.RIGHT)
        return sl
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    if v == "columns" and n <= 4:
        g = 0.35
        cw = (w - g * (n - 1)) / n
        fs = min(_fit(t, cw - 0.5, h - 2.0, 23, 13) for t in items)
        for i, it in enumerate(items):
            cx = x + i * (cw + g)
            _card(sl, th, cx, y + 0.1, cw, h - 0.2)
            _txt(sl, th, cx + 0.3, y + 0.3, cw - 0.6, 1.1, f"{i + 1:02d}", 48, "accent", True, th.font_head)
            _box(sl, th, cx + 0.35, y + 1.45, 0.7, 0.04, "accent2")
            _txt(sl, th, cx + 0.3, y + 1.7, cw - 0.6, h - 2.0, it, fs, "text", spacing=1.2)
    elif v == "nodes" and n <= 5:
        seg = w / n
        ly = y + h * 0.22
        _box(sl, th, x + seg / 2, ly + 0.5, w - seg, 0.05, "line")
        fs = min(_fit(t, seg - 0.4, h * 0.6, 20, 12) for t in items)
        for i, it in enumerate(items):
            cx = x + i * seg + seg / 2
            col = _col(th, i)
            _box(sl, th, cx - 0.5, ly, 1.0, 1.0, "bg", MSO_SHAPE.OVAL, line=col, line_w=3)
            _txt(sl, th, cx - 0.5, ly, 1.0, 1.0, f"{i + 1:02d}", 24, col, True, th.font_head,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, cx - seg / 2 + 0.2, ly + 1.25, seg - 0.4, h * 0.6, it, fs, "text", align=PP_ALIGN.CENTER,
                 spacing=1.2)
    elif v == "tabs":
        g = 0.2
        rh = min(1.1, (h - g * (n - 1)) / n)
        oy = y + (h - (rh * n + g * (n - 1))) / 2
        fs = min(_fit(t, w - 2.5, rh - 0.1, 22, 12) for t in items)
        for i, it in enumerate(items):
            ry = oy + i * (rh + g)
            col = _col(th, i)
            ind = i * 0.35
            _box(sl, th, x + ind, ry, w - ind, rh, col, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
            _txt(sl, th, x + ind + 0.25, ry, 1.0, rh, f"{i + 1:02d}", 26, _onc(th, col), True, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, x + ind + 1.3, ry, w - ind - 1.6, rh, it, fs, _onc(th, col), anchor=MSO_ANCHOR.MIDDLE)
    else:  # big_vertical
        rh = h / n
        fs = min(_fit(t, w - 2.0, rh - 0.1, 24, 12) for t in items)
        for i, it in enumerate(items):
            ry = y + i * rh
            _txt(sl, th, x, ry, 1.8, rh, f"{i + 1:02d}", min(60, rh * 52), "accent", False, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, x + 1.9, ry, w - 1.9, rh, it, fs, "text", anchor=MSO_ANCHOR.MIDDLE)
            _box(sl, th, x, ry + rh - 0.01, w, 0.015, "line")
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# USTUN / STATISTIKA DIAGRAMMASI — 6 variant
# ═══════════════════════════════════════════════════════════════════
def _style_axis(th, ax, size=12):
    ax.tick_labels.font.size = Pt(size)
    ax.tick_labels.font.color.rgb = th.c("muted")
    ax.tick_labels.font.name = th.font_body
    ax.format.line.color.rgb = RGBColor(*_hex(th.line))


def _native_chart(sl, th, ctype, pairs, unit, x, y, w, h):
    cd = CategoryChartData()
    cd.categories = [p[0] for p in pairs]
    cd.add_series(unit or " ", [p[1] for p in pairs])
    gf = sl.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ch.font.name = th.font_body
    ch.font.size = Pt(12)
    ch.font.color.rgb = th.c("muted")
    return ch


def _chart_side_panel(sl, th, pairs, unit, x, y, w, h, style="card"):
    vals = [p[1] for p in pairs]
    imax = vals.index(max(vals))
    if style == "card":
        _card(sl, th, x, y, w, h)
    else:
        _box(sl, th, x, y + 0.2, 0.02, h - 0.4, "line")
    px = x + 0.3
    _txt(sl, th, px, y + 0.3, w - 0.5, 0.35, "MAX", 11, "muted", True, spc=2)
    _txt(sl, th, px, y + 0.62, w - 0.5, 0.9, _fmt(max(vals)), 36, "accent", True, th.font_head, fit=True, min_size=18)
    _txt(sl, th, px, y + 1.45, w - 0.5, 0.5, f"{pairs[imax][0]} {unit}".strip(), 13, "text", fit=True)
    first, last = vals[0], vals[-1]
    if first:
        pct = (last - first) / abs(first) * 100
        _box(sl, th, px, y + 2.15, w - 0.6, 0.012, "line")
        _txt(sl, th, px, y + 2.3, w - 0.5, 0.35, "Δ", 13, "muted", True)
        _txt(sl, th, px, y + 2.6, w - 0.5, 0.8, f"{'+' if pct >= 0 else ''}{pct:.0f}%", 32,
             "accent2", True, th.font_head, fit=True, min_size=18)
        _txt(sl, th, px, y + 3.4, w - 0.5, 0.5, f"{pairs[0][0]} → {pairs[-1][0]}", 12, "muted", fit=True)


def slide_bar(prs, th, sd, num, total, topic):
    pairs = _pairs(sd.get("matn", ""))
    if len(pairs) < 2:
        return slide_list(prs, th, sd, num, total, topic)
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    unit = _clean(sd.get("birligi", ""))
    v = th.v["bar"]
    vals = [p[1] for p in pairs]
    vmax = max(max(vals), 1e-9)
    imax = vals.index(max(vals))
    if v in ("column", "bar_h", "line", "area"):
        pw = 3.0 if w > 9 else 2.6
        cw = w - pw - 0.4
        ctype = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar_h": XL_CHART_TYPE.BAR_CLUSTERED,
                 "line": XL_CHART_TYPE.LINE_MARKERS, "area": XL_CHART_TYPE.AREA}[v]
        ch = _native_chart(sl, th, ctype, pairs, unit, x, y, cw, h)
        plot = ch.plots[0]
        ser = plot.series[0]
        va = ch.value_axis
        va.has_major_gridlines = v != "bar_h"
        if va.has_major_gridlines:
            va.major_gridlines.format.line.color.rgb = RGBColor(*_hex(th.line))
        va.format.line.fill.background()
        _style_axis(th, va)
        _style_axis(th, ch.category_axis, 12)
        if v in ("column", "bar_h"):
            plot.gap_width = 55 if v == "column" else 45
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(13); dl.font.bold = True; dl.font.color.rgb = th.c("text")
            dl.number_format = "#,##0.##"; dl.number_format_is_linked = False
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            base = "accent2" if (th.decor == "classic" or th.dark) else "primary"
            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = th.c(base)
            pt = ser.points[imax]
            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = th.c("accent")
            if v == "bar_h":
                ch.category_axis.reverse_order = True
                va.visible = False
        elif v == "line":
            ser.smooth = False
            ser.format.line.color.rgb = th.c("accent"); ser.format.line.width = Pt(3.5)
            ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 11
            ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = th.c("bg")
            ser.marker.format.line.color.rgb = th.c("accent"); ser.marker.format.line.width = Pt(2.5)
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(13); dl.font.bold = True; dl.font.color.rgb = th.c("primary")
            dl.number_format = "#,##0.##"; dl.number_format_is_linked = False
            dl.position = XL_LABEL_POSITION.ABOVE
        else:  # area
            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = th.c("accent")
            _alpha_fill(ser.format.fill, 0.55)
            ser.format.line.color.rgb = th.c("primary"); ser.format.line.width = Pt(2.5)
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(12); dl.font.bold = True; dl.font.color.rgb = th.c("primary")
            dl.number_format = "#,##0.##"; dl.number_format_is_linked = False
        _chart_side_panel(sl, th, pairs, unit, x + cw + 0.4, y, pw, h,
                          "line" if th.decor in ("swiss", "luxe", "magazine") else "card")
    elif v == "progress":
        n = len(pairs)
        rh = min(0.95, h / n)
        oy = y + (h - rh * n) / 2
        lw = min(3.2, w * 0.28)
        tw = w - lw - 1.7
        for i, (k, val) in enumerate(pairs):
            ry = oy + i * rh
            col = th.accent if i == imax else (_col(th, i + 1) if th.dark else th.accent2)
            _txt(sl, th, x, ry, lw, rh, k, 16, "text", True, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=11)
            bh = min(0.36, rh * 0.42)
            by = ry + (rh - bh) / 2
            _box(sl, th, x + lw + 0.2, by, tw, bh, "card", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
            _box(sl, th, x + lw + 0.2, by, max(bh, tw * max(val, 0) / vmax), bh, col, MSO_SHAPE.ROUNDED_RECTANGLE,
                 radius=0.5)
            _txt(sl, th, x + lw + tw + 0.35, ry, 1.4, rh, _fmt(val) + ("%" if unit == "%" else ""), 18,
                 col, True, th.font_head, anchor=MSO_ANCHOR.MIDDLE)
        if unit and unit != "%":
            _txt(sl, th, x, y + h - 0.3, w, 0.3, unit, 11, "muted", align=PP_ALIGN.RIGHT)
    else:  # lollipop
        n = len(pairs)
        base = y + h - 0.65
        top = y + 0.55
        seg = w / n
        _box(sl, th, x, base, w, 0.025, "text")
        for i, (k, val) in enumerate(pairs):
            cx = x + i * seg + seg / 2
            hh = (base - top) * max(val, 0) / vmax
            col = th.accent if i == imax else th.accent2
            _box(sl, th, cx - 0.025, base - hh, 0.05, hh, col)
            d = 0.8 if i == imax else 0.64
            _box(sl, th, cx - d / 2, base - hh - d / 2, d, d, col, MSO_SHAPE.OVAL)
            _txt(sl, th, cx - d / 2, base - hh - d / 2, d, d, _fmt(val), 14 if d < 0.7 else 16, _onc(th, col), True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=9)
            _txt(sl, th, cx - seg / 2, base + 0.1, seg, 0.5, k, 13, "muted", align=PP_ALIGN.CENTER, fit=True, min_size=9)
        if unit:
            _txt(sl, th, x, y, 3, 0.35, unit, 12, "muted", True)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# DOIRA / TAQSIMOT — 6 variant
# ═══════════════════════════════════════════════════════════════════
def _legend(sl, th, pairs, tot, x, y, w, h, size=17):
    n = len(pairs)
    rh = min(0.8, h / n)
    oy = y + (h - rh * n) / 2
    for i, (k, v) in enumerate(pairs):
        ry = oy + i * rh
        col = _col(th, i)
        _box(sl, th, x, ry + rh / 2 - 0.13, 0.26, 0.26, col, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
        _txt(sl, th, x + 0.45, ry, w - 1.9, rh, k, size, "text", anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=11)
        _txt(sl, th, x + w - 1.4, ry, 1.4, rh, f"{v / tot * 100:.0f}%", size + 3, col, True, th.font_head,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            _box(sl, th, x, ry + rh - 0.01, w, 0.01, "line")


def _doughnut(sl, th, vals, colors, x, y, d, hole=58, sep=True):
    cd = CategoryChartData()
    cd.categories = [str(i) for i in range(len(vals))]
    cd.add_series(" ", vals)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(d), Inches(d), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    plot = ch.plots[0]
    try:
        hs = plot._element.find(qn("c:holeSize"))
        if hs is None:
            hs = etree.SubElement(plot._element, qn("c:holeSize"))
        hs.set("val", str(hole))
    except Exception:
        pass
    ser = plot.series[0]
    for i, c in enumerate(colors):
        p = ser.points[i]
        p.format.fill.solid(); p.format.fill.fore_color.rgb = _rgb(th, c)
        if sep:
            p.format.line.color.rgb = th.c("bg"); p.format.line.width = Pt(2)
        else:
            p.format.line.fill.background()
    return ch


def slide_pie(prs, th, sd, num, total, topic):
    pairs = [(k, max(v, 0)) for k, v in _pairs(sd.get("matn", ""))[:6]]
    if len(pairs) < 2 or sum(v for _, v in pairs) <= 0:
        return slide_list(prs, th, sd, num, total, topic)
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    tot = sum(v for _, v in pairs)
    v = th.v["pie"]
    big = max(range(len(pairs)), key=lambda i: pairs[i][1])
    if v in ("doughnut", "pie"):
        d = min(h, 5.0, w * 0.45)
        dy = y + (h - d) / 2
        if v == "doughnut":
            _doughnut(sl, th, [p[1] for p in pairs], [_col(th, i) for i in range(len(pairs))], x + 0.2, dy, d)
            _txt(sl, th, x + 0.2, dy + d / 2 - 0.55, d, 0.7, f"{pairs[big][1] / tot * 100:.0f}%", 34, "primary", True,
                 th.font_head, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, x + 0.2 + d * 0.2, dy + d / 2 + 0.12, d * 0.6, 0.5, pairs[big][0], 12, "muted",
                 align=PP_ALIGN.CENTER, fit=True)
        else:
            cd = CategoryChartData()
            cd.categories = [p[0] for p in pairs]
            cd.add_series(" ", [p[1] for p in pairs])
            gf = sl.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(x + 0.2), Inches(dy), Inches(d), Inches(d), cd)
            ch = gf.chart
            ch.has_legend = False
            ch.has_title = False
            plot = ch.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = False
            dl.show_percentage = True
            dl.number_format = "0%"
            dl.number_format_is_linked = False
            dl.font.size = Pt(14); dl.font.bold = True; dl.font.color.rgb = RGBColor(255, 255, 255)
            dl.position = XL_LABEL_POSITION.INSIDE_END
            ser = plot.series[0]
            for i in range(len(pairs)):
                p = ser.points[i]
                p.format.fill.solid(); p.format.fill.fore_color.rgb = _rgb(th, _col(th, i))
                p.format.line.color.rgb = th.c("bg"); p.format.line.width = Pt(2)
        _legend(sl, th, pairs, tot, x + d + 0.9, y, w - d - 0.9, h)
    elif v == "stacked100":
        by, bh = y + 0.3, 1.3
        cx = x
        for i, (k, val) in enumerate(pairs):
            sw = w * val / tot
            col = _col(th, i)
            _box(sl, th, cx, by, sw, bh, col)
            if sw > 0.9:
                _txt(sl, th, cx, by, sw, bh, f"{val / tot * 100:.0f}%", 22 if sw > 1.3 else 15, _onc(th, col), True,
                     th.font_head, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            cx += sw
        cw = w / min(len(pairs), 3)
        for i, (k, val) in enumerate(pairs):
            r, c = divmod(i, 3)
            lx, ly = x + c * cw, by + bh + 0.55 + r * 1.3
            _box(sl, th, lx, ly + 0.08, 0.08, 1.0, _col(th, i))
            _txt(sl, th, lx + 0.25, ly, cw - 0.4, 0.6, f"{val / tot * 100:.0f}%", 28, _col(th, i), True, th.font_head)
            _txt(sl, th, lx + 0.25, ly + 0.6, cw - 0.4, 0.5, k, 15, "text", fit=True, min_size=10)
    elif v == "waffle":
        size = min(h, 5.0)
        cell = size / 10
        counts = [round(val / tot * 100) for _, val in pairs]
        counts[big] += 100 - sum(counts)
        seq = []
        for i, c in enumerate(counts):
            seq += [i] * max(c, 0)
        seq = (seq + [big] * 100)[:100]
        ox, oy = x + 0.2, y + (h - size) / 2
        for idx in range(100):
            r, c = divmod(idx, 10)
            _box(sl, th, ox + c * cell + 0.03, oy + r * cell + 0.03, cell - 0.06, cell - 0.06,
                 _col(th, seq[idx]), MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
        _legend(sl, th, pairs, tot, ox + size + 0.8, y, w - size - 1.0, h)
    elif v == "bubbles":
        n = len(pairs)
        seg = w / n
        mx = max(p[1] for p in pairs)
        dmax = min(seg - 0.2, h - 1.3, 3.0)
        for i, (k, val) in enumerate(pairs):
            d = max(0.7, dmax * math.sqrt(val / mx))
            cx = x + i * seg + seg / 2
            cy = y + (h - 1.0) / 2
            col = _col(th, i)
            _box(sl, th, cx - d / 2, cy - d / 2, d, d, col, MSO_SHAPE.OVAL, opacity=0.92)
            _txt(sl, th, cx - d / 2, cy - d / 2, d, d, f"{val / tot * 100:.0f}%", max(12, min(30, d * 13)),
                 _onc(th, col), True, th.font_head, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, cx - seg / 2 + 0.1, y + h - 0.95, seg - 0.2, 0.8, k, 15, "text", align=PP_ALIGN.CENTER,
                 fit=True, min_size=10)
    else:  # rings
        n = len(pairs)
        cols = n if w / n >= 2.6 else math.ceil(n / 2)
        rows = math.ceil(n / cols)
        seg = w / cols
        rhh = h / rows
        d = min(seg - 0.35, rhh - 1.0, 2.8)
        for i, (k, val) in enumerate(pairs):
            r, c = divmod(i, cols)
            cx = x + c * seg + seg / 2
            dy = y + r * rhh + (rhh - d - 0.8) / 2
            pct = val / tot
            _doughnut(sl, th, [pct, 1 - pct], [_col(th, i), th.card], cx - d / 2, dy, d, hole=72, sep=False)
            _txt(sl, th, cx - d / 2, dy, d, d, f"{pct * 100:.0f}%", max(14, min(30, d * 11)), _col(th, i), True,
                 th.font_head, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, cx - seg / 2 + 0.1, dy + d + 0.15, seg - 0.2, 0.75, k, 15, "text", align=PP_ALIGN.CENTER,
                 fit=True, min_size=10)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# JADVAL — 4 variant
# ═══════════════════════════════════════════════════════════════════
def slide_table(prs, th, sd, num, total, topic):
    rows = [[_clean(c) for c in r.split("|")] for r in _clean(sd.get("matn", "")).split("\n") if "|" in r]
    rows = [r for r in rows if any(r)][:8]
    if len(rows) < 2:
        return slide_list(prs, th, sd, num, total, topic)
    nc = min(5, max(len(r) for r in rows))
    rows = [(r + [""] * nc)[:nc] for r in rows]
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    v = th.v["table"]
    longest = max(len(c) for r in rows for c in r)
    fs = 16 if longest < 25 and len(rows) <= 5 else 14 if longest < 45 else 12
    if v == "card_rows":
        g = 0.14
        rh = min(0.82, (h - g * (len(rows) - 1)) / len(rows))
        cw = w / nc
        for ri, r in enumerate(rows):
            ry = y + ri * (rh + g)
            if ri == 0:
                for ci, c in enumerate(r):
                    _txt(sl, th, x + ci * cw + 0.25, ry, cw - 0.35, rh, c, fs - 2, "accent", True,
                         align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, caps=True, spc=1)
                continue
            _box(sl, th, x, ry, w, rh, "card", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3,
                 line=None if th.dark else "line", line_w=0.75)
            _box(sl, th, x, ry, 0.12, rh, _col(th, ri - 1))
            for ci, c in enumerate(r):
                _txt(sl, th, x + ci * cw + 0.25, ry, cw - 0.35, rh, c, fs, "primary" if ci == 0 else "text", ci == 0,
                     align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=10)
    else:
        rh = min(0.72, h / len(rows))
        gf = sl.shapes.add_table(len(rows), nc, Inches(x), Inches(y + 0.1), Inches(w), Inches(rh * len(rows)))
        tbl = gf.table
        tp = gf._element.graphic.graphicData.tbl.tblPr
        tp.set("bandRow", "0"); tp.set("firstRow", "0")
        for ri, r in enumerate(rows):
            tbl.rows[ri].height = Inches(rh)
            for ci in range(nc):
                cell = tbl.cell(ri, ci)
                cell.margin_left = cell.margin_right = Inches(0.12)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                head = ri == 0
                if v == "banded":
                    fill = ("accent" if th.dark else "primary") if head else ("card" if ri % 2 == 0 else "bg")
                    color = _onc(th, fill) if head else ("primary" if ci == 0 else "text")
                    bold = head or ci == 0
                elif v == "header_col":
                    if ci == 0:
                        fill, color, bold = "primary", "#FFFFFF", True
                    elif head:
                        fill, color, bold = "accent", _onc(th, "accent"), True
                    else:
                        fill, color, bold = ("card" if ri % 2 == 0 else "bg"), "text", False
                else:  # minimal_lines
                    fill = "bg"
                    color = "accent" if head else ("primary" if ci == 0 else "text")
                    bold = head or ci == 0
                cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(th, fill)
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = ""
                run = p.add_run(); run.text = r[ci]
                run.font.size = Pt(fs + (1 if head else 0)); run.font.bold = bold
                run.font.name = th.font_body; run.font.color.rgb = _rgb(th, color)
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
        if v == "minimal_lines":
            for ri in range(len(rows) + 1):
                ly = y + 0.1 + ri * rh
                thick = ri in (0, 1, len(rows))
                _box(sl, th, x, ly - 0.005, w, 0.03 if thick else 0.01, "primary" if thick else "line")
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# TAQQOSLASH — 4 variant
# ═══════════════════════════════════════════════════════════════════
def slide_compare(prs, th, sd, num, total, topic, lb):
    raw = _clean(sd.get("matn", ""))
    if "|||" in raw:
        a, b = raw.split("|||", 1)
    else:
        it = _items(raw)
        half = (len(it) + 1) // 2
        a, b = "\n".join(it[:half]), "\n".join(it[half:])
    left, right = _items(a)[:6], _items(b)[:6]
    heads = (_clean(sd.get("chapNom")) or lb.get("afzallik", "A"), _clean(sd.get("ongNom")) or lb.get("kamchilik", "B"))
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    v = th.v["compare"]
    g = 0.5
    cw = (w - g) / 2
    cols = (("accent", "✓"), ("accent2", "✕"))
    if v == "colored_halves":
        for k, (items, head) in enumerate(((left, heads[0]), (right, heads[1]))):
            cx = x + k * (cw + g)
            col = cols[k][0]
            tc = _onc(th, col)
            _box(sl, th, cx, y, cw, h, col, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
            _txt(sl, th, cx + 0.4, y + 0.25, cw - 0.8, 0.8, head, 26, tc, True, th.font_head, fit=True, min_size=16)
            _box(sl, th, cx + 0.42, y + 1.1, 0.8, 0.04, tc)
            n = max(1, len(items))
            rh = (h - 1.4) / n
            fs = min([_fit(t, cw - 1.2, rh - 0.05, 20, 11) for t in items] or [18])
            for i, t in enumerate(items):
                ry = y + 1.3 + i * rh
                _txt(sl, th, cx + 0.4, ry, 0.4, rh, cols[k][1], 18, tc, True, anchor=MSO_ANCHOR.MIDDLE)
                _txt(sl, th, cx + 0.85, ry, cw - 1.2, rh, t, fs, tc, anchor=MSO_ANCHOR.MIDDLE)
    elif v == "split_line":
        _box(sl, th, x + w / 2 - 0.01, y, 0.02, h, "accent" if th.dark else "primary")
        for k, (items, head) in enumerate(((left, heads[0]), (right, heads[1]))):
            cx = x + k * (cw + g) + (0.1 if k else 0)
            _txt(sl, th, cx, y, cw - 0.2, 0.75, head, 28, cols[k][0], True, th.font_head, fit=True, min_size=16)
            n = max(1, len(items))
            rh = (h - 1.0) / n
            fs = min([_fit(t, cw - 0.8, rh - 0.05, 20, 11) for t in items] or [18])
            for i, t in enumerate(items):
                ry = y + 0.95 + i * rh
                _txt(sl, th, cx, ry, 0.5, rh, "—", 18, cols[k][0], True, anchor=MSO_ANCHOR.MIDDLE)
                _txt(sl, th, cx + 0.5, ry, cw - 0.8, rh, t, fs, "text", anchor=MSO_ANCHOR.MIDDLE)
    elif v == "table_compare":
        n = max(len(left), len(right), 1)
        hh = 0.75
        rh = min(0.85, (h - hh) / n)
        for k in range(2):
            _box(sl, th, x + k * (w / 2), y, w / 2 - 0.05, hh, cols[k][0])
            _txt(sl, th, x + k * (w / 2) + 0.3, y, w / 2 - 0.6, hh, heads[k], 20, _onc(th, cols[k][0]), True,
                 th.font_head, anchor=MSO_ANCHOR.MIDDLE, fit=True)
        for i in range(n):
            ry = y + hh + i * rh
            if i % 2 == 0:
                _box(sl, th, x, ry, w - 0.05, rh, "card")
            for k, items in enumerate((left, right)):
                if i < len(items):
                    cx = x + k * (w / 2)
                    _num_circle(sl, th, cx + 0.25, ry + rh / 2 - 0.2, 0.4, cols[k][1], _hx(th, cols[k][0]))
                    _txt(sl, th, cx + 0.85, ry, w / 2 - 1.1, rh, items[i], 17, "text", anchor=MSO_ANCHOR.MIDDLE,
                         fit=True, min_size=10)
        _box(sl, th, x + w / 2 - 0.03, y + hh, 0.015, rh * n, "line")
    else:  # cards_vs
        for k, (items, head) in enumerate(((left, heads[0]), (right, heads[1]))):
            cx = x + k * (cw + g)
            col = cols[k][0]
            _card(sl, th, cx, y, cw, h)
            _box(sl, th, cx, y, cw, 0.8, col)
            _txt(sl, th, cx + 0.3, y, cw - 0.6, 0.8, head, 21, _onc(th, col), True, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE, fit=True)
            n = max(1, len(items))
            rh = (h - 1.05) / n
            fs = min([_fit(t, cw - 1.2, rh - 0.05, 20, 11) for t in items] or [18])
            for i, t in enumerate(items):
                ry = y + 0.95 + i * rh
                _txt(sl, th, cx + 0.3, ry, 0.4, rh, cols[k][1], 18, col, True, anchor=MSO_ANCHOR.MIDDLE)
                _txt(sl, th, cx + 0.75, ry, cw - 1.0, rh, t, fs, "text", anchor=MSO_ANCHOR.MIDDLE)
        _num_circle(sl, th, x + cw + g / 2 - 0.38, y + h / 2 - 0.38, 0.76, "VS",
                    th.accent if th.dark else th.primary)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# STATISTIKA KARTALARI — 4 variant
# ═══════════════════════════════════════════════════════════════════
def slide_stats(prs, th, sd, num, total, topic):
    cards = []
    for line in _clean(sd.get("matn", "")).split("\n"):
        p = [q.strip() for q in line.split("|")]
        if len(p) >= 2 and p[1]:
            cards.append((p[0], p[1], p[2] if len(p) > 2 else ""))
    cards = cards[:4]
    if not cards:
        return slide_list(prs, th, sd, num, total, topic)
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    v = th.v["stats"]
    n = len(cards)
    if v == "hero_left" and n >= 2:
        hw = w * 0.45
        t, val, d = cards[0]
        fill = "card" if th.dark else "accent"
        tc = th.accent if th.dark else _onc(th, "accent")
        _box(sl, th, x, y, hw, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE if th.decor != "magazine" else MSO_SHAPE.RECTANGLE,
             radius=0.05)
        _txt(sl, th, x + 0.4, y + 0.3, hw - 0.8, h * 0.48, val, 80, tc, True, th.font_head, anchor=MSO_ANCHOR.BOTTOM,
             fit=True, min_size=36)
        _txt(sl, th, x + 0.45, y + h * 0.52, hw - 0.8, 0.7, t, 24, tc, True, fit=True, min_size=14)
        if d:
            _txt(sl, th, x + 0.45, y + h * 0.52 + 0.75, hw - 0.8, h * 0.48 - 1.0, d, 16, tc, fit=True, min_size=11)
        rx, rw = x + hw + 0.5, w - hw - 0.5
        rh = h / (n - 1)
        for i, (t, val, d) in enumerate(cards[1:]):
            ry = y + i * rh
            _txt(sl, th, rx, ry, rw * 0.45, rh, val, 40, _col(th, i + 1) if th.dark else "primary", True, th.font_head,
                 anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=20)
            _txt(sl, th, rx + rw * 0.47, ry, rw * 0.53, rh, f"{t}\n{d}".strip(), 16, "text", anchor=MSO_ANCHOR.MIDDLE,
                 fit=True, min_size=10)
            if i < n - 2:
                _box(sl, th, rx, ry + rh - 0.01, rw, 0.012, "line")
    elif v == "inline":
        cw = w / n
        top = y + h * 0.12
        for i, (t, val, d) in enumerate(cards):
            cx = x + i * cw
            _txt(sl, th, cx + 0.2, top, cw - 0.4, 1.5, val, 58, "accent", th.decor != "luxe", th.font_head,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, fit=True, min_size=24)
            _box(sl, th, cx + cw / 2 - 0.4, top + 1.65, 0.8, 0.03, "accent2" if th.dark else "primary")
            _txt(sl, th, cx + 0.25, top + 1.85, cw - 0.5, 0.6, t, 19, "primary", True, align=PP_ALIGN.CENTER,
                 fit=True, min_size=12)
            if d:
                _txt(sl, th, cx + 0.25, top + 2.5, cw - 0.5, 1.2, d, 14, "muted", align=PP_ALIGN.CENTER,
                     fit=True, min_size=10)
            if i < n - 1:
                _box(sl, th, cx + cw - 0.01, y + h * 0.15, 0.015, h * 0.7, "line")
    elif v == "circles":
        cw = w / n
        d = min(cw - 0.5, h - 1.6, 2.9)
        for i, (t, val, dd) in enumerate(cards):
            cx = x + i * cw + cw / 2
            col = _col(th, i)
            _box(sl, th, cx - d / 2, y + 0.1, d, d, col, MSO_SHAPE.OVAL, opacity=0.14)
            _box(sl, th, cx - d / 2, y + 0.1, d, d, None, MSO_SHAPE.OVAL, line=col, line_w=3)
            _txt(sl, th, cx - d / 2 + 0.1, y + 0.1, d - 0.2, d, val, 40, col, True, th.font_head,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=16)
            _txt(sl, th, cx - cw / 2 + 0.15, y + d + 0.3, cw - 0.3, 0.6, t, 18, "primary", True, align=PP_ALIGN.CENTER,
                 fit=True, min_size=11)
            if dd:
                _txt(sl, th, cx - cw / 2 + 0.15, y + d + 0.9, cw - 0.3, h - d - 1.0, dd, 13, "muted",
                     align=PP_ALIGN.CENTER, fit=True, min_size=10)
    else:  # cards
        g = 0.35
        cw = (w - g * (n - 1)) / n
        chh = min(h - 0.2, 4.2)
        cy = y + (h - chh) / 2
        for i, (t, val, d) in enumerate(cards):
            cx = x + i * (cw + g)
            col = _col(th, i) if th.decor in ("pastel", "gradient", "dark_geo", "startup") else th.accent
            _card(sl, th, cx, cy, cw, chh)
            _box(sl, th, cx, cy, cw, 0.1, col)
            _txt(sl, th, cx + 0.25, cy + 0.45, cw - 0.5, 1.4, val, 48 if n <= 3 else 40, col, True, th.font_head,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=20)
            _txt(sl, th, cx + 0.25, cy + 1.95, cw - 0.5, 0.7, t, 18, "primary", True, align=PP_ALIGN.CENTER,
                 fit=True, min_size=12)
            if d:
                _txt(sl, th, cx + 0.25, cy + 2.7, cw - 0.5, chh - 2.85, d, 14, "muted", align=PP_ALIGN.CENTER,
                     fit=True, min_size=10)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# BOSQICHLAR — 4 variant
# ═══════════════════════════════════════════════════════════════════
def slide_timeline(prs, th, sd, num, total, topic):
    steps = []
    for line in _clean(sd.get("matn", "")).split("\n"):
        line = re.sub(r"^\d{1,2}[.)]\s*", "", line.strip())
        if not line:
            continue
        sep = "|" if "|" in line else (":" if ":" in line else None)
        a, b = (line.split(sep, 1) if sep else (line, ""))
        steps.append((a.strip(), b.strip()))
    steps = steps[:6]
    if len(steps) < 2:
        return slide_list(prs, th, sd, num, total, topic)
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    n = len(steps)
    v = th.v["timeline"]
    if v == "vertical":
        rh = h / n
        lx = x + 0.35
        _box(sl, th, lx - 0.015, y + rh / 2, 0.03, h - rh, "muted" if th.dark else "line")
        tw = min(3.6, w * 0.35)
        for i, (a, b) in enumerate(steps):
            ry = y + i * rh
            col = _col(th, i) if th.dark else th.accent
            _box(sl, th, lx - 0.22, ry + rh / 2 - 0.22, 0.44, 0.44, col, MSO_SHAPE.OVAL, line="bg", line_w=3)
            _txt(sl, th, lx + 0.5, ry, tw, rh, a, 22, "primary", True, th.font_head, anchor=MSO_ANCHOR.MIDDLE,
                 fit=True, min_size=12)
            _txt(sl, th, lx + tw + 0.7, ry, w - tw - 1.1, rh, b, 19, "text", anchor=MSO_ANCHOR.MIDDLE, fit=True,
                 min_size=10)
    elif v == "chevrons":
        cw = w / n
        chh = 1.3
        cy = y + 0.3
        for i, (a, b) in enumerate(steps):
            cx = x + i * cw
            col = _col(th, i)
            _box(sl, th, cx, cy, cw + 0.15, chh, col, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
            _txt(sl, th, cx + 0.45, cy, cw - 0.7, chh, a, 17, _onc(th, col), True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=10)
            _txt(sl, th, cx + 0.15, cy + chh + 0.35, 0.8, 0.6, f"{i + 1:02d}", 26, col, True, th.font_head)
            if b:
                _txt(sl, th, cx + 0.15, cy + chh + 1.0, cw - 0.3, h - chh - 1.4, b, 18, "text", fit=True, min_size=10)
    elif v == "stairs":
        g = 0.15
        cw = (w - g * (n - 1)) / n
        step_h = (h - 1.0) / n
        for i, (a, b) in enumerate(steps):
            cx = x + i * (cw + g)
            bh = step_h * (i + 1) + 0.4
            by = y + h - bh
            col = _col(th, i)
            tc = _onc(th, col)
            _box(sl, th, cx, by, cw, bh, col, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
            _txt(sl, th, cx + 0.15, by + 0.1, cw - 0.3, 0.55, f"{i + 1:02d}", 22, tc, True, th.font_head)
            _txt(sl, th, cx + 0.15, by + 0.6, cw - 0.3, 0.75, a, 19, tc, True, fit=True, min_size=11)
            if b and bh > 1.9:
                _txt(sl, th, cx + 0.15, by + 1.35, cw - 0.3, bh - 1.45, b, 15, tc, fit=True, min_size=10)
            elif b:
                _txt(sl, th, cx, by - 1.0, cw, 0.95, b, 12, "text", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
                     fit=True, min_size=9)
    else:  # horizontal
        seg = w / n
        ly = y + max(0.5, (h - 3.2) / 2)
        _box(sl, th, x + seg / 2, ly + 0.27, w - seg, 0.05, "muted", opacity=0.5)
        for i, (a, b) in enumerate(steps):
            cx = x + i * seg + seg / 2
            col = _col(th, i) if th.decor in ("pastel", "gradient", "dark_geo") else th.accent
            _box(sl, th, cx - 0.3, ly, 0.6, 0.6, col, MSO_SHAPE.OVAL, line="bg", line_w=3)
            _txt(sl, th, cx - 0.3, ly, 0.6, 0.6, str(i + 1), 16, _onc(th, col), True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
            _txt(sl, th, cx - seg / 2 + 0.1, ly + 0.8, seg - 0.2, 0.9, a, 20, "primary", True, th.font_head,
                 align=PP_ALIGN.CENTER, fit=True, min_size=12)
            if b:
                _txt(sl, th, cx - seg / 2 + 0.12, ly + 1.7, seg - 0.24, 1.6, b, 16, "text", align=PP_ALIGN.CENTER,
                     fit=True, min_size=10)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# XULOSA — 4 variant
# ═══════════════════════════════════════════════════════════════════
def slide_conclusion(prs, th, sd, num, total, topic):
    sl, (x, y, w, h) = _start(prs, th, sd, num)
    items = _items(sd.get("matn", ""))[:7] or [_clean(sd.get("matn", "")) or "—"]
    v = th.v["concl"]
    if v == "banner" and len(items) >= 2:
        bh = 1.35
        if th.grad:
            _grad(sl, th, x, y, w, bh, th.grad[0], th.grad[1], 0)
            tc = "#FFFFFF"
        else:
            _box(sl, th, x, y, w, bh, "accent", MSO_SHAPE.ROUNDED_RECTANGLE if th.decor != "magazine" else MSO_SHAPE.RECTANGLE,
                 radius=0.1)
            tc = _onc(th, "accent")
        _txt(sl, th, x + 0.35, y, 0.8, bh, "★", 32, tc, True, anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, x + 1.1, y + 0.1, w - 1.4, bh - 0.2, items[0], 22, tc, True, anchor=MSO_ANCHOR.MIDDLE,
             fit=True, min_size=13)
        _list_render(sl, th, (x, y + bh + 0.3, w, h - bh - 0.3), items[1:], "rows", check=True)
    elif v == "cards" and 2 <= len(items) <= 6:
        _list_render(sl, th, (x, y, w, h), items, "cards", check=True)
    elif v == "numbered_big":
        _list_render(sl, th, (x, y, w, h), items, "numbered_big")
    else:
        _list_render(sl, th, (x, y, w, h), items, "rows", check=True)
    _footer(sl, th, num, total, topic)
    return sl


# ═══════════════════════════════════════════════════════════════════
# TITUL — 15 variant
# ═══════════════════════════════════════════════════════════════════
def _title_slide(prs, th: Theme, mavzu: str, ism: str, lb: dict):
    sl = prs.slides.add_slide(_blank(prs))
    v = th.v["title"]
    hf = th.font_head
    year = str(_joriy_yil())
    by = f"{lb.get('bajardi', 'Bajardi')}: {ism}" if ism else ""
    kind = lb.get("taqdimot", "Taqdimot")
    if v != "template_title":
        _bg(sl, th, "title_bg")

    def block(x, w, y_kind, y_title, h_title, y_by, align=PP_ALIGN.LEFT, size=44, rule=None):
        _txt(sl, th, x, y_kind, w, 0.45, kind, 14, "accent", True, align=align, caps=True, spc=3)
        _txt(sl, th, x, y_title, w, h_title, mavzu, size, "on_title", True, hf, align, MSO_ANCHOR.MIDDLE,
             fit=True, min_size=24, spacing=1.05)
        if rule:
            rx = x + (w / 2 - 0.9 if align == PP_ALIGN.CENTER else 0.02)
            _box(sl, th, rx, y_by - 0.3, 1.8 if align == PP_ALIGN.CENTER else 1.2, 0.05, rule)
        _txt(sl, th, x, y_by, w, 0.5, by, 20, "text", align=align)
        _txt(sl, th, x, y_by + 0.5, w, 0.4, year, 14, "muted", align=align)

    if v == "framed_center":
        _box(sl, th, 0, 0, SW, 0.35, "primary"); _box(sl, th, 0, 0.35, SW, 0.04, "accent")
        _box(sl, th, 0, SH - 0.35, SW, 0.35, "primary")
        block(1.0, SW - 2.0, 1.25, 1.75, 2.5, 4.75, align=PP_ALIGN.CENTER, size=42, rule="accent")
    elif v == "swiss_left":
        _box(sl, th, 0.8, 0.8, 0.55, 0.55, "accent")
        _txt(sl, th, 1.55, 0.83, 6, 0.5, kind, 14, "muted", True, caps=True, spc=4)
        _txt(sl, th, 0.75, 1.9, 11.2, 3.4, mavzu, 56, "primary", True, hf, anchor=MSO_ANCHOR.BOTTOM, fit=True,
             min_size=30, spacing=1.0)
        _box(sl, th, 0.8, 5.6, SW - 1.6, 0.02, "primary")
        _txt(sl, th, 0.75, 5.75, 8, 0.5, by, 18, "text")
        _txt(sl, th, SW - 3.0, 5.75, 2.2, 0.5, year, 18, "accent", True, align=PP_ALIGN.RIGHT)
    elif v == "split_panel":
        _bg(sl, th, "bg")
        _box(sl, th, 0, 0, 7.4, SH, "primary")
        _box(sl, th, 7.4, 0, 0.12, SH, "accent2")
        _box(sl, th, 8.4, 1.3, 4.0, 4.0, "accent", MSO_SHAPE.OVAL, opacity=0.18)
        _box(sl, th, 9.5, 2.4, 1.8, 1.8, "accent2", MSO_SHAPE.OVAL, opacity=0.9)
        _txt(sl, th, 0.8, 1.0, 6.2, 0.45, kind, 14, "#CFE3FA", True, caps=True, spc=3)
        _txt(sl, th, 0.75, 1.5, 6.3, 3.3, mavzu, 42, "#FFFFFF", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=24)
        _txt(sl, th, 0.8, 5.2, 6.2, 0.5, by, 19, "#FFFFFF")
        _txt(sl, th, 0.8, 5.7, 6.2, 0.4, year, 14, "#CFE3FA")
    elif v == "geo_rings":
        for i in range(1, 14):
            _box(sl, th, i * 1.0, 0, 0.005, SH, "#162036")
        for j in range(1, 8):
            _box(sl, th, 0, j * 1.0, SW, 0.005, "#162036")
        _box(sl, th, 8.4, 1.0, 5.6, 5.6, None, MSO_SHAPE.OVAL, line="accent", line_w=1.5)
        _box(sl, th, 9.4, 2.0, 3.6, 3.6, None, MSO_SHAPE.OVAL, line="accent2", line_w=1.0)
        _box(sl, th, 10.6, 3.2, 1.2, 1.2, "accent", MSO_SHAPE.OVAL, opacity=0.85)
        block(0.9, 7.2, 1.3, 1.8, 3.0, 5.25, rule="accent")
    elif v == "organic_blobs":
        _box(sl, th, 8.2, 2.3, 6.5, 6.5, "accent", MSO_SHAPE.OVAL, opacity=0.18)
        _box(sl, th, 10.0, 0.6, 3.6, 3.6, "accent2", MSO_SHAPE.OVAL, opacity=0.22)
        _box(sl, th, 9.3, 3.6, 3.0, 3.0, "primary", MSO_SHAPE.OVAL, opacity=0.85)
        _box(sl, th, -1.2, -1.2, 3.0, 3.0, "accent", MSO_SHAPE.OVAL, opacity=0.12)
        block(0.9, 7.8, 1.25, 1.75, 3.0, 5.1)
    elif v == "full_gradient":
        _grad(sl, th, 0, 0, SW, SH, th.grad[0], th.grad[1], 315)
        _box(sl, th, 8.8, -1.5, 6.0, 6.0, "#FFFFFF", MSO_SHAPE.OVAL, opacity=0.08)
        _box(sl, th, 10.5, 4.0, 4.0, 4.0, "#FFFFFF", MSO_SHAPE.OVAL, opacity=0.10)
        _box(sl, th, 0.9, 1.25, 2.6, 0.48, "#FFFFFF", MSO_SHAPE.ROUNDED_RECTANGLE, opacity=0.18, radius=0.5)
        _txt(sl, th, 0.9, 1.25, 2.6, 0.48, kind, 13, "#FFFFFF", True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             caps=True, spc=2)
        _txt(sl, th, 0.85, 1.95, 9.5, 2.9, mavzu, 46, "#FFFFFF", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=26)
        _txt(sl, th, 0.9, 5.15, 9, 0.5, by, 20, "#FFFFFF")
        _txt(sl, th, 0.9, 5.65, 9, 0.4, year, 14, "#F5D0FE")
    elif v == "collage":
        _box(sl, th, 8.6, 0.7, 4.0, 4.0, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, opacity=0.9, radius=0.35, rot=12)
        _box(sl, th, 10.3, 3.6, 2.6, 2.6, "accent2", MSO_SHAPE.ROUNDED_RECTANGLE, opacity=0.9, radius=0.35, rot=-8)
        _box(sl, th, 8.4, 4.6, 1.6, 1.6, "#9B8AFB", MSO_SHAPE.OVAL, opacity=0.85)
        _box(sl, th, 12.2, 0.4, 0.7, 0.7, "#FFC145", MSO_SHAPE.OVAL)
        block(0.9, 7.4, 1.3, 1.8, 3.0, 5.1)
    elif v == "ornament":
        _ornament_band(sl, th, 0, 0.55); _ornament_band(sl, th, SH - 0.55, 0.55)
        _star(sl, th, SW / 2, 1.55, 0.42, "accent2"); _star(sl, th, SW / 2, 1.55, 0.24, "accent")
        _txt(sl, th, 1.0, 2.15, SW - 2.0, 0.4, kind, 14, "accent", True, align=PP_ALIGN.CENTER, caps=True, spc=3)
        _txt(sl, th, 1.0, 2.55, SW - 2.0, 2.2, mavzu, 42, "on_title", True, hf, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
             fit=True, min_size=24)
        _box(sl, th, SW / 2 - 1.2, 4.95, 2.4, 0.04, "accent2")
        _txt(sl, th, 1.0, 5.15, SW - 2.0, 0.5, by, 20, "on_title", align=PP_ALIGN.CENTER)
        _txt(sl, th, 1.0, 5.65, SW - 2.0, 0.4, year, 14, "accent2", align=PP_ALIGN.CENTER)
    elif v == "luxe_frame":
        _box(sl, th, 0.4, 0.4, SW - 0.8, SH - 0.8, None, line="accent", line_w=1.25)
        _box(sl, th, 0.55, 0.55, SW - 1.1, SH - 1.1, None, line="accent2", line_w=0.5)
        for (cx, cy) in ((0.4, 0.4), (SW - 0.4, 0.4), (0.4, SH - 0.4), (SW - 0.4, SH - 0.4)):
            _box(sl, th, cx - 0.12, cy - 0.12, 0.24, 0.24, "accent", MSO_SHAPE.DIAMOND)
        _txt(sl, th, 1.2, 1.35, SW - 2.4, 0.45, kind, 14, "accent2", align=PP_ALIGN.CENTER, caps=True, spc=6)
        _txt(sl, th, 1.2, 1.9, SW - 2.4, 2.6, mavzu, 44, "on_title", False, hf, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
             fit=True, min_size=26)
        _box(sl, th, SW / 2 - 1.3, 4.75, 2.6, 0.02, "accent")
        _box(sl, th, SW / 2 - 0.09, 4.67, 0.18, 0.18, "accent", MSO_SHAPE.DIAMOND)
        _txt(sl, th, 1.2, 5.0, SW - 2.4, 0.5, by, 19, "text", align=PP_ALIGN.CENTER)
        _txt(sl, th, 1.2, 5.5, SW - 2.4, 0.4, year, 13, "muted", align=PP_ALIGN.CENTER, spc=3)
    elif v == "science_side":
        _dot_grid(sl, th, "#D5E6E3", step=0.5, size=0.03)
        _box(sl, th, 0, 0, 4.6, SH, "primary"); _box(sl, th, 4.6, 0, 0.1, SH, "accent")
        pts = [(1.4, 2.0), (2.6, 1.4), (3.4, 2.6), (2.2, 3.4), (1.2, 4.6), (3.1, 4.9)]
        for (a, b) in ((0, 1), (1, 2), (2, 3), (3, 0), (3, 4), (3, 5)):
            _line(sl, th, *pts[a], *pts[b], "accent")
        for i, (px, py) in enumerate(pts):
            _box(sl, th, px - 0.17, py - 0.17, 0.34, 0.34, "accent2" if i % 3 == 0 else "accent", MSO_SHAPE.OVAL)
        block(5.3, 7.3, 1.3, 1.8, 3.0, 5.2, rule="accent2")
    elif v == "diagonal":
        _box(sl, th, 8.2, -2.0, 9.0, 12.0, "accent", opacity=0.9, rot=20)
        _box(sl, th, 9.8, -2.0, 9.0, 12.0, "accent2", opacity=0.85, rot=20)
        _dot_grid(sl, th, "#3A2470", x0=0.4, y0=0.4, nx=11, ny=14, step=0.5, size=0.03)
        _txt(sl, th, 0.9, 1.3, 6.5, 0.45, "✦ " + kind, 14, "accent2", True, caps=True, spc=3)
        _txt(sl, th, 0.85, 1.8, 6.6, 3.2, mavzu, 44, "#FFFFFF", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=24)
        _txt(sl, th, 0.9, 5.25, 6.5, 0.5, by, 19, "#FFFFFF")
        _txt(sl, th, 0.9, 5.75, 6.5, 0.4, year, 14, "muted")
    elif v == "typographic":
        _box(sl, th, 0.75, 0.55, SW - 1.5, 0.04, "primary")
        _box(sl, th, 0.75, 0.64, SW - 1.5, 0.01, "primary")
        _txt(sl, th, 0.75, 0.75, 6, 0.4, kind.upper() + "  ·  " + year, 12, "primary", True, spc=3)
        _txt(sl, th, SW - 4.75, 0.75, 4, 0.4, "№ 01", 12, "accent", True, align=PP_ALIGN.RIGHT, spc=2)
        _txt(sl, th, 0.7, 1.3, SW - 1.4, 4.2, mavzu, 66, "primary", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True,
             min_size=34, spacing=0.98)
        _box(sl, th, 0.75, 5.75, SW - 1.5, 0.01, "primary")
        _box(sl, th, 0.75, 5.95, 0.12, 0.5, "accent")
        _txt(sl, th, 1.05, 5.92, 8, 0.55, by, 18, "text", italic=True, anchor=MSO_ANCHOR.MIDDLE)
    elif v == "bottom_band":
        _box(sl, th, 0, SH - 2.3, SW, 2.3, "primary")
        _box(sl, th, 0, SH - 2.3, SW, 0.1, "accent")
        _box(sl, th, SW - 3.0, 0.7, 2.2, 2.2, "accent", MSO_SHAPE.OVAL)
        _box(sl, th, SW - 1.8, 2.3, 1.0, 1.0, "accent2", MSO_SHAPE.OVAL)
        _box(sl, th, 0.9, 0.9, 1.7, 0.42, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        _txt(sl, th, 0.9, 0.9, 1.7, 0.42, kind, 12, "#FFFFFF", True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             caps=True, spc=1)
        _txt(sl, th, 0.85, 1.5, 9.4, 3.4, mavzu, 48, "primary", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=26)
        _txt(sl, th, 0.9, SH - 1.75, 8, 0.55, by, 20, "#FFFFFF", True)
        _txt(sl, th, 0.9, SH - 1.2, 8, 0.4, year, 14, "#AFC0E8")
    elif v == "circle_focus":
        _box(sl, th, 6.9, 0.55, 6.4, 6.4, "accent", MSO_SHAPE.OVAL, opacity=0.12)
        _box(sl, th, 7.7, 1.35, 4.8, 4.8, "accent", MSO_SHAPE.OVAL, opacity=0.2)
        _box(sl, th, 8.6, 2.25, 3.0, 3.0, "primary", MSO_SHAPE.OVAL)
        _box(sl, th, 11.4, 1.2, 0.8, 0.8, "accent2", MSO_SHAPE.OVAL)
        _txt(sl, th, 8.6, 2.25, 3.0, 3.0, year, 30, "#FFFFFF", True, hf, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _txt(sl, th, 0.9, 1.3, 6, 0.45, kind, 14, "accent", True, caps=True, spc=3)
        _txt(sl, th, 0.85, 1.8, 6.2, 3.1, mavzu, 42, "primary", True, hf, anchor=MSO_ANCHOR.MIDDLE, fit=True, min_size=24)
        _box(sl, th, 0.9, 5.05, 1.2, 0.05, "accent2")
        _txt(sl, th, 0.9, 5.25, 6, 0.5, by, 20, "text")
    elif v == "template_title":
        block(1.0, SW - 2.0, 1.45, 1.95, 2.4, 4.85, align=PP_ALIGN.CENTER, size=42, rule="accent")
    elif v == "chalk":
        _box(sl, th, 0, 0, SW, SH, None, line="#7A5230", line_w=18)
        _box(sl, th, 0.55, 0.55, SW - 1.1, SH - 1.1, None, line="#F5F1E3", line_w=0.5)
        _txt(sl, th, 1.0, 1.2, SW - 2, 0.5, "~ " + kind + " ~", 16, "accent", True, align=PP_ALIGN.CENTER, spc=3)
        _txt(sl, th, 1.0, 1.8, SW - 2, 2.7, mavzu, 46, "on_title", True, hf, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
             fit=True, min_size=26)
        _box(sl, th, SW / 2 - 2.0, 4.7, 4.0, 0.06, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        _txt(sl, th, 1.0, 4.95, SW - 2, 0.5, by, 20, "accent2", align=PP_ALIGN.CENTER)
        _txt(sl, th, 1.0, 5.45, SW - 2, 0.4, year, 14, "muted", align=PP_ALIGN.CENTER)
        _box(sl, th, 1.4, SH - 0.62, 1.0, 0.1, "#F5F1E3", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        _box(sl, th, 2.6, SH - 0.62, 0.5, 0.1, "accent", MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    return sl


def _thanks_slide(prs, th: Theme, lb: dict, ism: str):
    sl = prs.slides.add_slide(_blank(prs))
    v = th.v["title"]
    thanks = lb.get("rahmat", "E'tiboringiz uchun rahmat!")
    q = lb.get("savollar", "Savollaringiz bormi?")
    main, sub = "primary", "muted"
    if v in ("full_gradient", "diagonal"):
        _grad(sl, th, 0, 0, SW, SH, th.grad[0], th.grad[1], 315)
        main, sub = "#FFFFFF", "#F5D0FE"
    elif v in ("split_panel", "bottom_band", "science_side"):
        _bg(sl, th, "primary")
        main, sub = "#FFFFFF", "#CFE3FA"
        _box(sl, th, 0, SH - 0.12, SW, 0.12, "accent")
    elif v == "ornament":
        _bg(sl, th, "title_bg")
        main, sub = "on_title", "accent2"
        _ornament_band(sl, th, 0, 0.55); _ornament_band(sl, th, SH - 0.55, 0.55)
    elif v == "luxe_frame":
        _bg(sl, th, "title_bg")
        main, sub = "on_title", "muted"
        _box(sl, th, 0.4, 0.4, SW - 0.8, SH - 0.8, None, line="accent", line_w=1.25)
    elif v == "chalk":
        _bg(sl, th, "title_bg")
        main, sub = "on_title", "accent2"
        _box(sl, th, 0, 0, SW, SH, None, line="#7A5230", line_w=18)
    elif v == "typographic":
        _bg(sl, th, "bg")
        _box(sl, th, 0.75, 0.55, SW - 1.5, 0.04, "primary"); _box(sl, th, 0.75, SH - 0.6, SW - 1.5, 0.04, "primary")
    elif v == "template_title":
        pass            # shablonning o'z foni
    else:
        _bg(sl, th, "title_bg")
        if th.dark:
            _box(sl, th, SW / 2 - 3.4, SH / 2 - 3.4, 6.8, 6.8, None, MSO_SHAPE.OVAL, line="accent", line_w=1)
        else:
            _box(sl, th, SW / 2 - 3.2, SH / 2 - 3.2, 6.4, 6.4, "accent", MSO_SHAPE.OVAL, opacity=0.08)
    _txt(sl, th, 1.0, 2.35, SW - 2.0, 1.4, thanks, 46, main, True, th.font_head, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
         fit=True, min_size=28)
    _box(sl, th, SW / 2 - 0.8, 3.95, 1.6, 0.05, "#FFFFFF" if main == "#FFFFFF" else "accent")
    _txt(sl, th, 1.0, 4.2, SW - 2.0, 0.6, q, 22, sub, align=PP_ALIGN.CENTER)
    if ism:
        _txt(sl, th, 1.0, 4.85, SW - 2.0, 0.5, ism, 16, sub, align=PP_ALIGN.CENTER)
    return sl


# ═══════════════════════════════════════════════════════════════════
# ASOSIY FUNKSIYA
# ═══════════════════════════════════════════════════════════════════
DEFAULT_LABELS = {
    "uz": {"bajardi": "Bajardi", "rahmat": "E'tiboringiz uchun rahmat!", "savollar": "Savollaringiz bormi?",
           "taqdimot": "Taqdimot", "afzallik": "Afzalliklar", "kamchilik": "Kamchiliklar"},
    "ru": {"bajardi": "Выполнил(а)", "rahmat": "Спасибо за внимание!", "savollar": "Ваши вопросы?",
           "taqdimot": "Презентация", "afzallik": "Преимущества", "kamchilik": "Недостатки"},
    "en": {"bajardi": "Presented by", "rahmat": "Thank you for your attention!", "savollar": "Questions?",
           "taqdimot": "Presentation", "afzallik": "Advantages", "kamchilik": "Disadvantages"},
}

_RENDER = {
    "reja": slide_agenda, "xulosa": slide_conclusion, "ustun_diag": slide_bar, "doira_diag": slide_pie,
    "jadval_slayd": slide_table, "stat_karta": slide_stats, "bosqichlar": slide_timeline,
}


# ═══════════════════════════════════════════════════════════════════
# TAQDIMOT KONTENT TILLARI
# ═══════════════════════════════════════════════════════════════════
DOC_LANGS = {
    "uz":     {"nom": "🇺🇿 O'zbek (lotin)",   "cmd": "Barcha matnni O'ZBEK tilida, LOTIN yozuvida yoz."},
    "uz_cyr": {"nom": "🇺🇿 Ўзбек (кирилл)",   "cmd": "Барча матнни ЎЗБЕК тилида, КИРИЛЛ ёзувида ёз (Ўзбекистон кирилл алифбоси: ў, қ, ғ, ҳ)."},
    "kaa":    {"nom": "Qaraqalpaq tili",      "cmd": "Barlıq tekstti QARAQALPAQ tilinde, latın jazıwında jaz (á, ó, ú, ı, ń, ǵ háriplerin qollan). "
                                                     "Write ALL text in the KARAKALPAK language (Latin script)."},
    "ru":     {"nom": "🇷🇺 Русский",           "cmd": "Пиши ВЕСЬ текст ТОЛЬКО на РУССКОМ языке."},
    "en":     {"nom": "🇬🇧 English",           "cmd": "Write ALL text ONLY in ENGLISH."},
    "kk":     {"nom": "🇰🇿 Қазақ тілі",        "cmd": "Барлық мәтінді ҚАЗАҚ тілінде (кирилл) жаз. Write ALL text in KAZAKH (Cyrillic)."},
    "ky":     {"nom": "🇰🇬 Кыргыз тили",       "cmd": "Бардык текстти КЫРГЫЗ тилинде жаз. Write ALL text in KYRGYZ (Cyrillic)."},
    "tg":     {"nom": "🇹🇯 Тоҷикӣ",            "cmd": "Тамоми матнро бо забони ТОҶИКӢ (кириллӣ) нависед. Write ALL text in TAJIK (Cyrillic)."},
    "tr":     {"nom": "🇹🇷 Türkçe",            "cmd": "Tüm metni TÜRKÇE yaz. Write ALL text in TURKISH."},
    "de":     {"nom": "🇩🇪 Deutsch",           "cmd": "Schreibe den GESAMTEN Text auf DEUTSCH. Write ALL text in GERMAN."},
}
DEFAULT_LABELS.update({
    "uz_cyr": {"bajardi": "Бажарди", "rahmat": "Эътиборингиз учун раҳмат!", "savollar": "Саволларингиз борми?",
               "taqdimot": "Тақдимот", "afzallik": "Афзалликлар", "kamchilik": "Камчиликлар"},
    "kaa": {"bajardi": "Orınlaǵan", "rahmat": "Itibarıńız ushın raxmet!", "savollar": "Sorawlarıńız barma?",
            "taqdimot": "Prezentaciya", "afzallik": "Artıqmashılıqları", "kamchilik": "Kemshilikleri"},
    "kk": {"bajardi": "Орындаған", "rahmat": "Назарларыңызға рахмет!", "savollar": "Сұрақтарыңыз бар ма?",
           "taqdimot": "Презентация", "afzallik": "Артықшылықтары", "kamchilik": "Кемшіліктері"},
    "ky": {"bajardi": "Аткарган", "rahmat": "Көңүл бурганыңыз үчүн рахмат!", "savollar": "Суроолоруңуз барбы?",
           "taqdimot": "Презентация", "afzallik": "Артыкчылыктары", "kamchilik": "Кемчиликтери"},
    "tg": {"bajardi": "Иҷро кард", "rahmat": "Ташаккур барои диққататон!", "savollar": "Савол доред?",
           "taqdimot": "Презентатсия", "afzallik": "Бартариҳо", "kamchilik": "Камбудиҳо"},
    "tr": {"bajardi": "Hazırlayan", "rahmat": "İlginiz için teşekkürler!", "savollar": "Sorularınız var mı?",
           "taqdimot": "Sunum", "afzallik": "Avantajlar", "kamchilik": "Dezavantajlar"},
    "de": {"bajardi": "Vorgelegt von", "rahmat": "Vielen Dank für Ihre Aufmerksamkeit!", "savollar": "Fragen?",
           "taqdimot": "Präsentation", "afzallik": "Vorteile", "kamchilik": "Nachteile"},
})
_MANBA_LBL = {"uz": "Manba", "uz_cyr": "Манба", "kaa": "Derek", "ru": "Источник", "en": "Source", "kk": "Дереккөз",
              "ky": "Булак", "tg": "Манбаъ", "tr": "Kaynak", "de": "Quelle"}
for _k, _v in _MANBA_LBL.items():
    DEFAULT_LABELS.setdefault(_k, dict(DEFAULT_LABELS["uz"])).setdefault("manba", _v)


# ═══════════════════════════════════════════════════════════════════
# FOYDALANUVCHI SHABLONI (.pptx)
# ═══════════════════════════════════════════════════════════════════
def _blend(c1: str, c2: str, t: float) -> str:
    a, b = _hex(c1), _hex(c2)
    return "#" + "".join(f"{int(a[i] + (b[i] - a[i]) * t):02X}" for i in range(3))


def _lum(c: str) -> float:
    r, g, b = _hex(c)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _template_colors(prs) -> tuple[dict, dict]:
    """Shablon mavzusidan (theme1.xml) ranglar va shriftlarni o'qiydi."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    colors, fonts = {}, {}
    try:
        theme_part = prs.slide_master.part.part_related_by(RT.THEME)
        root = etree.fromstring(theme_part.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        cs = root.find(".//a:clrScheme", ns)
        for el in cs:
            name = etree.QName(el).localname
            c = el.find("a:srgbClr", ns)
            if c is None:
                c = el.find("a:sysClr", ns)
                val = c.get("lastClr") if c is not None else None
            else:
                val = c.get("val")
            if val:
                colors[name] = "#" + val.upper()
        for kind, key in (("majorFont", "head"), ("minorFont", "body")):
            lat = root.find(f".//a:{kind}/a:latin", ns)
            if lat is not None and lat.get("typeface") and not lat.get("typeface").startswith("+"):
                fonts[key] = lat.get("typeface")
    except Exception:
        pass
    return colors, fonts


def _template_bg(prs, colors: dict) -> str:
    """Shablon fonining asosiy rangi (taxminan)."""
    for obj in [prs.slide_master] + list(prs.slide_layouts)[:2]:
        try:
            bg = obj._element.find(".//" + qn("p:bg"))
            if bg is None:
                continue
            c = bg.find(".//" + qn("a:srgbClr"))
            if c is not None:
                return "#" + c.get("val").upper()
            sc = bg.find(".//" + qn("a:schemeClr"))
            if sc is not None:
                m = {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}.get(sc.get("val"), sc.get("val"))
                if m in colors:
                    return colors[m]
        except Exception:
            pass
    return colors.get("lt1", "#FFFFFF")


def _theme_from_template(prs) -> Theme:
    col, fonts = _template_colors(prs)
    bg = _template_bg(prs, col)
    dark = _lum(bg) < 110
    dk1, lt1 = col.get("dk1", "#1F2937"), col.get("lt1", "#FFFFFF")
    acc1, acc2 = col.get("accent1", "#2F5597"), col.get("accent2", "#ED7D31")
    text = lt1 if dark else dk1
    if _lum(text) > 110 and not dark:
        text = "#222222"
    if _lum(text) < 110 and dark:
        text = "#F2F2F2"
    primary = (col.get("lt2", text) if dark else col.get("dk2", dk1))
    if abs(_lum(primary) - _lum(bg)) < 90:
        primary = text
    chart = [col[k] for k in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6") if k in col] or \
        ["#2F5597", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]
    return Theme(
        key="custom", decor="template",
        names={"uz": "📎 O'z shabloningiz", "ru": "📎 Ваш шаблон", "en": "📎 Your template"},
        desc={"uz": "", "ru": "", "en": ""},
        v=dict(title="template_title", header="underline", text="quote_card", list="rows", agenda="columns",
               bar="column", pie="doughnut", table="banded", compare="cards_vs", stats="cards",
               timeline="horizontal", concl="check_rows"),
        bg=bg, title_bg=bg, primary=primary, accent=acc1, accent2=acc2, text=text,
        muted=_blend(text, bg, 0.45), card=_blend(bg, text, 0.07), on_title=primary,
        chart=chart, font_head=fonts.get("head", "Calibri"), font_body=fonts.get("body", "Calibri"),
        dark=dark, line=_blend(bg, text, 0.15),
    )


def _open_template(path: str):
    """Shablonni ochadi, slaydlarini o'chiradi va 16:9 (13.333 x 7.5) o'lchamga keltiradi."""
    prs = Presentation(path)
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        try:
            prs.part.drop_rel(sid.rId)
        except Exception:
            pass
        lst.remove(sid)
    W, H = Inches(SW), Inches(SH)
    if prs.slide_width != W or prs.slide_height != H:
        fx, fy = W / prs.slide_width, H / prs.slide_height
        for obj in [prs.slide_master] + list(prs.slide_layouts):
            for sh in obj.shapes:
                try:
                    if sh.left is not None:
                        sh.left, sh.top = int(sh.left * fx), int(sh.top * fy)
                        sh.width, sh.height = int(sh.width * fx), int(sh.height * fy)
                except Exception:
                    pass
        prs.slide_width, prs.slide_height = W, H
    return prs


def check_template(path: str) -> tuple[bool, str]:
    """Yuklangan .pptx shablon sifatida ishlatsa bo'ladimi."""
    try:
        prs = Presentation(path)
        if not list(prs.slide_layouts):
            return False, "no_layouts"
        return True, ""
    except Exception as e:
        return False, str(e)[:100]


_DATA_TIPO = {"ustun_diag", "doira_diag", "stat_karta", "jadval_slayd", "chiziq_diag"}


def _source_note(prs, th, sd, before, lb):
    """Raqamli slaydlarda ma'lumot manbasini kichik izoh sifatida ko'rsatadi (uydirma raqamlarning oldini olish)."""
    manba = _clean(sd.get("manba", ""))
    if not manba or len(prs.slides) <= before:
        return
    if sd.get("tipo") not in _DATA_TIPO and not sd.get("birligi"):
        return
    manba = manba if len(manba) <= 140 else manba[:137].rsplit(" ", 1)[0] + "…"
    sl = prs.slides[len(prs.slides) - 1]
    label = lb.get("manba", "Manba")
    _txt(sl, th, SW * 0.38, SH - 0.74, SW * 0.62 - 0.75, 0.24, f"{label}: {manba}", 8, "muted",
         italic=True, align=PP_ALIGN.RIGHT)


def build_presentation(slides: list, mavzu: str, ism: str, theme_key: str,
                       til: str = "uz", out_path: str = "taqdimot.pptx",
                       labels: dict | None = None, template_path: str | None = None) -> str:
    lb = dict(DEFAULT_LABELS.get(til, DEFAULT_LABELS["uz"]))
    if labels:
        lb.update({k: v for k, v in labels.items() if v})
    if template_path and os.path.exists(template_path):
        prs = _open_template(template_path)
        th = _theme_from_template(prs)
    else:
        th = get_theme(theme_key)
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    if not template_path:   # ba'zi shablonlarda (LibreOffice) core.xml boshqacha ulangan — tegmaymiz
        prs.core_properties.title = _clean(mavzu)[:200]
        prs.core_properties.author = _clean(ism)[:100] or "Academic Bot"
    total = len(slides) + 2
    _title_slide(prs, th, mavzu, ism, lb)
    for i, sd in enumerate(slides):
        if not isinstance(sd, dict):
            continue
        num = i + 2
        tipo = sd.get("tipo", "matn")
        before = len(prs.slides)
        try:
            if tipo in _RENDER:
                _RENDER[tipo](prs, th, sd, num, total, mavzu)
            elif tipo == "taqqoslash":
                slide_compare(prs, th, sd, num, total, mavzu, lb)
            elif tipo == "royxat" or len(_items(sd.get("matn", ""))) >= 4:
                slide_list(prs, th, sd, num, total, mavzu)
            else:
                slide_text(prs, th, sd, num, total, mavzu)
            _source_note(prs, th, sd, before, lb)
        except Exception:
            # Bitta slayd xato bersa ham taqdimot to'xtamasin: yarim chizilgan slaydni olib tashlaymiz
            lst = prs.slides._sldIdLst
            while len(prs.slides) > before:
                lst.remove(lst[-1])
            slide_text(prs, th, {"sarlavha": sd.get("sarlavha", ""), "matn": str(sd.get("matn", ""))},
                            num, total, mavzu)
    _thanks_slide(prs, th, lb, ism)
    prs.save(out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════════
# KO'RINISH RASMLARI
# ═══════════════════════════════════════════════════════════════════
SAMPLE = [
    {"sarlavha": "Reja", "tipo": "reja",
     "matn": "1. Sun'iy intellektning ta'limdagi o'rni\n2. Oliy ta'limda AI tajribasi\n3. Imkoniyatlar va xavflar"},
    {"sarlavha": "AI dan foydalanuvchi talabalar ulushi", "tipo": "ustun_diag", "birligi": "%",
     "izoh": "So'nggi besh yilda AI vositalaridan foydalanuvchi talabalar ulushi to'rt baravardan ko'proq oshdi.",
     "matn": "2020:18\n2021:27\n2022:41\n2023:63\n2024:78"},
    {"sarlavha": "Foydalanish sohalari", "tipo": "doira_diag",
     "izoh": "Talabalar AI dan eng ko'p uy vazifasi va ilmiy ishlarda foydalanadi.",
     "matn": "Uy vazifasi:38\nIlmiy ish:27\nTil o'rganish:21\nBoshqa:14"},
    {"sarlavha": "Joriy etish bosqichlari", "tipo": "bosqichlar",
     "izoh": "AI ni ta'limga joriy etish to'rt bosqichda amalga oshiriladi.",
     "matn": "Tahlil|Ehtiyojlarni o'rganish\nTanlash|Mos vositani tanlash\nO'qitish|Pedagoglarni tayyorlash\nBaholash|Natijani o'lchash"},
]


def _label(th: Theme) -> str:
    """Rasmga yoziladigan nom (emoji'siz)."""
    return re.sub(r"[^\w\s'ʻ’()-]", "", th.names["uz"]).strip()


def _onc_rgb(rgb) -> tuple:
    r, g, b = rgb
    return (20, 20, 30) if (0.299 * r + 0.587 * g + 0.114 * b) > 165 else (255, 255, 255)


def _font(size: int):
    from PIL import ImageFont
    for fp in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "DejaVuSans-Bold.ttf", "arialbd.ttf"):
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            pass
    return ImageFont.load_default()


def _badge(d, x, y, r, num: int, color):
    """Katta raqamli doira (tugmadagi raqam bilan bir xil)."""
    d.ellipse((x, y, x + 2 * r, y + 2 * r), fill=color, outline=(255, 255, 255), width=max(3, r // 8))
    t = str(num)
    f = _font(int(r * (1.15 if len(t) == 1 else 0.9)))
    bb = d.textbbox((0, 0), t, font=f)
    d.text((x + r - (bb[2] - bb[0]) / 2 - bb[0], y + r - (bb[3] - bb[1]) / 2 - bb[1]), t, fill=(255, 255, 255), font=f)


def make_previews(out_dir: str = ASSETS_DIR, keys: list | None = None) -> None:
    """
    Har bir shablon uchun ko'rinish rasmi: tepada KATTA raqam + nom (bot tugmasidagi bilan bir xil),
    ostida titul slaydi va 4 ta kontent slaydi. Oxirida barcha shablonlarning umumiy jadvali (_hammasi.jpg).
    """
    import subprocess
    import tempfile
    import fitz
    from PIL import Image, ImageDraw

    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    titles = {}
    for key in keys or THEME_ORDER:
        p = os.path.join(tmp, f"{key}.pptx")
        build_presentation(SAMPLE, "Sun'iy intellekt va zamonaviy ta'lim", "Aliyev Sardor", key, "uz", p)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, p],
                       check=True, capture_output=True, timeout=180)
        pdf = fitz.open(os.path.join(tmp, f"{key}.pdf"))
        imgs = []
        for i in range(5):
            pix = pdf[i].get_pixmap(matrix=fitz.Matrix(1.4, 1.4))
            imgs.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        titles[key] = imgs[0]
        th = THEMES[key]
        num = THEME_ORDER.index(key) + 1
        acc = _hex(th.accent)
        W, g, top = 1280, 14, 130
        big = imgs[0].resize((W, int(W * SH / SW)))
        sw = (W - 3 * g) // 4
        smalls = [im.resize((sw, int(sw * SH / SW))) for im in imgs[1:]]
        H = top + big.height + g + smalls[0].height + 30
        canvas = Image.new("RGB", (W + 60, H), (255, 255, 255))
        d = ImageDraw.Draw(canvas)
        d.rectangle((0, 0, W + 60, top - 10), fill=acc)
        _badge(d, 28, 12, 48, num, (20, 20, 30))
        d.text((150, 30), _label(th), fill=_onc_rgb(acc), font=_font(56))
        canvas.paste(big, (30, top))
        for i, sm in enumerate(smalls):
            canvas.paste(sm, (30 + i * (sw + g), top + big.height + g))
        canvas.save(os.path.join(out_dir, f"{key}.jpg"), quality=86)
        print("preview:", key)
    if not keys:
        make_overview(titles, out_dir)


def make_overview(titles: dict, out_dir: str = ASSETS_DIR) -> str:
    """Barcha shablonlar titul slaydlari bitta rasmda — har birida katta raqam va nom."""
    from PIL import Image, ImageDraw
    cols = 3
    tw = 600
    th_ = int(tw * SH / SW)
    lab = 64
    g = 22
    rows = math.ceil(len(THEME_ORDER) / cols)
    W = cols * tw + (cols + 1) * g
    H = rows * (th_ + lab) + (rows + 1) * g
    canvas = Image.new("RGB", (W, H), _hex("#ECEEF2"))
    d = ImageDraw.Draw(canvas)
    for i, key in enumerate(THEME_ORDER):
        r, c = divmod(i, cols)
        x = g + c * (tw + g)
        y = g + r * (th_ + lab + g)
        th = THEMES[key]
        canvas.paste(titles[key].resize((tw, th_)), (x, y))
        d.rectangle((x, y + th_, x + tw, y + th_ + lab), fill=_hex(th.accent))
        _badge(d, x + 14, y + th_ - 46, 44, i + 1, (20, 20, 30))
        fg = _onc_rgb(_hex(th.accent))
        d.text((x + 120, y + th_ + 12), _label(th), fill=fg, font=_font(36))
    path = os.path.join(out_dir, "_hammasi.jpg")
    canvas.save(path, quality=88)
    print("overview:", path)
    return path


def overview_path() -> str:
    return os.path.join(ASSETS_DIR, "_hammasi.jpg")


if __name__ == "__main__":
    import sys
    if "--previews" in sys.argv:
        make_previews()
    else:
        print("\n".join(f"{k}: {t.names['uz']}  {t.v}" for k, t in THEMES.items()))

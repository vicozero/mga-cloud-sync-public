import io
from datetime import datetime
from pathlib import Path
from typing import List, Sequence

from reportlab.lib.colors import Color, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as pdfcanvas

APP_W = 1240
APP_H = 1754
PAGE_W = APP_W
PAGE_H = APP_H
SCALE = 1.0

_LOGO_PATH = Path(__file__).resolve().parent / "assets" / "mga_logo.jpg"

_LOGO_IMAGE = None
try:
    _LOGO_IMAGE = ImageReader(str(_LOGO_PATH))
except Exception:
    pass

NAVY = HexColor("#0B1F33")
BLUE = HexColor("#1F4E79")
GOLD = HexColor("#F5B942")
GREEN = HexColor("#14A36C")
AMBER = HexColor("#EF9100")
RED = HexColor("#DA2330")
LIGHTGRAY = HexColor("#C0C0C0")
LABEL = HexColor("#687786")
ZEBRA = HexColor("#F8FAFC")
LINE = HexColor("#E1E7ED")
FOOTBG = HexColor("#EEF2F6")
GRAY = HexColor("#808080")
DKGRAY = HexColor("#434343")
CATBG = HexColor("#E2ECF7")
CARDGRAY = HexColor("#4B5B6B")
TIREBAND = HexColor("#EBF0F6")

_MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def _lerp(c1: Color, c2: Color, t: float) -> Color:
    return Color(
        c1.red + (c2.red - c1.red) * t,
        c1.green + (c2.green - c1.green) * t,
        c1.blue + (c2.blue - c1.blue) * t,
    )


def _tint(c: Color, alpha: float) -> Color:
    return _lerp(Color(1, 1, 1), c, alpha)


def _fmt_number(value: float) -> str:
    if value == int(value):
        return str(int(value))
    return f"{value:.2f}"


def _fmt_ts() -> str:
    now = datetime.now()
    return f"{now.day} de {_MESES[now.month - 1]} de {now.year} - {now:%H:%M}"


def _fmt_dm() -> str:
    now = datetime.now()
    return f"{now.day:02d}/{now.month:02d}/{now.year} {now:%H:%M}"


def _fmt_date_short(iso: str) -> str:
    parts = str(iso).split("-")
    if len(parts) == 3:
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return iso


def _status_colors(status: str) -> Color:
    if status == "DISPONIBLE":
        return HexColor("#19A974")
    if status == "OPERATIVA":
        return HexColor("#DC8400")
    return HexColor("#E24646")


def _mech_status_color(status: str) -> Color:
    if status == "Disponible":
        return GREEN
    if status == "Operativa":
        return AMBER
    if status == "Pendiente":
        return BLUE
    return RED


def _tire_status_color(status: str) -> Color:
    if status == "Bueno":
        return GREEN
    if status == "Aceptable":
        return GOLD
    if status == "Regular":
        return AMBER
    return RED


class AppPdf:
    def __init__(self, metadata: str):
        self._buf = io.BytesIO()
        self._c = pdfcanvas.Canvas(self._buf, pagesize=(PAGE_W, PAGE_H))
        self._c.setTitle(metadata)

    def new_page(self):
        self._c.showPage()

    def finish(self) -> bytes:
        self._c.save()
        return self._buf.getvalue()

    def rect(self, x, y, w, h, color: Color):
        self._c.setFillColor(color)
        self._c.rect(x * SCALE, PAGE_H - (y + h) * SCALE, w * SCALE, h * SCALE, stroke=0, fill=1)

    def rr(self, x, y, w, h, r, color: Color):
        self._c.setFillColor(color)
        self._c.roundRect(x * SCALE, PAGE_H - (y + h) * SCALE, w * SCALE, h * SCALE, r * SCALE, stroke=0, fill=1)

    def img(self, image: ImageReader, x, y, w, h):
        self._c.drawImage(
            image, x * SCALE, PAGE_H - (y + h) * SCALE, w * SCALE, h * SCALE,
            preserveAspectRatio=False, mask="auto",
        )

    def glow_circle(self, cx, cy, r, color: Color):
        self._c.setFillAlpha(0.10)
        self._c.setFillColor(color)
        self._c.circle(cx * SCALE, PAGE_H - cy * SCALE, r * SCALE, stroke=0, fill=1)
        self._c.setFillAlpha(1.0)

    def width(self, text: str, size: float, bold: bool = False) -> float:
        return stringWidth(text, "Helvetica-Bold" if bold else "Helvetica", size)

    def wrap(self, text: str, maxw: float, size: float) -> List[str]:
        if text is None or not str(text).strip():
            return [""]
        out: List[str] = []
        line = ""
        for word in str(text).strip().split():
            test = word if not line else line + " " + word
            if self.width(test, size) <= maxw:
                line = test
            else:
                if line:
                    out.append(line)
                line = word
        if line:
            out.append(line)
        return out

    def fit(self, x, y, text: str, size: float, color: Color, bold: bool, maxw: float):
        t = "" if text is None else str(text)
        while len(t) > 2 and self.width(t, size, bold) > maxw:
            t = t[:-1]
        if t != text:
            t = t.strip() + "..."
        self.txt(x, y, t, size, color, bold)

    def txt(self, x, y, text: str, size: float, color: Color, bold: bool = False):
        self._c.setFont("Helvetica-Bold" if bold else "Helvetica", size * SCALE)
        self._c.setFillColor(color)
        self._c.drawString(x * SCALE, PAGE_H - y * SCALE, "" if text is None else str(text))


def _gradient(ac: AppPdf, w: float, h: float, c1: Color, c2: Color, steps: int = 18):
    step = h / steps
    for i in range(steps):
        ac.rect(0, i * step, w, step + 0.05, _lerp(c1, c2, i / steps))


def _logo(ac: AppPdf, left: float, top: float, size: float):
    ac.rr(left, top, size, size, size * 0.125, HexColor("#FFFFFF"))
    if _LOGO_IMAGE is None:
        ac.txt(left + size * 0.22, top + size * 0.62, "MGA", size * 0.34, GOLD, True)
        return
    inset = size * 0.045
    ac.img(_LOGO_IMAGE, left + inset, top + inset, size - 2 * inset, size - 2 * inset)


def _equipment_kpis(ac: AppPdf, data: List[Sequence], header_h: int, left: int):
    available = sum(1 for e in data if e[3] == "DISPONIBLE")
    operating = sum(1 for e in data if e[3] == "OPERATIVA")
    down = len(data) - available - operating
    band_bottom = 245
    usable = APP_W - left * 2
    gap = 10
    kw = (usable - gap * 3) / 4
    kt = header_h + 8
    kb = band_bottom - 8
    values = [
        ("DISPONIBLES", available, GREEN),
        ("OPERATIVOS", operating, AMBER),
        ("FUERA SERV.", down, RED),
        ("TOTAL", len(data), BLUE),
    ]
    for i, (label, value, color) in enumerate(values):
        x = left + (kw + gap) * i
        ac.rr(x, kt, kw, kb - kt, 9, HexColor("#FFFFFF"))
        ac.rr(x, kt, 5, kb - kt, 4, color)
        ac.txt(x + 14, kt + 23, str(value), 18, NAVY, True)
        ac.fit(x + 14, kb - 7, label, 9, LABEL, True, kw - 22)


def _equipment_header(ac: AppPdf, data: List[Sequence], project: str, page: int):
    _gradient(ac, APP_W, 185, NAVY, BLUE)
    ac.glow_circle(APP_W - 35, 0, 150, HexColor("#FFFFFF"))
    _logo(ac, 48, 22, 120)
    tx = 48 + 120 + 24
    ac.fit(tx, 64, "DISPONIBILIDAD DE EQUIPOS", 34, HexColor("#FFFFFF"), True, APP_W - tx - 48)
    ac.fit(tx, 105, project, 21, GOLD, True, APP_W - tx - 48)
    ac.txt(tx, 145, _fmt_ts(), 17, LIGHTGRAY, False)
    ac.rect(0, 185, APP_W, 60, HexColor("#F4F7FB"))
    _equipment_kpis(ac, data, 185, 48)


def _scoreboard(ac: AppPdf, left: int, top: int, kw: float, bottom: float, label: str, value: str, color: Color):
    ac.rr(left, top, kw, bottom - top, 9, _tint(HexColor("#FFFFFF"), 0.96))
    ac.rr(left, top, 5, bottom - top, 4, color)
    ac.fit(left + 14, top + 27, value, 18, NAVY, True, kw - 22)
    ac.fit(left + 14, bottom - 7, label, 9, LABEL, True, kw - 22)


def _generic_header(ac: AppPdf, title: str, project: str, kpis: List[tuple]):
    header_h = 250
    _gradient(ac, APP_W, header_h, NAVY, BLUE)
    ac.glow_circle(APP_W - 35, 0, 150, HexColor("#FFFFFF"))
    _logo(ac, 48, 22, 112)
    tx = 48 + 112 + 24
    ac.fit(tx, 63, title, 31, HexColor("#FFFFFF"), True, APP_W - tx - 48)
    ac.fit(tx, 104, project, 20, GOLD, True, APP_W - tx - 48)
    ac.txt(tx, 142, "Generado: " + _fmt_ts(), 17, HexColor("#DAE5EF"), False)
    usable = APP_W - 96
    gap = 10
    y = 174
    bottom = header_h - 10
    cw = (usable - gap * (len(kpis) - 1)) / len(kpis)
    for i, (label, value, color) in enumerate(kpis):
        _scoreboard(ac, 48 + (cw + gap) * i, y, cw, bottom, label, value, color)


def _footer(ac: AppPdf, left_text: str, page: int):
    ac.rect(0, APP_H - 45, APP_W, 45, FOOTBG)
    ac.txt(48, APP_H - 28, left_text, 15, GRAY, False)
    ac.txt(APP_W - 115, APP_H - 28, f"Página {page}", 15, GRAY, False)


def _tire_footer(ac: AppPdf, left_text: str, page: int):
    ac.rect(0, APP_H - 45, APP_W, 45, FOOTBG)
    ac.txt(48, APP_H - 28, left_text, 15, GRAY, False)
    ac.txt(APP_W - 115, APP_H - 28, f"Pagina {page}", 15, GRAY, False)


# ---------------------------------------------------------------------------
# DISPONIBILIDAD DE EQUIPOS  (ReportPainter)
# ---------------------------------------------------------------------------
def _equipment_columns() -> List[float]:
    usable = APP_W - 96
    return [usable * 0.27, usable * 0.11, usable * 0.15, usable * 0.20, usable * 0.27]


def _equipment_table_header(ac: AppPdf, l: float, r: float, y: int):
    cols = _equipment_columns()
    ac.rr(l, y, r - l, 34, 10, NAVY)
    ac.txt(l + 12, y + 23, "EQUIPO", 14, HexColor("#FFFFFF"), True)
    ac.txt(l + cols[0] + 10, y + 23, "NO. ECO", 14, HexColor("#FFFFFF"), True)
    ac.txt(l + cols[0] + cols[1] + 10, y + 23, "HORÓMETRO", 14, HexColor("#FFFFFF"), True)
    ac.txt(l + cols[0] + cols[1] + cols[2] + 10, y + 23, "CONDICIÓN", 14, HexColor("#FFFFFF"), True)
    ac.txt(l + cols[0] + cols[1] + cols[2] + cols[3] + 10, y + 23, "OBSERVACIONES", 14, HexColor("#FFFFFF"), True)
    return y + 40


def _category_bar(ac: AppPdf, l: float, r: float, y: int, category: str):
    ac.rr(l, y, r - l, 42, 8, CATBG)
    ac.rr(l, y, 7, 42, 4, RED)
    ac.txt(l + 20, y + 29, category, 18, BLUE, True)
    return y + 42


def _sup_bar(ac: AppPdf, l: float, r: float, y: int, label: str):
    name = str(label).strip() if str(label).strip() else "SIN SUPERVISOR"
    ac.rr(l, y, r - l, 40, 8, HexColor("#F7EFD6"))
    ac.rr(l, y, 7, 40, 4, GOLD)
    ac.fit(l + 20, y + 26, "Supervisor: " + name, 17, BLUE, True, r - l - 40)
    return y + 40


def _group_by(data: List[Sequence], key_index: int) -> List[tuple]:
    order: List[str] = []
    bucket: dict = {}
    for x in data:
        k = str(x[key_index]).strip() if str(x[key_index]).strip() else ""
        if k not in bucket:
            bucket[k] = []
            order.append(k)
        bucket[k].append(x)
    return [(k, bucket[k]) for k in order]


def _equipment_row(ac: AppPdf, cols: List[float], e: Sequence, l: float, r: float, y: int, h: int, index: int):
    line = 24.0
    if index % 2 == 1:
        ac.rect(l, y, r - l, h, ZEBRA)
    ac.rect(l, y, r - l, 1, LINE)
    ac.fit(l + 14, y + line, e[1], 17, NAVY, True, cols[0] - 24)
    ac.fit(l + cols[0] + 10, y + line, e[2], 16, DKGRAY, False, cols[1] - 18)
    hm = str(e[5]) if e[5] else ""
    if hm:
        hsz = 12.0
        hmax = cols[2] - 8
        th = hm
        while len(th) > 2 and ac.width(th, hsz, True) > hmax:
            th = th[:-1]
        if th != hm:
            th = th.strip() + "..."
        hw = ac.width(th, hsz, True)
        hx = l + cols[0] + cols[1] + 10
        hbox_half = hsz / 2 - 1
        ac.rr(hx - 4, y + line - hsz / 2 + 1, hw + 8, hsz + 6, 6, _tint(BLUE, 0.16))
        ac.txt(hx, y + hbox_half, th, hsz, BLUE, True)
    color = _status_colors(e[3])
    badge_l = l + cols[0] + cols[1] + cols[2] + 7
    badge_r = badge_l + cols[3] - 14
    badge_top = y + 5
    badge_bottom = min(y + h - 5, y + line + 7)
    ac.rr(badge_l, badge_top, badge_r - badge_l, badge_bottom - badge_top, 8, _tint(color, 0.118))
    ac.fit(badge_l + 8, y + line, e[3], 15, color, True, badge_r - badge_l - 16)
    notes = ac.wrap(e[4], cols[4] - 22, 15)
    ny = y + line
    for note in notes:
        ac.txt(l + cols[0] + cols[1] + cols[2] + cols[3] + 10, ny, note, 15, DKGRAY, False)
        ny += line


def _equipment_row_height(ac: AppPdf, cols: List[float], e: Sequence) -> int:
    lines = max(1, len(ac.wrap(e[4], cols[4] - 22, 15)))
    return max(46, lines * 24 + 18)


def build_equipment_pdf(project: str, rows: List[Sequence]) -> bytes:
    data = [(r[0], r[1], r[2], r[3], r[4] or "", r[5] or "") for r in rows]
    ac = AppPdf(project if project else "Equipos")
    cols = _equipment_columns()
    l = 48
    r = APP_W - 48
    start = 0
    total = len(data)
    page = 1
    if total == 0:
        _equipment_header(ac, data, project, 1)
        _equipment_table_header(ac, l, r, 260)
        t = "Sin equipos registrados en la base central"
        ac.txt((APP_W - ac.width(t, 22, True)) / 2, 420, t, 22, LABEL, True)
        _footer(ac, "Disponibilidad de Equipos", 1)
        return ac.finish()
    while start < total:
        if page > 1:
            ac.new_page()
        _equipment_header(ac, data, project, page)
        y = _equipment_table_header(ac, l, r, 260)
        i = start
        while i < total:
            e = data[i]
            new_category = i == start or e[0] != data[i - 1][0]
            cat_h = 42 if new_category else 0
            rh = _equipment_row_height(ac, cols, e)
            if i > start and y + cat_h + rh > APP_H - 55:
                break
            if new_category:
                y = _category_bar(ac, l, r, y, e[0])
            _equipment_row(ac, cols, e, l, r, y, rh, i)
            y += rh
            i += 1
        _footer(ac, "Disponibilidad de Equipos", page)
        start = max(i, start + 1)
        page += 1
    return ac.finish()


# ---------------------------------------------------------------------------
# CONSUMO DE LUBRICANTES  (ConsumptionPainter) - 4 columnas como el APK
# ---------------------------------------------------------------------------
def _consumption_columns() -> List[float]:
    usable = APP_W - 96
    return [usable * 0.25, usable * 0.30, usable * 0.16, usable * 0.29]


def _consumption_table_header(ac: AppPdf, cols: List[float]):
    ac.rr(48, 270, APP_W - 96, 36, 10, NAVY)
    x = 48
    for i, h in enumerate(["EQUIPO", "LUBRICANTE", "CANTIDAD", "OBSERVACIONES"]):
        ac.txt(x + 10, 294, h, 14, HexColor("#FFFFFF"), True)
        x += cols[i]
    return 270 + 36 + 6


def build_consumptions_pdf(project: str, rows: List[Sequence]) -> bytes:
    data = list(rows)
    ac = AppPdf(project if project else "Consumos")
    cols = _consumption_columns()
    line = 25.0
    total_l = sum(float(r[4]) for r in data if str(r[5]) == "L")
    total_kg = sum(float(r[4]) for r in data if str(r[5]) == "kg")
    kpis = [
        ("REGISTROS", str(len(data)), BLUE),
        ("TOTAL LITROS", f"{_fmt_number(total_l)} L", GREEN),
        ("TOTAL KG", f"{_fmt_number(total_kg)} kg", AMBER),
    ]
    if not data:
        _generic_header(ac, "CONSUMO DE LUBRICANTES", project, kpis)
        _consumption_table_header(ac, cols)
        ac.txt(348, 420, "Sin consumos registrados en la base central", 22, LABEL, True)
        _footer(ac, "MGA - Control de lubricantes", 1)
        return ac.finish()

    def row_h(x) -> int:
        notes = ac.wrap(x[6], cols[3] - 22, 15)
        return max(int(line * 2 + 14), (len(notes) + 1) * int(line) + 12)

    items: List[tuple] = []
    for k, group in _group_by(data, 7):
        items.append(("bar", k))
        for x in group:
            items.append(("row", x))

    start = 0
    page = 1
    row_no = 0
    while start < len(items):
        if page > 1:
            ac.new_page()
        _generic_header(ac, "CONSUMO DE LUBRICANTES", project, kpis)
        y = _consumption_table_header(ac, cols)
        i = start
        while i < len(items):
            kind, val = items[i]
            if kind == "bar":
                nxt = row_h(items[i + 1][1])
                if i > start and y + 40 + nxt > APP_H - 55:
                    break
                y = _sup_bar(ac, 48, APP_W - 48, y, val)
                i += 1
                continue
            x = val
            rh = row_h(x)
            if i > start and y + rh > APP_H - 55:
                break
            if row_no % 2 == 1:
                ac.rect(48, y, APP_W - 96, rh, ZEBRA)
            ac.rect(48, y, APP_W - 96, 1, LINE)
            ac.fit(48 + 12, y + line, x[1], 17, NAVY, True, cols[0] - 22)
            ac.fit(48 + 12, y + line * 2, x[2], 14, DKGRAY, False, cols[0] - 22)
            ac.fit(48 + cols[0] + 10, y + line, x[3], 17, NAVY, True, cols[1] - 20)
            al = 48 + cols[0] + cols[1] + 7
            ar = al + cols[2] - 14
            ac.rr(al, y + 5, ar - al, min(rh - 10, y + line + 8 - (y + 5)), 8, _tint(GREEN, 0.118))
            ac.fit(al + 8, y + line, f"{_fmt_number(float(x[4]))} {x[5]}", 16, GREEN, True, ar - al - 16)
            meta = ""
            if x[8]:
                meta += f"Turno {x[8]}  ·  "
            meta += _fmt_date_short(x[0])
            if x[7]:
                meta += "  ·  " + str(x[7])
            ny = y + line
            ac.fit(48 + cols[0] + cols[1] + cols[2] + 10, ny, meta, 12, BLUE, False, cols[3] - 22)
            ny += line
            for note in ac.wrap(x[6], cols[3] - 22, 15):
                ac.txt(48 + cols[0] + cols[1] + cols[2] + 10, ny, note, 15, DKGRAY, False)
                ny += line
            y += rh
            i += 1
            row_no += 1
        _footer(ac, "MGA - Control de lubricantes", page)
        start = max(i, start + 1)
        page += 1
    return ac.finish()


# ---------------------------------------------------------------------------
# REPORTE MECÁNICO  (MechanicPainter)
# ---------------------------------------------------------------------------
def _mech_card(ac: AppPdf, x: Sequence, l: float, right: float, y: int, h: int, font: float, line: float):
    ac.rr(l, y, right - l, h, 14, HexColor("#FFFFFF"))
    status = _mech_status_color(x[9])
    ac.rr(l, y, 7, h, 4, status)
    ac.fit(l + 22, y + 34, f"{x[0]} · {x[1]}", font + 3, NAVY, True, right - l - 210)
    badge_r = right - 16
    badge_l = badge_r - 170
    ac.rr(badge_l, y + 12, badge_r - badge_l, 31, 9, _tint(status, 0.118))
    ac.fit(badge_l + 10, y + 33, str(x[9]).upper(), font - 3, status, True, badge_r - badge_l - 20)
    info = f"Mecánico: {x[3]}  ·  {x[5]}"
    if x[4]:
        info += "  ·  Horómetro: " + str(x[4])
    if len(x) > 2 and x[2]:
        info += "  ·  " + _fmt_date_short(x[2])
    if len(x) > 11 and x[11]:
        info += "  ·  Supervisor: " + str(x[11])
    ac.fit(l + 22, y + 62, info, font - 2, CARDGRAY, False, right - l - 44)
    cy = y + 90
    cy = _mech_section(ac, "FALLA REPORTADA", x[6], l + 22, right - 18, cy, font, line)
    cy = _mech_section(ac, "TRABAJO REALIZADO", x[7], l + 22, right - 18, cy, font, line)
    cy = _mech_section(ac, "REFACCIONES", x[8] if x[8] else "Sin refacciones registradas", l + 22, right - 18, cy, font, line)
    if x[10]:
        ac.txt(right - 285, y + h - 18, "EVIDENCIA FOTOGRÁFICA ADJUNTA", font - 4, BLUE, True)


def _mech_section(ac: AppPdf, label: str, value: str, l: float, r: float, y: int, font: float, line: float) -> int:
    ac.txt(l, y, label, font - 4, BLUE, True)
    cy = y + line
    for s in ac.wrap(value, r - l, font - 2):
        ac.txt(l, cy, s, font - 2, DKGRAY, False)
        cy += line
    return int(cy + 8)


def build_mechanic_pdf(project: str, rows: List[Sequence]) -> bytes:
    data = list(rows)
    ac = AppPdf(project if project else "Reportes mecánicos")
    correctives = sum(1 for r in data if r[5] == "Correctivo")
    pending = sum(1 for r in data if str(r[9]) in ("Pendiente", "Fuera de servicio"))
    kpis = [
        ("REPORTES", str(len(data)), BLUE),
        ("CORRECTIVOS", str(correctives), AMBER),
        ("PENDIENTES", str(pending), RED),
    ]
    font = 16.0
    line = 23.0
    l = 48
    right = APP_W - 48
    if not data:
        _generic_header(ac, "REPORTE MECÁNICO", project, kpis)
        ac.txt(354, 420, "Sin reportes mecánicos registrados en la base central", 22, LABEL, True)
        _footer(ac, "MGA - Reporte mecánico", 1)
        return ac.finish()

    def card_h(r) -> int:
        text_w = right - l - 34
        failure = len(ac.wrap(r[6], text_w, font - 2))
        work = len(ac.wrap(r[7], text_w, font - 2))
        parts = len(ac.wrap(r[8] if r[8] else "Sin refacciones registradas", text_w, font - 2))
        return int(210 + (failure + work + parts) * line)

    items: List[tuple] = []
    for k, group in _group_by(data, 11):
        items.append(("bar", k))
        for x in group:
            items.append(("row", x))

    start = 0
    page = 1
    while start < len(items):
        if page > 1:
            ac.new_page()
        _generic_header(ac, "REPORTE MECÁNICO", project, kpis)
        y = 270
        i = start
        while i < len(items):
            kind, val = items[i]
            if kind == "bar":
                nxt = card_h(items[i + 1][1])
                if i > start and y + 40 + nxt > APP_H - 60:
                    break
                y = _sup_bar(ac, l, right, y, val)
                i += 1
                continue
            r = val
            rh = card_h(r)
            if i > start and y + rh > APP_H - 60:
                break
            _mech_card(ac, r, l, right, y, rh, font, line)
            y += rh + 12
            i += 1
        _footer(ac, "MGA - Reporte mecánico", page)
        start = max(i, start + 1)
        page += 1
    return ac.finish()


# ---------------------------------------------------------------------------
# REVISION DE LLANTAS - REPORTE GENERAL  (TirePainter.drawGeneral)
# ---------------------------------------------------------------------------
def _tire_header(ac: AppPdf, project: str):
    _gradient(ac, APP_W, 190, NAVY, BLUE)
    ac.glow_circle(APP_W - 35, 0, 150, HexColor("#FFFFFF"))
    _logo(ac, 48, 24, 112)
    tx = 185
    ac.fit(tx, 72, "REPORTE GENERAL DE REVISION DE LLANTAS", 30, HexColor("#FFFFFF"), True, APP_W - 245)
    ac.fit(tx, 110, project, 18, GOLD, True, APP_W - 245)
    ac.txt(tx, 142, "Generado: " + _fmt_ts(), 14, LIGHTGRAY, False)


def _tire_kpi(ac: AppPdf, x: float, y: float, w: float, label: str, value: str, color: Color):
    ac.rr(x, y, w, 84, 10, HexColor("#F8FAFC"))
    ac.rr(x, y, 6, 84, 4, color)
    ac.fit(x + 16, y + 45, value, 17, NAVY, True, w - 28)
    ac.fit(x + 16, y + 67, label, 9, LABEL, True, w - 28)


def _tire_pct(value: int, total: int) -> str:
    return "0" if total == 0 else f"{value} / {round(100 * value / total)}%"


def _tire_table_head(ac: AppPdf, l: float, y: float, cols: List[float], labels: List[str]):
    total_w = sum(cols)
    ac.rr(l, y, total_w, 42, 9, NAVY)
    x = l
    for i, label in enumerate(labels):
        ac.fit(x + 8, y + 27, label, 12, HexColor("#FFFFFF"), True, cols[i] - 14)
        x += cols[i]
    return y + 42


def _tire_band(ac: AppPdf, l: float, r: float, y: int, equipment: str, eco: str, hour_meter: str):
    ac.rect(l, y, r - l, 40, TIREBAND)
    ac.rect(l, y, 5, 40, BLUE)
    ac.fit(l + 16, y + 26, equipment, 14, NAVY, True, 400)
    ac.txt(l + 430, y + 26, "Serie/ID: " + eco, 12, DKGRAY, False)
    hm = "-" if not hour_meter else str(hour_meter) + " h"
    ac.fit(r - 250, y + 26, "Horometro: " + hm, 12, DKGRAY, False, 242)


def _tire_row(ac: AppPdf, l: float, r: float, y: int, cols: List[float], detail: Sequence, index: int):
    if index % 2 == 1:
        ac.rect(l, y, r - l, 32, ZEBRA)
    ac.rect(l, y, r - l, 1, LINE)
    x = l + 8
    ac.fit(x, y + 21, detail[0], 12, NAVY, True, cols[0] - 16)
    x += cols[0]
    sc = _tire_status_color(detail[1])
    ac.rr(x, y + 5, 130, 22, 6, _tint(sc, 0.13))
    ac.fit(x + 10, y + 20, detail[1], 11, sc, True, 110)
    x += cols[1]
    vida = "-" if not detail[2] else str(detail[2]) + " %"
    ac.txt(x + 10, y + 21, vida, 13, NAVY, True)
    x += cols[2]
    ac.fit(x + 6, y + 21, detail[3], 11, DKGRAY, False, cols[3] - 12)


def build_tires_pdf(project: str, rows: List[Sequence]) -> bytes:
    ac = AppPdf(project if project else "Inspecciones de llantas")
    margin = 48
    w = APP_W
    h = APP_H
    cols = [240, 260, 160, w - margin * 2 - 660]

    seen = set()
    records = []
    for r in rows:
        equipment = str(r[1] or "")
        if equipment in seen:
            continue
        seen.add(equipment)
        records.append(r)

    counts = [0, 0, 0, 0]
    for r in records:
        for d in r[8]:
            s = d[1]
            if s == "Bueno":
                counts[0] += 1
            elif s == "Aceptable":
                counts[1] += 1
            elif s == "Regular":
                counts[2] += 1
            else:
                counts[3] += 1
    tire_total = sum(counts)

    idx = 0
    page = 1
    first = True
    if not records:
        _tire_header(ac, project)
        gap = 12
        kw = (w - margin * 2 - gap * 5) / 6
        y = 220
        _tire_kpi(ac, margin, y, kw, "EQUIPOS", "0", BLUE)
        _tire_kpi(ac, margin + (kw + gap), y, kw, "LLANTAS", "0", NAVY)
        _tire_kpi(ac, margin + (kw + gap) * 2, y, kw, "BUENAS", "0", GREEN)
        _tire_kpi(ac, margin + (kw + gap) * 3, y, kw, "ACEPTABLES", "0", GOLD)
        _tire_kpi(ac, margin + (kw + gap) * 4, y, kw, "REGULARES", "0", AMBER)
        _tire_kpi(ac, margin + (kw + gap) * 5, y, kw, "MALAS", "0", RED)
        y += 112
        ac.txt(margin, y, "DETALLE POR EQUIPO", 18, NAVY, True)
        y += 24
        y = _tire_table_head(ac, margin, y, cols, ["Posicion", "Estado", "Vida Util", "Observaciones"])
        ac.txt(margin + 120, 430, "Sin revisiones de llantas en la base central", 22, LABEL, True)
        _tire_footer(ac, "MGA - Revision de llantas general", 1)
        return ac.finish()

    while idx < len(records) or first:
        if page > 1:
            ac.new_page()
        _tire_header(ac, project)
        y = 220
        if first:
            gap = 12
            kw = (w - margin * 2 - gap * 5) / 6
            _tire_kpi(ac, margin, y, kw, "EQUIPOS", str(len(records)), BLUE)
            _tire_kpi(ac, margin + (kw + gap), y, kw, "LLANTAS", str(tire_total), NAVY)
            _tire_kpi(ac, margin + (kw + gap) * 2, y, kw, "BUENAS", _tire_pct(counts[0], tire_total), GREEN)
            _tire_kpi(ac, margin + (kw + gap) * 3, y, kw, "ACEPTABLES", _tire_pct(counts[1], tire_total), GOLD)
            _tire_kpi(ac, margin + (kw + gap) * 4, y, kw, "REGULARES", _tire_pct(counts[2], tire_total), AMBER)
            _tire_kpi(ac, margin + (kw + gap) * 5, y, kw, "MALAS", _tire_pct(counts[3], tire_total), RED)
            y += 112
            ac.txt(margin, y, "DETALLE POR EQUIPO", 18, NAVY, True)
            y += 24
        else:
            y = 214
            ac.txt(margin, y - 14, "DETALLE POR EQUIPO (continuacion)", 14, NAVY, True)
        y = _tire_table_head(ac, margin, y, cols, ["Posicion", "Estado", "Vida Util", "Observaciones"])
        start_idx = idx
        while idx < len(records):
            r = records[idx]
            details = r[8]
            needed = 40 + len(details) * 32 + 12
            if idx > start_idx and y + needed > h - 58:
                break
            _tire_band(ac, margin, w - margin, y, str(r[1] or ""), str(r[0] or ""), str(r[3] or ""))
            y += 40
            for i, d in enumerate(details):
                _tire_row(ac, margin, w - margin, y, cols, d, i)
                y += 32
            y += 12
            idx += 1
        _tire_footer(ac, "MGA - Revision de llantas general", page)
        first = False
        page += 1
    return ac.finish()
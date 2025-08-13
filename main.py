# main.py
from fastapi import FastAPI, Request, Form, Response
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage  # <-- ÚJ: képek beszúrásához

from datetime import datetime, time
from io import BytesIO
import os
import uuid

# PIL (Pillow) lehet, hogy nincs telepítve: óvatos import + fallback
try:
    from PIL import Image as PILImage, ImageDraw, ImageFont
    HAS_PIL = True
except Exception:
    HAS_PIL = False

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- helpers for merged cells ----------
def merged_ranges(ws):
    return [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]

def in_range(rng, r, c):
    r1, c1, r2, c2 = rng
    return (r1 <= r <= r2) and (c1 <= c <= c2)

def block_of(ws, r, c):
    for rng in merged_ranges(ws):
        if in_range(rng, r, c):
            return rng
    return (r, c, r, c)

def top_left_of_block(ws, r, c):
    r1, c1, _, _ = block_of(ws, r, c)
    return r1, c1

def set_text(ws, r, c, text, wrap=False, align_left=False, valign_top=False):
    rr, cc = top_left_of_block(ws, r, c)
    cell = ws.cell(row=rr, column=cc)
    cell.value = text
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal=("left" if align_left else "center"),
        vertical=("top" if valign_top else cell.alignment.vertical or "center"),
    )

def set_text_addr(ws, addr, text, *, wrap=False, horizontal="left", vertical="center"):
    cell = ws[addr]
    rr, cc = top_left_of_block(ws, cell.row, cell.column)
    tgt = ws.cell(row=rr, column=cc)
    tgt.value = text
    tgt.alignment = Alignment(wrap_text=wrap, horizontal=horizontal, vertical=vertical)

# ---------- time & hours ----------
def parse_hhmm(s: str) -> time | None:
    s = (s or "").strip()
    if not s:
        return None
    hh, mm = s.split(":")
    return time(int(hh), int(mm))

def overlap_minutes(a1: time, a2: time, b1: time, b2: time) -> int:
    dt = datetime(2000,1,1)
    A1 = dt.replace(hour=a1.hour, minute=a1.minute)
    A2 = dt.replace(hour=a2.hour, minute=a2.minute)
    B1 = dt.replace(hour=b1.hour, minute=b1.minute)
    B2 = dt.replace(hour=b2.hour, minute=b2.minute)
    start = max(A1, B1)
    end   = min(A2, B2)
    if end <= start:
        return 0
    return int((end - start).total_seconds() // 60)

# ---- módosítva: opcionális pause percben (60 alapértelmezés, vagy 30) ----
def hours_with_breaks(beg: time | None, end: time | None, pause_min: int = 60) -> float:
    if not beg or not end:
        return 0.0
    dt = datetime(2000,1,1)
    start = dt.replace(hour=beg.hour, minute=beg.minute)
    finish = dt.replace(hour=end.hour, minute=end.minute)
    if finish <= start:
        return 0.0

    total_min = int((finish - start).total_seconds() // 60)

    if pause_min >= 60:
        # csak a valós átfedést vonjuk le a 09:00–09:15 és 12:00–12:45 sávokból
        minus = overlap_minutes(beg, end, time(9,0), time(9,15)) \
              + overlap_minutes(beg, end, time(12,0), time(12,45))
    else:
        # félórás opció: fix 30 percet vonunk (de legfeljebb a teljes tartamot)
        minus = min(total_min, 30)

    return max(0.0, (total_min - minus) / 60.0)

# ---------- table helpers ----------
def find_header_positions(ws):
    pos = {}
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=120):
        for cell in row:
            v = cell.value
            if isinstance(v, str):
                t = v.strip()
                if t == "Name":
                    pos["name_col"] = cell.column
                    header_row = cell.row
                if t == "Vorname":
                    pos["vorname_col"] = cell.column
                if "Ausweis" in t or "Kennzeichen" in t:
                    pos["ausweis_col"] = cell.column
                if t == "Beginn":
                    pos["beginn_col"] = cell.column
                    pos["subheader_row"] = cell.row
                if t == "Ende":
                    pos["ende_col"] = cell.column
                if "Anzahl Stunden" in t:
                    pos["stunden_col"] = cell.column
                # ÚJ: Vorhaltung oszlop felismerése
                if t.lower().startswith("vorhaltung") or "beauftragtes gerät" in t.lower():
                    pos["vorhaltung_col"] = cell.column
        if all(k in pos for k in ["name_col","vorname_col","ausweis_col","beginn_col","ende_col","stunden_col","subheader_row"]):
            # nem feltétlen várjuk meg a vorhaltung_col-t, ha nincs, nem kötelező
            break
    pos["data_start_row"] = pos.get("subheader_row", header_row) + 1
    return pos

def find_total_cells(ws, stunden_col):
    """
    Visszaad:
      - right_of_label: a 'Gesamtstunden' felirat melletti kis cella (amit ürítünk)
      - stunden_total:  ugyanazon a soron a 'Anzahl Stunden' oszlop alatti cella (ebbe írjuk az összeget)
    """
    total_row = None
    right_of_label = None
    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if isinstance(cell.value, str) and "Gesamtstunden" in cell.value:
                total_row = cell.row
                r_neighbor, c_neighbor = total_row, cell.column + 1
                rr, cc = top_left_of_block(ws, r_neighbor, c_neighbor)
                right_of_label = (rr, cc)
                break
        if total_row:
            break

    stunden_total = None
    if total_row:
        rr, cc = top_left_of_block(ws, total_row, stunden_col)
        stunden_total = (rr, cc)

    return right_of_label, stunden_total

def find_big_description_block(ws):
    best = None
    for (r1,c1,r2,c2) in merged_ranges(ws):
        height = r2 - r1 + 1
        width  = c2 - c1 + 1
        if r1 >= 6 and height >= 4 and width >= 4:
            area = height * width
            if not best or area > best[-1]:
                best = (r1, c1, r2, c2, area)
    if best:
        r1, c1, r2, c2, _ = best
        return (r1, c1, r2, c2)
    return (6, 1, 20, 8)

# ---- ÚJ: Beschreibung → kép render + beszúrás
def _excel_col_width_to_px(width):
    """
    Excel oszlopszélesség becslése pixelekre.
    1 egység kb. 7 px (szokásos közelítés).
    """
    try:
        return int(round((width or 8.43) * 7))
    except Exception:
        return 60  # fallback

def _excel_row_height_to_px(height_points):
    """
    Pont → pixel (96 DPI ~ 1.333 px / pont).
    """
    try:
        pts = height_points if height_points else 15  # default row height ~15 pt
        return int(round(pts * 96.0 / 72.0))
    except Exception:
        return 20

def _get_block_pixel_size(ws, r1, c1, r2, c2):
    # szélesség px
    w_px = 0
    for c in range(c1, c2 + 1):
        cd = ws.column_dimensions.get(ws.cell(row=1, column=c).column_letter)
        w_excel = cd.width if cd and cd.width is not None else 8.43
        w_px += _excel_col_width_to_px(w_excel)
    # magasság px
    h_px = 0
    for r in range(r1, r2 + 1):
        rd = ws.row_dimensions.get(r)
        h_pt = rd.height if rd and rd.height is not None else 15
        h_px += _excel_row_height_to_px(h_pt)
    # kicsi belső margó a képhez
    return max(50, w_px - 8), max(40, h_px - 8)

def _render_text_image(text, w_px, h_px):
    """
    Kép generálása a leírásból (fehér háttér, fekete szöveg, lágy sortördelés).
    Ha nincs PIL, visszaad None-t.
    """
    if not HAS_PIL:
        return None

    # Betűkészlet: próbálunk egy elterjedtet, majd fallback a defaultra
    font = None
    for fp in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
               "/Library/Fonts/Arial.ttf",
               "arial.ttf"):
        try:
            font = ImageFont.truetype(fp, 14)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()

    img = PILImage.new("RGB", (w_px, h_px), "white")
    draw = ImageDraw.Draw(img)

    # Szöveg tördelése a rendelkezésre álló szélességre
    padding = 10
    max_w = w_px - padding * 2
    max_h = h_px - padding * 2
    words = (text or "").replace("\r\n", "\n").replace("\r", "\n").split()
    lines = []
    line = ""

    def text_w(s): 
        # textbbox pontosabb, de fallback a textlength-re
        try:
            bbox = draw.textbbox((0,0), s, font=font)
            return bbox[2] - bbox[0]
        except Exception:
            return int(draw.textlength(s, font=font))

    for w in words:
        test = (line + " " + w).strip()
        if text_w(test) <= max_w:
            line = test
        else:
            if line:
                lines.append(line)
            line = w
    if line:
        lines.append(line)

    # sorok kirajzolása (nem engedjük túlfolyni a blokkot)
    y = padding
    line_h = (font.size + 6)
    for ln in lines:
        if y + line_h > padding + max_h:
            break
        draw.text((padding, y), ln, fill="black", font=font)
        y += line_h

    # mentés memóriába
    bio = BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio

def insert_description_as_image(ws, r1, c1, r2, c2, text):
    """
    A megadott blokkban (r1..r2, c1..c2) képet szúr be a szöveg helyett.
    Ha PIL nem elérhető, visszatér False-szal (és a hívó fél tehet fallbacket).
    """
    if not HAS_PIL:
        return False

    # blokk méret pixelekben
    w_px, h_px = _get_block_pixel_size(ws, r1, c1, r2, c2)

    img_buf = _render_text_image(text, w_px, h_px)
    if not img_buf:
        return False

    # képként beszúrni
    xl_img = XLImage(img_buf)
    # Horgony a blokk bal felső cellájához
    anchor_cell = ws.cell(row=r1, column=c1).coordinate
    xl_img.anchor = anchor_cell
    ws.add_image(xl_img)

    # a cellák értékét üresre tesszük, hogy ne üssön össze a képpel
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).value = None
    return True

# ---------- routes ----------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    basf_beauftragter: str = Form(""),
    # a régi „geraet” mezőt már nem használjuk a fejléchez, helyette soronkénti Vorhaltung van
    geraet: str = Form(""),
    beschreibung: str = Form(""),
    break_minutes: int = Form(60),
    # dolgozók + ÚJ: vorhaltung1..5
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""), vorhaltung1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""), vorhaltung2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""), vorhaltung3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""), vorhaltung4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""), vorhaltung5: str = Form(""),
):
    wb = load_workbook(os.path.join(os.getcwd(), "GP-t.xlsx"))
    ws = wb.active

    # --- Felső mezők: CSAK a dátum formázása változik (német, szövegként) ---
    date_text = datum
    try:
        dt = datetime.strptime(datum.strip(), "%Y-%m-%d")
        date_text = dt.strftime("%d.%m.%Y")   # pl. 11.08.2025
    except Exception:
        pass

    set_text_addr(ws, "B2", date_text, horizontal="left")
    set_text_addr(ws, "B3", bau,        horizontal="left")
    if (basf_beauftragter or "").strip():
        set_text_addr(ws, "E3", basf_beauftragter, horizontal="left")

    # --- Beschreibung: KÉP beszúrása a nagy blokkra; ha nem megy, fallback a régi szövegre ---
    r1, c1, r2, c2 = find_big_description_block(ws)
    inserted = insert_description_as_image(ws, r1, c1, r2, c2, beschreibung)
    if not inserted:
        # fallback: a régi megoldás (szöveg + wrap + fix sormagasság)
        for r in range(r1, r2 + 1):
            ws.row_dimensions[r].height = 22
        set_text(ws, r1, c1, beschreibung, wrap=True, align_left=True, valign_top=True)

    # --- Dolgozók és órák + Vorhaltung oszlop ---
    pos = find_header_positions(ws)
    row = pos["data_start_row"]
    vorhaltung_col = pos.get("vorhaltung_col", None)

    workers = []
    for i in range(1, 5 + 1):
        vn = locals().get(f"vorname{i}", "") or ""
        nn = locals().get(f"nachname{i}", "") or ""
        aw = locals().get(f"ausweis{i}", "") or ""
        bg = locals().get(f"beginn{i}", "") or ""
        en = locals().get(f"ende{i}", "") or ""
        vh = locals().get(f"vorhaltung{i}", "") or ""
        if not (vn or nn or aw or bg or en or vh):
            continue
        workers.append((vn, nn, aw, bg, en, vh))

    total_hours = 0.0
    for (vn, nn, aw, bg, en, vh) in workers:
        set_text(ws, row, pos["name_col"], nn, wrap=False, align_left=True)
        set_text(ws, row, pos["vorname_col"], vn, wrap=False, align_left=True)
        set_text(ws, row, pos["ausweis_col"], aw, wrap=False, align_left=True)
        set_text(ws, row, pos["beginn_col"], bg, wrap=False, align_left=True)
        set_text(ws, row, pos["ende_col"], en, wrap=False, align_left=True)

        # Vorhaltung az adott sorban, ha van ilyen oszlop a sablonban
        if vorhaltung_col and (vh or "").strip():
            set_text(ws, row, vorhaltung_col, vh, wrap=True, align_left=True, valign_top=True)

        hb = parse_hhmm(bg)
        he = parse_hhmm(en)
        h = round(hours_with_breaks(hb, he, int(break_minutes)), 2)
        total_hours += h
        set_text(ws, row, pos["stunden_col"], h, wrap=False, align_left=True)
        row += 1

    # --- Összóraszám: jobb oldali nagy dobozban; a kis mezőt ürítjük ---
    right_of_label, stunden_total = find_total_cells(ws, pos["stunden_col"])
    if stunden_total:
        tr, tc = stunden_total
        set_text(ws, tr, tc, round(total_hours, 2), wrap=False, align_left=True)
    if right_of_label:
        rr, rc = right_of_label
        set_text(ws, rr, rc, "", wrap=False, align_left=True)

    # ---- Fix méretű válasz, Content-Length-cel (marad a te módosításod) ----
    bio = BytesIO()
    wb.save(bio)
    data = bio.getvalue()
    fname = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{fname}"',
        "Content-Length": str(len(data)),
        "Cache-Control": "no-store",
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

# main.py
from fastapi import FastAPI, Request, Form, Response
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

from PIL import Image as PILImage, ImageDraw, ImageFont

from datetime import datetime, time
from io import BytesIO
import os
import uuid
import math
import textwrap

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

# ---- opcionális pause percben (60 alapértelmezés, vagy 30) ----
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
        minus = overlap_minutes(beg, end, time(9,0), time(9,15)) \
              + overlap_minutes(beg, end, time(12,0), time(12,45))
    else:
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
                if t.lower().startswith("vorhaltung") or "beauftragtes gerät" in t.lower():
                    pos["vorhaltung_col"] = cell.column
        if all(k in pos for k in ["name_col","vorname_col","ausweis_col","beginn_col","ende_col","stunden_col","subheader_row"]):
            break
    pos["data_start_row"] = pos.get("subheader_row", header_row) + 1
    return pos

def find_total_cells(ws, stunden_col):
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

# ---------- pixel becslések & kép-beillesztés (Beschreibung blokk) ----------
def _col_width_px(ws, c: int) -> int:
    # Excel szélesség (character units) -> px közelítés
    cd = ws.column_dimensions.get(get_column_letter(c))
    w = getattr(cd, "width", None)
    if w is None:
        w = 8.43  # Excel alap
    return int(w * 7 + 5)

def _row_height_px(ws, r: int) -> int:
    # Pont -> px (96 DPI)
    rd = ws.row_dimensions.get(r)
    h_pt = getattr(rd, "height", None) or 15.0  # Excel alap ~15pt
    return int(h_pt * 96 / 72)

def _get_block_pixel_size(ws, r1, c1, r2, c2):
    width = sum(_col_width_px(ws, c) for c in range(c1, c2 + 1))
    height = sum(_row_height_px(ws, r) for r in range(r1, r2 + 1))
    return width, height

def insert_description_as_image(ws, r1, c1, r2, c2, text: str) -> bool:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not text.strip():
        return False

    # blokk méret px-ben, belső margó
    W, H = _get_block_pixel_size(ws, r1, c1, r2, c2)
    pad = 10
    W_i = max(50, W - 2 * pad)
    H_i = max(30, H - 2 * pad)

    # fehér vászon
    img = PILImage.new("RGB", (max(1, W), max(1, H)), "white")
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None

    # becsült karakter/sor szám a szélességre
    if font:
        avg_char_px = max(6, int(draw.textlength("ABCDEFGHIJKLMNOPQRSTUVWXYZ", font=font) / 26))
        max_chars = max(10, int(W_i / avg_char_px))
        line_h = max(12, int(font.getbbox("Ay")[3] - font.getbbox("Ay")[1] + 4))
    else:
        max_chars = max(10, int(W_i / 7))
        line_h = 16

    lines = []
    for paragraph in text.split("\n"):
        paragraph = paragraph.strip()
        if not paragraph:
            lines.append("")
            continue
        wrapped = textwrap.wrap(paragraph, width=max_chars, break_long_words=False)
        if not wrapped:
            lines.append("")
        else:
            lines.extend(wrapped)

    # szöveg rajzolása
    x = pad
    y = pad
    for ln in lines:
        if y + line_h > H - pad:
            break  # ne folyjon ki a blokkból
        draw.text((x, y), ln, fill="black", font=font)
        y += line_h

    # kép beillesztése a blokk bal felső sarkába
    xl_img = XLImage(img)
    xl_img.width, xl_img.height = W, H
    anchor = f"{get_column_letter(c1)}{r1}"
    ws.add_image(xl_img, anchor)

    # a cella értékét ürítsük, ne zavarjon
    rr, cc = top_left_of_block(ws, r1, c1)
    ws.cell(row=rr, column=cc).value = ""
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

    # --- Felső mezők: dátum szövegként, német formátumban ---
    date_text = datum
    try:
        dt = datetime.strptime(datum.strip(), "%Y-%m-%d")
        date_text = dt.strftime("%d.%m.%Y")
    except Exception:
        pass

    set_text_addr(ws, "B2", date_text, horizontal="left")
    set_text_addr(ws, "B3", bau,        horizontal="left")
    if (basf_beauftragter or "").strip():
        set_text_addr(ws, "E3", basf_beauftragter, horizontal="left")

    # --- Beschreibung: nagy blokk + kép-beillesztés (Numbers/Safari kompatibilis) ---
    r1, c1, r2, c2 = find_big_description_block(ws)
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = 22  # sorok maradjanak egységesek

    # próbáljuk képként
    try:
        inserted = insert_description_as_image(ws, r1, c1, r2, c2, beschreibung)
    except Exception:
        inserted = False

    # ha valamiért nem sikerült, essünk vissza a sima wrap-elt szövegre
    if not inserted:
        set_text(ws, r1, c1, beschreibung, wrap=True, align_left=True, valign_top=True)

    # --- Dolgozók és órák + Vorhaltung oszlop ---
    pos = find_header_positions(ws)
    row = pos["data_start_row"]
    vorhaltung_col = pos.get("vorhaltung_col", None)

    workers = []
    for i in range(1, 6):
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

    # ---- fix méretű válasz, Content-Length-cel (iOS letöltési gond ellen) ----
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

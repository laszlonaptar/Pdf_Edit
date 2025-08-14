# main.py
from fastapi import FastAPI, Request, Form, Response
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

from datetime import datetime, time
from io import BytesIO
import os
import uuid
import textwrap
import traceback

# Képgeneráláshoz
try:
    from PIL import Image as PILImage, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except Exception as e:
    PIL_AVAILABLE = False
    print("IMG: PIL not available ->", repr(e))

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

# --- Fix Beschreibung-blokk: A6–G15 ---
def find_description_block(ws):
    # A = 1, G = 7; 6. sortól 15. sorig
    return (6, 1, 15, 7)

# ---------- Bild (Beschreibung) helpers ----------
def _excel_col_width_to_pixels(width):
    if width is None:
        width = 8.43  # Excel default
    return int(round(7 * width + 5))

def _excel_row_height_to_pixels(height):
    if height is None:
        height = 15.0  # Excel default ~ 20px
    return int(round(height * 96 / 72))

def _get_block_pixel_size(ws, r1, c1, r2, c2):
    w_px = 0
    for c in range(c1, c2 + 1):
        letter = get_column_letter(c)
        cd = ws.column_dimensions.get(letter)
        w_px += _excel_col_width_to_pixels(getattr(cd, "width", None))
    h_px = 0
    for r in range(r1, r2 + 1):
        rd = ws.row_dimensions.get(r)
        h_px += _excel_row_height_to_pixels(getattr(rd, "height", None))
    w_px = max(40, w_px)
    h_px = max(40, h_px)
    return w_px, h_px

def _make_description_image(text, w_px, h_px):
    # FEHÉR háttér – pontosan a blokk mérete
    img = PILImage.new("RGB", (w_px, h_px), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    font = None
    for name in ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"]:
        try:
            font = ImageFont.truetype(name, 14)
            break
        except Exception:
            font = None
    if font is None:
        font = ImageFont.load_default()

    # belső margók: bal=12, jobb=0, fent=10, lent=10
    pad_left, pad_top, pad_right, pad_bottom = 12, 10, 0, 10
    avail_w = max(10, w_px - (pad_left + pad_right))
    avail_h = max(10, h_px - (pad_top + pad_bottom))

    # bátrabb tördelés: kb. 90% átlag-karakterszélességgel számolunk -> hosszabb sorok
    sample = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,:;-"
    avg_w = max(6, sum(draw.textlength(ch, font=font) for ch in sample) / len(sample))
    avg_w_eff = avg_w * 0.90
    max_chars_per_line = max(10, int(avail_w / avg_w_eff))

    paragraphs = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = []
    for para in paragraphs:
        if not para:
            lines.append("")
            continue
        lines.extend(
            textwrap.wrap(
                para,
                width=max_chars_per_line,
                replace_whitespace=False,
                break_long_words=False,
            )
        )

    ascent, descent = font.getmetrics()
    line_h = ascent + descent + 4
    max_lines = max(1, int(avail_h // line_h))

    if len(lines) > max_lines:
        lines = lines[:max_lines - 1] + ["…"]

    x = pad_left
    y = pad_top
    for ln in lines:
        draw.text((x, y), ln, fill=(0, 0, 0), font=font)
        y += line_h

    return img

def _xlimage_from_pil(pil_img):
    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    return XLImage(buf)

def insert_description_as_image(ws, r1, c1, r2, c2, text):
    """Próbál képet beszúrni. Siker log + True, hiba esetén kivétel log és False."""
    if not PIL_AVAILABLE:
        print("IMG: PIL not available at runtime")
        return False
    try:
        text_s = (text or "")
        print(f"IMG: will insert, text_len={len(text_s)} at {get_column_letter(c1)}{r1}-{get_column_letter(c2)}{r2}")
        w_px, h_px = _get_block_pixel_size(ws, r1, c1, r2, c2)
        # a kép alját 8%-kal rövidítjük, hogy ne érjen bele az alatta lévő vonalba
        h_px = int(h_px * 0.92)
        print(f"IMG: block px size = {w_px}x{h_px}")
        pil_img = _make_description_image(text_s, w_px, h_px)
        xlimg = _xlimage_from_pil(pil_img)
        anchor = f"{get_column_letter(c1)}{r1}"
        ws.add_image(xlimg, anchor)
        set_text(ws, r1, c1, "", wrap=False, align_left=True, valign_top=True)
        print(f"IMG: inserted ok at {anchor}, size={w_px}x{h_px}")
        return True
    except Exception as e:
        print("IMG: insert FAILED ->", repr(e))
        traceback.print_exc()
        return False

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
    geraet: str = Form(""),
    beschreibung: str = Form(""),
    break_minutes: int = Form(60),
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""), vorhaltung1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""), vorhaltung2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""), vorhaltung3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""), vorhaltung4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""), vorhaltung5: str = Form(""),
):
    wb = load_workbook(os.path.join(os.getcwd(), "GP-t.xlsx"))
    ws = wb.active

    # --- Felső mezők ---
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

    # --- Beschreibung blokk ---
    r1, c1, r2, c2 = find_description_block(ws)

    # Sormagasság rögzítése (ha nincs beállítva)
    for r in range(r1, r2 + 1):
        if ws.row_dimensions.get(r) is None or ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 22

    inserted = False
    text_in = (beschreibung or "").strip()
    print(f"IMG: incoming beschreibung len={len(text_in)}")
    if text_in:
        inserted = insert_description_as_image(ws, r1, c1, r2, c2, text_in)

    # Ha nincs kép (hiba/üres), essünk vissza szövegre, hogy mindig legyen tartalom
    if not inserted:
        print("IMG: fallback to plain text (wrap+top)")
        set_text(ws, r1, c1, text_in, wrap=True, align_left=True, valign_top=True)

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

        if vorhaltung_col and (vh or "").strip():
            set_text(ws, row, vorhaltung_col, vh, wrap=True, align_left=True, valign_top=True)

        hb = parse_hhmm(bg)
        he = parse_hhmm(en)
        h = round(hours_with_breaks(hb, he, int(break_minutes)), 2)
        total_hours += h
        set_text(ws, row, pos["stunden_col"], h, wrap=False, align_left=True)
        row += 1

    # --- Összóraszám ---
    right_of_label, stunden_total = find_total_cells(ws, pos["stunden_col"])
    if stunden_total:
        tr, tc = stunden_total
        set_text(ws, tr, tc, round(total_hours, 2), wrap=False, align_left=True)
    if right_of_label:
        rr, rc = right_of_label
        set_text(ws, rr, rc, "", wrap=False, align_left=True)

    # ---- Fix méretű válasz, Content-Length-cel ----
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

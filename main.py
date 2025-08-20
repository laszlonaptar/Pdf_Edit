# main.py — FastAPI app with Excel export and true print-preview PDF via LibreOffice (fallback to ReportLab)

from fastapi import FastAPI, Request, Form, Response
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

from datetime import datetime, time
from io import BytesIO
import os, uuid, textwrap, traceback, subprocess, tempfile, shutil

# --- ReportLab for PDF fallback ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors

# Képgeneráláshoz
try:
    from PIL import Image as PILImage, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except Exception as e:
    PIL_AVAILABLE = False
    print("IMG: PIL not available ->", repr(e))

# -------- App + statikusok, sablonok --------
app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# -------- Egyszerű jelszóvédelem (Render env var-okkal) --------
APP_USERNAME = os.getenv("APP_USERNAME", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
SESSION_SECRET = os.getenv("SESSION_SECRET", "dev-secret-change-me")

app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET, same_site="lax")

def _is_authed(request: Request) -> bool:
    return bool(request.session.get("auth_ok") is True)

def _login_page(msg: str = "", next_path: str = "/") -> str:
    return f"""<!doctype html>
    <html lang="de"><head>
    <meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Login</title>
    <link rel="stylesheet" href="/static/style.css">
    <style>.login-card{{max-width:420px;margin:4rem auto;padding:1.25rem}}
    .muted{{color:#666;font-size:.9rem}}.err{{color:#b00020;margin:.5rem 0}}</style>
    </head><body><main class="container"><section class="card login-card">
      <h1>Anmeldung</h1>{("<div class='err'>"+msg+"</div>" if msg else "")}
      <form method="post" action="/login">
        <input type="hidden" name="next" value="{next_path}">
        <div class="field"><label for="u">Benutzername</label>
        <input id="u" name="username" type="text" required></div>
        <div class="field"><label for="p">Passwort</label>
        <input id="p" name="password" type="password" required></div>
        <div class="actions"><button class="btn primary" type="submit">Anmelden</button></div>
      </form><p class="muted">Zugriff ist passwortgeschützt.</p></section></main></body></html>"""

@app.get("/login", response_class=HTMLResponse)
async def login_form(request: Request, next: str = "/"):
    if _is_authed(request):
        return RedirectResponse(next or "/", status_code=303)
    return HTMLResponse(_login_page(next_path=next))

@app.post("/login", response_class=HTMLResponse)
async def login_submit(request: Request, username: str = Form(...), password: str = Form(...), next: str = Form("/")):
    if APP_USERNAME and APP_PASSWORD:
        if username == APP_USERNAME and password == APP_PASSWORD:
            request.session["auth_ok"] = True
            return RedirectResponse(next or "/", status_code=303)
        else:
            return HTMLResponse(_login_page("Falscher Benutzername oder Passwort.", next), status_code=401)
    else:
        request.session["auth_ok"] = True
        return RedirectResponse(next or "/", status_code=303)

@app.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)

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
        minus = overlap_minutes(beg, end, time(9,0), time(9,15)) + overlap_minutes(beg, end, time(12,0), time(12,45))
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
    return (6, 1, 15, 7)  # A6..G15

# ---------- pixel helpers ----------
def _excel_col_width_to_pixels(width):
    if width is None:
        width = 8.43
    return int(round(7 * width + 5))

def _excel_row_height_to_pixels(height):
    if height is None:
        height = 15.0
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

def _get_col_pixel_width(ws, col_index):
    letter = get_column_letter(col_index)
    cd = ws.column_dimensions.get(letter)
    return _excel_col_width_to_pixels(getattr(cd, "width", None))

# ---------- (ÚJ) oszlopszélességek a teljes A4 kitöltéséhez ----------
def _set_column_widths_for_print(ws, pos):
    widths = {
        pos["name_col"]:     18.0,  # Name
        pos["vorname_col"]:  14.0,  # Vorname
        pos["ausweis_col"]:  14.0,  # Ausweis/Kennzeichen
        pos["beginn_col"]:   10.0,  # Beginn
        pos["ende_col"]:     10.0,  # Ende
        pos["stunden_col"]:  12.0,  # Anzahl Stunden
    }
    if "vorhaltung_col" in pos and pos["vorhaltung_col"]:
        widths[pos["vorhaltung_col"]] = 28.0  # Vorhaltung / Gerät

    for col_idx, width in widths.items():
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].width = width

# ---------- image (Beschreibung) ----------
LEFT_INSET_PX = 25
BOTTOM_CROP   = 0.92

def _make_description_image(text, w_px, h_px):
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

    pad_left, pad_top, pad_right, pad_bottom = 12, 10, 0, 10
    avail_w = max(10, w_px - (pad_left + pad_right))
    avail_h = max(10, h_px - (pad_top + pad_bottom))

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
        lines.extend(textwrap.wrap(para, width=max_chars_per_line, replace_whitespace=False, break_long_words=False))

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

# extra biztonsági jobb oldali ráhagyás (QuickLook/Excel kerekítések ellen)
RIGHT_SAFE = 14  # px

def insert_description_as_image(ws, r1, c1, r2, c2, text):
    if not PIL_AVAILABLE:
        print("IMG: PIL not available at runtime")
        return False
    try:
        text_s = (text or "")
        print(f"IMG: will insert, text_len={len(text_s)} at {get_column_letter(c1)}{r1}-{get_column_letter(c2)}{r2}")

        # Teljes A..G blokk mérete (px)
        block_w_px, block_h_px = _get_block_pixel_size(ws, r1, c1, r2, c2)
        # A oszlop (bal keret) szélessége (px)
        colA_w_px = _get_col_pixel_width(ws, 1)

        # Rendelkezésre álló terület a képnek (px) – bal insett + jobb biztonsági ráhagyás levonva
        avail_w = block_w_px - colA_w_px - LEFT_INSET_PX - RIGHT_SAFE
        avail_h = int(block_h_px * BOTTOM_CROP)

        # Alsó korlátok
        avail_w = max(60, avail_w)
        avail_h = max(40, avail_h)

        # Kép generálása PONTOSAN ekkora méretre
        pil_img = _make_description_image(text_s, avail_w, avail_h)
        xlimg = _xlimage_from_pil(pil_img)
        # (dupla biztosítás) – rögzítjük a méretet px-ben
        xlimg.width = avail_w
        xlimg.height = avail_h

        # Anchor a B6-ba (így a bal oldali keretet nem fedi le)
        anchor = f"{get_column_letter(c1 + 1)}{r1}"  # B6
        ws.add_image(xlimg, anchor)

        # A6 cellát ürítjük, ne legyen átfedés
        set_text(ws, r1, c1, "", wrap=False, align_left=True, valign_top=True)
        return True
    except Exception as e:
        print("IMG: insert FAILED ->", repr(e))
        traceback.print_exc()
        return False

# ---- shared: workbook fill ----
def build_workbook(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers):
    wb = load_workbook(os.path.join(os.getcwd(), "GP-t.xlsx"))
    ws = wb.active

    # Felső mezők
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

    # Beschreibung blokk
    r1, c1, r2, c2 = find_description_block(ws)
    for r in range(r1, r2 + 1):
        if ws.row_dimensions.get(r) is None or ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 22

    inserted = False
    text_in = (beschreibung or "").strip()
    if text_in:
        inserted = insert_description_as_image(ws, r1, c1, r2, c2, text_in)
    if not inserted:
        set_text(ws, r1, c1+1, text_in, wrap=True, align_left=True, valign_top=True)

    # Dolgozók
    pos = find_header_positions(ws)

    # Oszlopszélességek A4-hez
    try:
        _set_column_widths_for_print(ws, pos)
    except Exception as _e:
        print("WIDTH SET WARN:", repr(_e))

    row = pos["data_start_row"]
    vorhaltung_col = pos.get("vorhaltung_col", None)

    total_hours = 0.0
    for (vn, nn, aw, bg, en, vh) in workers:
        if not any([vn, nn, aw, bg, en, vh]):
            continue
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

    right_of_label, stunden_total = find_total_cells(ws, pos["stunden_col"])
    if stunden_total:
        tr, tc = stunden_total
        set_text(ws, tr, tc, round(total_hours, 2), wrap=False, align_left=True)
    if right_of_label:
        rr, rc = right_of_label
        set_text(ws, rr, rc, "", wrap=False, align_left=True)

    # ---------- Nyomtatási beállítások (A4, teljes oldal, 1×1 oldalra illesztés) ----------
    try:
        # A4 fekvő
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = 9  # A4

        # Skálázás: pontosan 1 oldal széles × 1 oldal magas
        ws.page_setup.scale = None
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        # fitToPage jelölés (openpyxl külön property-n)
        if hasattr(ws, "sheet_properties") and hasattr(ws.sheet_properties, "pageSetUpPr"):
            ws.sheet_properties.pageSetUpPr.fitToPage = True

        # Margók minimalizálása (inch)
        ws.page_margins.left   = 0.2
        ws.page_margins.right  = 0.2
        ws.page_margins.top    = 0.2
        ws.page_margins.bottom = 0.2
        ws.page_margins.header = 0
        ws.page_margins.footer = 0

        # Header/Footer letiltása – kompatibilisen (régi openpyxl esetén se dobjon hibát)
        try:
            if hasattr(ws, "oddHeader") and hasattr(ws, "oddFooter"):
                ws.oddHeader.left.text = ws.oddHeader.center.text = ws.oddHeader.right.text = ""
                ws.oddFooter.left.text = ws.oddFooter.center.text = ws.oddFooter.right.text = ""
            if hasattr(ws, "header_footer"):
                ws.header_footer.differentFirst = False
                ws.header_footer.differentOddEven = False
        except Exception:
            pass  # nem kritikus

        # Nyomtatási terület: ténylegesen használt utolsó sorig
        last_data_row = 1
        for r in range(1, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                last_data_row = r

        last_col = ws.max_column
        ws.print_area = f"A1:{get_column_letter(last_col)}{last_data_row}"

        # Ne középre igazítsuk (különben optikailag kisebbnek hat)
        ws.print_options.horizontalCentered = False
        ws.print_options.verticalCentered = False
    except Exception as e:
        print("PRINT SETUP WARN:", repr(e))

    return wb

# ---------- routes ----------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    if not _is_authed(request):
        return RedirectResponse("/login?next=/", status_code=303)
    return templates.TemplateResponse("index.html", {"request": request})

def _collect_workers(formdict):
    def g(key): return (formdict.get(key) or "").strip()
    workers = []
    for i in range(1, 6):
        workers.append((g(f"vorname{i}"), g(f"nachname{i}"), g(f"ausweis{i}"),
                        g(f"beginn{i}"), g(f"ende{i}"), g(f"vorhaltung{i}")))
    return workers

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...), bau: str = Form(...),
    basf_beauftragter: str = Form(""), geraet: str = Form(""),
    beschreibung: str = Form(""), break_minutes: int = Form(60),
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""), vorhaltung1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""), vorhaltung2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""), vorhaltung3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""), vorhaltung4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""), vorhaltung5: str = Form(""),
):
    if not _is_authed(request):
        return RedirectResponse("/login?next=/", status_code=303)

    workers = _collect_workers(locals())
    wb = build_workbook(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers)

    bio = BytesIO()
    wb.save(bio)
    data = bio.getvalue()
    fname = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{fname}"',
        "Content-Length": str(len(data)),
        "Cache-Control": "no-store",
    }
    return Response(content=data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

# --- PDF előnézet route megmarad (gomb nélkül nem zavar), LO-val a lapbeállításokat követi ---
def _has_soffice() -> bool:
    return shutil.which("soffice") is not None

def _reportlab_preview_pdf(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers) -> bytes:
    pagesize = landscape(A4)
    W, H = pagesize
    margin = 18 * mm
    bio = BytesIO()
    c = canvas.Canvas(bio, pagesize=pagesize)

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    c.drawString(margin, H - margin, "Leistungsnachweis – PDF Vorschau")

    date_text = datum
    try:
        dt = datetime.strptime(datum.strip(), "%Y-%m-%d")
        date_text = dt.strftime("%d.%m.%Y")
    except Exception:
        pass

    c.setFont("Helvetica", 11)
    y = H - margin - 10*mm
    c.drawString(margin, y, f"Datum: {date_text}")
    y -= 6 * mm
    c.drawString(margin, y, f"Bau: {bau}")
    y -= 6 * mm
    if (basf_beauftragter or "").strip():
        c.drawString(margin, y, f"BASF-Beauftragter: {basf_beauftragter}")
        y -= 8 * mm
    else:
        y -= 2 * mm

    box_x = margin
    box_w = W - 2*margin
    box_h = 52 * mm
    box_y = H - margin - 40*mm

    c.setStrokeColor(colors.black)
    c.rect(box_x, box_y - box_h, box_w, box_h, stroke=1, fill=0)

    styles = getSampleStyleSheet()
    style = ParagraphStyle(
        "Besch",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=11,
        leading=13.5,
        alignment=TA_LEFT,
    )

    beschr = (beschreibung or "").replace("\r\n", "\n").replace("\r", "\n")
    beschr_html = "<br/>".join(beschr.split("\n"))
    para = Paragraph(beschr_html, style)

    frame = Frame(box_x + 8*mm, (box_y - box_h) + 8*mm,
                  box_w - 16*mm, box_h - 16*mm, showBoundary=0)
    frame.addFromList([para], c)

    y_tab = box_y - 12*mm
    c.setFont("Helvetica-Bold", 10)
    headers = ["Name", "Vorname", "Ausweis", "Beginn", "Ende", "Anzahl Stunden", "Vorhaltung"]
    col_widths = [35*mm, 35*mm, 28*mm, 20*mm, 20*mm, 28*mm, 40*mm]
    x = margin
    for htxt, w in zip(headers, col_widths):
        c.drawString(x, y_tab, htxt)
        x += w

    c.setFont("Helvetica", 10)
    def row(values, yrow):
        x = margin
        for val, w in zip(values, col_widths):
            c.drawString(x, yrow, str(val or ""))
            x += w

    def hhmm(t): return parse_hhmm(t) if t else None

    total_hours = 0.0
    y_tab -= 6*mm
    for (vn, nn, aw, bg, en, vh) in workers:
        if not any([vn, nn, aw, bg, en, vh]):
            continue
        hb = hhmm(bg); he = hhmm(en)
        h = round(hours_with_breaks(hb, he, int(break_minutes)), 2)
        total_hours += h
        row([nn, vn, aw, bg, en, h, vh], y_tab)
        y_tab -= 6*mm

    y_tab -= 4 * mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y_tab, f"Gesamtstunden: {round(total_hours, 2)}")

    c.setFont("Helvetica", 8)
    c.setFillColor(colors.grey)
    c.drawRightString(W - margin, margin - 2*mm, "PDF Vorschau – (ReportLab Fallback)")

    c.showPage(); c.save()
    return bio.getvalue()

@app.post("/generate_pdf")
async def generate_pdf(
    request: Request,
    datum: str = Form(...), bau: str = Form(...),
    basf_beauftragter: str = Form(""), geraet: str = Form(""),
    beschreibung: str = Form(""), break_minutes: int = Form(60),
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""), vorhaltung1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""), vorhaltung2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""), vorhaltung3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""), vorhaltung4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""), vorhaltung5: str = Form(""),
):
    if not _is_authed(request):
        return RedirectResponse("/login?next=/", status_code=303)

    workers = _collect_workers(locals())
    wb = build_workbook(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers)

    if _has_soffice():
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, f"ln_{uuid.uuid4().hex[:8]}.xlsx")
            pdf_outdir = tmpdir
            wb.save(xlsx_path)

            filter_data = '{"UsePageSettings":true,"ScaleToPagesX":1,"ScaleToPagesY":1}'
            cmd = [
                "soffice","--headless","--nologo","--nodefault","--nolockcheck","--nofirststartwizard",
                "--convert-to", f"pdf:calc_pdf_Export:{filter_data}",
                "--outdir", pdf_outdir,
                xlsx_path
            ]
            try:
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
                pdf_path = xlsx_path[:-5] + ".pdf"
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
            except Exception as e:
                print("LibreOffice conversion failed ->", repr(e))
                pdf_bytes = _reportlab_preview_pdf(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers)
    else:
        pdf_bytes = _reportlab_preview_pdf(datum, bau, basf_beauftragter, beschreibung, break_minutes, workers)

    fname = f"leistungsnachweis_preview_{uuid.uuid4().hex[:6]}.pdf"
    headers = {
        "Content-Disposition": f'attachment; filename="{fname}"',
        "Content-Length": str(len(pdf_bytes)),
        "Cache-Control": "no-store",
    }
    return Response(content=pdf_bytes, media_type="application/pdf", headers=headers)

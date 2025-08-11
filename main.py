# main.py
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment

from datetime import datetime, time
from io import BytesIO
import os
import uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- helpers: merged-cell safe writing ----------

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

def right_neighbor_block(ws, r, c):
    cur = block_of(ws, r, c)
    _, _, _, cur_max_c = cur
    candidates = []
    for (r1, c1, r2, _) in merged_ranges(ws):
        if r1 <= r <= r2 and c1 > cur_max_c:
            candidates.append((c1, r1, c1))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    _, rr, cc = candidates[0]
    return rr, cc

def set_text(ws, r, c, text, wrap=False, align_left=False, valign_top=False, number_format=None):
    rr, cc = top_left_of_block(ws, r, c)
    cell = ws.cell(row=rr, column=cc)
    cell.value = text
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal=("left" if align_left else "center"),
        vertical=("top" if valign_top else cell.alignment.vertical or "center"),
    )
    if number_format:
        cell.number_format = number_format

def put_value_right_of_label(ws, label_text, value, wrap=False, align_left=False, valign_top=False, number_format=None):
    found = None
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.strip() == label_text:
                found = (cell.row, cell.column)
                break
        if found:
            break
    if not found:
        return False
    r, c = found
    neigh = right_neighbor_block(ws, r, c)
    if not neigh:
        return False
    nr, nc = neigh
    set_text(ws, nr, nc, value, wrap=wrap, align_left=align_left, valign_top=valign_top, number_format=number_format)
    return True

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

def hours_with_breaks(beg: time | None, end: time | None) -> float:
    if not beg or not end:
        return 0.0
    dt = datetime(2000,1,1)
    start = dt.replace(hour=beg.hour, minute=beg.minute)
    finish = dt.replace(hour=end.hour, minute=end.minute)
    if finish <= start:
        return 0.0
    total_min = int((finish - start).total_seconds() // 60)
    minus = overlap_minutes(beg, end, time(9,0), time(9,15)) + overlap_minutes(beg, end, time(12,0), time(12,45))
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
        if all(k in pos for k in ["name_col","vorname_col","ausweis_col","beginn_col","ende_col","stunden_col","subheader_row"]):
            break
    pos["data_start_row"] = pos.get("subheader_row", header_row) + 1
    return pos

def find_total_cell(ws):
    """
    Keresd meg a 'Gesamtstunden' felirat sorában a JOBB SZÉLSŐ összevont blokkot
    (vagy ha nincs összevonás, a legutolsó cellát), és annak bal-felső celláját add vissza.
    Így a tényleges összesítő dobozba írunk.
    """
    label_row = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and "Gesamtstunden" in v:
                label_row = cell.row
                # nem törünk ki azonnal – ha több van, az utolsót akarjuk (alsó blokk)
    if not label_row:
        return None

    # jobb szélső összevont blokk kiválasztása
    rightmost = None
    rightmost_c2 = -1
    for rng in ws.merged_cells.ranges:
        r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if r1 <= label_row <= r2 and c2 > rightmost_c2:
            rightmost_c2 = c2
            rightmost = (r1, c1)

    if rightmost:
        return rightmost

    # fallback: ha nincs merge, a sor utolsó cellája
    last_col = 0
    for cell in ws[label_row]:
        if cell.column > last_col:
            last_col = cell.column
    return (label_row, last_col)

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
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""),
):
    wb = load_workbook(os.path.join(os.getcwd(), "GP-t.xlsx"))
    ws = wb.active

    # --- Dátum német formátumban, szövegként ---
    date_text = datum
    try:
        dt = datetime.strptime(datum.strip(), "%Y-%m-%d")
        date_text = dt.strftime("%d.%m.%Y")
    except Exception:
        pass
    put_value_right_of_label(ws, "Datum der Leistungsausführung:", date_text, align_left=True)

    put_value_right_of_label(ws, "Bau und Ausführungsort:", bau, align_left=True)
    if (basf_beauftragter or "").strip():
        put_value_right_of_label(ws, "BASF-Beauftragter, Org.-Code:", basf_beauftragter, align_left=True)
    if (geraet or "").strip():
        put_value_right_of_label(ws, "Vorhaltung / beauftragtes Gerät / Fahrzeug:", geraet, align_left=True)

    # Beschreibung – sor-magasság + tördelés
    r1, c1, r2, c2 = find_big_description_block(ws)
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = 22
    set_text(ws, r1, c1, beschreibung, wrap=True, align_left=True, valign_top=True)

    # Dolgozók
    pos = find_header_positions(ws)
    row = pos["data_start_row"]

    workers = []
    for i in range(1, 5+1):
        vn = locals().get(f"vorname{i}", "") or ""
        nn = locals().get(f"nachname{i}", "") or ""
        aw = locals().get(f"ausweis{i}", "") or ""
        bg = locals().get(f"beginn{i}", "") or ""
        en = locals().get(f"ende{i}", "") or ""
        if not (vn or nn or aw or bg or en):
            continue
        workers.append((vn, nn, aw, bg, en))

    total_hours = 0.0
    for (vn, nn, aw, bg, en) in workers:
        set_text(ws, row, pos["name_col"], nn, wrap=False, align_left=True)
        set_text(ws, row, pos["vorname_col"], vn, wrap=False, align_left=True)
        set_text(ws, row, pos["ausweis_col"], aw, wrap=False, align_left=True)
        set_text(ws, row, pos["beginn_col"], bg, wrap=False, align_left=True)
        set_text(ws, row, pos["ende_col"], en, wrap=False, align_left=True)
        hb = parse_hhmm(bg)
        he = parse_hhmm(en)
        h = round(hours_with_breaks(hb, he), 2)
        total_hours += h
        set_text(ws, row, pos["stunden_col"], h, wrap=False, align_left=True)
        row += 1

    tot_cell = find_total_cell(ws)
    if tot_cell:
        tr, tc = tot_cell
        set_text(ws, tr, tc, round(total_hours, 2), wrap=False, align_left=True)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{fname}"'}
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

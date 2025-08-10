# main.py
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment

from datetime import datetime, time, timedelta
from io import BytesIO
import os
import uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- helpers: merged-cell safe writing ----------

def merged_ranges(ws):
    # list of (min_row, min_col, max_row, max_col)
    return [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]

def in_range(rng, r, c):
    r1, c1, r2, c2 = rng
    return (r1 <= r <= r2) and (c1 <= c <= c2)

def block_of(ws, r, c):
    """Return (r1,c1,r2,c2) of the merged block containing (r,c),
       or the single-cell block if not merged."""
    for rng in merged_ranges(ws):
        if in_range(rng, r, c):
            return rng
    return (r, c, r, c)

def top_left_of_block(ws, r, c):
    r1, c1, _, _ = block_of(ws, r, c)
    return r1, c1

def right_neighbor_block(ws, r, c):
    """On the SAME row as (r,c), find the nearest merged block strictly to the right
       of the current block. Return its top-left (rr,cc)."""
    cur = block_of(ws, r, c)
    _, _, _, cur_max_c = cur
    candidates = []
    for (r1, c1, r2, c2) in merged_ranges(ws):
        if r1 <= r <= r2 and c1 > cur_max_c:
            candidates.append((c1, r1, c1))  # sort by left edge
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    _, rr, cc = candidates[0]
    return rr, cc

def set_text(ws, r, c, text, wrap=False, align_left=False, valign_top=False):
    rr, cc = top_left_of_block(ws, r, c)
    cell = ws.cell(row=rr, column=cc)
    cell.value = text
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal=("left" if align_left else "center"),
        vertical=("top" if valign_top else cell.alignment.vertical or "center"),
    )

def put_value_right_of_label(ws, label_text, value, wrap=False, align_left=False, valign_top=False):
    """Find a cell whose STRIPPED value == label_text; put value into the nearest
       merged block to the right on the same row."""
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
    set_text(ws, nr, nc, value, wrap=wrap, align_left=align_left, valign_top=valign_top)
    return True

# ---------- time & hours ----------

def parse_hhmm(s: str) -> time:
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

def hours_with_breaks(beg: time, end: time) -> float:
    if not beg or not end:
        return 0.0
    dt = datetime(2000,1,1)
    start = dt.replace(hour=beg.hour, minute=beg.minute)
    finish = dt.replace(hour=end.hour, minute=end.minute)
    if finish <= start:
        return 0.0
    total_min = int((finish - start).total_seconds() // 60)

    # breaks: 09:00–09:15 (15m), 12:00–12:45 (45m)
    b1s, b1e = time(9,0),  time(9,15)
    b2s, b2e = time(12,0), time(12,45)
    minus = overlap_minutes(beg, end, b1s, b1e) + overlap_minutes(beg, end, b2s, b2e)

    return max(0.0, (total_min - minus) / 60.0)

# ---------- table helpers ----------

def find_header_positions(ws):
    """Return dict with columns for each header + data_start_row."""
    pos = {}
    # 1) find 'Name' and 'Vorname'
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
        # when we found the main things and the subheader row, we can stop early
        if all(k in pos for k in ["name_col","vorname_col","ausweis_col","beginn_col","ende_col","stunden_col","subheader_row"]):
            break

    # rows where data starts: one row below the subheader ("Beginn/Ende" sor)
    pos["data_start_row"] = pos.get("subheader_row", header_row) + 1
    return pos

def find_total_cell(ws):
    """Find 'Gesamtstunden' label and return column right to it"""
    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and "Gesamtstunden" in v:
                r, c = cell.row, cell.column
                # value is usually in the next merged block to the right
                neigh = right_neighbor_block(ws, r, c)
                if neigh:
                    return neigh
                else:
                    return (r, c+1)
    return None

def find_big_description_block(ws):
    """Heurisztika: legnagyobb (vagy az egyik legnagyobb) több-soros összevont blokk a lap teteje alatt."""
    big = None
    for (r1,c1,r2,c2) in merged_ranges(ws):
        height = r2 - r1 + 1
        width  = c2 - c1 + 1
        if r1 >= 6 and height >= 4 and width >= 4:
            area = height * width
            if not big or area > big[-1]:
                big = (r1, c1, r2, c2, area)
    if big:
        r1, c1, _, _, _ = big
        return (r1, c1)
    # visszaesés: 6,1
    return (6, 1)

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
    # 1. dolgozó (kötelező)
    vorname1: str = Form(""),
    nachname1: str = Form(""),
    ausweis1: str = Form(""),
    beginn1: str = Form(""),
    ende1: str = Form(""),
    # opcionális továbbiak – ha vannak a formban
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""),
):
    # --- open template
    template_path = os.path.join(os.getcwd(), "GP-t.xlsx")
    wb = load_workbook(template_path)
    ws = wb.active

    # --- Top area: date / bau / beauftragter
    put_value_right_of_label(ws, "Datum der Leistungsausführung:", datum)
    put_value_right_of_label(ws, "Bau und Ausführungsort:", bau)
    if (basf_beauftragter or "").strip():
        put_value_right_of_label(ws, "BASF-Beauftragter, Org.-Code:", basf_beauftragter)

    # --- Beschreibung: nagy jegyzetblokk – felülre igazítva, sortöréssel
    desc_r, desc_c = find_big_description_block(ws)
    set_text(ws, desc_r, desc_c, beschreibung, wrap=True, align_left=True, valign_top=True)

    # --- dolgozók beírása
    pos = find_header_positions(ws)
    row = pos["data_start_row"]

    workers = []
    for i in range(1, 6):
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
        # Name (vezetéknév) – a sablonban ez a „Name” oszlop
        set_text(ws, row, pos["name_col"], nn, wrap=False, align_left=True)
        # Vorname
        set_text(ws, row, pos["vorname_col"], vn, wrap=False, align_left=True)
        # Ausweis
        set_text(ws, row, pos["ausweis_col"], aw, wrap=False, align_left=True)
        # Beginn / Ende
        set_text(ws, row, pos["beginn_col"], bg, wrap=False, align_left=True)
        set_text(ws, row, pos["ende_col"], en, wrap=False, align_left=True)

        hb = parse_hhmm(bg)
        he = parse_hhmm(en)
        h = round(hours_with_breaks(hb, he), 2)
        total_hours += h
        # Anzahl Stunden
        set_text(ws, row, pos["stunden_col"], h, wrap=False, align_left=True)
        row += 1

    # Gesamtstunden
    tot_cell = find_total_cell(ws)
    if tot_cell:
        tr, tc = tot_cell
        set_text(ws, tr, tc, round(total_hours, 2), wrap=False, align_left=True)

    # --- stream back
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{fname}"'}
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

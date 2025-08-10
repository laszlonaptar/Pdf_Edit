from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, time

import io
import uuid
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- Helpers for merged cells & label-based placement ----------

def merged_bounds(ws, r, c):
    """Return (min_row, min_col, max_row, max_col) of the merged block containing (r,c) or itself."""
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            return rng.min_row, rng.min_col, rng.max_row, rng.max_col
    return r, c, r, c

def right_of_label_cell(ws, label_text):
    """Find the cell that contains label_text (exact match or startswith),
    then return the top-left cell **right of the whole merged label block**."""
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            v = cell.value if cell.value is not None else ""
            if isinstance(v, str) and (v.strip() == label_text or v.strip().startswith(label_text)):
                r0, c0, r1, c1 = merged_bounds(ws, cell.row, cell.column)
                return r0, c1 + 1
    raise ValueError(f'Label not found: "{label_text}"')

def set_text(ws, r, c, text, wrap=False, left=True, top=True):
    cell = ws.cell(r, c)
    cell.value = text
    align = Alignment(
        horizontal="left" if left else "center",
        vertical="top" if top else "center",
        wrap_text=wrap
    )
    cell.alignment = align

def increase_rows_height(ws, r_start, r_end, height=22):
    for r in range(r_start, r_end + 1):
        ws.row_dimensions[r].height = height

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def hours_minus_breaks(t1: time, t2: time) -> float:
    """Compute hours between t1..t2 minus fixed breaks (09:00–09:15, 12:00–12:45)."""
    dt0 = datetime.combine(datetime.today(), t1)
    dt1 = datetime.combine(datetime.today(), t2)
    if dt1 < dt0:
        dt1 = dt1.replace(day=dt1.day + 1)

    total = (dt1 - dt0).total_seconds() / 3600.0

    def overlap(beg_h, beg_m, end_h, end_m):
        b0 = datetime.combine(datetime.today(), time(beg_h, beg_m))
        b1 = datetime.combine(datetime.today(), time(end_h, end_m))
        a = max(dt0, b0)
        b = min(dt1, b1)
        return max(0.0, (b - a).total_seconds() / 3600.0)

    pause = overlap(9, 0, 9, 15) + overlap(12, 0, 12, 45)
    return max(0.0, round(total - pause, 2))

def find_header_row_and_columns(ws):
    """Detect the header row with 'Name' and return a dict of relevant column indexes."""
    header_row = None
    cols = {}
    for row in ws.iter_rows(min_row=1, max_row=80):
        texts = [ (cell.column, str(cell.value).strip() if cell.value else "") for cell in row ]
        labels = [t for _, t in texts]
        if "Name" in labels and "Vorname" in labels:
            # likely the header row
            header_row = row[0].row
            for col_idx, txt in texts:
                if "Name" == txt:
                    cols["name"] = col_idx
                if "Vorname" == txt:
                    cols["vorname"] = col_idx
                if "Ausweis" in txt:
                    cols["ausweis"] = col_idx
                if "Beginn" in txt:
                    cols["beginn"] = col_idx
                if "Ende" in txt:
                    cols["ende"] = col_idx
                if "Anzahl Stunden" in txt:
                    cols["stunden"] = col_idx
            break
    if not header_row or not cols:
        raise RuntimeError("Konnte die Kopfzeile der Mitarbeitertabelle nicht sicher erkennen.")
    return header_row, cols

# ---------- Routes ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    basf: str = Form(""),  # BASF-Beauftragter (optional a formban)
    beschreibung: str = Form(...),

    nachname1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),

    # Ha lesz több dolgozó: nachname2, vorname2, stb. ugyanígy bővíthető
):
    # 1) Excel sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # 2) Dátum és Bau: mindig a címke blokkjának jobb szélére írunk
    try:
        r_date, c_date = right_of_label_cell(ws, "Datum der Leistungsausführung:")
        set_text(ws, r_date, c_date, datum, wrap=False)
    except Exception:
        pass

    try:
        r_bau, c_bau = right_of_label_cell(ws, "Bau und Ausführungsort:")
        set_text(ws, r_bau, c_bau, bau, wrap=False)
    except Exception:
        pass

    # BASF-Beauftragter – E3 környéke, de inkább label alapján:
    try:
        r_basf, c_basf = right_of_label_cell(ws, "BASF-Beauftragter, Org.-Code:")
        if basf:
            set_text(ws, r_basf, c_basf, basf, wrap=False)
    except Exception:
        # ha nincs ilyen label a sablonban, átugorjuk
        pass

    # 3) Beschreibung: A6–G15 egy nagy, összevont terület – bal-felső cellába írunk, sortörés + topline.
    # Megpróbáljuk label nélkül, konkrét pozícióval, mert ezt így egyeztettük:
    A6, G15 = (6, 1), (15, 7)
    set_text(ws, A6[0], A6[1], beschreibung, wrap=True, left=True, top=True)
    increase_rows_height(ws, A6[0], G15[0], height=22)  # hogy biztosan ne vágódjon le

    # 4) Dolgozó sor – fejlécek alapján megtaláljuk az oszlopokat és a következő üres sort
    header_row, cols = find_header_row_and_columns(ws)
    first_data_row = header_row + 1

    # megkeressük az első üres sort a „Name” oszlopban
    r = first_data_row
    while ws.cell(r, cols["name"]).value not in (None, "") and r < header_row + 40:
        r += 1

    # per-dolgozó óraszám
    try:
        b1 = parse_hhmm(beginn1)
        e1 = parse_hhmm(ende1)
        stunden_1 = hours_minus_breaks(b1, e1)
    except Exception:
        stunden_1 = ""

    # értékek kiírása a megfelelő oszlopokba
    set_text(ws, r, cols["name"], nachname1, wrap=False)
    set_text(ws, r, cols["vorname"], vorname1, wrap=False)
    set_text(ws, r, cols["ausweis"], ausweis1, wrap=False)
    set_text(ws, r, cols["beginn"], beginn1, wrap=False)
    set_text(ws, r, cols["ende"], ende1, wrap=False)
    if "stunden" in cols:
        set_text(ws, r, cols["stunden"], stunden_1, wrap=False)

    # 5) Gesammtstunden alsó mező: összegezhetünk, de nálad már működik — ha kell, itt is beírható
    # (meghagyom most; a sablonod alján már rendben számoltatjuk/írjuk.)

    # 6) Mentés memóriába és letöltés
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    filename = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return FileResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

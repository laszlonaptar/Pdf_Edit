from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, time
import io, uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- merged-cell helpers ----------

def merged_bounds(ws, r, c):
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            return rng.min_row, rng.min_col, rng.max_row, rng.max_col
    return r, c, r, c

def set_text(ws, r, c, text, wrap=False, left=True, top=True):
    r0, c0, *_ = merged_bounds(ws, r, c)  # mindig a blokk bal-felső cellájába írunk
    cell = ws.cell(r0, c0)
    cell.value = text
    cell.alignment = Alignment(
        horizontal="left" if left else "center",
        vertical="top" if top else "center",
        wrap_text=wrap
    )

def right_of_label_cell(ws, label_text):
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            v = cell.value if cell.value is not None else ""
            if isinstance(v, str) and (v.strip() == label_text or v.strip().startswith(label_text)):
                r0, c0, r1, c1 = merged_bounds(ws, cell.row, cell.column)
                return r0, c1 + 1
    raise ValueError(f'Label not found: {label_text}')

def increase_rows_height(ws, r_start, r_end, height=22):
    for r in range(r_start, r_end + 1):
        ws.row_dimensions[r].height = height

def is_merged_not_top(ws, r, c):
    r0, c0, r1, c1 = merged_bounds(ws, r, c)
    return (r0, c0) != (r, c)

# ---------- time helpers ----------

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def hours_minus_breaks(t1: time, t2: time) -> float:
    dt0 = datetime.combine(datetime.today(), t1)
    dt1 = datetime.combine(datetime.today(), t2)
    if dt1 < dt0:
        dt1 = dt1.replace(day=dt1.day + 1)
    total = (dt1 - dt0).total_seconds() / 3600.0

    def overlap(bh, bm, eh, em):
        b0 = datetime.combine(datetime.today(), time(bh, bm))
        b1 = datetime.combine(datetime.today(), time(eh, em))
        a = max(dt0, b0); b = min(dt1, b1)
        return max(0.0, (b - a).total_seconds() / 3600.0)

    pause = overlap(9, 0, 9, 15) + overlap(12, 0, 12, 45)
    return max(0.0, round(total - pause, 2))

# ---------- table header detection ----------

def find_header_row_and_columns(ws):
    header_row = None
    cols = {}
    for row in ws.iter_rows(min_row=1, max_row=80):
        labels = {cell.column: (str(cell.value).strip() if cell.value else "") for cell in row}
        if "Name" in labels.values() and "Vorname" in labels.values():
            header_row = row[0].row
            for col_idx, txt in labels.items():
                if txt == "Name":
                    cols["name"] = col_idx
                if txt == "Vorname":
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
        raise RuntimeError("Kopfzeile der Mitarbeitertabelle nicht gefunden.")
    return header_row, cols

def first_free_data_row(ws, header_row, name_col, vorname_col):
    r = header_row + 1
    # ugorjunk át minden olyan sort, ami még fejléchez/összevonáshoz tartozik
    while True:
        if is_merged_not_top(ws, r, name_col) or is_merged_not_top(ws, r, vorname_col):
            r += 1
            continue
        if (ws.cell(r, name_col).value in (None, "")) and (ws.cell(r, vorname_col).value in (None, "")):
            return r
        r += 1

# ---------- routes ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    basf: str = Form(""),
    beschreibung: str = Form(...),

    nachname1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
):
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Fejléc mezők (label -> érték a label blokk jobb oldalán)
    try:
        r, c = right_of_label_cell(ws, "Datum der Leistungsausführung:")
        set_text(ws, r, c, datum)
    except Exception:
        pass
    try:
        r, c = right_of_label_cell(ws, "Bau und Ausführungsort:")
        set_text(ws, r, c, bau)
    except Exception:
        pass
    try:
        r, c = right_of_label_cell(ws, "BASF-Beauftragter, Org.-Code:")
        if basf:
            set_text(ws, r, c, basf)
    except Exception:
        pass

    # Beschreibung – A6:G15 blokk teteje, sortörés + nagyobb sormagasság
    set_text(ws, 6, 1, beschreibung, wrap=True, left=True, top=True)
    increase_rows_height(ws, 6, 15, height=22)

    # Dolgozó sor
    header_row, cols = find_header_row_and_columns(ws)
    r = first_free_data_row(ws, header_row, cols["name"], cols["vorname"])

    # óraszám számítás a fix szünetekkel
    try:
        stunden = hours_minus_breaks(parse_hhmm(beginn1), parse_hhmm(ende1))
    except Exception:
        stunden = ""

    set_text(ws, r, cols["name"],    nachname1)
    set_text(ws, r, cols["vorname"], vorname1)
    set_text(ws, r, cols["ausweis"], ausweis1)
    set_text(ws, r, cols["beginn"],  beginn1)
    set_text(ws, r, cols["ende"],    ende1)
    if "stunden" in cols:
        set_text(ws, r, cols["stunden"], stunden)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    fname = f"leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    return FileResponse(out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'})

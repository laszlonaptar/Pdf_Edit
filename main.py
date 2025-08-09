from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime, time
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import os
import uuid

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_PATH = "GP-t.xlsx"  # must exist in repo root at deploy

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

def parse_time(hhmm: str):
    try:
        return datetime.strptime(hhmm, "%H:%M").time()
    except Exception:
        return None

def overlap_minutes(start: time, end: time, bstart: time, bend: time) -> int:
    # compute overlap in minutes between [start,end) and [bstart,bend)
    s1 = start.hour*60 + start.minute
    e1 = end.hour*60 + end.minute
    s2 = bstart.hour*60 + bstart.minute
    e2 = bend.hour*60 + bend.minute
    left = max(s1, s2)
    right = min(e1, e2)
    return max(0, right - left)

def minutes_to_hours(mins: int) -> float:
    return round(mins / 60.0, 2)

def find_label_cell(ws, labels):
    # Search first ~50 rows and 20 columns for any label variant
    for row in ws.iter_rows(min_row=1, max_row=50, min_col=1, max_col=20):
        for cell in row:
            val = str(cell.value).strip() if cell.value is not None else ""
            if val in labels:
                return cell.row, cell.column
    return None, None

def top_left_of_merged(ws, row, col):
    for mr in ws.merged_cells.ranges:
        if (row, col) in mr.cells:
            return mr.min_row, mr.min_col
    return row, col

def safe_write(ws, row, col, value, wrap=False):
    # Redirect to top-left if merged
    r, c = top_left_of_merged(ws, row, col)
    ws.cell(r, c, value)
    if wrap:
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

def merge_and_write(ws, range_ref, value):
    # Always unmerge then merge to be safe
    try:
        if range_ref in [str(rng) for rng in ws.merged_cells.ranges]:
            ws.unmerge_cells(range_ref)
    except Exception:
        pass
    ws.merge_cells(range_ref)
    tl = ws[range_ref].cell(1,1)
    tl.value = value
    tl.alignment = Alignment(wrap_text=True, vertical="top")

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    geraet: str = Form(""),
    beschreibung: str = Form(""),
    # dynamic workers - we accept up to 5; frontend ensures consistent arrays
    mitarbeiter_vorname: str = Form(...),
    mitarbeiter_nachname: str = Form(...),
    ausweis: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
    # optional extra workers (2..5)
    mitarbeiter_vorname2: str = Form(""),
    mitarbeiter_nachname2: str = Form(""),
    ausweis2: str = Form(""),
    beginn2: str = Form(""),
    ende2: str = Form(""),
    mitarbeiter_vorname3: str = Form(""),
    mitarbeiter_nachname3: str = Form(""),
    ausweis3: str = Form(""),
    beginn3: str = Form(""),
    ende3: str = Form(""),
    mitarbeiter_vorname4: str = Form(""),
    mitarbeiter_nachname4: str = Form(""),
    ausweis4: str = Form(""),
    beginn4: str = Form(""),
    ende4: str = Form(""),
    mitarbeiter_vorname5: str = Form(""),
    mitarbeiter_nachname5: str = Form(""),
    ausweis5: str = Form(""),
    beginn5: str = Form(""),
    ende5: str = Form(""),
):
    if not os.path.exists(TEMPLATE_PATH):
        return PlainTextResponse("Template GP-t.xlsx not found on server.", status_code=500)

    # Build workers list
    raw = [
        (mitarbeiter_vorname, mitarbeiter_nachname, ausweis, beginn, ende),
        (mitarbeiter_vorname2, mitarbeiter_nachname2, ausweis2, beginn2, ende2),
        (mitarbeiter_vorname3, mitarbeiter_nachname3, ausweis3, beginn3, ende3),
        (mitarbeiter_vorname4, mitarbeiter_nachname4, ausweis4, beginn4, ende4),
        (mitarbeiter_vorname5, mitarbeiter_nachname5, ausweis5, beginn5, ende5),
    ]
    workers = []
    for v, n, a, b, e in raw:
        if (v or n) and b and e:
            t1 = parse_time(b)
            t2 = parse_time(e)
            if t1 and t2 and (t2 > t1):
                # subtract fixed breaks
                minutes = (t2.hour*60+t2.minute) - (t1.hour*60+t1.minute)
                minutes -= overlap_minutes(t1, t2, time(9,0), time(9,15))
                minutes -= overlap_minutes(t1, t2, time(12,0), time(12,45))
                hours = minutes_to_hours(max(0, minutes))
            else:
                hours = 0.0
            workers.append({
                "name": f"{v.strip()} {n.strip()}".strip(),
                "ausweis": a.strip(),
                "beginn": b,
                "ende": e,
                "stunden": hours,
            })

    # open template
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Fill by labels (robust to cell changes)
    # Datum
    r, c = find_label_cell(ws, {"Dátum", "Datum", "Datum:", "Dátum:", "Date"})
    if r:
        safe_write(ws, r, c+1, datum)  # write to the cell next to the label

    # Bau/Projekt
    r, c = find_label_cell(ws, {"Bau", "Projekt", "Bau:", "Projekt:"})
    if r:
        safe_write(ws, r, c+1, bau)

    # Gerät (optional)
    r, c = find_label_cell(ws, {"Gerät", "Gép/Eszköz", "Gerät:", "Gép/Eszköz:"})
    if r:
        safe_write(ws, r, c+1, geraet)

    # Description into A6:G15 as a single merged, wrapped cell (per our agreement)
    try:
        if beschreibung.strip():
            merge_and_write(ws, "A6:G15", beschreibung.strip())
    except Exception as e:
        # fail soft
        pass

    # Workers summary block: write compactly under H6 (adjust if needed)
    start_row = 6
    start_col = 9  # I = 9 (A=1), so column I/J area on the right-hand side
    # headers
    ws.cell(start_row, start_col, "Mitarbeiter")
    ws.cell(start_row, start_col+1, "Ausweis")
    ws.cell(start_row, start_col+2, "Beginn")
    ws.cell(start_row, start_col+3, "Ende")
    ws.cell(start_row, start_col+4, "Stunden")
    for i, w in enumerate(workers, start=1):
        ws.cell(start_row+i, start_col, w["name"])
        ws.cell(start_row+i, start_col+1, w["ausweis"])
        ws.cell(start_row+i, start_col+2, w["beginn"])
        ws.cell(start_row+i, start_col+3, w["ende"])
        ws.cell(start_row+i, start_col+4, w["stunden"])

    total_hours = sum(w["stunden"] for w in workers)
    ws.cell(start_row+len(workers)+2, start_col+3, "Összesen:")
    ws.cell(start_row+len(workers)+2, start_col+4, total_hours)

    # Save unique filename
    out_name = f"GP-t_filled_{uuid.uuid4().hex[:8]}.xlsx"
    wb.save(out_name)

    return FileResponse(
        path=out_name,
        filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

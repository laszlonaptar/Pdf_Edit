from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List, Optional
from datetime import datetime, time, timedelta
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- break handling ----------
def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    # work in minutes from midnight
    def t2m(t: time) -> int:
        return t.hour * 60 + t.minute
    a1, a2, b1, b2 = t2m(a_start), t2m(a_end), t2m(b_start), t2m(b_end)
    left = max(a1, b1)
    right = min(a2, b2)
    return max(0, right - left)

def net_hours(begin: str, end: str) -> float:
    b = parse_hhmm(begin)
    e = parse_hhmm(end)
    total = (datetime.combine(datetime.today(), e) - datetime.combine(datetime.today(), b)).total_seconds() / 3600.0
    # fixed breaks: 09:00-09:15 (0.25h) and 12:00-12:45 (0.75h)
    brk = 0.0
    brk += overlap_minutes(b, e, time(9, 0), time(9, 15)) / 60.0
    brk += overlap_minutes(b, e, time(12, 0), time(12, 45)) / 60.0
    return max(0.0, round(total - brk, 2))

# ---------- Excel helpers ----------
def ensure_merge(ws, min_row, min_col, max_row, max_col):
    # merge if not already merged exactly this range
    rng = f"{openpyxl.utils.get_column_letter(min_col)}{min_row}:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
    for m in ws.merged_cells.ranges:
        if m.coord == rng:
            return
    ws.merge_cells(rng)

def top_left_of_merge(ws, r, c):
    # return (row,col) of the top-left cell for (r,c) considering merged regions
    for m in ws.merged_cells.ranges:
        if (r, c) in m:
            return (m.min_row, m.min_col)
    return (r, c)

def set_cell(ws, r, c, value, wrap=False, align_left=False):
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = value
    if wrap or align_left:
        cell.alignment = Alignment(
            wrap_text=wrap,
            horizontal="left" if align_left else cell.alignment.horizontal if cell.alignment else "general",
            vertical="top"
        )

def find_label_cell(ws, label_text: str):
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == label_text:
                return (cell.row, cell.column)
    return None

def right_region_top_left(ws, row: int, col: int):
    # find merged regions on this row, pick the one immediately to the right of (row,col)
    regions = []
    for m in ws.merged_cells.ranges:
        if m.min_row == m.max_row == row:
            regions.append(m)
    regions.sort(key=lambda m: m.min_col)
    for i, m in enumerate(regions):
        if m.min_col <= col <= m.max_col:
            # pick next region if exists
            if i + 1 < len(regions):
                return (regions[i+1].min_row, regions[i+1].min_col)
            break
    # fallback: next cell on the right
    return (row, col + 1)

@app.get("/", response_class=HTMLResponse)
def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    bf: Optional[str] = Form(None),
    geraet: Optional[str] = Form(None),
    taetigkeit: str = Form(...),
    vorname: List[str] = Form(...),
    nachname: List[str] = Form(...),
    ausweis: List[str] = Form(...),
    beginn: List[str] = Form(...),
    ende: List[str] = Form(...)
):
    # compute hours per worker + total
    stunden = []
    total_hours = 0.0
    for b, e in zip(beginn, ende):
        h = net_hours(b, e)
        stunden.append(h)
        total_hours += h
    total_hours = round(total_hours, 2)

    # load template
    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # 1) long description in A6:G15, left aligned + wrap (we explicitly (re)merge to be safe)
    ensure_merge(ws, 6, 1, 15, 7)  # A6:G15
    set_cell(ws, 6, 1, taetigkeit, wrap=True, align_left=True)

    # 2) fill fields next to labels (robust: find label, write to the next merged region on the right)
    if (pos := find_label_cell(ws, "Datum der Leistungsausführung:")):
        r, c = right_region_top_left(ws, pos[0], pos[1])
        set_cell(ws, r, c, datum)

    if (pos := find_label_cell(ws, "Bau und Ausführungsort:")):
        r, c = right_region_top_left(ws, pos[0], pos[1])
        set_cell(ws, r, c, bau)

    if bf:
        if (pos := find_label_cell(ws, "BASF-Beauftragter, Org.-Code:")):
            r, c = right_region_top_left(ws, pos[0], pos[1])
            set_cell(ws, r, c, bf)

    if geraet:
        # ha van "Vorhaltung / beauftragtes Gerät / Fahrzeug" cím a táblázat alján, nem piszkáljuk
        pass

    # 3) Workers table: locate header "Name" then write rows below
    header_row = None
    name_col = None
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "Name":
                header_row = cell.row
                name_col = cell.column
                break
        if header_row:
            break

    if header_row and name_col:
        # deduce other columns from headers on the same row
        cols = { "Name": None, "Vorname": None, "Ausweis- Nr.\noder\nKennzeichen": None,
                 "Beginn": None, "Ende": None,
                 "Anzahl Stunden\n(ohne Pausen)": None }
        for cell in ws[header_row]:
            val = cell.value.strip() if isinstance(cell.value, str) else None
            if val in cols:
                cols[val] = cell.column

        first_row = header_row + 2  # a mintában a fejléc alatt van egy elválasztó sor
        for i in range(len(vorname)):
            r = first_row + i
            # name
            if cols["Name"]: set_cell(ws, r, cols["Name"], nachname[i])
            # vorname
            if cols["Vorname"]: set_cell(ws, r, cols["Vorname"], vorname[i])
            # ausweis
            if cols["Ausweis- Nr.\noder\nKennzeichen"]: set_cell(ws, r, cols["Ausweis- Nr.\noder\nKennzeichen"], ausweis[i])
            # beginn / ende
            if cols["Beginn"]: set_cell(ws, r, cols["Beginn"], beginn[i])
            if cols["Ende"]: set_cell(ws, r, cols["Ende"], ende[i])
            # stunden
            if cols["Anzahl Stunden\n(ohne Pausen)"]: set_cell(ws, r, cols["Anzahl Stunden\n(ohne Pausen)"], stunden[i])

        # összesített óraszám – a táblázat jobb felső összesítő mezőjébe:
        # megkeressük ugyanennek a fejléctextnek a sorában a cellát, és az alatta lévő összesítő mezőbe írunk
        if cols["Anzahl Stunden\n(ohne Pausen)"]:
            sum_cell_row = header_row - 1  # a legtöbb sablonban az összesítő a fejléctől egy sorral feljebb, ugyanazon oszlopban
            set_cell(ws, sum_cell_row, cols["Anzahl Stunden\n(ohne Pausen)"], total_hours)

    # return as xlsx download
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"Arbeitsnachweis_{datum.replace('.', '-')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

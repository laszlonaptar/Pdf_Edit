from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List, Optional
from datetime import datetime, time
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# --------------------- idő / szünetek ---------------------
def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    to_min = lambda t: t.hour * 60 + t.minute
    a1, a2, b1, b2 = to_min(a_start), to_min(a_end), to_min(b_start), to_min(b_end)
    left, right = max(a1, b1), min(a2, b2)
    return max(0, right - left)

def net_hours(begin: str, end: str) -> float:
    b = parse_hhmm(begin)
    e = parse_hhmm(end)
    total = (datetime.combine(datetime.today(), e) - datetime.combine(datetime.today(), b)).total_seconds() / 3600.0
    brk = 0.0
    brk += overlap_minutes(b, e, time(9, 0), time(9, 15)) / 60.0
    brk += overlap_minutes(b, e, time(12, 0), time(12, 45)) / 60.0
    return max(0.0, round(total - brk, 2))

# --------------------- Excel segédek ---------------------
def norm(s: Optional[str]) -> str:
    return (s or "").strip().replace("\n", " ").replace("  ", " ")

def ensure_merge(ws, min_row, min_col, max_row, max_col):
    from openpyxl.utils import get_column_letter
    rng = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    for m in ws.merged_cells.ranges:
        if m.coord == rng:
            return
    ws.merge_cells(rng)

def top_left_of_merge(ws, r, c):
    for m in ws.merged_cells.ranges:
        if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
            return (m.min_row, m.min_col)
    return (r, c)

def set_cell(ws, r, c, value, wrap=False, align_left=False):
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = value
    if wrap or align_left:
        cell.alignment = Alignment(
            wrap_text=wrap,
            horizontal="left" if align_left else (cell.alignment.horizontal if cell.alignment else "general"),
            vertical="top"
        )

def find_cell_eq(ws, text: str):
    t = norm(text)
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and norm(cell.value) == t:
                return (cell.row, cell.column)
    return None

def find_cell_contains(ws, part: str):
    p = norm(part).lower()
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and p in norm(cell.value).lower():
                return (cell.row, cell.column)
    return None

def next_merged_region_right(ws, row: int, col: int):
    """A megadott (row,col) cella merge-régióját megkeresi, és visszaadja a KÖZVETLEN jobb oldali régió bal-felső pontját."""
    # gyűjtsd össze a sor összes merge-régióját
    regions = [m for m in ws.merged_cells.ranges if m.min_row == m.max_row == row]
    regions.sort(key=lambda m: m.min_col)
    # melyikben ül a címke?
    idx = None
    for i, m in enumerate(regions):
        if m.min_col <= col <= m.max_col:
            idx = i
            break
    if idx is None:
        return (row, col + 1)
    # jobb oldali szomszéd
    if idx + 1 < len(regions):
        right = regions[idx + 1]
        return (right.min_row, right.min_col)
    return (row, regions[idx].max_col + 1)

def set_value_right_of_label(ws, label_text: str, value: str):
    pos = find_cell_contains(ws, label_text)
    if not pos:
        return False
    r_label, c_label = pos
    r_target, c_target = next_merged_region_right(ws, r_label, c_label)
    set_cell(ws, r_target, c_target, value, wrap=False, align_left=True)
    return True

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
    # órák
    stunden = []
    total_hours = 0.0
    for b, e in zip(beginn, ende):
        h = net_hours(b, e)
        stunden.append(h)
        total_hours += h
    total_hours = round(total_hours, 2)

    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # Leírás blokk (A6:G15), magasabb sorok
    ensure_merge(ws, 6, 1, 15, 7)
    set_cell(ws, 6, 1, taetigkeit, wrap=True, align_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 28  # nagyobb, hogy a 2. sor se vágódjon le

    # Fejléc mezők – mindig a címke utáni MERGED blokkba írunk
    set_value_right_of_label(ws, "Datum der Leistungsausführung", datum)
    set_value_right_of_label(ws, "Bau und Ausführungsort", bau)
    if bf:
        set_value_right_of_label(ws, "BASF-Beauftragter", bf)

    # Dolgozói táblázat oszlopok felderítése
    pos_name   = find_cell_eq(ws, "Name")
    pos_vor    = find_cell_eq(ws, "Vorname")
    pos_ausw   = find_cell_contains(ws, "Ausweis")
    pos_beginn = find_cell_eq(ws, "Beginn")
    pos_ende   = find_cell_eq(ws, "Ende")
    pos_std    = find_cell_contains(ws, "Anzahl Stunden")

    col_name = pos_name[1] if pos_name else None
    col_vor  = pos_vor[1]  if pos_vor  else None
    col_ausw = pos_ausw[1] if pos_ausw else None
    col_beg  = pos_beginn[1] if pos_beginn else None
    col_end  = pos_ende[1]   if pos_ende else None
    col_std  = pos_std[1]    if pos_std  else None

    rows = [p[0] for p in [pos_name, pos_vor, pos_ausw, pos_beginn, pos_ende, pos_std] if p]
    data_start = (max(rows) + 1) if rows else 1

    for i in range(len(vorname)):
        r = data_start + i
        if col_name: set_cell(ws, r, col_name, nachname[i])
        if col_vor:  set_cell(ws, r, col_vor,  vorname[i])
        if col_ausw: set_cell(ws, r, col_ausw, ausweis[i])
        if col_beg:  set_cell(ws, r, col_beg,  beginn[i])
        if col_end:  set_cell(ws, r, col_end,  ende[i])
        if col_std:  set_cell(ws, r, col_std,  stunden[i])

    # Gesamtstunden
    if col_std:
        pos_total = find_cell_contains(ws, "Gesamtstunden")
        if pos_total:
            set_cell(ws, pos_total[0], col_std, total_hours)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"Arbeitsnachweis_{datum.replace('.', '-')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

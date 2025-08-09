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

# ---------- idő / szünetek ----------
def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    to_min = lambda t: t.hour * 60 + t.minute
    a1, a2, b1, b2 = to_min(a_start), to_min(a_end), to_min(b_start), to_min(b_end)
    left, right = max(a1, b1), min(a2, b2)
    return max(0, right - left)

def net_hours(begin: str, end: str) -> float:
    b = parse_hhmm(begin); e = parse_hhmm(end)
    total = (datetime.combine(datetime.today(), e) - datetime.combine(datetime.today(), b)).total_seconds() / 3600.0
    brk = 0.0
    brk += overlap_minutes(b, e, time(9, 0),  time(9, 15)) / 60.0
    brk += overlap_minutes(b, e, time(12, 0), time(12, 45)) / 60.0
    return max(0.0, round(total - brk, 2))

# ---------- Excel segédek ----------
def norm(s: Optional[str]) -> str:
    return (s or "").strip().replace("\n", " ").replace("  ", " ")

def ensure_merge(ws, r1, c1, r2, c2):
    from openpyxl.utils import get_column_letter
    coord = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
    for m in ws.merged_cells.ranges:
        if m.coord == coord:
            return
    ws.merge_cells(coord)

def in_merge(m, r, c):
    return m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col

def top_left_of_merge(ws, r, c):
    for m in ws.merged_cells.ranges:
        if in_merge(m, r, c):
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

def merged_regions_on_row(ws, row: int):
    regs = [m for m in ws.merged_cells.ranges if m.min_row == m.max_row == row]
    regs.sort(key=lambda m: m.min_col)
    return regs

def set_value_between_labels(ws, left_label: str, right_label: str, value: str):
    """Bal címke ÉS jobb címke közötti merge-blokkba ír."""
    pos_left  = find_cell_contains(ws, left_label)
    pos_right = find_cell_contains(ws, right_label)
    if not pos_left or not pos_right or pos_left[0] != pos_right[0]:
        return False
    row = pos_left[0]
    regs = merged_regions_on_row(ws, row)
    # melyik régióban ülnek a címkék?
    idx_left = idx_right = None
    for i, m in enumerate(regs):
        if in_merge(m, *pos_left):  idx_left  = i
        if in_merge(m, *pos_right): idx_right = i
    if idx_left is None or idx_right is None or idx_right - idx_left < 2:
        # nincs köztes blokk -> essünk vissza a bal címke utáni blokkra
        target = regs[idx_left + 1] if (idx_left is not None and idx_left + 1 < len(regs)) else None
    else:
        # pontosan a kettő közti első köztes blokk
        target = regs[idx_left + 1]
    if not target:
        return False
    set_cell(ws, target.min_row, target.min_col, value, align_left=True)
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
    total = 0.0
    for b, e in zip(beginn, ende):
        h = net_hours(b, e)
        stunden.append(h)
        total += h
    total = round(total, 2)

    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # Leírás blokk (A6:G15)
    ensure_merge(ws, 6, 1, 15, 7)
    set_cell(ws, 6, 1, taetigkeit, wrap=True, align_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 28

    # Fejléc: bal–köztes–jobb minta
    set_value_between_labels(ws, "Datum der Leistungsausführung", "Einzelauftrags", datum)
    set_value_between_labels(ws, "Bau und Ausführungsort", "Einzelauftrags", bau)
    if bf:
        # itt is a „Einzelauftrags” előtti köztes mezőbe írunk
        ok = set_value_between_labels(ws, "BASF-Beauftragter", "Einzelauftrags", bf)
        if not ok:
            # ha nem találja a mintát, essen vissza a sima „következő blokkra” írásra
            pos = find_cell_contains(ws, "BASF-Beauftragter")
            if pos:
                regs = merged_regions_on_row(ws, pos[0])
                for i, m in enumerate(regs):
                    if in_merge(m, *pos) and i + 1 < len(regs):
                        set_cell(ws, regs[i + 1].min_row, regs[i + 1].min_col, bf, align_left=True)
                        break

    # Dolgozói táblázat oszlopok
    c_name   = (find_cell_eq(ws, "Name") or (0,0))[1]
    c_vor    = (find_cell_eq(ws, "Vorname") or (0,0))[1]
    c_ausw   = (find_cell_contains(ws, "Ausweis") or (0,0))[1]
    c_beginn = (find_cell_eq(ws, "Beginn") or (0,0))[1]
    c_ende   = (find_cell_eq(ws, "Ende") or (0,0))[1]
    c_std    = (find_cell_contains(ws, "Anzahl Stunden") or (0,0))[1]
    header_rows = [p[0] for p in filter(None, [
        find_cell_eq(ws, "Name"),
        find_cell_eq(ws, "Vorname"),
        find_cell_contains(ws, "Ausweis"),
        find_cell_eq(ws, "Beginn"),
        find_cell_eq(ws, "Ende"),
        find_cell_contains(ws, "Anzahl Stunden"),
    ])]
    start_row = max(header_rows) + 1 if header_rows else 1

    for i in range(len(vorname)):
        r = start_row + i
        if c_name:   set_cell(ws, r, c_name,   nachname[i])
        if c_vor:    set_cell(ws, r, c_vor,    vorname[i])
        if c_ausw:   set_cell(ws, r, c_ausw,   ausweis[i])
        if c_beginn: set_cell(ws, r, c_beginn, beginn[i])
        if c_ende:   set_cell(ws, r, c_ende,   ende[i])
        if c_std:    set_cell(ws, r, c_std,    stunden[i])

    # Gesamtstunden
    pos_total_label = find_cell_contains(ws, "Gesamtstunden")
    if pos_total_label and c_std:
        set_cell(ws, pos_total_label[0], c_std, total)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"Arbeitsnachweis_{datum.replace('.', '-')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

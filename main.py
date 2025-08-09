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

# ---------- breaks & time ----------
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

# ---------- Excel helpers ----------
def ensure_merge(ws, min_row, min_col, max_row, max_col):
    from openpyxl.utils import get_column_letter
    rng = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    for m in ws.merged_cells.ranges:
        if m.coord == rng:
            return
    ws.merge_cells(rng)

def top_left_of_merge(ws, r, c):
    # FIX: numerikus ellenőrzés, nem "(r,c) in m"
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

def find_label_cell(ws, label_text: str):
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == label_text:
                return (cell.row, cell.column)
    return None

def right_region_top_left(ws, row: int, col: int):
    # ugyanazon a soron lévő merge tartományok közül a következő blokknak a bal felső cellája
    regions = [m for m in ws.merged_cells.ranges if m.min_row == m.max_row == row]
    regions.sort(key=lambda m: m.min_col)
    for i, m in enumerate(regions):
        if m.min_col <= col <= m.max_col:
            if i + 1 < len(regions):
                return (regions[i+1].min_row, regions[i+1].min_col)
            break
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
    # óraszámok
    stunden = []
    total_hours = 0.0
    for b, e in zip(beginn, ende):
        h = net_hours(b, e)
        stunden.append(h)
        total_hours += h
    total_hours = round(total_hours, 2)

    # sablon betöltése
    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # 1) napi leírás A6:G15
    ensure_merge(ws, 6, 1, 15, 7)  # A6:G15
    set_cell(ws, 6, 1, taetigkeit, wrap=True, align_left=True)

    # 2) címke melletti mezők
    if (pos := find_label_cell(ws, "Datum der Leistungsausführung:")):
        r, c = right_region_top_left(ws, pos[0], pos[1])
        set_cell(ws, r, c, datum)

    if (pos := find_label_cell(ws, "Bau und Ausführungsort:")):
        r, c = right_region_top_left(ws, pos[0], pos[1])
        set_cell(ws, r, c, bau)

    if bf and (pos := find_label_cell(ws, "BASF-Beauftragter, Org.-Code:")):
        r, c = right_region_top_left(ws, pos[0], pos[1])
        set_cell(ws, r, c, bf)

    # 3) dolgozói tábla
    header_row = None
    cols = {}
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "Name":
                header_row = cell.row
                # fejlécek oszlopai
                for hc in ws[header_row]:
                    val = hc.value.strip() if isinstance(hc.value, str) else None
                    if val:
                        cols[val] = hc.column
                break
        if header_row:
            break

    # fejlécek kulcsai (sablon szövege alapján)
    KEY_NAME = "Name"
    KEY_VOR = "Vorname"
    KEY_AUSW = "Ausweis- Nr.\noder\nKennzeichen"
    KEY_BEG = "Beginn"
    KEY_END = "Ende"
    KEY_STD = "Anzahl Stunden\n(ohne Pausen)"

    if header_row:
        first_row = header_row + 2  # elválasztó sorral számolunk
        for i in range(len(vorname)):
            r = first_row + i
            if KEY_NAME in cols: set_cell(ws, r, cols[KEY_NAME], nachname[i])
            if KEY_VOR in cols: set_cell(ws, r, cols[KEY_VOR], vorname[i])
            if KEY_AUSW in cols: set_cell(ws, r, cols[KEY_AUSW], ausweis[i])
            if KEY_BEG in cols: set_cell(ws, r, cols[KEY_BEG], beginn[i])
            if KEY_END in cols: set_cell(ws, r, cols[KEY_END], ende[i])
            if KEY_STD in cols: set_cell(ws, r, cols[KEY_STD], stunden[i])

        # összesített óraszám: a fejléc oszlopában egy sorral feljebb
        if KEY_STD in cols:
            sum_row = header_row - 1
            set_cell(ws, sum_row, cols[KEY_STD], total_hours)

    # letöltés
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"Arbeitsnachweis_{datum.replace('.', '-')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

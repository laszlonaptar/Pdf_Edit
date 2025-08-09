from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime, time, timedelta
import uuid
import os

app = FastAPI()

# statikus fájlok + HTML sablon
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# --------- SEGÉDFÜGGVÉNYEK (merge-kezelés, időszámítás) ---------
def top_left_of_merge(ws, r: int, c: int):
    """Ha (r,c) egy összevont tartományon belül van, adja vissza a bal-felső cella koordinátáját."""
    for m in ws.merged_cells.ranges:
        if (r, c) in m:
            return m.min_row, m.min_col
    return r, c


def set_cell(ws, r: int, c: int, value, wrap=False, align_left=False):
    """Írás összevont cellák figyelembevételével."""
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(row=r0, column=c0)
    cell.value = value
    if wrap or align_left:
        cell.alignment = Alignment(
            wrap_text=True if wrap else cell.alignment.wrap_text,
            horizontal="left" if align_left else cell.alignment.horizontal,
            vertical="top",
        )


def parse_hhmm(s: str) -> time | None:
    if not s:
        return None
    try:
        return datetime.strptime(s.strip(), "%H:%M").time()
    except Exception:
        return None


def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    """Két időintervallum átfedése percekben (zárt-balos, nyílt-jobbos logika elég ide)."""
    dt0 = datetime(2000, 1, 1)
    A1 = dt0.replace(hour=a_start.hour, minute=a_start.minute)
    A2 = dt0.replace(hour=a_end.hour, minute=a_end.minute)
    B1 = dt0.replace(hour=b_start.hour, minute=b_start.minute)
    B2 = dt0.replace(hour=b_end.hour, minute=b_end.minute)
    start = max(A1, B1)
    end = min(A2, B2)
    if end <= start:
        return 0
    return int((end - start).total_seconds() // 60)


def worked_minutes_with_breaks(beg: time, end: time) -> int:
    """Nettó percek levonva a fix szüneteket: 09:00–09:15 (15p) és 12:00–12:45 (45p)."""
    if not beg or not end:
        return 0
    if end <= beg:
        return 0
    total = (datetime.combine(datetime.min, end) - datetime.combine(datetime.min, beg)).seconds // 60
    # fix szünetek
    total -= overlap_minutes(beg, end, time(9, 0), time(9, 15))
    total -= overlap_minutes(beg, end, time(12, 0), time(12, 45))
    return max(total, 0)


def fmt_hours(mins: int) -> str:
    h = mins // 60
    m = mins % 60
    return f"{h}:{m:02d}"


# --------- BEKÜLDÉS ---------
@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    taetigkeit: str = Form(""),
    basf_beauftragter: str = Form(""),
    # dolgozók (max 5)
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""),
    beginn1: str = Form(""),  ende1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""),
    beginn2: str = Form(""),  ende2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""),
    beginn3: str = Form(""),  ende3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""),
    beginn4: str = Form(""),  ende4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""),
    beginn5: str = Form(""),  ende5: str = Form(""),
):
    # minimális validáció
    if not bau or not datum:
        return HTMLResponse(
            content='{"detail":"Kötelező mezők hiányoznak: projekt/bau és dátum."}',
            status_code=400,
            media_type="application/json",
        )

    # dolgozók összegyűjtése
    workers_raw = [
        (vorname1, nachname1, ausweis1, beginn1, ende1),
        (vorname2, nachname2, ausweis2, beginn2, ende2),
        (vorname3, nachname3, ausweis3, beginn3, ende3),
        (vorname4, nachname4, ausweis4, beginn4, ende4),
        (vorname5, nachname5, ausweis5, beginn5, ende5),
    ]
    workers = []
    total_minutes = 0
    for v, n, a, b, e in workers_raw:
        if (v or n or a or b or e):  # van-e bármilyen adat
            beg = parse_hhmm(b)
            end = parse_hhmm(e)
            mins = worked_minutes_with_breaks(beg, end) if (beg and end) else 0
            total_minutes += mins
            workers.append({
                "name": f"{v.strip()} {n.strip()}".strip(),
                "ausweis": a.strip(),
                "beginn": b.strip(),
                "ende": e.strip(),
                "mins": mins
            })

    # Excel sablon betöltése
    template_path = "GP-t.xlsx"
    if not os.path.exists(template_path):
        return HTMLResponse(
            content='{"detail":"Hiányzik a GP-t.xlsx sablon a gyökérben."}',
            status_code=500,
            media_type="application/json",
        )

    wb = load_workbook(template_path)
    ws = wb.active

    # --- FEJLÉC MEZŐK (összevont cellákra figyelve) ---
    # Dátum -> D3
    set_cell(ws, 3, 4, datum)
    # Bau -> D4
    set_cell(ws, 4, 4, bau)
    # BASF-Beauftragter -> E3
    if basf_beauftragter:
        set_cell(ws, 3, 5, basf_beauftragter)

    # --- Tätigkeiten (A6:G15 egy nagy összevont blokk) ---
    # biztos ami biztos: egyesítjük (idempotens, ha már egyesítve van, nem gond)
    try:
        ws.merge_cells(start_row=6, start_column=1, end_row=15, end_column=7)
    except Exception:
        pass
    set_cell(ws, 6, 1, taetigkeit, wrap=True, align_left=True)

    # --- Dolgozók blokk (pozíciók feltételezve, ha eltér: finomhangoljuk) ---
    # Kiindulás: név -> A18..A22, Ausweis -> C18..C22, Beginn -> E18..E22, Ende -> F18..F22, Óra -> G18..G22
    start_row = 18
    for idx, w in enumerate(workers[:5]):
        r = start_row + idx
        set_cell(ws, r, 1, w["name"])          # A: név
        set_cell(ws, r, 3, w["ausweis"])       # C: Ausweis
        set_cell(ws, r, 5, w["beginn"])        # E: Beginn
        set_cell(ws, r, 6, w["ende"])          # F: Ende
        set_cell(ws, r, 7, fmt_hours(w["mins"]))  # G: ledolgozott idő

    # --- Összesített munkaidő (G16) ---
    set_cell(ws, 16, 7, fmt_hours(total_minutes))

    # ideiglenes fájlnév és mentés
    out_name = f"GP-t_filled_{uuid.uuid4().hex[:8]}.xlsx"
    out_path = os.path.join("/tmp", out_name)
    wb.save(out_path)

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=out_name,
    )

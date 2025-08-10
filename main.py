from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from datetime import datetime, time
from io import BytesIO
import uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ----------------- merged-cell safe helpers -----------------

def a1(c: int, r: int) -> str:
    """(r,c) -> 'A1' formátum (openpyxl ezt várja)."""
    return f"{get_column_letter(c)}{r}"

def find_cell_by_text(ws, text: str):
    for r in range(1, 41):
        for c in range(1, 21):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == text.strip():
                return r, c
    return None

def merged_top_left(ws, r, c):
    """Ha (r,c) merged tartományban van, add vissza annak bal-felső sarkát."""
    coord = a1(c, r)
    for rng in ws.merged_cells.ranges:
        if coord in rng:                 # <-- A1 cím ellenőrzés
            return rng.min_row, rng.min_col
    return r, c

def merged_block_right_top_left(ws, r, c):
    """A címke-blokk UTÁNI (jobbra lévő) értékblokk bal-felső cellája."""
    r0, c0 = merged_top_left(ws, r, c)

    # keresd meg a címke saját merged blokkját
    label_rng = None
    coord_label = a1(c, r)
    for rng in ws.merged_cells.ranges:
        if coord_label in rng:
            label_rng = rng
            break

    if label_rng:
        rr, cc = r0, label_rng.max_col + 1
    else:
        rr, cc = r, c + 1

    return merged_top_left(ws, rr, cc)

def set_value(ws, r, c, value, wrap=False, valign_top=False, halign_left=True):
    rr, cc = merged_top_left(ws, r, c)
    cell = ws.cell(rr, cc)
    cell.value = value
    cell.alignment = Alignment(
        wrap_text=wrap,
        vertical="top" if valign_top else None,
        horizontal="left" if halign_left else None,
    )
    return cell

def set_value_right_of_label(ws, label_text: str, value, wrap=False):
    pos = find_cell_by_text(ws, label_text)
    if not pos:
        return False
    r, c = pos
    rr, cc = merged_block_right_top_left(ws, r, c)
    set_value(ws, rr, cc, value, wrap=wrap, valign_top=False, halign_left=True)
    return True

# ----------------- time helpers -----------------

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a0: time, a1: time, b0: time, b1: time) -> int:
    to_dt = lambda t: datetime.combine(datetime(2000,1,1).date(), t)
    start = max(to_dt(a0), to_dt(b0))
    end   = min(to_dt(a1), to_dt(b1))
    return max(0, int((end - start).total_seconds() // 60))

def worked_hours_with_breaks(beg: str, end: str) -> float:
    t0 = parse_hhmm(beg); t1 = parse_hhmm(end)
    if t1 <= t0:
        return 0.0
    total = (datetime.combine(datetime.today(), t1) -
             datetime.combine(datetime.today(), t0)).total_seconds() / 60.0
    pauses = [(time(9,0), time(9,15)), (time(12,0), time(12,45))]
    ded = sum(overlap_minutes(t0, t1, p0, p1) for p0, p1 in pauses)
    return round(max(0.0, (total - ded) / 60.0), 2)

# ----------------- routes -----------------

@app.get("/")
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    bf: str = Form(""),
    beschreibung: str = Form(...),

    nachname1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),

    nachname2: str = Form(None), vorname2: str = Form(None),
    ausweis2: str = Form(None),  beginn2: str = Form(None), ende2: str = Form(None),
    nachname3: str = Form(None), vorname3: str = Form(None),
    ausweis3: str = Form(None),  beginn3: str = Form(None), ende3: str = Form(None),
    nachname4: str = Form(None), vorname4: str = Form(None),
    ausweis4: str = Form(None),  beginn4: str = Form(None), ende4: str = Form(None),
    nachname5: str = Form(None), vorname5: str = Form(None),
    ausweis5: str = Form(None),  beginn5: str = Form(None), ende5: str = Form(None),
):
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Fejléc mezők
    set_value_right_of_label(ws, "Datum der Leistungsausführung:", datum)
    set_value_right_of_label(ws, "Bau und Ausführungsort:", bau)
    if bf:
        set_value_right_of_label(ws, "BASF-Beauftragter, Org.-Code:", bf)

    # Beschreibung (A6:G15 blokk tetejére, tördelve)
    set_value(ws, 6, 1, beschreibung, wrap=True, valign_top=True, halign_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 22

    # Táblázat oszlopok
    headers = {
        "name": "Name",
        "vorname": "Vorname",
        "ausweis": "Ausweis- Nr.\noder\nKennzeichen",
        "beginn": "Beginn",
        "ende": "Ende",
        "stunden": "Anzahl Stunden\n(ohne Pausen)",
    }
    cols = {}
    header_row = None
    for key, label in headers.items():
        pos = find_cell_by_text(ws, label)
        if not pos and key == "ausweis":
            pos = find_cell_by_text(ws, "Ausweis- Nr.")
        if pos:
            r, c = merged_top_left(ws, *pos)
            cols[key] = c
            header_row = r
    if header_row is None:
        header_row = 0
    start_row = header_row + 1

    # Dolgozók listája
    raw = [
        (nachname1, vorname1, ausweis1, beginn1, ende1),
        (nachname2, vorname2, ausweis2, beginn2, ende2),
        (nachname3, vorname3, ausweis3, beginn3, ende3),
        (nachname4, vorname4, ausweis4, beginn4, ende4),
        (nachname5, vorname5, ausweis5, beginn5, ende5),
    ]
    workers = [w for w in raw if all(w)]

    total_hours = 0.0
    r = start_row
    for (nachn, vorn, ausw, beg, end_) in workers:
        h = worked_hours_with_breaks(beg, end_)
        total_hours += h
        if cols.get("name"):    set_value(ws, r, cols["name"], nachn)
        if cols.get("vorname"): set_value(ws, r, cols["vorname"], vorn)
        if cols.get("ausweis"): set_value(ws, r, cols["ausweis"], ausw)
        if cols.get("beginn"):  set_value(ws, r, cols["beginn"], beg)
        if cols.get("ende"):    set_value(ws, r, cols["ende"], end_)
        if cols.get("stunden"): set_value(ws, r, cols["stunden"], h)
        r += 1

    # Gesamtstunden
    if not set_value_right_of_label(ws, "Gesamtstunden", total_hours):
        if cols.get("stunden"):
            set_value(ws, r, cols["stunden"], total_hours)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"Leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename=\"{fname}\"'}
    )

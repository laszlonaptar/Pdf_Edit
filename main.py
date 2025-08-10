from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from datetime import datetime, time, timedelta
from io import BytesIO
import uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------- Helpers for merged-cell-safe writing ----------

def find_cell_by_text(ws, text: str):
    """Find exact matching cell by value (scan the first ~40 rows, 20 cols)."""
    for r in range(1, 41):
        for c in range(1, 21):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == text.strip():
                return r, c
    return None

def merged_top_left(ws, r, c):
    """If (r,c) in a merged range, return its top-left; else return (r,c)."""
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            return rng.min_row, rng.min_col
    return r, c

def merged_block_right_top_left(ws, r, c):
    """
    Given a label starting cell (r,c) which may be merged, return the
    top-left of the NEXT merged block right next to it (value cell).
    """
    # current block
    r0, c0 = merged_top_left(ws, r, c)
    # find its merged range
    cur = None
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            cur = rng
            break
    if cur:
        # immediate cell to the right of current block
        rr, cc = r0, cur.max_col + 1
    else:
        rr, cc = r, c + 1

    # if that right neighbor itself is part of a merged range,
    # normalize to its top-left.
    return merged_top_left(ws, rr, cc)

def set_value(ws, r, c, value, wrap=False, valign_top=False, halign_left=True):
    """Write to cell (r,c), being robust for merged cells (write to top-left)."""
    rr, cc = merged_top_left(ws, r, c)
    cell = ws.cell(rr, cc)
    cell.value = value
    align = Alignment(
        wrap_text=wrap,
        vertical="top" if valign_top else None,
        horizontal="left" if halign_left else None,
    )
    cell.alignment = align
    return cell

def set_value_right_of_label(ws, label_text: str, value, wrap=False):
    """Find label cell by text, then write value into the merged block on its right."""
    pos = find_cell_by_text(ws, label_text)
    if not pos:
        return False
    r, c = pos
    rr, cc = merged_block_right_top_left(ws, r, c)
    set_value(ws, rr, cc, value, wrap=wrap, valign_top=False, halign_left=True)
    return True

# ---------- Time helpers ----------

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    """Return overlap in minutes between two [start,end) intervals in the same day."""
    to_dt = lambda t: datetime.combine(datetime(2000,1,1).date(), t)
    a0, a1 = to_dt(a_start), to_dt(a_end)
    b0, b1 = to_dt(b_start), to_dt(b_end)
    start = max(a0, b0)
    end = min(a1, b1)
    return max(0, int((end - start).total_seconds() // 60))

def worked_hours_with_breaks(beg: str, end: str) -> float:
    """Compute worked hours minus fixed breaks (09:00-09:15 and 12:00-12:45)."""
    t0 = parse_hhmm(beg)
    t1 = parse_hhmm(end)
    if t1 <= t0:
        # next day guard: treat as no time
        return 0.0
    total_min = (datetime.combine(datetime.today(), t1) -
                 datetime.combine(datetime.today(), t0)).total_seconds() / 60.0

    brks = [(time(9, 0), time(9, 15)), (time(12, 0), time(12, 45))]
    deducted = 0
    for b0, b1 in brks:
        deducted += overlap_minutes(t0, t1, b0, b1)

    hours = max(0.0, (total_min - deducted) / 60.0)
    return round(hours, 2)

# ---------- Routes ----------

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

    # minimum 1 dolgozó (1. sor)
    nachname1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),

    # opcionális 2–5. dolgozó (ha a frontend később hozzáadja)
    nachname2: str = Form(None),
    vorname2: str = Form(None),
    ausweis2: str = Form(None),
    beginn2: str = Form(None),
    ende2: str = Form(None),

    nachname3: str = Form(None),
    vorname3: str = Form(None),
    ausweis3: str = Form(None),
    beginn3: str = Form(None),
    ende3: str = Form(None),

    nachname4: str = Form(None),
    vorname4: str = Form(None),
    ausweis4: str = Form(None),
    beginn4: str = Form(None),
    ende4: str = Form(None),

    nachname5: str = Form(None),
    vorname5: str = Form(None),
    ausweis5: str = Form(None),
    beginn5: str = Form(None),
    ende5: str = Form(None),
):
    # 1) sablon megnyitása
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # 2) Fejléc mezők (címkék melletti érték cellák)
    # Dátum
    set_value_right_of_label(ws, "Datum der Leistungsausführung:", datum)
    # Bau
    set_value_right_of_label(ws, "Bau und Ausführungsort:", bau)
    # BASF-Beauftragter
    if bf:
        set_value_right_of_label(ws, "BASF-Beauftragter, Org.-Code:", bf)

    # 3) Beschreibung (A6:G15 feltételezve)
    #   – sortörés + top align; némi sormagasság biztos amiatt, hogy ne vágódjon
    set_value(ws, 6, 1, beschreibung, wrap=True, valign_top=True, halign_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 22  # kis extra hely

    # 4) Táblázat oszlopfejlécek pozíciói
    # Keresünk header címkéket a táblázat első sorában („Name”, „Vorname”, …)
    # és ezek oszlopába írunk a KÖVETKEZŐ sor(ok)ba.
    def col_of_header(header_text: str):
        pos = find_cell_by_text(ws, header_text)
        if not pos:
            return None
        # ha a header összevont, a bal felső oszlop az igazi
        r, c = merged_top_left(ws, *pos)
        return r, c

    hdrs = {
        "name": "Name",
        "vorname": "Vorname",
        "ausweis": "Ausweis- Nr.\noder\nKennzeichen",
        "beginn": "Beginn",
        "ende": "Ende",
        "stunden": "Anzahl Stunden\n(ohne Pausen)",
    }

    # oszlopok felvétele
    cols = {}
    for key, label in hdrs.items():
        pos = find_cell_by_text(ws, label)
        if not pos:
            # néhány sablonban az „Ausweis” sor máshogy tördel
            if key == "ausweis":
                alt = "Ausweis- Nr."
                pos = find_cell_by_text(ws, alt)
                if not pos:
                    pos = find_cell_by_text(ws, "Ausweis Nr.")
        if pos:
            r, c = merged_top_left(ws, *pos)
            cols[key] = c
            header_row = r
        else:
            cols[key] = None

    # az első kitöltendő adat sor a header alatti sor
    start_row = header_row + 1 if 'name' in cols and cols['name'] else header_row + 1

    # 5) Dolgozók listája összeállítás (csak a nem üres mezőkkel)
    raw_workers = [
        (nachname1, vorname1, ausweis1, beginn1, ende1),
        (nachname2, vorname2, ausweis2, beginn2, ende2),
        (nachname3, vorname3, ausweis3, beginn3, ende3),
        (nachname4, vorname4, ausweis4, beginn4, ende4),
        (nachname5, vorname5, ausweis5, beginn5, ende5),
    ]
    workers = []
    for w in raw_workers:
        if w[0] and w[1] and w[2] and w[3] and w[4]:
            workers.append(w)

    total_hours = 0.0
    r = start_row
    for (nachn, vorn, ausw, beg, end_) in workers:
        h = worked_hours_with_breaks(beg, end_)
        total_hours += h
        if cols.get("name"):    set_value(ws, r, cols["name"], nachn, halign_left=True)
        if cols.get("vorname"): set_value(ws, r, cols["vorname"], vorn, halign_left=True)
        if cols.get("ausweis"): set_value(ws, r, cols["ausweis"], ausw, halign_left=True)
        if cols.get("beginn"):  set_value(ws, r, cols["beginn"], beg, halign_left=True)
        if cols.get("ende"):    set_value(ws, r, cols["ende"], end_, halign_left=True)
        if cols.get("stunden"): set_value(ws, r, cols["stunden"], h, halign_left=True)
        r += 1

    # 6) Gesamtstunden – megpróbáljuk a „Gesamtstunden” felirat melletti cellába tenni
    # (ha nincs ilyen, akkor az utolsó sor „Stunden” oszlopába)
    if not set_value_right_of_label(ws, "Gesamtstunden", total_hours):
        if cols.get("stunden"):
            set_value(ws, r, cols["stunden"], total_hours, halign_left=True)

    # 7) Visszaadás
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"Leistungsnachweis_{uuid.uuid4().hex[:8]}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

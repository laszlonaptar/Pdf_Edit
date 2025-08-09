from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

from datetime import datetime, time, timedelta
import uuid
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_PATH = "GP-t.xlsx"

# ---------- helpers ----------

def find_label_cell(ws, label_text: str, max_row=20, max_col=20):
    """Megkeresi a cellát, amely pontosan a kapott feliratot tartalmazza."""
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() == label_text.strip():
                return r, c
    return None

def right_value_cell_of_label(ws, label_text: str):
    """Visszaadja a címkecella sorában a tőle jobbra eső köv. cellát,
    és ha mergelt tartomány, a tartomány bal-felső celláját."""
    rc = find_label_cell(ws, label_text)
    if not rc:
        return None
    r, c = rc
    # alapértelmezetten a szomszéd jobbra
    target_r, target_c = r, c + 1

    # ha mergelt tartomány felső-bal cellája máshol van, ugorjunk oda
    for m in ws.merged_cells.ranges:
        if (target_r, target_c) in m:
            target_r = m.min_row
            target_c = m.min_col
            break
    return target_r, target_c

def write_cell(ws, r, c, value, wrap=False, align_left=False):
    cell = ws.cell(row=r, column=c)
    cell.value = value
    cell.alignment = Alignment(
        wrap_text=wrap or cell.alignment.wrap_text,
        horizontal=("left" if align_left else cell.alignment.horizontal)
    )

def write_into_range(ws, min_row, min_col, max_row, max_col, text):
    """Tartomány (pl. A6:G15) bal-felső cellájába ír, wrap + balra igazítással."""
    # egyesítve legyen – ha nincs, nem baj, a bal-felsőbe írunk
    cell = ws.cell(row=min_row, column=min_col)
    cell.value = text
    cell.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def hours_without_breaks(start: time, end: time) -> float:
    """Összóra fix szünetek levonásával (09:00–09:15 és 12:00–12:45)."""
    dt0 = datetime(2000,1,1, start.hour, start.minute)
    dt1 = datetime(2000,1,1, end.hour, end.minute)
    if dt1 < dt0:
        dt1 += timedelta(days=1)

    total = (dt1 - dt0).total_seconds() / 3600.0

    def overlap(a0, a1, b0, b1):
        x0 = max(a0, b0)
        x1 = min(a1, b1)
        return max(0.0, (x1 - x0).total_seconds()/3600.0)

    # szünetek
    br_morn_0 = datetime(2000,1,1,9,0)
    br_morn_1 = datetime(2000,1,1,9,15)
    br_lunch_0 = datetime(2000,1,1,12,0)
    br_lunch_1 = datetime(2000,1,1,12,45)

    total -= overlap(dt0, dt1, br_morn_0, br_morn_1)  # 0.25h max
    total -= overlap(dt0, dt1, br_lunch_0, br_lunch_1)  # 0.75h max

    return round(max(total, 0.0), 2)

# ---------- pages ----------

@app.get("/")
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# ---------- Excel generation ----------

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    beschreibung: str = Form(...),

    vorname1: str = Form(...),
    nachname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),

    vorname2: str = Form(None),
    nachname2: str = Form(None),
    ausweis2: str = Form(None),
    beginn2: str = Form(None),
    ende2: str = Form(None),

    vorname3: str = Form(None),
    nachname3: str = Form(None),
    ausweis3: str = Form(None),
    beginn3: str = Form(None),
    ende3: str = Form(None),

    vorname4: str = Form(None),
    nachname4: str = Form(None),
    ausweis4: str = Form(None),
    beginn4: str = Form(None),
    ende4: str = Form(None),

    vorname5: str = Form(None),
    nachname5: str = Form(None),
    ausweis5: str = Form(None),
    beginn5: str = Form(None),
    ende5: str = Form(None),

    basf_beauftragter: str = Form(None),
    geraet: str = Form(None),
):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # 1) Fejlécmezők – címkék alapján
    # Dátum
    pos = right_value_cell_of_label(ws, "Datum der Leistungsausführung:")
    if pos:
        write_cell(ws, pos[0], pos[1], datum)

    # Bau
    pos = right_value_cell_of_label(ws, "Bau und Ausführungsort:")
    if pos:
        write_cell(ws, pos[0], pos[1], bau)

    # BASF-Beauftragter
    if basf_beauftragter:
        pos = right_value_cell_of_label(ws, "BASF-Beauftragter, Org.-Code:")
        if pos:
            write_cell(ws, pos[0], pos[1], basf_beauftragter)

    # 2) Napi leírás – A6:G15
    write_into_range(ws, 6, 1, 15, 7, beschreibung)

    # 3) Dolgozók
    # A táblázat első sora jellemzően a fejléc után indul.
    # A korábbi mintád alapján az első adat sor rögzített: sor 23 környéke helyett
    # stabilan megcélozzuk a látható adatblokkot:
    START_ROW = 23  # ha csúszna, ezt az 1 értéket kell majd hangolni
    COL_NAME = 2
    COL_VORNAME = 3
    COL_AUSWEIS = 4
    COL_BEGINN = 6
    COL_ENDE = 7
    COL_STUNDEN = 8
    COL_GERAET = 10

    mitarbeiter = []
    for i in range(1, 6):
        vn = locals().get(f"vorname{i}")
        nn = locals().get(f"nachname{i}")
        aw = locals().get(f"ausweis{i}")
        bg = locals().get(f"beginn{i}")
        en = locals().get(f"ende{i}")
        if nn and vn and aw and bg and en:
            mitarbeiter.append((nn, vn, aw, bg, en))

    total_hours = 0.0
    for idx, (nn, vn, aw, bg, en) in enumerate(mitarbeiter, start=0):
        r = START_ROW + idx
        ws.cell(row=r, column=COL_NAME).value = nn
        ws.cell(row=r, column=COL_VORNAME).value = vn
        ws.cell(row=r, column=COL_AUSWEIS).value = aw
        ws.cell(row=r, column=COL_BEGINN).value = bg
        ws.cell(row=r, column=COL_ENDE).value = en

        h = hours_without_breaks(parse_hhmm(bg), parse_hhmm(en))
        total_hours += h
        ws.cell(row=r, column=COL_STUNDEN).value = h
        if geraet:
            ws.cell(row=r, column=COL_GERAET).value = geraet

    # 4) Összes óraszám (táblázat alján – a te sablonodban egy "Gesamtstunden" cella jobb oldalán)
    # Keresünk egy "Gesamtstunden" feliratot, és a szomszéd jobbra cellába írunk:
    pos_total = right_value_cell_of_label(ws, "Gesamtstunden")
    if pos_total:
        write_cell(ws, pos_total[0], pos_total[1], total_hours)

    # 5) Mentés és visszaadás
    out_name = f"ausgabe_{uuid.uuid4().hex}.xlsx"
    wb.save(out_name)
    return FileResponse(out_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Leistungsnachweis.xlsx")

# main.py
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from io import BytesIO
from datetime import datetime, time
import re
from typing import List, Tuple, Dict

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

app = FastAPI()

# Statikus és template kiszolgálás
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


# =========================
# Segédfüggvények
# =========================
def parse_hhmm(v: str) -> time:
    v = (v or "").strip()
    return datetime.strptime(v, "%H:%M").time()

def minutes_between(a: time, b: time) -> int:
    """ percek a és b között (feltételezzük: ugyanaz a nap) """
    da = datetime.combine(datetime.today().date(), a)
    db = datetime.combine(datetime.today().date(), b)
    return int((db - da).total_seconds() // 60)

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    """ két intervallum (a és b) átfedése percekben """
    start = max(a_start, b_start)
    end = min(a_end, b_end)
    if end <= start:
        return 0
    return minutes_between(start, end)

def total_minutes_with_breaks(intervals: List[Tuple[time, time]]) -> int:
    """
    Több dolgozó időintervallumából ([(start,end), ...]) kiszámítja az össz perceket,
    és levonja a fix szüneteket (09:00–09:15 és 12:00–12:45).
    """
    total = 0
    # fix szünetek
    br1 = (time(9, 0), time(9, 15))    # 0.25 óra
    br2 = (time(12, 0), time(12, 45))  # 0.75 óra

    for start, end in intervals:
        if end <= start:
            continue
        mins = minutes_between(start, end)
        mins -= overlap_minutes(start, end, br1[0], br1[1])
        mins -= overlap_minutes(start, end, br2[0], br2[1])
        total += max(0, mins)
    return total

def top_left_of_merge(ws, row: int, col: int) -> Tuple[int, int]:
    """
    Ha (row, col) egy összevont (merged) tartományban van, adja vissza a bal-felső
    cella (row, col) koordinátáját; különben saját magát.
    (Ez javítja a "MergedCell ... read-only" hibát.)
    """
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col
    return row, col

def set_cell_rc(ws, row: int, col: int, value):
    r, c = top_left_of_merge(ws, row, col)
    ws.cell(row=r, column=c, value=value)

def set_cell_a1(ws, a1: str, value):
    # Pl. "D6" -> (row, col)
    col_letters = "".join([ch for ch in a1 if ch.isalpha()]).upper()
    row_digits = "".join([ch for ch in a1 if ch.isdigit()])
    col = column_index_from_string(col_letters)
    row = int(row_digits)
    set_cell_rc(ws, row, col, value)

def collect_workers_from_form(form: Dict[str, str]):
    """
    Az űrlap mezői várhatóan ilyenek:
      vorname1, nachname1, ausweis1, beginn1, ende1
      vorname2, nachname2, ... stb.
    Visszaad: lista dict-ekkel és a számított percekkel.
    """
    # Gyűjtsük össze az indexeket a kulcsok végéről
    idxs = set()
    pat = re.compile(r"(vorname|nachname|ausweis|beginn|ende)(\d+)$")
    for k in form.keys():
        m = pat.match(k)
        if m:
            idxs.add(int(m.group(2)))
    idxs = sorted(list(idxs))

    workers = []
    intervals = []
    for i in idxs:
        vn = form.get(f"vorname{i}", "").strip()
        nn = form.get(f"nachname{i}", "").strip()
        az = form.get(f"ausweis{i}", "").strip()
        b  = form.get(f"beginn{i}", "").strip()
        e  = form.get(f"ende{i}", "").strip()
        if not (vn or nn or az or b or e):
            continue

        start_t = parse_hhmm(b) if b else None
        end_t   = parse_hhmm(e) if e else None

        work_mins = 0
        if start_t and end_t and end_t > start_t:
            work_mins = total_minutes_with_breaks([(start_t, end_t)])
            intervals.append((start_t, end_t))

        workers.append({
            "vorname": vn,
            "nachname": nn,
            "ausweis": az,
            "beginn": b,
            "ende": e,
            "mins": work_mins
        })

    # összesített idő (összes dolgozó)
    total_mins = 0
    if intervals:
        total_mins = total_minutes_with_breaks(intervals)

    return workers, total_mins


# =========================
# POZÍCIÓK – KÖNNYEN ÁTÍRHATÓ
# =========================
# Itt tudod később finoman hangolni, melyik adat melyik cellába menjen.
POS = {
    "date":      "D4",     # dátum
    "bau":       "B4",     # Bau
    "basf_org":  "G4",     # BASF-Beauftragter, Org.-Code (ha van)
    "geraet":    "B5",     # Gép/Eszköz (ha van)
    "desc":      "A6",     # napi leírás (A6:G15 összevont tartomány tetejére írunk)
    "sum_hours": "G17",    # összesített óraszám (óra.formátumban, pl. 7.50)
    # Dolgozói sorok (max 5). TETSZŐLEGESEN ÁTÍRHATÓ!
    # Itt feltételezek egy táblát, ahol:
    #   név a B18..B22, igazolvány a C18..C22, kezdés D18..D22, vég E18..E22, óra F18..F22
    "workers": {
        1: {"name": "B18", "id": "C18", "start": "D18", "end": "E18", "hours": "F18"},
        2: {"name": "B19", "id": "C19", "start": "D19", "end": "E19", "hours": "F19"},
        3: {"name": "B20", "id": "C20", "start": "D20", "end": "E20", "hours": "F20"},
        4: {"name": "B21", "id": "C21", "start": "D21", "end": "E21", "hours": "F21"},
        5: {"name": "B22", "id": "C22", "start": "D22", "end": "E22", "hours": "F22"},
    }
}


# =========================
# ROUTE-OK
# =========================
@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate_excel")
async def generate_excel(request: Request):
    """
    Bemenet: a jelenlegi űrlap mezői (dátum, bau, basf_org, geraet, leírás, dolgozók...).
    Kimenet: kitöltött Excel letöltése.
    """
    form = dict((await request.form()).items())

    datum = form.get("datum", "").strip()
    bau = form.get("bau", "").strip()
    basf_org = form.get("basf_org", "").strip()
    geraet = form.get("geraet", "").strip()
    beschreibung = form.get("beschreibung", "").strip()

    # Dolgozók, összóra
    workers, total_mins = collect_workers_from_form(form)
    total_hours = round(total_mins / 60.0, 2)

    # Excel sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Fejléc adatok
    if datum:
        set_cell_a1(ws, POS["date"], datum)
    if bau:
        set_cell_a1(ws, POS["bau"], bau)
    if basf_org:
        set_cell_a1(ws, POS["basf_org"], basf_org)
    if geraet:
        set_cell_a1(ws, POS["geraet"], geraet)

    # Napi leírás – az A6:G15 tartomány tetejére írunk (merge safe)
    if beschreibung:
        set_cell_a1(ws, POS["desc"], beschreibung)

    # Dolgozói sorok (max 5)
    for i, w in enumerate(workers[:5], start=1):
        slot = POS["workers"][i]
        full_name = (w["vorname"] + " " + w["nachname"]).strip()
        if full_name:
            set_cell_a1(ws, slot["name"], full_name)
        if w["ausweis"]:
            set_cell_a1(ws, slot["id"], w["ausweis"])
        if w["beginn"]:
            set_cell_a1(ws, slot["start"], w["beginn"])
        if w["ende"]:
            set_cell_a1(ws, slot["end"], w["ende"])
        # egyéni óraszám (ha csak egy intervallum van / dolgozó)
        if w["mins"] > 0:
            set_cell_a1(ws, slot["hours"], f"{round(w['mins']/60.0, 2):.2f}")

    # Összesített óraszám
    set_cell_a1(ws, POS["sum_hours"], f"{total_hours:.2f}")

    # Mentés memóriába és visszaküldés
    bio = BytesIO()
    # Fájlnév: GP-t_filled_YYYYMMDD.xlsx (ha nincs dátum, timestamp)
    try:
        dn = datetime.strptime(datum, "%Y-%m-%d").strftime("%Y%m%d")
    except Exception:
        dn = datetime.now().strftime("%Y%m%d_%H%M%S")

    out_name = f"GP-t_filled_{dn}.xlsx"
    wb.save(bio)
    bio.seek(0)

    headers = {
        "Content-Disposition": f'attachment; filename="{out_name}"'
    }
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

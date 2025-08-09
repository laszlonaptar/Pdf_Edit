from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from datetime import datetime, time
from io import BytesIO
from typing import List, Dict, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import coordinate_to_tuple

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


# ---------- Segédek ----------
def parse_time(s: str) -> time:
    s = (s or "").strip().replace(".", ":")
    return datetime.strptime(s, "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    a0, a1 = a_start.hour * 60 + a_start.minute, a_end.hour * 60 + a_end.minute
    b0, b1 = b_start.hour * 60 + b_start.minute, b_end.hour * 60 + b_end.minute
    return max(0, min(a1, b1) - max(a0, b0))

def net_minutes(start: time, end: time) -> int:
    total = overlap_minutes(start, end, start, end)
    total -= overlap_minutes(start, end, time(9, 0), time(9, 15))
    total -= overlap_minutes(start, end, time(12, 0), time(12, 45))
    return max(0, total)

def find_text(ws, text_substr: str) -> Optional[Tuple[int, int]]:
    t = text_substr.lower()
    for row in ws.iter_rows(1, ws.max_row, 1, ws.max_column):
        for cell in row:
            v = "" if cell.value is None else str(cell.value).strip()
            if t in v.lower():
                return cell.row, cell.column
    return None

def find_header_row_and_cols(ws, header_names: Dict[str, List[str]]) -> Tuple[int, Dict[str, int]]:
    for r in range(1, ws.max_row + 1):
        row_vals = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        hits = {}
        for key, variants in header_names.items():
            for c, val in enumerate(row_vals, start=1):
                if any(v in val for v in (v.lower() for v in variants)):
                    hits[key] = c
                    break
        if len(hits) >= 3:
            return r, hits
    raise RuntimeError("Nem találom a dolgozói táblázat fejlécét a sablonban.")

def top_left_of_merge(ws, row: int, col: int) -> Tuple[int, int]:
    """Ha a cella egy merge tartományban van, visszaadja a bal-felső cellát; különben saját magát."""
    for rng in ws.merged_cells.ranges:
        if (row, col) in rng:
            return rng.min_row, rng.min_col
    return row, col

def set_cell(ws, row: int, col: int, value, alignment: Alignment | None = None):
    r, c = top_left_of_merge(ws, row, col)
    cell = ws.cell(r, c)
    cell.value = value
    if alignment:
        cell.alignment = alignment

def set_cell_by_addr(ws, addr: str, value, alignment: Alignment | None = None):
    r, c = coordinate_to_tuple(addr)
    set_cell(ws, r, c, value, alignment)

def write_wrapped(ws, addr: str, text: str, indent: int = 1):
    set_cell_by_addr(ws, addr, text, Alignment(wrap_text=True, horizontal="left", vertical="top", indent=indent))


# ---------- Frontend ----------
@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# ---------- Excel generálás ----------
@app.post("/generate_excel")
async def generate_excel(request: Request):
    form = await request.form()

    def g(key: str, default: str = "") -> str:
        return str(form.get(key, default)).strip()

    datum = g("datum") or g("date")
    bau = g("bau") or g("bauort") or g("projekt")
    bf = g("bf") or g("beauftragter") or g("basf")
    geraet = g("geraet") or g("vorhaltung")
    beschreibung = g("beschreibung") or g("taetigkeit") or g("was")

    vorn = form.getlist("vorname[]") or form.getlist("vorname") or []
    nachn = form.getlist("nachname[]") or form.getlist("nachname") or []
    ausw = form.getlist("ausweis[]") or form.getlist("ausweis") or []
    begn = form.getlist("beginn[]") or form.getlist("beginn") or []
    ende = form.getlist("ende[]") or form.getlist("ende") or []

    if not vorn and any(k.startswith("vorname") for k in form.keys()):
        for i in range(1, 6):
            if f"vorname{i}" in form:
                vorn.append(g(f"vorname{i}"))
                nachn.append(g(f"nachname{i}"))
                ausw.append(g(f"ausweis{i}"))
                begn.append(g(f"beginn{i}"))
                ende.append(g(f"ende{i}"))

    workers = []
    N = max(len(vorn), len(nachn), len(ausw), len(begn), len(ende))
    for i in range(N):
        workers.append({
            "vorname": vorn[i] if i < len(vorn) else "",
            "nachname": nachn[i] if i < len(nachn) else "",
            "ausweis": ausw[i] if i < len(ausw) else "",
            "beginn": begn[i] if i < len(begn) else "",
            "ende": ende[i] if i < len(ende) else "",
        })

    # --- sablon ---
    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # --- fejrészek (biztonságos írás merge mellett is) ---
    pos_date = find_text(ws, "Datum der Leistungs")
    if pos_date:
        set_cell(ws, pos_date[0], pos_date[1] + 1, datum)
    pos_bau = find_text(ws, "Bau und Ausführungsort")
    if pos_bau:
        set_cell(ws, pos_bau[0], pos_bau[1] + 1, bau)
    pos_bf = find_text(ws, "BASF-Beauftragter")
    if pos_bf:
        set_cell(ws, pos_bf[0], pos_bf[1] + 1, bf)
    pos_vorh = find_text(ws, "Vorhaltung")
    if pos_vorh and geraet:
        set_cell(ws, pos_vorh[0] + 1, pos_vorh[1], geraet)

    # --- Leírás A6:G15 bal-fentről, tördelve ---
    write_wrapped(ws, "A6", beschreibung or "", indent=1)

    # --- dolgozói táblázat keresése ---
    header_map = {
        "name": ["name"],
        "vorname": ["vorname"],
        "ausweis": ["ausweis", "kennzeichen"],
        "beginn": ["beginn"],
        "ende": ["ende"],
        "stunden": ["anzahl stunden", "stunden"],
    }
    header_row, cols = find_header_row_and_cols(ws, header_map)
    data_start_row = header_row + 2

    total_minutes = 0
    for idx, w in enumerate(workers[:5]):
        r = data_start_row + idx
        if "name" in cols:
            set_cell(ws, r, cols["name"], w["nachname"], Alignment(horizontal="left", vertical="center"))
        if "vorname" in cols:
            set_cell(ws, r, cols["vorname"], w["vorname"], Alignment(horizontal="left", vertical="center"))
        if "ausweis" in cols:
            set_cell(ws, r, cols["ausweis"], w["ausweis"], Alignment(horizontal="left", vertical="center"))

        try:
            t_start = parse_time(w["beginn"])
            t_end = parse_time(w["ende"])
            if "beginn" in cols:
                set_cell(ws, r, cols["beginn"], t_start.strftime("%H:%M"), Alignment(horizontal="center", vertical="center"))
            if "ende" in cols:
                set_cell(ws, r, cols["ende"], t_end.strftime("%H:%M"), Alignment(horizontal="center", vertical="center"))

            mins = net_minutes(t_start, t_end)
            total_minutes += mins
            hours = round(mins / 60.0, 2)
            if "stunden" in cols:
                set_cell(ws, r, cols["stunden"], hours, Alignment(horizontal="center", vertical="center"))
        except Exception:
            pass

    pos_total = find_text(ws, "Gesamtstunden")
    if pos_total:
        set_cell(ws, pos_total[0], pos_total[1] + 1, round(total_minutes / 60.0, 2),
                 Alignment(horizontal="center", vertical="center"))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"Arbeitsnachweis_{datum or datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

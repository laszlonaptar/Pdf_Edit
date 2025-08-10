# main.py
from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from typing import Optional, Tuple, Dict, List
from datetime import datetime, time, timedelta
import io
import uuid
import re
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_XLSX = os.getenv("TEMPLATE_XLSX", "GP-t.xlsx")


# ---------- helpers for merged-cell safe addressing ----------

def in_range(r, c, cr) -> bool:
    return cr.min_row <= r <= cr.max_row and cr.min_col <= c <= cr.max_col

def merged_block(ws, r, c):
    for cr in ws.merged_cells.ranges:
        if in_range(r, c, cr):
            return cr
    return None

def right_adjacent_block(ws, r, c):
    """Return the merged range immediately to the right of the block containing (r,c)."""
    cur = merged_block(ws, r, c)
    end_col = cur.max_col if cur else c
    target_min_col = end_col + 1
    # exact neighbor merged range?
    for cr in ws.merged_cells.ranges:
        if cr.min_row <= r <= cr.max_row and cr.min_col == target_min_col:
            return cr
    # if the neighbor is a single (nem összevont) cell-sáv
    return None

def top_left_of(cr) -> Tuple[int, int]:
    return cr.min_row, cr.min_col

def set_text(ws, r, c, text, wrap=True, align_left=True):
    cr = merged_block(ws, r, c)
    rr = cr.min_row if cr else r
    cc = cr.min_col if cr else c
    cell = ws.cell(rr, cc)
    cell.value = text
    cell.alignment = Alignment(wrap_text=wrap, horizontal=("left" if align_left else None), vertical="top")


# ---------- label-based placement ----------

def find_label(ws, needle: str) -> Optional[Tuple[int, int]]:
    norm = lambda s: re.sub(r"\s+", " ", str(s or "")).strip()
    want = norm(needle)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if norm(ws.cell(r, c).value) == want:
                return r, c
    return None

def put_right_of_label(ws, label: str, value: str):
    pos = find_label(ws, label)
    if not pos:
        return
    r, c = pos
    # jobbra lévő blokk top-left
    right = right_adjacent_block(ws, r, c)
    if right:
        rr, cc = top_left_of(right)
    else:
        cur = merged_block(ws, r, c)
        cc = (cur.max_col + 1) if cur else (c + 1)
        rr = r
    set_text(ws, rr, cc, value, wrap=False, align_left=True)


# ---------- table header detection & writing ----------

def normalize_header(s: str) -> str:
    s = re.sub(r"\s+", " ", s or "").strip().lower()
    s = s.replace("ausweis- nr.", "ausweis-nr.")
    s = s.replace("ausweis - nr.", "ausweis-nr.")
    return s

def find_table_headers(ws) -> Tuple[int, Dict[str, int]]:
    wanted = {
        "name": ["name"],
        "vorname": ["vorname"],
        "ausweis": ["ausweis", "kennzeichen"],
        "beginn": ["beginn"],
        "ende": ["ende"],
        "stunden": ["anzahl stunden", "ohne pausen"],
    }
    for r in range(1, ws.max_row + 1):
        cols: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            val = normalize_header(ws.cell(r, c).value)
            if not val:
                continue
            for key, needles in wanted.items():
                if key in cols:
                    continue
                if all(any(n in val for n in needles) for needles in [needles := wanted[key]]):
                    cols[key] = c
        # egyszerűbb: külön-külön keresés
        # javítva:
        cols = {}
        for c in range(1, ws.max_column + 1):
            val = normalize_header(ws.cell(r, c).value)
            if "name" == val:
                cols["name"] = c
            if "vorname" in val:
                cols["vorname"] = c
            if "ausweis" in val or "kennzeichen" in val:
                cols["ausweis"] = c
            if "beginn" in val:
                cols["beginn"] = c
            if "ende" in val:
                cols["ende"] = c
            if "anzahl stunden" in val:
                cols["stunden"] = c
        if {"name","vorname","ausweis","beginn","ende","stunden"}.issubset(cols.keys()):
            return r, cols
    raise RuntimeError("Fejléc sor nem található a táblában.")

def first_data_row(ws, header_row: int) -> int:
    # A fejléc alatt jellemzően még egy sor a „[hh:mm] / Begin / Ende” felosztás.
    return header_row + 2


# ---------- time & breaks ----------

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    to_min = lambda t: t.hour * 60 + t.minute
    a0, a1, b0, b1 = to_min(a_start), to_min(a_end), to_min(b_start), to_min(b_end)
    inter = max(0, min(a1, b1) - max(a0, b0))
    return inter

def net_hours(start: str, end: str) -> float:
    s, e = parse_hhmm(start), parse_hhmm(end)
    total = (datetime.combine(datetime.today(), e) - datetime.combine(datetime.today(), s)).total_seconds() / 3600.0
    # szünetek
    br = 0.0
    br += overlap_minutes(s, e, time(9, 0), time(9, 15)) / 60.0  # 0.25h max
    br += overlap_minutes(s, e, time(12, 0), time(12, 45)) / 60.0  # 0.75h max
    return max(0.0, round(total - br, 2))


# ---------- routes ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    beschreibung: str = Form(""),
    basf_beauftragter: str = Form(""),
    geraet: str = Form(""),
    # 1..5 dolgozó
    vorname1: str = Form(""), nachname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""),
    vorname2: str = Form(""), nachname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""),
    vorname3: str = Form(""), nachname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""),
    vorname4: str = Form(""), nachname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""),
    vorname5: str = Form(""), nachname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""),
):
    # --- load template
    wb = load_workbook(TEMPLATE_XLSX)
    ws = wb.active

    # --- header fields by labels (left sections)
    put_right_of_label(ws, "Datum der Leistungsausführung:", datum)
    put_right_of_label(ws, "Bau und Ausführungsort:", bau)

    # BASF-Beauftragter (label variációk)
    for lab in [
        "BASF-Beauftragter, Org.-Code:",
        "BASF-Beauftragter, Org.-Code:",
        "BASF-Beauftragter, Org.-Code",
    ]:
        if find_label(ws, lab):
            put_right_of_label(ws, lab, basf_beauftragter)
            break

    # Gerät / Fahrzeug (ha van ilyen címke; ha nincs, kihagyjuk)
    for lab in [
        "Vorhaltung / beauftragtes Gerät / Fahrzeug",
        "Vorhaltung / beauftragtes Gerät / Fahrzeug:",
    ]:
        if find_label(ws, lab):
            put_right_of_label(ws, lab, geraet)
            break

    # --- long description: keresünk egy nagy, széles összevont blokkot a fejléc alatt
    bau_lbl = find_label(ws, "Bau und Ausführungsort:")
    if bau_lbl:
        start_row = bau_lbl[0] + 1
    else:
        start_row = 6
    big = None
    for cr in ws.merged_cells.ranges:
        if cr.min_row >= start_row and (cr.max_row - cr.min_row) >= 2 and (cr.max_col - cr.min_col) >= 6:
            big = cr
            break
    if big:
        rr, cc = big.min_row, big.min_col
        set_text(ws, rr, cc, beschreibung, wrap=True, align_left=True)

    # --- table headers & rows
    header_row, cols = find_table_headers(ws)
    row = first_data_row(ws, header_row)

    employees = []
    raw = [
        (vorname1, nachname1, ausweis1, beginn1, ende1),
        (vorname2, nachname2, ausweis2, beginn2, ende2),
        (vorname3, nachname3, ausweis3, beginn3, ende3),
        (vorname4, nachname4, ausweis4, beginn4, ende4),
        (vorname5, nachname5, ausweis5, beginn5, ende5),
    ]
    for v, n, a, b, e in raw:
        if (v or n) and a and b and e:
            employees.append((v.strip(), n.strip(), a.strip(), b.strip(), e.strip()))

    total_hours = 0.0
    for v, n, a, b, e in employees:
        # külön cellákba: vezetéknév = Name, keresztnév = Vorname
        set_text(ws, row, cols["name"], n, wrap=False, align_left=True)
        set_text(ws, row, cols["vorname"], v, wrap=False, align_left=True)
        set_text(ws, row, cols["ausweis"], a, wrap=False, align_left=True)
        set_text(ws, row, cols["beginn"], b, wrap=False, align_left=True)
        set_text(ws, row, cols["ende"], e, wrap=False, align_left=True)
        nh = net_hours(b, e)
        total_hours += nh
        set_text(ws, row, cols["stunden"], f"{nh:.2f}", wrap=False, align_left=True)
        row += 1

    # Gesamtstunden – megpróbáljuk megtalálni a „Gesamtstunden” felirat mellé eső értékcellát
    for r in range(header_row, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(r, c).value).strip().lower() == "gesamtstunden":
                put_right_of_label(ws, "Gesamtstunden", f"{total_hours:.2f}")
                # ha a klasszikus címke-form nem működik, akkor a következő cellába írunk
                try:
                    ws.cell(r, c + 1).value = f"{total_hours:.2f}"
                except Exception:
                    pass
                r = ws.max_row + 1
                break

    # --- output
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"leistungsnachweis_{uuid.uuid4().hex}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

# main.py
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from datetime import datetime, time, timedelta
from io import BytesIO
import re
import uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---------------------------
# Excel helper függvények
# ---------------------------

def merged_bounds(ws, r: int, c: int):
    """Ha (r,c) összevont blokkban van, adja vissza (min_row, min_col, max_row, max_col),
    különben (r,c,r,c).  (A1 címkével ellenőrzünk – ez javítja a TypeError hibát.)"""
    coord = f"{get_column_letter(c)}{r}"
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return rng.min_row, rng.min_col, rng.max_row, rng.max_col
    return r, c, r, c

def top_left_of(ws, r: int, c: int):
    """Az adott cella összevont blokkjának bal-felső sarka."""
    r0, c0, _, _ = merged_bounds(ws, r, c)
    return r0, c0

def right_block_top_left(ws, r: int, c: int):
    """Az (r,c) cella összevont blokkjától jobbra eső blokk bal-felső sarka."""
    _, _, _, max_c = merged_bounds(ws, r, c)
    target_r, target_c = r, max_c + 1
    return top_left_of(ws, target_r, target_c)

def set_text(ws, r: int, c: int, text: str, wrap=False, align_left=True, align_top=True):
    """Szöveg beírása az adott cella/blokk bal-felső cellájába."""
    r0, c0 = top_left_of(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = text
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal=("left" if align_left else "center"),
        vertical=("top" if align_top else "center"),
    )

def find_label(ws, needle: str):
    """Visszaadja az első olyan cella (r,c) pozícióját, amelynek szövege
    tartalmazza a needle-t (kis/nagybetű érzéketlen)."""
    rx = re.compile(re.escape(needle), re.IGNORECASE)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and rx.search(v):
                return r, c
    return None

def set_value_right_of_label(ws, label_text: str, value: str):
    """Címkét megkeres, majd az attól jobbra lévő érték-cellába ír (összevont blokkokkal számol)."""
    pos = find_label(ws, label_text)
    if not pos:
        return False
    r, c = pos
    rr, cc = right_block_top_left(ws, r, c)
    set_text(ws, rr, cc, value, wrap=False, align_left=True, align_top=True)
    return True

def find_lower_table_mapping(ws):
    """
    Megkeresi az alsó táblázat oszlopait:
    - Name
    - Vorname
    - Ausweis (bármely 'Ausweis' előfordulás)
    - Beginn
    - Ende
    - Anzahl Stunden (olyan cella, ami ezt tartalmazza)
    Visszatér: (header_row, dict)
    """
    name_col = vorname_col = ausweis_col = beg_col = end_col = stunden_col = None
    header_row = None

    # előbb keressük a 'Beginn' és 'Ende' feliratokat (egyértelműek)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            txt = v.strip().lower()
            if txt == "beginn":
                beg_col = c
                header_row = r
            if txt == "ende":
                end_col = c
                header_row = r

    # a fejléc sor legyen a megtalált 'Beginn'/'Ende' sor
    if header_row is None:
        # fallback: keresünk 'Name' sort és azt tekintjük fejlécnek
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip().lower() == "name":
                    header_row = r
                    break
            if header_row:
                break

    # Oszlopok: Name / Vorname / Ausweis / Anzahl Stunden
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if not isinstance(v, str):
            continue
        t = v.lower()
        if t.strip() == "name":
            name_col = c
        elif t.strip() == "vorname":
            vorname_col = c
        elif "ausweis" in t:
            ausweis_col = c
        elif "anzahl stunden" in t:
            stunden_col = c

    mapping = {
        "name": name_col,
        "vorname": vorname_col,
        "ausweis": ausweis_col,
        "beginn": beg_col,
        "ende": end_col,
        "stunden": stunden_col,
    }
    return header_row, mapping

# ---------------------------
# Idő és óraszám számítás
# ---------------------------

def parse_hhmm(s: str) -> time:
    return datetime.strptime(s.strip(), "%H:%M").time()

def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    """Két idősáv átfedése percben."""
    base = datetime(2000, 1, 1)
    s1 = datetime.combine(base, a_start)
    e1 = datetime.combine(base, a_end)
    s2 = datetime.combine(base, b_start)
    e2 = datetime.combine(base, b_end)
    s = max(s1, s2)
    e = min(e1, e2)
    return max(0, int((e - s).total_seconds() // 60))

def compute_hours_with_breaks(beg: str, end: str) -> float:
    """Összóraszám számítása fix szünetekkel: 09:00–09:15 és 12:00–12:45."""
    t1 = parse_hhmm(beg)
    t2 = parse_hhmm(end)
    total = (datetime.combine(datetime.today(), t2) - datetime.combine(datetime.today(), t1)).total_seconds() / 3600.0
    # szünetek
    m = 0
    m += overlap_minutes(t1, t2, time(9, 0), time(9, 15))
    m += overlap_minutes(t1, t2, time(12, 0), time(12, 45))
    return max(0.0, round(total - m / 60.0, 2))

# ---------------------------
# Web
# ---------------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(request: Request):
    form = await request.form()
    # Alap mezők – több névváltozatot is elfogadunk
    datum = form.get("datum") or form.get("date") or ""
    bau = form.get("bau") or form.get("projekt") or ""
    beauftragter = form.get("bf") or form.get("beauftragter") or ""
    beschreibung = form.get("beschreibung") or form.get("taetigkeit") or form.get("taetigkeiten") or ""

    # Dolgozók dinamikusan (1..10)
    workers = []
    for i in range(1, 11):
        ln = form.get(f"nachname{i}") or form.get(f"name{i}") or form.get(f"mitarbeiter_nachname_{i}")
        fn = form.get(f"vorname{i}") or form.get(f"mitarbeiter_vorname_{i}")
        aid = form.get(f"ausweis{i}") or form.get(f"ausweisnummer_{i}") or form.get(f"ausweis_nr{i}")
        beg = form.get(f"beginn{i}") or form.get(f"start{i}")
        end = form.get(f"ende{i}") or form.get(f"stop{i}")
        if any([ln, fn, aid, beg, end]):
            # üres nevet ne írjunk
            ln = ln or ""
            fn = fn or ""
            aid = aid or ""
            beg = beg or ""
            end = end or ""
            hours = compute_hours_with_breaks(beg, end) if beg and end else 0.0
            workers.append({"ln": ln, "fn": fn, "id": aid, "beg": beg, "end": end, "hours": hours})

    # Excel sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # 1) Dátum, Bau (bal oldali értékmezők), és BASF-Beauftragter (E3 blokk jobb oldalára)
    set_value_right_of_label(ws, "Datum der Leistungsausführung", datum)
    set_value_right_of_label(ws, "Bau und Ausführungsort", bau)
    set_value_right_of_label(ws, "BASF-Beauftragter", beauftragter)

    # 2) Leírás – nagy, több soros terület. Ha létezik A6:G15, oda írunk, különben
    # a Bau sor alatti ~10 sorba összevonunk és oda.
    area_r1, area_c1, area_r2, area_c2 = 6, 1, 15, 7
    try:
        # próbáljuk meg ezt az ismert blokkot
        ws.merge_cells(start_row=area_r1, start_column=area_c1, end_row=area_r2, end_column=area_c2)
    except Exception:
        pass
    set_text(ws, area_r1, area_c1, beschreibung, wrap=True, align_left=True, align_top=True)

    # 3) Alsó tábla – oszlopok felderítése és feltöltése
    header_row, cols = find_lower_table_mapping(ws)
    if header_row and all([cols["name"], cols["vorname"], cols["ausweis"], cols["beginn"], cols["ende"], cols["stunden"]]):
        data_row = header_row + 1
        r = data_row
        total_hours = 0.0
        for w in workers:
            set_text(ws, r, cols["name"], w["ln"])
            set_text(ws, r, cols["vorname"], w["fn"])
            set_text(ws, r, cols["ausweis"], w["id"])
            set_text(ws, r, cols["beginn"], w["beg"])
            set_text(ws, r, cols["ende"], w["end"])
            hrs_str = f"{w['hours']:.2f}".replace(".", ",") if w["hours"] % 1 else f"{int(w['hours'])},00"
            set_text(ws, r, cols["stunden"], hrs_str, align_left=True)
            total_hours += w["hours"]
            r += 1

        # 4) Gesamtstunden – a táblázat alján “Gesamtstunden” felirat mellett (jobbra) írjuk
        pos_total = find_label(ws, "Gesamtstunden")
        if pos_total:
            tr, tc = right_block_top_left(ws, *pos_total)
            tot_str = f"{total_hours:.2f}".replace(".", ",") if total_hours % 1 else f"{int(total_hours)},00"
            set_text(ws, tr, tc, tot_str, align_left=True)

    # 5) Mentés memóriába és visszaküldés
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"GP_t_filled_{uuid.uuid4().hex[:8]}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"'
    }
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

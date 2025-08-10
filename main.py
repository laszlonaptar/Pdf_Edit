from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime, time, timedelta
import io, uuid, os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_PATH = "GP-t.xlsx"

# -------------------- helper: merged / write-safe --------------------

def top_left_of_merge(ws, r, c):
    """Ha (r,c) egy merge-en belül van, visszaadja a merge bal-felső celláját (r0,c0).
       Ha nincs merge-ben, az eredeti (r,c)-t adja vissza."""
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            return rng.min_row, rng.min_col
    return r, c

def next_merged_block_right(ws, r, c):
    """Megkeresi ugyanazon a soron a (r,c)-t tartalmazó merge blokkot (vagy egyedülálló cellát),
       és visszaadja a KÖVETKEZŐ összevont blokk bal-felső celláját jobbra. Ha nincs, None."""
    # 1) melyik blokkban ül a (r,c)?
    this_min_c = c
    for rng in ws.merged_cells.ranges:
        if (r, c) in rng:
            this_min_c = rng.min_col
            this_max_c = rng.max_col
            break
    else:
        # nincs merge-ben -> a "blokk" önmaga, max_c = c
        this_max_c = c

    # 2) soron lévő blokkok felderítése (min_col szerinti sorrend)
    blocks = []
    occupied = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row:
            blocks.append((rng.min_col, rng.min_row, rng.max_col, rng.max_row))
            for cc in range(rng.min_col, rng.max_col + 1):
                occupied.add(cc)
    # olyan oszlopokat is tekintsünk külön "blokknak", ahol NINCS merge
    max_col = ws.max_column
    col = 1
    while col <= max_col:
        if col not in occupied:
            # önálló cella-blokk
            blocks.append((col, r, col, r))
        col += 1

    # 3) a blokkokat rendezzük a min_col alapján és válasszuk ki a "következőt"
    blocks.sort(key=lambda x: x[0])
    # jelenlegi blokk min_col-ját nézzük (this_min_c), és keressük az utána következőt
    for min_c, min_r, max_c, max_r in blocks:
        if min_c > this_max_c:
            return (min_r, min_c, max_r, max_c)
    return None

def write_in_block(ws, r, c, value, wrap=False, align_left=True):
    """A (r,c) cella BLOKKJÁNAK bal-felső cellájába ír (ha merge-ben van),
       különben közvetlenül a cellába. Beállítja az igazítást/tördelést."""
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = value
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal="left" if align_left else "center",
        vertical="top" if wrap else "center"
    )

# -------------------- domain segédek --------------------

def hhmm_to_dt(hhmm: str) -> time:
    hhmm = hhmm.strip()
    if not hhmm:
        return None
    parts = hhmm.replace(".", ":").split(":")
    h = int(parts[0])
    m = int(parts[1]) if len(parts) > 1 else 0
    return time(hour=h, minute=m)

def hours_between(beg: time, end: time) -> float:
    dt0 = datetime.combine(datetime.today(), beg)
    dt1 = datetime.combine(datetime.today(), end)
    if dt1 < dt0:
        dt1 += timedelta(days=1)
    delta = dt1 - dt0
    return round(delta.total_seconds() / 3600.0, 2)

def subtract_breaks(total: float, beg: time, end: time) -> float:
    """Fix szünetek: 09:00–09:15 (0.25h) és 12:00–12:45 (0.75h) ha belelóg a munkaidőbe."""
    def overlap(b1: time, e1: time, b2: time, e2: time) -> float:
        dt = datetime.today()
        s1 = datetime.combine(dt, b1); e1d = datetime.combine(dt, e1)
        s2 = datetime.combine(dt, b2); e2d = datetime.combine(dt, e2)
        if e1d < s1: e1d += timedelta(days=1)
        if e2d < s2: e2d += timedelta(days=1)
        start = max(s1, s2); end = min(e1d, e2d)
        secs = (end - start).total_seconds()
        return max(0.0, secs/3600.0)

    b1 = overlap(beg, end, time(9,0), time(9,15))
    b2 = overlap(beg, end, time(12,0), time(12,45))
    return round(max(0.0, total - b1 - b2), 2)

# -------------------- UI --------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# -------------------- fő logika --------------------

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(""),
    basf: str = Form(""),
    auftrag: str = Form(""),
    beschreibung: str = Form(""),

    name1: str = Form(""),
    vorname1: str = Form(""),
    ausweis1: str = Form(""),
    beginn1: str = Form(""),
    ende1: str = Form(""),

    name2: str = Form(""),
    vorname2: str = Form(""),
    ausweis2: str = Form(""),
    beginn2: str = Form(""),
    ende2: str = Form(""),

    name3: str = Form(""),
    vorname3: str = Form(""),
    ausweis3: str = Form(""),
    beginn3: str = Form(""),
    ende3: str = Form(""),

    name4: str = Form(""),
    vorname4: str = Form(""),
    ausweis4: str = Form(""),
    beginn4: str = Form(""),
    ende4: str = Form(""),

    name5: str = Form(""),
    vorname5: str = Form(""),
    ausweis5: str = Form(""),
    beginn5: str = Form(""),
    ende5: str = Form(""),
):
    # minimális validáció
    if not datum:
        return JSONResponse({"detail":"Hiányzik a dátum."}, status_code=400)

    try:
        wb = load_workbook(TEMPLATE_PATH)
    except Exception as e:
        return JSONResponse({"detail": f"Nem tudom megnyitni a sablont: {e}"}, status_code=500)

    ws = wb.active

    # ---------- 1) Fejrészek: label -> next block right ----------
    def put_by_label(label_text: str, value: str):
        # megkeressük a felirat celláját
        for row in ws.iter_rows(min_row=1, max_row=20, values_only=False):
            for cell in row:
                if str(cell.value).strip() == label_text:
                    r, c = cell.row, cell.column
                    nxt = next_merged_block_right(ws, r, c)
                    if not nxt:
                        # ha nincs következő blokk, írjunk a felirat jobb SZOMSZÉDJÁBA
                        write_in_block(ws, r, c+1, value, wrap=False, align_left=True)
                        return True
                    r0, c0, r1, c1 = nxt
                    write_in_block(ws, r0, c0, value, wrap=False, align_left=True)
                    return True
        return False

    put_by_label("Datum der Leistungsausführung:", datum)
    put_by_label("Bau und Ausführungsort:", bau)
    put_by_label("BASF-Beauftragter, Org.-Code:", basf)
    put_by_label("Einzelauftrags-Nr. (Avisor) oder Best.-Nr. (sonstige):", auftrag)

    # ---------- 2) Beschreibung: a bal oldali nagy, vonalas blokk ----------
    # A te sablonodban ez az A6–G15 közös blokk (ezt korábban egyeztettük).
    # Írjunk az A6 bal-felső cellába; tördeléssel és nagy sor­magassággal.
    write_in_block(ws, 6, 1, beschreibung, wrap=True, align_left=True)
    # emeljük meg a sorok magasságát, hogy ne vágja le az alsó pixelsort
    for r in range(6, 16):
        ws.row_dimensions[r].height = 28  # láthatóbb sorok

    # ---------- 3) Dolgozói sorok ----------
    # A táblázat első adat sora NÁLAD a fejléc alatt kezdődik.
    # A képek alapján ez kb. a 21. sor körül van; állítsuk be:
    START_ROW = 21  # ha máshol van, szólj, átírjuk 1 számmal
    cols = {
        "name": 2,        # Vezetéknév / “Name”
        "vorname": 4,     # Keresztnév / “Vorname”
        "ausweis": 6,     # “Ausweis-Nr.”
        "beginn": 8,      # Kezdés “Beginn”
        "ende": 9,        # Vége “Ende”
        "stunden": 11,    # “Anzahl Stunden (ohne Pausen)”
    }

    def put_worker(idx: int, nachname: str, vorname: str, ausweis: str, b: str, e: str):
        if not (nachname or vorname or ausweis or b or e):
            return 0.0
        r = START_ROW + (idx - 1)
        # mindig a blokk bal-felső cellájába írunk
        write_in_block(ws, r, cols["name"], nachname, wrap=False, align_left=True)
        write_in_block(ws, r, cols["vorname"], vorname, wrap=False, align_left=True)
        write_in_block(ws, r, cols["ausweis"], ausweis, wrap=False, align_left=True)

        tb = hhmm_to_dt(b) if b else None
        te = hhmm_to_dt(e) if e else None
        if tb:
            write_in_block(ws, r, cols["beginn"], b, wrap=False, align_left=True)
        if te:
            write_in_block(ws, r, cols["ende"], e, wrap=False, align_left=True)

        hours = 0.0
        if tb and te:
            gross = hours_between(tb, te)
            net = subtract_breaks(gross, tb, te)
            hours = net
            write_in_block(ws, r, cols["stunden"], f"{net:.2f}", wrap=False, align_left=True)
        return hours

    total = 0.0
    total += put_worker(1, name1, vorname1, ausweis1, beginn1, ende1)
    total += put_worker(2, name2, vorname2, ausweis2, beginn2, ende2)
    total += put_worker(3, name3, vorname3, ausweis3, beginn3, ende3)
    total += put_worker(4, name4, vorname4, ausweis4, beginn4, ende4)
    total += put_worker(5, name5, vorname5, ausweis5, beginn5, ende5)

    # összóra a táblázat alján – a képed alapján az alsó “Gesamtstunden” cellába
    # Tegyük fel, hogy ez a táblázat alatti bal oldali blokk (pl. B??? oszlop),
    # írd be ugyanoda, ahova eddig is került: keressük meg a “Gesamtstunden” feliratot.
    def put_total_by_label(label_text: str, value: str):
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if str(cell.value).strip() == label_text:
                    r, c = cell.row, cell.column
                    nxt = next_merged_block_right(ws, r, c)
                    if nxt:
                        r0, c0, r1, c1 = nxt
                        write_in_block(ws, r0, c0, value, wrap=False, align_left=True)
                        return True
                    else:
                        write_in_block(ws, r, c+1, value, wrap=False, align_left=True)
                        return True
        return False

    put_total_by_label("Gesamtstunden", f"{total:.2f}")

    # mentés memóriába és visszaadás
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    filename = f"leistungsnachweis_{uuid.uuid4().hex}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"'
    }
    return FileResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

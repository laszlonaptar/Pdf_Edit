from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime, time, timedelta
import io, uuid

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_PATH = "GP-t.xlsx"

# -------------------- helper: coords / merged handling --------------------

def coord(r: int, c: int) -> str:
    return f"{get_column_letter(c)}{r}"

def top_left_of_merge(ws, r, c):
    here = coord(r, c)
    for rng in ws.merged_cells.ranges:
        if here in rng:
            return rng.min_row, rng.min_col
    return r, c

def next_merged_block_right(ws, r, c):
    this_min_c = c
    this_max_c = c
    here = coord(r, c)
    for rng in ws.merged_cells.ranges:
        if here in rng:
            this_min_c = rng.min_col
            this_max_c = rng.max_col
            break

    blocks = []
    occupied_cols = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row:
            blocks.append((rng.min_col, rng.min_row, rng.max_col, rng.max_row))
            for cc in range(rng.min_col, rng.max_col + 1):
                occupied_cols.add(cc)

    max_col = ws.max_column
    for col in range(1, max_col + 1):
        if col not in occupied_cols:
            blocks.append((col, r, col, r))

    blocks.sort(key=lambda x: x[0])

    for min_c, min_r, max_c, max_r in blocks:
        if min_c > this_max_c:
            return (min_r, min_c, max_r, max_c)
    return None

def write_in_block(ws, r, c, value, wrap=False, align_left=True):
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = value
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal="left" if align_left else "center",
        vertical="top" if wrap else "center",
    )

# -------------------- time utils --------------------

def hhmm_to_dt(hhmm: str) -> time | None:
    hhmm = (hhmm or "").strip()
    if not hhmm:
        return None
    p = hhmm.replace(".", ":").split(":")
    h = int(p[0])
    m = int(p[1]) if len(p) > 1 else 0
    return time(h, m)

def hours_between(beg: time, end: time) -> float:
    d0 = datetime.combine(datetime.today(), beg)
    d1 = datetime.combine(datetime.today(), end)
    if d1 < d0:
        d1 += timedelta(days=1)
    return round((d1 - d0).total_seconds() / 3600.0, 2)

def subtract_breaks(total: float, beg: time, end: time) -> float:
    def ov(b1, e1, b2, e2):
        t = datetime.today()
        s1, e1d = datetime.combine(t, b1), datetime.combine(t, e1)
        s2, e2d = datetime.combine(t, b2), datetime.combine(t, e2)
        if e1d < s1: e1d += timedelta(days=1)
        if e2d < s2: e2d += timedelta(days=1)
        sec = (min(e1d, e2d) - max(s1, s2)).total_seconds()
        return max(0.0, sec/3600.0)
    b1 = ov(beg, end, time(9,0), time(9,15))
    b2 = ov(beg, end, time(12,0), time(12,45))
    return round(max(0.0, total - b1 - b2), 2)

# -------------------- routes --------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(""),
    basf: str = Form(""),
    auftrag: str = Form(""),
    beschreibung: str = Form(""),

    name1: str = Form(""), vorname1: str = Form(""), ausweis1: str = Form(""), beginn1: str = Form(""), ende1: str = Form(""),
    name2: str = Form(""), vorname2: str = Form(""), ausweis2: str = Form(""), beginn2: str = Form(""), ende2: str = Form(""),
    name3: str = Form(""), vorname3: str = Form(""), ausweis3: str = Form(""), beginn3: str = Form(""), ende3: str = Form(""),
    name4: str = Form(""), vorname4: str = Form(""), ausweis4: str = Form(""), beginn4: str = Form(""), ende4: str = Form(""),
    name5: str = Form(""), vorname5: str = Form(""), ausweis5: str = Form(""), beginn5: str = Form(""), ende5: str = Form(""),
):
    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
    except Exception as e:
        return JSONResponse({"detail": f"Sablon hiba: {e}"}, status_code=500)

    # -------- fejrészek --------
    def put_by_label(label_text: str, value: str):
        for row in ws.iter_rows(min_row=1, max_row=30, values_only=False):
            for cell in row:
                if str(cell.value).strip() == label_text:
                    r, c = cell.row, cell.column
                    nxt = next_merged_block_right(ws, r, c)
                    if nxt:
                        r0, c0, r1, c1 = nxt
                        write_in_block(ws, r0, c0, value, wrap=False, align_left=True)
                    else:
                        write_in_block(ws, r, c+1, value, wrap=False, align_left=True)
                    return True
        return False

    put_by_label("Datum der Leistungsausführung:", datum)
    put_by_label("Bau und Ausführungsort:", bau)
    put_by_label("BASF-Beauftragter, Org.-Code:", basf)
    put_by_label("Einzelauftrags-Nr. (Avisor) oder Best.-Nr. (sonstige):", auftrag)

    # -------- leírás --------
    write_in_block(ws, 6, 1, beschreibung, wrap=True, align_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 28

    # -------- dolgozók --------
    START_ROW = 21
    cols = {"name": 2, "vorname": 4, "ausweis": 6, "beginn": 8, "ende": 9, "stunden": 11}

    def put_worker(i, nachname, vorname, ausweis, b, e):
        if not (nachname or vorname or ausweis or b or e):
            return 0.0
        r = START_ROW + (i - 1)
        write_in_block(ws, r, cols["name"], nachname, False, True)
        write_in_block(ws, r, cols["vorname"], vorname, False, True)
        write_in_block(ws, r, cols["ausweis"], ausweis, False, True)

        tb = hhmm_to_dt(b) if b else None
        te = hhmm_to_dt(e) if e else None
        if tb: write_in_block(ws, r, cols["beginn"], b, False, True)
        if te: write_in_block(ws, r, cols["ende"], e, False, True)

        net = 0.0
        if tb and te:
            gross = hours_between(tb, te)
            net = subtract_breaks(gross, tb, te)
            write_in_block(ws, r, cols["stunden"], f"{net:.2f}", False, True)
        return net

    total = 0.0
    total += put_worker(1, name1, vorname1, ausweis1, beginn1, ende1)
    total += put_worker(2, name2, vorname2, ausweis2, beginn2, ende2)
    total += put_worker(3, name3, vorname3, ausweis3, beginn3, ende3)
    total += put_worker(4, name4, vorname4, ausweis4, beginn4, ende4)
    total += put_worker(5, name5, vorname5, ausweis5, beginn5, ende5)

    # összóra
    def put_total(label_text: str, value: str):
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if str(cell.value).strip() == label_text:
                    r, c = cell.row, cell.column
                    nxt = next_merged_block_right(ws, r, c)
                    if nxt:
                        r0, c0, r1, c1 = nxt
                        write_in_block(ws, r0, c0, value, False, True)
                    else:
                        write_in_block(ws, r, c+1, value, False, True)
                    return True
        return False

    put_total("Gesamtstunden", f"{total:.2f}")

    # -------- visszaküldés: NEM FileResponse, hanem Response a BytesIO tartalmával --------
    out = io.BytesIO()
    wb.save(out)
    content = out.getvalue()
    filename = f"leistungsnachweis_{uuid.uuid4().hex}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

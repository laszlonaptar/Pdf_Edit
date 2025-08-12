from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, time, timedelta
import os
import io

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

TEMPLATE_XLSX = "GP-t.xlsx"

# ---------- Excel helpers ----------
def find_text_cell(ws, text: str):
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            v = "" if cell.value is None else str(cell.value).strip()
            if v == text.strip():
                return cell.row, cell.column
    return None

def top_left_of_merge(ws, r, c):
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= r <= rng.max_row) and (rng.min_col <= c <= rng.max_col):
            return rng.min_row, rng.min_col
    return r, c

def set_cell(ws, r, c, value, wrap=False, align_left=False, v_top=False):
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(row=r0, column=c0)
    cell.value = value
    if wrap or align_left or v_top:
        cell.alignment = Alignment(
            wrap_text=bool(wrap),
            horizontal=("left" if align_left else None),
            vertical=("top" if v_top else None),
        )

def right_value_cell_of_label(ws, label_text: str, offset_cols: int = 1, offset_rows: int = 0):
    pos = find_text_cell(ws, label_text)
    if not pos:
        return None
    r, c = pos
    return r + offset_rows, c + offset_cols

# ---------- time utils ----------
def parse_hhmm(s: str) -> time | None:
    s = (s or "").strip()
    if not s:
        return None
    hh, mm = s.split(":")
    return time(int(hh), int(mm))

def diff_minutes(beg: time | None, end: time | None) -> int:
    if not beg or not end:
        return 0
    dt = datetime(2000,1,1)
    b = dt.replace(hour=beg.hour, minute=beg.minute)
    e = dt.replace(hour=end.hour, minute=end.minute)
    if e <= b:
        return 0
    return int((e - b).total_seconds() // 60)

# ---------- routes ----------
@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str | None = Form(None),
    basf_beauftragter: str | None = Form(None),
    beschreibung: str = Form(...),
    break_minutes: str | None = Form(None),
    # workers (max 5)
    vorname1: str | None = Form(None), nachname1: str | None = Form(None),
    ausweis1: str | None = Form(None), beginn1: str | None = Form(None), ende1: str | None = Form(None),
    vorname2: str | None = Form(None), nachname2: str | None = Form(None),
    ausweis2: str | None = Form(None), beginn2: str | None = Form(None), ende2: str | None = Form(None),
    vorname3: str | None = Form(None), nachname3: str | None = Form(None),
    ausweis3: str | None = Form(None), beginn3: str | None = Form(None), ende3: str | None = Form(None),
    vorname4: str | None = Form(None), nachname4: str | None = Form(None),
    ausweis4: str | None = Form(None), beginn4: str | None = Form(None), ende4: str | None = Form(None),
    vorname5: str | None = Form(None), nachname5: str | None = Form(None),
    ausweis5: str | None = Form(None), beginn5: str | None = Form(None), ende5: str | None = Form(None),
):
    if not os.path.exists(TEMPLATE_XLSX):
        return JSONResponse({"detail": "Hiányzik a GP-t.xlsx sablon a gyökérben."}, status_code=500)

    try:
        wb = load_workbook(TEMPLATE_XLSX)
        ws = wb.active

        # Fejléc-mezők (a sablonban szereplő pontos feliratokra keresünk)
        # Dátum
        pos_date = right_value_cell_of_label(ws, "Datum der Leistungsausführung:", 1)
        if pos_date:
            set_cell(ws, pos_date[0], pos_date[1], datum)

        # Bau
        pos_bau = right_value_cell_of_label(ws, "Bau und Ausführungsort:", 1)
        if pos_bau:
            set_cell(ws, pos_bau[0], pos_bau[1], bau or "")

        # BASF-Beauftragter, Org.-Code
        pos_basf = right_value_cell_of_label(ws, "BASF-Beauftragter, Org.-Code:", 1)
        if pos_basf and basf_beauftragter:
            set_cell(ws, pos_basf[0], pos_basf[1], basf_beauftragter)

        # Beschreibung a nagy szövegblokk bal felső cellájába (pl. A6)
        set_cell(ws, 6, 1, beschreibung, wrap=True, align_left=True, v_top=True)

        # Munkások beírása – külön „Name” (vezetéknév) és „Vorname” oszlop
        name_col_cell = find_text_cell(ws, "Name")
        vorname_col_cell = find_text_cell(ws, "Vorname")
        ausweis_col_cell = None
        # a fejlécben az oszlopfelirat hosszabb
        for label in ["Ausweis- Nr.", "Ausweis- Nr. oder Kennzeichen", "Ausweis-Nr. / Kennzeichen"]:
            ausweis_col_cell = ausweis_col_cell or find_text_cell(ws, label)

        # kezdősor: a fejléc alatti első adat sor
        header_row = max(
            (name_col_cell[0] if name_col_cell else 0),
            (vorname_col_cell[0] if vorname_col_cell else 0),
            (ausweis_col_cell[0] if ausweis_col_cell else 0),
        )
        data_row = (header_row or 18) + 1

        col_name = name_col_cell[1] if name_col_cell else 2
        col_vor = vorname_col_cell[1] if vorname_col_cell else 3
        col_aus = ausweis_col_cell[1] if ausweis_col_cell else 6

        total_hours = 0.0
        bm = int(break_minutes or "60")  # 60 alapból, 30 ha pipálva volt

        for i in range(1, 6):
            vn = (locals().get(f"vorname{i}") or "").strip()
            nn = (locals().get(f"nachname{i}") or "").strip()
            aw = (locals().get(f"ausweis{i}") or "").strip()
            bg = parse_hhmm(locals().get(f"beginn{i}") or "")
            en = parse_hhmm(locals().get(f"ende{i}") or "")

            if not any([vn, nn, aw, bg, en]):
                continue

            # nevek + ausweis a saját oszlopukba
            if nn:
                set_cell(ws, data_row, col_name, nn, align_left=True)
            if vn:
                set_cell(ws, data_row, col_vor, vn, align_left=True)
            if aw:
                set_cell(ws, data_row, col_aus, aw, align_left=True)

            # óraszám számítása: (Ende-Beginn) - break
            mins = diff_minutes(bg, en)
            hours = max(0.0, (mins - bm) / 60.0)
            total_hours += round(hours, 2)

            # ha van "Anzahl Stunden" oszlop a táblában, oda is írhatsz:
            hours_header = find_text_cell(ws, "Anzahl Stunden")
            if hours_header:
                set_cell(ws, data_row, hours_header[1], round(hours, 2), align_left=True)

            # Beginn / Ende oszlopokba (ha megtaláljuk a fejlécet)
            beg_header = find_text_cell(ws, "Beginn")
            end_header = find_text_cell(ws, "Ende")
            if beg_header and bg:
                set_cell(ws, data_row, beg_header[1], f"{bg.hour:02d}:{bg.minute:02d}", align_left=True)
            if end_header and en:
                set_cell(ws, data_row, end_header[1], f"{en.hour:02d}:{en.minute:02d}", align_left=True)

            data_row += 1

        # Gesamtstunden cella kitöltése
        pos_total_label = find_text_cell(ws, "Gesamtstunden")
        if pos_total_label:
            set_cell(ws, pos_total_label[0], pos_total_label[1] + 1, round(total_hours, 2), align_left=True)

        # mentés és küldés
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        filename = f'leistungsnachweis_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
        headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return StreamingResponse(
            bio,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except Exception as e:
        return JSONResponse({"detail": f"Hiba a generálásnál: {e}"}, status_code=500)

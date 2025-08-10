# main.py
from fastapi import FastAPI, Form, Request
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional, List, Tuple
from datetime import datetime
import io

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
def index():
    return FileResponse("static/index.html")


# ---------------- Excel utilok ----------------
def A1(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"

def in_merge(ws, r: int, c: int):
    """Ha cella merge-ben van, visszaadja a CellRange-et, különben None."""
    coord = A1(c, r)
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return rng
    return None

def top_left(ws, r: int, c: int) -> Tuple[int, int]:
    rng = in_merge(ws, r, c)
    if rng:
        return rng.min_row, rng.min_col
    return r, c

def write_cell(ws, r: int, c: int, value, wrap=False, align_left=False, vtop=True):
    r0, c0 = top_left(ws, r, c)
    cell = ws.cell(row=r0, column=c0)
    cell.value = value
    align = cell.alignment or Alignment()
    cell.alignment = Alignment(
        wrap_text=True if wrap else align.wrapText,
        horizontal="left" if align_left else align.horizontal,
        vertical="top" if vtop else align.vertical,
    )

def find_label_cell(ws, needle: str, rows=(1, 120), cols=(1, 20)):
    """Rész-ill. egyezést is elfogad. (kis/nagy mindegy)"""
    n = needle.lower()
    r1, r2 = rows
    c1, c2 = cols
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            txt = cell.value
            if isinstance(txt, str) and n in txt.strip().lower():
                return cell.row, cell.column
    return None

def value_cell_right_of_label(ws, label_text: str):
    """A címke tartományának jobb széle UTÁNI bal-felső cellát adja vissza."""
    pos = find_label_cell(ws, label_text)
    if not pos:
        return None
    r, c = pos
    rng = in_merge(ws, r, c)
    right_col = (rng.max_col + 1) if rng else (c + 1)
    return r, right_col
# ------------------------------------------------


@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    beauftragter: Optional[str] = Form(None),
    beschreibung: Optional[str] = Form(""),
    # dolgozók
    vorname1: Optional[str] = Form(None), nachname1: Optional[str] = Form(None),
    ausweis1: Optional[str] = Form(None),  beginn1: Optional[str] = Form(None), ende1: Optional[str] = Form(None),
    vorname2: Optional[str] = Form(None), nachname2: Optional[str] = Form(None),
    ausweis2: Optional[str] = Form(None),  beginn2: Optional[str] = Form(None), ende2: Optional[str] = Form(None),
    vorname3: Optional[str] = Form(None), nachname3: Optional[str] = Form(None),
    ausweis3: Optional[str] = Form(None),  beginn3: Optional[str] = Form(None), ende3: Optional[str] = Form(None),
    vorname4: Optional[str] = Form(None), nachname4: Optional[str] = Form(None),
    ausweis4: Optional[str] = Form(None),  beginn4: Optional[str] = Form(None), ende4: Optional[str] = Form(None),
    vorname5: Optional[str] = Form(None), nachname5: Optional[str] = Form(None),
    ausweis5: Optional[str] = Form(None),  beginn5: Optional[str] = Form(None), ende5: Optional[str] = Form(None),
    gesamtstunden: Optional[str] = Form(None),
):
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Fejléc mezők – rugalmas címke keresés
    # DÁTUM
    pos_date = value_cell_right_of_label(ws, "datum der leistungs")
    if pos_date:
        try:
            d = datetime.strptime(datum, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            d = datum
        write_cell(ws, pos_date[0], pos_date[1], d)

    # BAU (a sablonban „Bau und Ausführungsort:”)
    pos_bau = value_cell_right_of_label(ws, "bau und ausführungsort")
    if pos_bau:
        write_cell(ws, pos_bau[0], pos_bau[1], bau)

    # BASF-Beauftragter (gyakran „BASF-Beauftragter, Org.-Code:”)
    pos_beauf = value_cell_right_of_label(ws, "basf-beauftragter")
    if pos_beauf and beauftragter:
        write_cell(ws, pos_beauf[0], pos_beauf[1], beauftragter)

    # Napi leírás: A6:G15 – több sor, balra zárt
    write_cell(ws, 6, 1, beschreibung or "", wrap=True, align_left=True)
    for r in range(6, 16):
        ws.row_dimensions[r].height = 22  # hogy ne vágja le az alsó sort

    # Dolgozók – fejléc keresés tág tartományban
    header_row = None
    name_col = vorname_col = ausweis_col = beginn_col = ende_col = None

    for row in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=20):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            t = cell.value.strip().lower()
            if t == "name":
                header_row = cell.row
                name_col = cell.column
            elif t == "vorname":
                vorname_col = cell.column
            elif "ausweis" in t and ("nr" in t or "nummer" in t or "kennzeichen" in t):
                ausweis_col = cell.column
            elif t == "beginn":
                beginn_col = cell.column
            elif t == "ende":
                ende_col = cell.column
        if header_row and (ausweis_col or beginn_col or ende_col):
            break

    # összeállítjuk a dolgozók listáját
    raw = [
        (vorname1, nachname1, ausweis1, beginn1, ende1),
        (vorname2, nachname2, ausweis2, beginn2, ende2),
        (vorname3, nachname3, ausweis3, beginn3, ende3),
        (vorname4, nachname4, ausweis4, beginn4, ende4),
        (vorname5, nachname5, ausweis5, beginn5, ende5),
    ]
    workers = []
    for v, n, a, b, e in raw:
        if any([v, n, a, b, e]):
            workers.append({
                "name": (f"{(n or '').strip()} {(v or '').strip()}").strip(),
                "ausweis": (a or "").strip(),
                "beginn": (b or "").strip(),
                "ende": (e or "").strip(),
            })

    if header_row and workers:
        r = header_row + 1
        for w in workers:
            if name_col:
                write_cell(ws, r, name_col, w["name"])
            if ausweis_col:
                write_cell(ws, r, ausweis_col, w["ausweis"])
            if beginn_col:
                write_cell(ws, r, beginn_col, w["beginn"])
            if ende_col:
                write_cell(ws, r, ende_col, w["ende"])
            r += 1

    # Összóraszám – ha küldöd az űrlapról
    if gesamtstunden:
        pos_total = value_cell_right_of_label(ws, "gesamtstunden") \
                 or value_cell_right_of_label(ws, "összes munkaóra")
        if pos_total:
            write_cell(ws, pos_total[0], pos_total[1], gesamtstunden)

    # Visszaadás
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"GP-t_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

# main.py
from fastapi import FastAPI, Form, Request
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional, List
from datetime import datetime
import io

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

app = FastAPI()

# statikus fájlok
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
def index():
    # a meglévő statikus index.html-t szolgáljuk ki
    return FileResponse("static/index.html")


# ---------- Excel utilok (merged cell safe írás) ----------
def _coord(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"

def top_left_of_merge(ws, row: int, col: int):
    """Ha (row,col) egy összevont tartományban van, visszaadja annak bal-felső celláját,
    különben (row,col)-t."""
    for rng in ws.merged_cells.ranges:
        # openpyxl itt cella-koordináta STRING-et vár (pl. "C5"), nem tuple-t
        if _coord(col, row) in rng:
            return rng.min_row, rng.min_col
    return row, col

def set_cell(ws, row: int, col: int, value, wrap: bool = False, align_left: bool = False):
    r0, c0 = top_left_of_merge(ws, row, col)
    cell = ws.cell(row=r0, column=c0)
    cell.value = value
    if wrap or align_left:
        cell.alignment = Alignment(
            wrap_text=True if wrap else (cell.alignment.wrapText if cell.alignment else False),
            horizontal="left" if align_left else (cell.alignment.horizontal if cell.alignment else None),
            vertical="top",
        )

def find_label(ws, text: str, search_rows=(1, 30), search_cols=(1, 10)):
    """Megkeresi a cellát, ami pontosan a megadott szöveget tartalmazza."""
    r1, r2 = search_rows
    c1, c2 = search_cols
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            val = cell.value
            if isinstance(val, str) and val.strip() == text.strip():
                return cell.row, cell.column
    return None

def right_value_cell_of_label(ws, label_text: str):
    """Visszaadja a címke melletti (ugyanazon sorban, jobbra eső) érték-cella bal-felső koordinátáját."""
    pos = find_label(ws, label_text)
    if not pos:
        return None
    r, c = pos
    # közvetlen jobb oldali cella (ha összevont, bal-felsőt kérjük le)
    r0, c0 = top_left_of_merge(ws, r, c + 1)
    return r0, c0
# ---------------------------------------------------------


@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    # fejlécek
    datum: str = Form(...),                  # pl. "2025-08-09"
    bau: str = Form(...),
    beauftragter: Optional[str] = Form(None),
    beschreibung: Optional[str] = Form(""),
    # dolgozó 1 kötelező mezők (a Te űrlapod szerint)
    vorname1: Optional[str] = Form(None),
    nachname1: Optional[str] = Form(None),
    ausweis1: Optional[str] = Form(None),
    beginn1: Optional[str] = Form(None),     # "HH:MM"
    ende1: Optional[str] = Form(None),       # "HH:MM"
    # opcionálisan a 2–5. dolgozó (ha az űrlap küldi)
    vorname2: Optional[str] = Form(None), nachname2: Optional[str] = Form(None),
    ausweis2: Optional[str] = Form(None), beginn2: Optional[str] = Form(None), ende2: Optional[str] = Form(None),
    vorname3: Optional[str] = Form(None), nachname3: Optional[str] = Form(None),
    ausweis3: Optional[str] = Form(None), beginn3: Optional[str] = Form(None), ende3: Optional[str] = Form(None),
    vorname4: Optional[str] = Form(None), nachname4: Optional[str] = Form(None),
    ausweis4: Optional[str] = Form(None), beginn4: Optional[str] = Form(None), ende4: Optional[str] = Form(None),
    vorname5: Optional[str] = Form(None), nachname5: Optional[str] = Form(None),
    ausweis5: Optional[str] = Form(None), beginn5: Optional[str] = Form(None), ende5: Optional[str] = Form(None),
    gesamtstunden: Optional[str] = Form(None),  # ha a frontenden kiszámolod és elküldöd
):
    # --- Excel sablon betöltése ---
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # --- Fejléc mezők a címkék alapján (nem fix koordináta!) ---
    # Dátum a "Datum der Leistungsausführung:" címke MELLÉ
    pos_date = right_value_cell_of_label(ws, "Datum der Leistungsausführung:")
    if pos_date:
        # dátum formázás (ha YYYY-MM-DD jön)
        try:
            d = datetime.strptime(datum, "%Y-%m-%d").strftime("%d.%m.%Y")
        except Exception:
            d = datum  # ha nem ISO formátum, hagyjuk, ahogy jön
        set_cell(ws, pos_date[0], pos_date[1], d)

    # Projekt/Bau a címke mellé
    pos_bau = right_value_cell_of_label(ws, "Projekt/Bau:")
    if pos_bau:
        set_cell(ws, pos_bau[0], pos_bau[1], bau)

    # BASF-Beauftragter a címke mellé
    pos_beauf = right_value_cell_of_label(ws, "BASF-Beauftragter:")
    if not pos_beauf:
        # egyes sablonokban sima kötőjellel szerepelhet
        pos_beauf = right_value_cell_of_label(ws, "BASF-Beauftragter:")
    if pos_beauf and beauftragter:
        set_cell(ws, pos_beauf[0], pos_beauf[1], beauftragter)

    # --- „Mit csináltunk ma” leírás (A6:G15 összevont terület) ---
    # Ezt korábban közösen így állítottuk be.
    set_cell(ws, 6, 1, beschreibung or "", wrap=True, align_left=True)

    # --- Dolgozók kiírása ---
    # Az űrlap nevei a Te oldaladon: vorname{i}, nachname{i}, ausweis{i}, beginn{i}, ende{i}
    workers = []
    raw = [
        (vorname1, nachname1, ausweis1, beginn1, ende1),
        (vorname2, nachname2, ausweis2, beginn2, ende2),
        (vorname3, nachname3, ausweis3, beginn3, ende3),
        (vorname4, nachname4, ausweis4, beginn4, ende4),
        (vorname5, nachname5, ausweis5, beginn5, ende5),
    ]
    for v, n, a, b, e in raw:
        if any([v, n, a, b, e]):
            workers.append({
                "name": f"{(n or '').strip()} {(v or '').strip()}".strip(),
                "ausweis": (a or "").strip(),
                "beginn": (b or "").strip(),
                "ende": (e or "").strip()
            })

    # A sablon táblázati része nálad már működött; itt egy kíméletes írás:
    # Keressük meg a „Name” és „Ausweis-Nr.” fejlécet, és az az alatti sorokba írunk.
    header_row = None
    name_col = None
    ausweis_col = None
    beginn_col = None
    ende_col = None

    # keresési tartomány ésszerűen
    for row in ws.iter_rows(min_row=10, max_row=40, min_col=1, max_col=12):
        for cell in row:
            txt = (cell.value or "")
            if isinstance(txt, str):
                t = txt.strip().lower()
                if t in ("name", "arbeiter", "mitarbeiter"):
                    header_row = cell.row
                    name_col = cell.column
                if t in ("ausweis-nr.", "ausweis", "ausweisnr."):
                    ausweis_col = cell.column
                if t in ("beginn", "arbeitsbeginn", "start"):
                    beginn_col = cell.column
                if t in ("ende", "arbeitsende", "stop"):
                    ende_col = cell.column
        if header_row:
            break

    if header_row:
        row = header_row + 1
        for w in workers:
            if name_col:
                set_cell(ws, row, name_col, w["name"])
            if ausweis_col:
                set_cell(ws, row, ausweis_col, w["ausweis"])
            if beginn_col:
                set_cell(ws, row, beginn_col, w["beginn"])
            if ende_col:
                set_cell(ws, row, ende_col, w["ende"])
            row += 1
    # Ha nem találtunk fejrécet, NEM írunk semmit a táblázati részbe,
    # így nem borítjuk fel a már működő pozicionálást.

    # --- Összóraszám: ha a frontenden kiszámoltad és küldöd, megpróbáljuk címke alapján beírni ---
    if gesamtstunden:
        pos_total = right_value_cell_of_label(ws, "Összes munkaóra:") or \
                    right_value_cell_of_label(ws, "Gesamtstunden:") or \
                    right_value_cell_of_label(ws, "Összes munkaóra")
        if pos_total:
            set_cell(ws, pos_total[0], pos_total[1], gesamtstunden)

    # --- Visszaadás memória-pufferből ---
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"GP-t_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

from fastapi import FastAPI, Form
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime
import io

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

TEMPLATE_PATH = Path("GP-t.xlsx")  # sablon a repo gyökerében

def a1(r:int, c:int)->str:
    return f"{get_column_letter(c)}{r}"

def merge_top_left(ws, r:int, c:int):
    cell = a1(r,c)
    for rng in ws.merged_cells.ranges:
        if cell in rng:
            return rng.min_row, rng.min_col
    return r,c

def set_cell(ws, r:int, c:int, value:str, wrap=False, left=False):
    r0,c0 = merge_top_left(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = value
    if wrap or left:
        cell.alignment = Alignment(
            wrap_text=True if wrap else cell.alignment.wrap_text,
            horizontal="left" if left else cell.alignment.horizontal
        )

def right_value_cell_of_label(ws, label:str):
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            txt = cell.value.strip() if isinstance(cell.value,str) else cell.value
            if txt == label:
                rr, cc = cell.row, cell.column + 1
                return merge_top_left(ws, rr, cc)
    return None

@app.get("/")
async def root():
    return FileResponse("static/index.html")

@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bau: str = Form(...),
    basf_beauftragter: str = Form(""),
    geraet: str = Form(""),
    beschreibung: str = Form(...),

    vorname1: str = Form(None), nachname1: str = Form(None), ausweis1: str = Form(None), beginn1: str = Form(None), ende1: str = Form(None),
    vorname2: str = Form(None), nachname2: str = Form(None), ausweis2: str = Form(None), beginn2: str = Form(None), ende2: str = Form(None),
    vorname3: str = Form(None), nachname3: str = Form(None), ausweis3: str = Form(None), beginn3: str = Form(None), ende3: str = Form(None),
    vorname4: str = Form(None), nachname4: str = Form(None), ausweis4: str = Form(None), beginn4: str = Form(None), ende4: str = Form(None),
    vorname5: str = Form(None), nachname5: str = Form(None), ausweis5: str = Form(None), beginn5: str = Form(None), ende5: str = Form(None),
):
    if not TEMPLATE_PATH.exists():
        return PlainTextResponse("Hiányzik a sablon: GP-t.xlsx", status_code=500)

    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Dátum
        pos = right_value_cell_of_label(ws, "Datum der Leistungsausführung:")
        if pos: set_cell(ws, pos[0], pos[1], datum)

        # Bau
        pos = right_value_cell_of_label(ws, "Bau:")
        if pos: set_cell(ws, pos[0], pos[1], bau)

        # BASF-Beauftragter
        pos = right_value_cell_of_label(ws, "BASF-Beauftragter:")
        if pos: set_cell(ws, pos[0], pos[1], basf_beauftragter)

        # Gerät/Fahrzeug
        pos = right_value_cell_of_label(ws, "Gerät/Fahrzeug:")
        if pos and geraet: set_cell(ws, pos[0], pos[1], geraet)

        # Leírás blokk (A6:G15 bal felső cella)
        set_cell(ws, 6, 1, beschreibung, wrap=True, left=True)

        # Dolgozók kiírása a korábbi, működő sorokra (B18..B22 és F18..F22)
        workers = [
            (vorname1, nachname1, ausweis1),
            (vorname2, nachname2, ausweis2),
            (vorname3, nachname3, ausweis3),
            (vorname4, nachname4, ausweis4),
            (vorname5, nachname5, ausweis5),
        ]
        start_row = 18
        for i, (v, n, a) in enumerate(workers):
            if not (v or n or a): continue
            row = start_row + i
            name = " ".join([p for p in [v, n] if p])
            if name: set_cell(ws, row, 2, name)   # B
            if a:    set_cell(ws, row, 6, a)      # F

        # Mentés memóriába
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        try:
            d = datetime.strptime(datum, "%Y-%m-%d").strftime("%Y%m%d")
        except Exception:
            d = datum.replace(".", "").replace("-", "").replace("/", "")
        fname = f"Tagesbericht_{bau}_{d}.xlsx".replace(" ", "_")

        return FileResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=fname,
        )
    except Exception as e:
        return PlainTextResponse(f"Generálási hiba: {e}", status_code=500)

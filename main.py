from fastapi import FastAPI, Form
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from pathlib import Path
import io
from datetime import datetime

app = FastAPI()

# statikus fájlok
app.mount("/static", StaticFiles(directory="static"), name="static")

TEMPLATE_PATH = Path("GP-t.xlsx")   # a repo gyökerében


# ---------- Merge-segédek (JAVÍTVA) ----------

def coord(r: int, c: int) -> str:
    """(row, col) -> 'A1'"""
    return f"{get_column_letter(c)}{r}"

def is_in_merge(ws, r: int, c: int):
    """Igaz, ha az adott cella bármelyik összevont tartomány része."""
    cell_str = coord(r, c)
    for rng in ws.merged_cells.ranges:
        if cell_str in rng:
            return rng  # visszaadjuk a CellRange-et
    return None

def top_left_of_merge(ws, r: int, c: int):
    """Ha (r,c) merge-ben van, adja vissza a merge bal felső (min_row, min_col) celláját, különben (r,c)."""
    rng = is_in_merge(ws, r, c)
    if rng:
        return rng.min_row, rng.min_col
    return r, c

def right_value_cell_of_label(ws, label_text: str):
    """
    Megkeresi a label cellát a lapon (egyező szöveggel),
    és visszaadja a TŐLE JOBBRA lévő értékmező bal felső celláját (merge-et is figyelembe véve).
    """
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            val = (cell.value or "").strip() if isinstance(cell.value, str) else cell.value
            if val == label_text:
                r, c = cell.row, cell.column
                # azonnal jobbra lépünk egy oszlopot
                target_r, target_c = r, c + 1
                # ha ez merge része, vegyük a merge bal felső sarkát
                target_r, target_c = top_left_of_merge(ws, target_r, target_c)
                return target_r, target_c
    return None


# ---------- írás cellába (wrap, balra zárás stb.) ----------

def set_cell(ws, r: int, c: int, text: str, *, wrap=False, align_left=False):
    """Szöveg beírása; ha merge-ben vagyunk, a bal felső cellába írunk."""
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(row=r0, column=c0)
    cell.value = text
    if wrap or align_left:
        cell.alignment = Alignment(
            wrap_text=True if wrap else cell.alignment.wrap_text,
            horizontal="left" if align_left else cell.alignment.horizontal
        )


# ---------- HTTP ----------

@app.get("/")
async def index():
    return FileResponse("static/index.html")


@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bau: str = Form(...),
    basf_beauftragter: str = Form(""),
    geraet: str = Form(""),
    beschreibung: str = Form(...),

    # dolgozók (max 5 – csak ami tényleg jön a formból, azt írjuk)
    vorname1: str = Form(None), nachname1: str = Form(None), ausweis1: str = Form(None),
    beginn1: str = Form(None),  ende1: str = Form(None),

    vorname2: str = Form(None), nachname2: str = Form(None), ausweis2: str = Form(None),
    beginn2: str = Form(None),  ende2: str = Form(None),

    vorname3: str = Form(None), nachname3: str = Form(None), ausweis3: str = Form(None),
    beginn3: str = Form(None),  ende3: str = Form(None),

    vorname4: str = Form(None), nachname4: str = Form(None), ausweis4: str = Form(None),
    beginn4: str = Form(None),  ende4: str = Form(None),

    vorname5: str = Form(None), nachname5: str = Form(None), ausweis5: str = Form(None),
    beginn5: str = Form(None),  ende5: str = Form(None),
):
    if not TEMPLATE_PATH.exists():
        return PlainTextResponse("Hiányzik a sablon: GP-t.xlsx", status_code=500)

    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # ---- Fejléc mezők címkék alapján ----
        # Dátum
        pos = right_value_cell_of_label(ws, "Datum der Leistungsausführung:")
        if pos:
            # Ha a sablonban az értékmező egy sorral lejjebb van, ezen tudunk finomítani.
            set_cell(ws, pos[0], pos[1], datum)

        # Bau
        pos = right_value_cell_of_label(ws, "Bau:")
        if pos:
            set_cell(ws, pos[0], pos[1], bau)

        # BASF-Beauftragter – sokszor E3 (összevont), de címke alapján keressük
        pos = right_value_cell_of_label(ws, "BASF-Beauftragter:")
        if pos:
            set_cell(ws, pos[0], pos[1], basf_beauftragter)

        # Gerät/Fahrzeug – ha van a sablonon címke, oda; ha nincs, kihagyjuk csendben
        pos = right_value_cell_of_label(ws, "Gerät/Fahrzeug:")
        if pos and geraet:
            set_cell(ws, pos[0], pos[1], geraet)

        # ---- Leírás: A6:G15 merge blokk bal felső cellája ----
        # (A sablonban ez az a nagy, több soros terület.)
        set_cell(ws, 6, 1, beschreibung, wrap=True, align_left=True)

        # ---- Dolgozók — név + Ausweis beírása a sablon megfelelő soraira ----
        # Itt a sablon szerinti konkrét helyekre írunk. Ha máshol vannak,
        # csak az indexek finomhangolása kell.
        # Példa elhelyezés:
        # 1. dolgozó: Név: B18, Ausweis: F18
        # 2.:         B19,           F19
        # 3.:         B20,           F20
        # 4.:         B21,           F21
        # 5.:         B22,           F22
        worker_slots = [
            (vorname1, nachname1, ausweis1),
            (vorname2, nachname2, ausweis2),
            (vorname3, nachname3, ausweis3),
            (vorname4, nachname4, ausweis4),
            (vorname5, nachname5, ausweis5),
        ]
        start_row = 18
        for idx, (v, n, a) in enumerate(worker_slots, start=0):
            if not (v or n or a):
                continue
            row = start_row + idx
            full_name = " ".join([p for p in [v, n] if p])
            if full_name:
                set_cell(ws, row, 2, full_name)   # B-col
            if a:
                set_cell(ws, row, 6, a)           # F-col

        # (Munkaidők és összóra a frontenden számolódnak; ha mégis cellába kell írni,
        # ugyanitt bővíthető.)

        # ---- Válasz: XLSX bájtsor visszaadása ----
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Fájlnév: bau + dátum
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

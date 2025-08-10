from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
import os
import io

app = FastAPI()

# Serve static
app.mount("/static", StaticFiles(directory="static"), name="static")

INDEX_PATH = os.path.join("static", "index.html")
TEMPLATE_XLSX = os.path.join("GP-t.xlsx")

def find_text_cell(ws, text: str):
    """Find the first cell that matches text exactly (stripped). Returns (row, col) or None."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            v = "" if cell.value is None else str(cell.value).strip()
            if v == text.strip():
                return cell.row, cell.column
    return None

def top_left_of_merge(ws, r, c):
    """Given a cell coordinate (r,c), if it lies within a merged range, return that range's top-left.
    Otherwise return (r,c)."""
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= r <= rng.max_row) and (rng.min_col <= c <= rng.max_col):
            return rng.min_row, rng.min_col
    return r, c

def set_cell(ws, r, c, value, wrap=False, align_left=False, v_top=False):
    """Write value safely even if (r,c) is inside a merged range."""
    r0, c0 = top_left_of_merge(ws, r, c)
    cell = ws.cell(row=r0, column=c0)
    cell.value = value
    if wrap or align_left or v_top:
        cell.alignment = Alignment(
            wrap_text=bool(wrap),
            horizontal=("left" if align_left else None),
            vertical=("top" if v_top else None)
        )

def right_value_cell_of_label(ws, label_text: str, offset_cols: int = 1, offset_rows: int = 0):
    """Find a label cell by exact text and return the coordinate to its right (with offsets)."""
    pos = find_text_cell(ws, label_text)
    if not pos:
        return None
    r, c = pos
    return r + offset_rows, c + offset_cols

@app.get("/", response_class=HTMLResponse)
async def root():
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return f.read()

@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bau: str | None = Form(None),
    projekt: str | None = Form(None),
    basf: str | None = Form(None),
    beschreibung: str = Form(...),
    total_hours: str | None = Form(None),
    # up to 5 workers (optional); empty ones are ignored
    vorname1: str | None = Form(None),
    nachname1: str | None = Form(None),
    ausweis1: str | None = Form(None),
    vorname2: str | None = Form(None),
    nachname2: str | None = Form(None),
    ausweis2: str | None = Form(None),
    vorname3: str | None = Form(None),
    nachname3: str | None = Form(None),
    ausweis3: str | None = Form(None),
    vorname4: str | None = Form(None),
    nachname4: str | None = Form(None),
    ausweis4: str | None = Form(None),
    vorname5: str | None = Form(None),
    nachname5: str | None = Form(None),
    ausweis5: str | None = Form(None),
):
    # Basic validation for required fields
    if not bau and not projekt:
        return JSONResponse({"detail": 'Hiányzó mező: "bau" vagy "projekt".'}, status_code=400)

    if not os.path.exists(TEMPLATE_XLSX):
        return JSONResponse({"detail": "Hiányzik a GP-t.xlsx sablon a gyökérben."}, status_code=500)

    try:
        wb = load_workbook(TEMPLATE_XLSX)
        ws = wb.active

        # 1) Dátum (label: "Datum der Leistungsausführung:")
        target = right_value_cell_of_label(ws, "Datum der Leistungsausführung:", offset_cols=1)
        if target:
            set_cell(ws, target[0], target[1], datum)

        # 2) Bau/Projekt (label próbák)
        target_bp = (right_value_cell_of_label(ws, "Bau:", 1) or
                     right_value_cell_of_label(ws, "Projekt:", 1) or
                     right_value_cell_of_label(ws, "Projekt / Bau:", 1))
        if target_bp:
            set_cell(ws, target_bp[0], target_bp[1], bau or projekt or "")

        # 3) BASF-Beauftragter (label: "BASF-Beauftragter:")
        if basf:
            target_basf = right_value_cell_of_label(ws, "BASF-Beauftragter:", 1)
            if target_basf:
                set_cell(ws, target_basf[0], target_basf[1], basf)

        # 4) Tevékenység szöveg az A6:G15 blokkba (bal-felső A6)
        #    A sablonod már merge-ölve van, mi a bal-felsőre írunk, sortöréssel, balra-zártan.
        set_cell(ws, 6, 1, beschreibung, wrap=True, align_left=True, v_top=True)

        # 5) Össz óraszám (ha jön a front-endtől)
        if total_hours:
            # keresünk egy "Összesen" feliratot, vagy a típikus helyre írunk (példa: G18)
            pos_total = (right_value_cell_of_label(ws, "Összesen:", 1) or
                         right_value_cell_of_label(ws, "Gesamtstunden:", 1))
            if pos_total:
                set_cell(ws, pos_total[0], pos_total[1], total_hours)

        # 6) Dolgozók – ideiglenes, generikus elhelyezés (ha nincs fix címke a sablonban)
        # Ha a sablonban van címke (pl. "Name", "Ausweis-Nr."), ide érdemes igazítani később.
        workers = []
        for i in range(1, 6):
            vn = locals().get(f"vorname{i}")
            nn = locals().get(f"nachname{i}")
            au = locals().get(f"ausweis{i}")
            if (vn and vn.strip()) or (nn and nn.strip()) or (au and au.strip()):
                workers.append({"vor": vn or "", "nach": nn or "", "ausweis": au or ""})

        # Próbáljuk címkék mellett elhelyezni, különben fallback egy tipikus blokkra.
        name_header = find_text_cell(ws, "Name:")
        ausweis_header = find_text_cell(ws, "Ausweis-Nr.:")
        start_row = None
        col_name = None
        col_ausw = None
        if name_header and ausweis_header:
            # a fejléc alatt kezdünk
            start_row = max(name_header[0], ausweis_header[0]) + 1
            col_name = name_header[1]
            col_ausw = ausweis_header[1]
        else:
            # fallback: tegyük a neveket a B19-től lefelé és az Ausweis-t az F19-től (hangolható)
            start_row = 19
            col_name = 2
            col_ausw = 6

        r = start_row
        for w in workers:
            full = (w["nach"] + " " + w["vor"]).strip()
            if full:
                set_cell(ws, r, col_name, full)
            if w["ausweis"]:
                set_cell(ws, r, col_ausw, w["ausweis"])
            r += 1

        # Kész fájl memória-pufferbe
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        # Fájlnév
        safe_date = datum.replace("/", "-").replace(".", "-")
        filename = f"GP-t_filled_{safe_date or 'heute'}.xlsx"

        return FileResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename=filename)
    except Exception as e:
        return JSONResponse({"detail": f"Hiba a generálásnál: {e}"}, status_code=500)

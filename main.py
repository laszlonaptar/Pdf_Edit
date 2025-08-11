from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
import tempfile

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# ---- helper: írás összevont cellákba biztonságosan ----
def set_merged_safe(ws, addr: str, value):
    cell = ws[addr]
    r, c = cell.row, cell.column
    wrote = False
    for rng in ws.merged_cells.ranges:
        # ha az adott cím egy merge-blokk része, a bal-felső cellába írunk
        if (r, c) in rng:
            tl = ws.cell(row=rng.min_row, column=rng.min_col)
            tl.value = value
            wrote = True
            break
    if not wrote:
        cell.value = value

@app.get("/", response_class=HTMLResponse)
async def get_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    basf_beauftragter: str = Form(""),
    beschreibung: str = Form(...),
    nachname1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
    pause1: str = Form(""),
    vorhaltung1: str = Form(""),
    nachname2: str = Form(""), vorname2: str = Form(""), ausweis2: str = Form(""),
    beginn2: str = Form(""), ende2: str = Form(""), pause2: str = Form(""), vorhaltung2: str = Form(""),
    nachname3: str = Form(""), vorname3: str = Form(""), ausweis3: str = Form(""),
    beginn3: str = Form(""), ende3: str = Form(""), pause3: str = Form(""), vorhaltung3: str = Form(""),
    nachname4: str = Form(""), vorname4: str = Form(""), ausweis4: str = Form(""),
    beginn4: str = Form(""), ende4: str = Form(""), pause4: str = Form(""), vorhaltung4: str = Form(""),
    nachname5: str = Form(""), vorname5: str = Form(""), ausweis5: str = Form(""),
    beginn5: str = Form(""), ende5: str = Form(""), pause5: str = Form(""), vorhaltung5: str = Form("")
):
    # Német dátum formátum (TT.MM.JJJJ)
    try:
        datum_dt = datetime.strptime(datum, "%Y-%m-%d")
        datum_formatted = datum_dt.strftime("%d.%m.%Y")
    except ValueError:
        datum_formatted = datum

    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Felső mezők – merged-safe
    set_merged_safe(ws, "B3", datum_formatted)        # Datum der Leistungsausführung
    set_merged_safe(ws, "B4", bau)                    # Bau und Ausführungsort
    if (basf_beauftragter or "").strip():
        set_merged_safe(ws, "G3", basf_beauftragter)  # BASF-Beauftragter, Org.-Code

    # Beschreibung: A6–G15 soronként (nem merged, soronként külön cellába)
    lines = beschreibung.split("\n")
    row = 6
    for line in lines:
        if row > 15:
            break
        ws[f"A{row}"] = line
        row += 1

    # Dolgozók
    employees = [
        (nachname1, vorname1, ausweis1, beginn1, ende1, pause1, vorhaltung1),
        (nachname2, vorname2, ausweis2, beginn2, ende2, pause2, vorhaltung2),
        (nachname3, vorname3, ausweis3, beginn3, ende3, pause3, vorhaltung3),
        (nachname4, vorname4, ausweis4, beginn4, ende4, pause4, vorhaltung4),
        (nachname5, vorname5, ausweis5, beginn5, ende5, pause5, vorhaltung5),
    ]

    start_row = 17
    for i, emp in enumerate(employees):
        if emp[0] and emp[1]:  # van vezetéknév és keresztnév
            r = start_row + i
            ws[f"A{r}"] = emp[0]   # Nachname
            ws[f"B{r}"] = emp[1]   # Vorname
            ws[f"C{r}"] = emp[2]   # Ausweis
            ws[f"D{r}"] = emp[3]   # Beginn
            ws[f"E{r}"] = emp[4]   # Ende
            ws[f"F{r}"] = emp[5]   # Pause (ha van)
            ws[f"G{r}"] = emp[6]   # Vorhaltung (ha van)

    # Mentés és visszaküldés
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()

    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="arbeitsnachweis.xlsx"
    )

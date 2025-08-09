from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import uvicorn
from openpyxl import load_workbook
from datetime import datetime
import os
import tempfile

app = FastAPI()

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

# Főoldal
@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Excel generálás
@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
    mitarbeiter_nachname_1: str = Form(...),
    mitarbeiter_vorname_1: str = Form(...),
    ausweisnummer_1: str = Form(...),
    mitarbeiter_nachname_2: str = Form(""),
    mitarbeiter_vorname_2: str = Form(""),
    ausweisnummer_2: str = Form(""),
    mitarbeiter_nachname_3: str = Form(""),
    mitarbeiter_vorname_3: str = Form(""),
    ausweisnummer_3: str = Form(""),
    mitarbeiter_nachname_4: str = Form(""),
    mitarbeiter_vorname_4: str = Form(""),
    ausweisnummer_4: str = Form(""),
    mitarbeiter_nachname_5: str = Form(""),
    mitarbeiter_vorname_5: str = Form(""),
    ausweisnummer_5: str = Form("")
):
    # Kötelező mezők ellenőrzése
    if not bau or not mitarbeiter_nachname_1:
        return JSONResponse(status_code=400, content={"detail": "Kötelező mezők hiányoznak: projekt/bau és legalább egy dolgozó (név)."})

    try:
        # Sablon betöltése
        wb = load_workbook("GP-t.xlsx")
        ws = wb.active

        # Mezők kitöltése (itt még nem pozicionálunk újra, csak működjön minden)
        ws["B2"] = datum
        ws["B3"] = bau
        ws["B4"] = bf
        ws["A6"] = beschreibung

        # Munkaidő számítása
        total_hours = calculate_hours(beginn, ende)
        ws["H2"] = beginn
        ws["I2"] = ende
        ws["J2"] = total_hours

        # Dolgozók adatai
        employees = [
            (mitarbeiter_nachname_1, mitarbeiter_vorname_1, ausweisnummer_1),
            (mitarbeiter_nachname_2, mitarbeiter_vorname_2, ausweisnummer_2),
            (mitarbeiter_nachname_3, mitarbeiter_vorname_3, ausweisnummer_3),
            (mitarbeiter_nachname_4, mitarbeiter_vorname_4, ausweisnummer_4),
            (mitarbeiter_nachname_5, mitarbeiter_vorname_5, ausweisnummer_5)
        ]

        row = 20
        for nachname, vorname, ausweis in employees:
            if nachname.strip():
                ws[f"A{row}"] = nachname
                ws[f"B{row}"] = vorname
                ws[f"C{row}"] = ausweis
                row += 1

        # Ideiglenes fájl mentése
        tmp_dir = tempfile.mkdtemp()
        file_path = os.path.join(tmp_dir, "arbeitsnachweis.xlsx")
        wb.save(file_path)

        return FileResponse(file_path, filename="arbeitsnachweis.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})


def calculate_hours(start_time: str, end_time: str) -> float:
    """Munkaidő számítása a fix szünetek levonásával."""
    fmt = "%H:%M"
    start = datetime.strptime(start_time, fmt)
    end = datetime.strptime(end_time, fmt)
    hours = (end - start).seconds / 3600

    # Reggeli szünet 9:00–9:15
    if start <= datetime.strptime("09:00", fmt) < end:
        hours -= 0.25
    # Ebédszünet 12:00–12:45
    if start <= datetime.strptime("12:00", fmt) < end:
        hours -= 0.75

    return round(hours, 2)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)

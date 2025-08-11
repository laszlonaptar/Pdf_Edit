from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
import tempfile
import os

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


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
    nachname2: str = Form(""),
    vorname2: str = Form(""),
    ausweis2: str = Form(""),
    beginn2: str = Form(""),
    ende2: str = Form(""),
    pause2: str = Form(""),
    vorhaltung2: str = Form(""),
    nachname3: str = Form(""),
    vorname3: str = Form(""),
    ausweis3: str = Form(""),
    beginn3: str = Form(""),
    ende3: str = Form(""),
    pause3: str = Form(""),
    vorhaltung3: str = Form(""),
    nachname4: str = Form(""),
    vorname4: str = Form(""),
    ausweis4: str = Form(""),
    beginn4: str = Form(""),
    ende4: str = Form(""),
    pause4: str = Form(""),
    vorhaltung4: str = Form(""),
    nachname5: str = Form(""),
    vorname5: str = Form(""),
    ausweis5: str = Form(""),
    beginn5: str = Form(""),
    ende5: str = Form(""),
    pause5: str = Form(""),
    vorhaltung5: str = Form("")
):
    # Német dátum formátumra konvertálás (TT.MM.JJJJ)
    try:
        datum_dt = datetime.strptime(datum, "%Y-%m-%d")
        datum_formatted = datum_dt.strftime("%d.%m.%Y")
    except ValueError:
        datum_formatted = datum  # Ha nem sikerül konvertálni, marad az eredeti

    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Adatok beírása
    ws["B3"] = datum_formatted
    ws["B4"] = bau
    ws["G3"] = basf_beauftragter

    # Munka leírása (A6–G15 sorok feltöltése balról jobbra)
    beschreibung_lines = beschreibung.split("\n")
    row = 6
    for line in beschreibung_lines:
        if row > 15:
            break
        ws[f"A{row}"] = line
        row += 1

    # Dolgozók adatainak beírása
    employees = [
        (nachname1, vorname1, ausweis1, beginn1, ende1, pause1, vorhaltung1),
        (nachname2, vorname2, ausweis2, beginn2, ende2, pause2, vorhaltung2),
        (nachname3, vorname3, ausweis3, beginn3, ende3, pause3, vorhaltung3),
        (nachname4, vorname4, ausweis4, beginn4, ende4, pause4, vorhaltung4),
        (nachname5, vorname5, ausweis5, beginn5, ende5, pause5, vorhaltung5),
    ]

    start_row = 17
    for i, emp in enumerate(employees):
        if emp[0] and emp[1]:  # Ha van vezetéknév és keresztnév
            ws[f"A{start_row+i}"] = emp[0]
            ws[f"B{start_row+i}"] = emp[1]
            ws[f"C{start_row+i}"] = emp[2]
            ws[f"D{start_row+i}"] = emp[3]
            ws[f"E{start_row+i}"] = emp[4]
            ws[f"F{start_row+i}"] = emp[5]
            ws[f"G{start_row+i}"] = emp[6]

    # Ideiglenes fájl mentése
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()

    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="arbeitsnachweis.xlsx"
    )

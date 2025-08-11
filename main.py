from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from datetime import datetime
from openpyxl import load_workbook
import os
import tempfile

app = FastAPI()

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau_ort: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
    gesamtstunden: str = Form(...),
    vorhaltung: str = Form(""),
    nachname_1: str = Form(...),
    vorname_1: str = Form(...),
    ausweis_1: str = Form(...),
    nachname_2: str = Form(""),
    vorname_2: str = Form(""),
    ausweis_2: str = Form(""),
    nachname_3: str = Form(""),
    vorname_3: str = Form(""),
    ausweis_3: str = Form(""),
    nachname_4: str = Form(""),
    vorname_4: str = Form(""),
    ausweis_4: str = Form(""),
    nachname_5: str = Form(""),
    vorname_5: str = Form(""),
    ausweis_5: str = Form(""),
):
    # Betöltjük a sablont
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    def put_value_right_of_label(ws, label, value, align_left=True):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value == label:
                    target_cell = ws.cell(row=cell.row, column=cell.column + 1)
                    target_cell.value = value
                    if align_left:
                        target_cell.alignment = cell.alignment
                    return

    # Dátum német formátumban (TT.MM.JJJJ)
    try:
        dt = datetime.strptime(datum.strip(), "%Y-%m-%d")
        date_text = dt.strftime("%d.%m.%Y")  # pl. 11.08.2025
    except Exception:
        date_text = datum

    put_value_right_of_label(ws, "Datum der Leistungsausführung:", date_text, align_left=True)

    # Egyéb mezők
    put_value_right_of_label(ws, "Bau und Ausführungsort:", bau_ort)
    put_value_right_of_label(ws, "Bauleiter / Fachbauleiter:", bf)
    put_value_right_of_label(ws, "Vorhaltung / beauftragtes Gerät / Fahrzeug:", vorhaltung)

    # Munka leírása (A6-G15 cellák kitöltése soronként)
    lines = beschreibung.split("\n")
    row_index = 6
    for line in lines:
        if row_index > 15:
            break
        ws[f"A{row_index}"] = line
        row_index += 1

    # Dolgozói adatok
    workers = [
        (nachname_1, vorname_1, ausweis_1),
        (nachname_2, vorname_2, ausweis_2),
        (nachname_3, vorname_3, ausweis_3),
        (nachname_4, vorname_4, ausweis_4),
        (nachname_5, vorname_5, ausweis_5),
    ]

    start_row = 28
    for i, (nachname, vorname, ausweis) in enumerate(workers):
        if nachname or vorname or ausweis:
            ws[f"A{start_row + i}"] = nachname
            ws[f"B{start_row + i}"] = vorname
            ws[f"C{start_row + i}"] = ausweis

    # Munkaidő adatok
    put_value_right_of_label(ws, "Beginn:", beginn)
    put_value_right_of_label(ws, "Ende:", ende)
    put_value_right_of_label(ws, "Gesamtstunden:", gesamtstunden)

    # Ideiglenes fájl mentése
    tmp_dir = tempfile.gettempdir()
    output_path = os.path.join(tmp_dir, "arbeitsnachweis.xlsx")
    wb.save(output_path)

    return FileResponse(output_path, filename="arbeitsnachweis.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

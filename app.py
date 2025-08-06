from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bauort: str = Form(...),
    bf: str = Form(...),
    taetigkeit: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
    geraet: str = Form(None),
    name1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...)
):
    # Sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Alapadatok beírása
    ws["C4"] = datum
    ws["C5"] = bauort
    ws["C6"] = bf
    ws["C7"] = taetigkeit

    # Dolgozó 1
    ws["B11"] = name1
    ws["C11"] = vorname1
    ws["D11"] = ausweis1

    # Munkaidő
    ws["E11"] = beginn
    ws["F11"] = ende

    # Időtartam kiszámítása
    fmt = "%H:%M"
    start = datetime.strptime(beginn, fmt)
    end = datetime.strptime(ende, fmt)
    total = (end - start).seconds / 3600

    # Szünetek levonása
    if start <= datetime.strptime("09:15", fmt) and end >= datetime.strptime("09:00", fmt):
        total -= 0.25
    if start <= datetime.strptime("12:45", fmt) and end >= datetime.strptime("12:00", fmt):
        total -= 0.75

    ws["G11"] = round(total, 2)

    if geraet:
        ws["B21"] = geraet

    output_path = "arbeitsnachweis.xlsx"
    wb.save(output_path)
    return FileResponse(output_path, filename=output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

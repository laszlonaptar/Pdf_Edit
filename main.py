from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
import shutil
import os
from datetime import datetime
from typing import Optional

app = FastAPI()

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

# Főoldal – index.html betöltése
@app.get("/", response_class=HTMLResponse)
async def read_index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Fájl feltöltés oldal – csak sablon teszteléshez
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    file_location = f"temp/{file.filename}"
    os.makedirs("temp", exist_ok=True)
    with open(file_location, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return {"info": f"File saved at {file_location}"}

# Kitöltés és letöltés
@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bauort: str = Form(...),
    geraet: Optional[str] = Form(""),
    arbeiter_name: str = Form(...),
    arbeiter_vorname: str = Form(...),
    arbeiter_ausweis: str = Form(...),
    taetigkeit: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
):
    template_path = "GP-t.xlsx"
    wb = load_workbook(template_path)
    ws = wb.active

    # Dátum és alapadatok
    ws["C2"] = datum
    ws["C3"] = bauort
    ws["C4"] = geraet

    # Munkaidő-számítás szünetekkel
    def parse_time(t: str) -> datetime:
        return datetime.strptime(t, "%H:%M")

    start = parse_time(beginn)
    end = parse_time(ende)
    pause = 0

    if start <= parse_time("09:00") < end:
        pause += 0.25
    if start <= parse_time("12:00") < end:
        pause += 0.75

    total_hours = (end - start).seconds / 3600 - pause
    total_hours = max(total_hours, 0)

    # Kitöltés
    ws["A7"] = arbeiter_name
    ws["B7"] = arbeiter_vorname
    ws["C7"] = arbeiter_ausweis
    ws["D7"] = taetigkeit
    ws["E7"] = beginn
    ws["F7"] = ende
    ws["G7"] = round(total_hours, 2)
    ws["H7"] = round(total_hours, 2)

    output_path = "output/generated.xlsx"
    os.makedirs("output", exist_ok=True)
    wb.save(output_path)

    return FileResponse(output_path, filename="Arbeitsnachweis.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

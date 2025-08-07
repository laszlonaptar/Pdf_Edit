from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import shutil
from openpyxl import load_workbook

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/")
async def form_get(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    name1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...)
):
    shutil.copy("GP-t.xlsx", "output.xlsx")
    wb = load_workbook("output.xlsx")
    ws = wb.active

    # Egyszerű mezők
    ws["D6"] = datum
    ws["D7"] = projekt
    ws["D8"] = bf
    ws["A6"] = beschreibung

    ws["A17"] = name1
    ws["B17"] = vorname1
    ws["C17"] = ausweis1
    ws["D17"] = beginn1
    ws["E17"] = ende1

    # Összóra számítás szünetek levonásával
    from datetime import datetime

    def parse_time(t):
        return datetime.strptime(t, "%H:%M")

    start = parse_time(beginn1)
    end = parse_time(ende1)
    total_hours = (end - start).seconds / 3600

    break_time = 0.0
    if start <= parse_time("09:15") and end >= parse_time("09:00"):
        break_time += 0.25
    if start <= parse_time("12:45") and end >= parse_time("12:00"):
        break_time += 0.75

    total_hours -= break_time
    ws["F17"] = round(total_hours, 2)
    ws["G17"] = round(total_hours, 2)

    wb.save("output.xlsx")
    return FileResponse("output.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Arbeitsnachweis.xlsx")

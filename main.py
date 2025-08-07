from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
import io

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/")
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bauort: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    geraet: str = Form(""),
    vorname1: str = Form(...),
    nachname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
    vorname2: str = Form(None),
    nachname2: str = Form(None),
    ausweis2: str = Form(None),
    beginn2: str = Form(None),
    ende2: str = Form(None)
):
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    ws["D6"] = datum
    ws["D7"] = bauort
    ws["D8"] = bf
    ws["B11"] = beschreibung
    ws["B21"] = geraet

    def calc_hours(start, end):
        if not start or not end:
            return ""
        start_time = datetime.strptime(start, "%H:%M")
        end_time = datetime.strptime(end, "%H:%M")
        total = (end_time - start_time).seconds / 3600

        pause = 0
        if start_time <= datetime.strptime("09:00", "%H:%M") < end_time:
            pause += 0.25
        if start_time <= datetime.strptime("12:00", "%H:%M") < end_time:
            pause += 0.75
        return round(total - pause, 2)

    ws["B13"] = nachname1
    ws["C13"] = vorname1
    ws["D13"] = ausweis1
    ws["E13"] = beginn1
    ws["F13"] = ende1
    ws["G13"] = calc_hours(beginn1, ende1)

    if vorname2:
        ws["B14"] = nachname2
        ws["C14"] = vorname2
        ws["D14"] = ausweis2
        ws["E14"] = beginn2
        ws["F14"] = ende2
        ws["G14"] = calc_hours(beginn2, ende2)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="arbeitsnachweis.xlsx")
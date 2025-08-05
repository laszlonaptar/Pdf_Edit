import os
import openpyxl
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.responses import RedirectResponse
from datetime import datetime

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_TEMPLATE_PATH = os.path.join(BASE_DIR, "GP-t.xlsx")


@app.get("/", response_class=HTMLResponse)
async def form_get(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate", response_class=FileResponse)
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    was_wurde_gemacht: str = Form(...),
    geraet: str = Form(""),  # nem kötelező
    arbeiter_nachname: str = Form(...),
    arbeiter_vorname: str = Form(...),
    arbeiter_ausweis: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
):
    wb = openpyxl.load_workbook(EXCEL_TEMPLATE_PATH)
    ws = wb.active

    ws["E4"] = datum
    ws["E5"] = projekt
    ws["E6"] = bf
    ws["E7"] = was_wurde_gemacht
    ws["E8"] = geraet
    ws["B13"] = arbeiter_nachname
    ws["C13"] = arbeiter_vorname
    ws["D13"] = arbeiter_ausweis
    ws["E13"] = beginn
    ws["F13"] = ende

    # Munkaidő számítás
    time_format = "%H:%M"
    try:
        start_dt = datetime.strptime(beginn, time_format)
        end_dt = datetime.strptime(ende, time_format)
        duration = (end_dt - start_dt).total_seconds() / 3600.0

        pause = 0.0
        if start_dt.time() < datetime.strptime("09:15", time_format).time() and end_dt.time() > datetime.strptime("09:00", time_format).time():
            pause += 0.25
        if start_dt.time() < datetime.strptime("12:45", time_format).time() and end_dt.time() > datetime.strptime("12:00", time_format).time():
            pause += 0.75

        worked_hours = max(duration - pause, 0)
        ws["G13"] = round(worked_hours, 2)
        ws["H13"] = round(worked_hours, 2)
    except Exception as e:
        ws["G13"] = "Hiba"
        ws["H13"] = "Hiba"

    output_filename = os.path.join(BASE_DIR, "arbeitsnachweis_kitoltve.xlsx")
    wb.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="arbeitsnachweis_kitoltve.xlsx"
    )

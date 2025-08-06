
from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import shutil

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    basf: str = Form(...),
    beschreibung: str = Form(...),
    nachname: list[str] = Form(...),
    vorname: list[str] = Form(...),
    ausweis: list[str] = Form(...),
    beginn: list[str] = Form(...),
    ende: list[str] = Form(...)
):
    template_path = Path("GP-t.xlsx")
    output_path = Path("generated.xlsx")
    wb = load_workbook(template_path)
    ws = wb.active

    ws["B4"] = datum
    ws["I4"] = projekt
    ws["B5"] = basf
    ws["B7"] = beschreibung

    for i in range(len(nachname)):
        row = 10 + i
        ws[f"B{row}"] = nachname[i]
        ws[f"C{row}"] = vorname[i]
        ws[f"D{row}"] = ausweis[i]
        ws[f"E{row}"] = beginn[i]
        ws[f"F{row}"] = ende[i]

        b = datetime.strptime(beginn[i], "%H:%M")
        e = datetime.strptime(ende[i], "%H:%M")
        total = (e - b).seconds / 3600

        if b <= datetime.strptime("09:00", "%H:%M") <= e:
            total -= 0.25
        if b <= datetime.strptime("12:00", "%H:%M") <= e:
            total -= 0.75

        ws[f"G{row}"] = round(total, 2)

    wb.save(output_path)
    return FileResponse(output_path, filename="Arbeitsnachweis.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

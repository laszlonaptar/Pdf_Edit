
from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime
import openpyxl
import shutil

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def form_get(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    geraet: str = Form(""),
    was_gemacht: str = Form(...),
    name1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
):

    shutil.copy("GP-t.xlsx", "output.xlsx")
    wb = openpyxl.load_workbook("output.xlsx")
    ws = wb.active

    # Cellákba írás (nem összevont cellába!)
    ws["D6"].value = datum
    ws["D7"].value = projekt
    ws["D8"].value = bf
    ws["D9"].value = geraet

    # "Was wurde gemacht?" mező: egy cellába (A6)
    ws["A6"].value = was_gemacht

    # Munkaóra számítás
    def berechne_stunden(beginn, ende):
        fmt = "%H:%M"
        start = datetime.strptime(beginn, fmt)
        end = datetime.strptime(ende, fmt)
        total = (end - start).seconds / 3600

        # szünetek levonása
        if start <= datetime.strptime("09:00", fmt) <= end:
            total -= 0.25
        if start <= datetime.strptime("12:00", fmt) <= end:
            total -= 0.75
        return max(total, 0)

    gesamtstunden = berechne_stunden(beginn1, ende1)

    # Dolgozói adatok (első dolgozó)
    ws["A17"].value = name1
    ws["B17"].value = vorname1
    ws["C17"].value = ausweis1
    ws["E17"].value = beginn1
    ws["F17"].value = ende1
    ws["G17"].value = gesamtstunden

    wb.save("output.xlsx")
    return FileResponse("output.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Arbeitsnachweis.xlsx")

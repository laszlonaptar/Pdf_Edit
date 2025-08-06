from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
import os

app = FastAPI()

# Statikus fájlok mappa
app.mount("/static", StaticFiles(directory="static"), name="static")

# HTML sablon mappa
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate_excel/")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    bf: str = Form(...),
    was_wurde_gemacht: str = Form(...),
    vorname_1: str = Form(...),
    nachname_1: str = Form(...),
    ausweis_1: str = Form(...),
    beginn_1: str = Form(...),
    ende_1: str = Form(...),
    geraet: str = Form(""),
):
    try:
        # Excel sablon betöltése
        wb = load_workbook("GP-t.xlsx")
        ws = wb.active

        # Alapadatok beírása
        ws["E7"] = datum
        ws["E8"] = bau
        ws["E9"] = bf
        ws["B12"] = was_wurde_gemacht
        ws["B14"] = f"{vorname_1} {nachname_1}"
        ws["F14"] = ausweis_1
        ws["H14"] = beginn_1
        ws["I14"] = ende_1
        ws["B29"] = geraet

        # Munkaórák kiszámítása
        def parse_time(t: str) -> datetime:
            return datetime.strptime(t.strip(), "%H:%M")

        start = parse_time(beginn_1)
        end = parse_time(ende_1)
        pause = 0

        # Automatikus szünetlevonás
        if start <= datetime.strptime("09:00", "%H:%M") < end:
            pause += 0.25
        if start <= datetime.strptime("12:00", "%H:%
# FastAPI alkalmazás itt jön majd

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime, time
from io import BytesIO
from openpyxl import load_workbook
from typing import List

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


def calculate_hours(start: str, end: str) -> float:
    """Munkaidő számítás szünetek levonásával."""
    if not start or not end:
        return 0.0

    start_time = datetime.strptime(start, "%H:%M")
    end_time = datetime.strptime(end, "%H:%M")
    work_duration = (end_time - start_time).total_seconds() / 3600

    # Szünetek
    breaks = 0.0
    if start_time.time() <= time(9, 15) and end_time.time() >= time(9, 0):
        breaks += 0.25
    if start_time.time() <= time(12, 45) and end_time.time() >= time(12, 0):
        breaks += 0.75

    return max(work_duration - breaks, 0)


@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    taetigkeit: str = Form(...),
    vorhaltung: str = Form(""),
    name: List[str] = Form(...),
    vorname: List[str] = Form(...),
    ausweis: List[str] = Form(...),
    beginn: List[str] = Form(...),
    ende: List[str] = Form(...)
):
    # Excel sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Fejléc adatok
    ws["A1"] = datum
    ws["B2"] = projekt
    ws["G2"] = bf

    if vorhaltung:
        ws["A3"] = vorhaltung

    # Tevékenység szöveg tördelve A6-G15 cellák közé
    lines = taetigkeit.split("\n")
    for i, line in enumerate(lines[:10]):
        ws[f"A{6+i}"] = line

    # Dolgozói adatok feltöltése
    start_row = 17
    for i in range(len(name)):
        if not name[i]:
            continue
        ws[f"A{start_row + i}"] = name[i]
        ws[f"B{start_row + i}"] = vorname[i]
        ws[f"C{start_row + i}"] = ausweis[i]
        ws[f"D{start_row + i}"] = beginn[i]
        ws[f"E{start_row + i}"] = ende[i]

        total_hours = calculate_hours(beginn[i], ende[i])
        ws[f"F{start_row + i}"] = total_hours
        ws[f"G{start_row + i}"] = total_hours  # Gesamtstunden

    # Excel letöltés
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"arbeitsnachweis_{datum}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

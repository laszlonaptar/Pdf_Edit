from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from datetime import datetime, timedelta
import shutil
import os

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
async def serve_form():
    with open("static/index.html", encoding="utf-8") as f:
        return HTMLResponse(f.read())

@app.post("/generate")
async def generate_excel(
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    vorname1: str = Form(...),
    nachname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
    geraet: str = Form("")
):
    template_path = "GP-t.xlsx"
    filename = f"Arbeitsnachweis_{nachname1}_{datum}.xlsx"
    output_path = f"/tmp/{filename}"
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb.active

    ws["C4"] = datum
    ws["C5"] = projekt
    ws["K5"] = bf
    ws["C6"] = beschreibung
    ws["C34"] = geraet
    ws["B10"] = nachname1
    ws["C10"] = vorname1
    ws["D10"] = ausweis1
    ws["E10"] = beginn1
    ws["F10"] = ende1

    def parse_time(time_str):
        return datetime.strptime(time_str, "%H:%M")

    start = parse_time(beginn1)
    end = parse_time(ende1)
    break_time = timedelta(0)
    if start < datetime.strptime("09:15", "%H:%M") and end > datetime.strptime("09:00", "%H:%M"):
        break_time += timedelta(minutes=15)
    if start < datetime.strptime("12:45", "%H:%M") and end > datetime.strptime("12:00", "%H:%M"):
        break_time += timedelta(minutes=45)
    worked = end - start - break_time
    worked_hours = round(worked.total_seconds() / 3600, 2)
    ws["G10"] = worked_hours
    ws["H10"] = worked_hours

    wb.save(output_path)
    return FileResponse(output_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

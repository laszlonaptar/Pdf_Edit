from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import shutil
import uuid
import os
from openpyxl import load_workbook
from datetime import datetime

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_FILE = BASE_DIR / "GP-t.xlsx"
OUTPUT_DIR = BASE_DIR / "generated_excels"
OUTPUT_DIR.mkdir(exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def serve_form():
    with open("static/index.html", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bauort: str = Form(...),
    bf: str = Form(...),
    was: str = Form(...),
    geraet: str = Form(""),
    arbeiter_vn: list[str] = Form(...),
    arbeiter_nn: list[str] = Form(...),
    arbeiter_ausweis: list[str] = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
):
    excel = load_workbook(TEMPLATE_FILE)
    sheet = excel.active

    sheet["E6"] = datum
    sheet["E8"] = bauort
    sheet["E10"] = bf
    sheet["B14"] = was
    sheet["B22"] = geraet

    beginn_dt = datetime.strptime(beginn, "%H:%M")
    ende_dt = datetime.strptime(ende, "%H:%M")
    pause = 0
    if beginn_dt <= datetime.strptime("09:00", "%H:%M") <= ende_dt:
        pause += 0.25
    if beginn_dt <= datetime.strptime("12:00", "%H:%M") <= ende_dt:
        pause += 0.75
    hours = round((ende_dt - beginn_dt).seconds / 3600 - pause, 2)

    start_row = 14
    for i in range(len(arbeiter_vn)):
        row = start_row + i
        sheet[f"B{row}"] = arbeiter_nn[i]
        sheet[f"C{row}"] = arbeiter_vn[i]
        sheet[f"D{row}"] = arbeiter_ausweis[i]
        sheet[f"K{row}"] = beginn
        sheet[f"L{row}"] = ende
        sheet[f"M{row}"] = hours
        sheet[f"N{row}"] = hours

    output_file = OUTPUT_DIR / f"{uuid.uuid4()}.xlsx"
    excel.save(output_file)

    return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="arbeitsnachweis.xlsx")

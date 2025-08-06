from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from datetime import datetime
import os

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def serve_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    bauort: str = Form(...),
    bf: str = Form(...),
    was: str = Form(...),
    arbeiter_vn: str = Form(...),
    arbeiter_nn: str = Form(...),
    arbeiter_ausweis: str = Form(...),
    beginn: str = Form(...),
    ende: str = Form(...),
    geraet: str = Form(None)
):
    workbook = load_workbook("GP-t.xlsx")
    sheet = workbook.active

    datum = datetime.today().strftime("%d.%m.%Y")
    sheet["D4"] = datum
    sheet["D5"] = bauort
    sheet["D6"] = bf
    sheet["A10"] = was

    sheet["A13"] = arbeiter_nn
    sheet["B13"] = arbeiter_vn
    sheet["C13"] = arbeiter_ausweis
    sheet["D13"] = beginn
    sheet["E13"] = ende

    if geraet:
        sheet["A32"] = geraet

    output_filename = "generate_excel.xlsx"
    workbook.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_filename
    )

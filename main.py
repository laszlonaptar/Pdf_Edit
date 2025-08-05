
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import FileResponse
import openpyxl
from datetime import datetime
import os

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel/")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    geraet: str = Form(""),
    mitarbeiter1: str = Form(...),
    stunden1: str = Form(...),
):
    workbook = openpyxl.load_workbook("GP-t.xlsx")
    sheet = workbook.active

    sheet["C6"] = datum
    sheet["C7"] = projekt
    sheet["C8"] = geraet
    sheet["C11"] = mitarbeiter1
    sheet["L11"] = float(stunden1.replace(",", "."))

    today = datetime.today().strftime("%Y-%m-%d")
    filename = f"arbeitsnachweis_{today}.xlsx"
    workbook.save(filename)

    return FileResponse(path=filename, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

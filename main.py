
from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.cors import CORSMiddleware
import openpyxl
import shutil

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.mount("/", StaticFiles(directory=".", html=True), name="static")

@app.post("/generate")
async def generate(
    datum: str = Form(...),
    ort: str = Form(...),
    beschreibung: str = Form(...),
    name_0: str = Form(None), vorname_0: str = Form(None), ausweis_0: str = Form(None), stunden_0: str = Form(None),
    name_1: str = Form(None), vorname_1: str = Form(None), ausweis_1: str = Form(None), stunden_1: str = Form(None),
    name_2: str = Form(None), vorname_2: str = Form(None), ausweis_2: str = Form(None), stunden_2: str = Form(None),
    name_3: str = Form(None), vorname_3: str = Form(None), ausweis_3: str = Form(None), stunden_3: str = Form(None),
    name_4: str = Form(None), vorname_4: str = Form(None), ausweis_4: str = Form(None), stunden_4: str = Form(None),
):
    shutil.copy("arbeitsnachweis_leeres_layout.xls", "arbeitsnachweis_filled.xls")
    wb = openpyxl.load_workbook("arbeitsnachweis_filled.xls")
    ws = wb.active

    ws["A1"] = datum
    ws["A2"] = ort
    ws["A3"] = beschreibung

    mitarbeiter = [
        (name_0, vorname_0, ausweis_0, stunden_0),
        (name_1, vorname_1, ausweis_1, stunden_1),
        (name_2, vorname_2, ausweis_2, stunden_2),
        (name_3, vorname_3, ausweis_3, stunden_3),
        (name_4, vorname_4, ausweis_4, stunden_4),
    ]

    for i, (name, vorname, ausweis, stunden) in enumerate(mitarbeiter):
        base_row = 5 + i
        if name: ws[f"A{base_row}"] = name
        if vorname: ws[f"B{base_row}"] = vorname
        if ausweis: ws[f"C{base_row}"] = ausweis
        if stunden: ws[f"D{base_row}"] = stunden

    wb.save("arbeitsnachweis_filled.xls")
    return FileResponse("arbeitsnachweis_filled.xls", media_type="application/vnd.ms-excel", filename="arbeitsnachweis.xls")

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import io
from datetime import datetime, timedelta
import openpyxl

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate", response_class=StreamingResponse)
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bauort: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    geraet: str = Form(""),
    arbeiter_vorname: list[str] = Form(...),
    arbeiter_nachname: list[str] = Form(...),
    arbeiter_ausweis: list[str] = Form(...),
    arbeiter_beginn: list[str] = Form(...),
    arbeiter_ende: list[str] = Form(...),
    gesamtstunden: str = Form(...)
):
    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    ws["D4"] = datum
    ws["D5"] = bauort
    ws["D6"] = bf
    ws["D7"] = beschreibung
    ws["D23"] = gesamtstunden
    ws["D24"] = geraet

    for i in range(len(arbeiter_vorname)):
        ws[f"A{10 + i}"] = arbeiter_nachname[i]
        ws[f"B{10 + i}"] = arbeiter_vorname[i]
        ws[f"C{10 + i}"] = arbeiter_ausweis[i]
        ws[f"D{10 + i}"] = arbeiter_beginn[i]
        ws[f"E{10 + i}"] = arbeiter_ende[i]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"arbeitsnachweis_{datum}.xlsx"
    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

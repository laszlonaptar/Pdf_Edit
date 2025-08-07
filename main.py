from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    datum: str = Form(...),
    bauort: str = Form(...),
    bf: str = Form(...),
    taetigkeit: str = Form(...),
    vorname1: str = Form(...),
    nachname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
    total_hours: str = Form(None)
):
    # sablon betöltés
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # DÁTUM/BF/BAUORT: mindhárom helyen a tartomány bal felső cellájába írjunk
    # Általános szabály: mindig a merged range top-left celláját írd
    # — itt példaként a következő pozíciókat használjuk:
    #   Datum: B2   Bauort: B3   BF: F2  (ha ezek máshol vannak a sablonban, állítsd át)
    ws["B2"].value = datum
    ws["B3"].value = bauort
    ws["F2"].value = bf

    # Napi leírás: A6:G15 összevont blokk bal felső cellája A6
    ws["A6"].value = taetigkeit
    ws["A6"].alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")

    # Személy sor: (mintaként) név mezők a táblázat megfelelő oszlopaiba
    # Itt állítsd a saját sablonodhoz a cellákat:
    ws["B18"].value = nachname1
    ws["C18"].value = vorname1
    ws["D18"].value = ausweis1
    ws["E18"].value = beginn1
    ws["F18"].value = ende1
    if total_hours:
        ws["G18"].value = total_hours

    # Ha van összesített óramező a lapon, töltsük ki (pl. B24)
    if total_hours:
        ws["B24"].value = total_hours

    # excel visszaküldése
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"Arbeitsnachweis_{datum}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        },
    )

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import openpyxl
from io import BytesIO

app = FastAPI()
app.mount("/static", StaticFiles(directory="."), name="static")
templates = Jinja2Templates(directory=".")

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(request: Request):
    form = await request.form()
    data = form._dict
    datum = data.get("datum")
    bauort = data.get("bauort")
    bauleiter = data.get("bauleiter")
    beschreibung = data.get("beschreibung")
    geraet = data.get("geraet", "")
    gesamtstunden = data.get("gesamtstunden")

    arbeiter = []
    i = 1
    while True:
        if f"vorname{i}" in data:
            arbeiter.append({
                "vorname": data.get(f"vorname{i}"),
                "nachname": data.get(f"nachname{i}"),
                "ausweis": data.get(f"ausweis{i}"),
                "beginn": data.get(f"beginn{i}"),
                "ende": data.get(f"ende{i}")
            })
            i += 1
        else:
            break

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nachweis"
    ws.append(["Datum", datum])
    ws.append(["Bauort", bauort])
    ws.append(["Bauleiter", bauleiter])
    ws.append(["Beschreibung", beschreibung])
    ws.append(["Vorhaltung/Gerät", geraet])
    ws.append(["Gesamtstunden", gesamtstunden])
    ws.append([])
    ws.append(["Arbeiter", "Vorname", "Nachname", "Ausweis-Nr.", "Beginn", "Ende"])

    for idx, a in enumerate(arbeiter, 1):
        ws.append([idx, a["vorname"], a["nachname"], a["ausweis"], a["beginn"], a["ende"]])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename="arbeitsnachweis.xlsx")

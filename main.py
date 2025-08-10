from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
import uvicorn

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

def fix_beschreibung_block(ws):
    """
    A6–G15 blokk: engedjük a sortörést, balra + felülre igazítunk,
    és megemeljük a sorok magasságát, hogy ne vágódjon le a 2. sor.
    """
    top_left = ws.cell(row=6, column=1)  # A6
    top_left.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')

    for r in range(6, 16):  # sorok 6..15
        ws.row_dimensions[r].height = 22  # állítható: 22 → 24–26, ha még vág

@app.get("/")
async def form_get(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    bf: str = Form(...),
    beschreibung: str = Form(...),
    name1: str = Form(...),
    vorname1: str = Form(...),
    ausweis1: str = Form(...),
    beginn1: str = Form(...),
    ende1: str = Form(...),
    gesamtstunden1: str = Form(...),
):
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Adatok beírása
    ws["B2"] = datum
    ws["E2"] = bau
    ws["G2"] = bf
    ws["A6"] = beschreibung

    ws["A17"] = name1
    ws["B17"] = vorname1
    ws["C17"] = ausweis1
    ws["E17"] = beginn1
    ws["F17"] = ende1
    ws["G17"] = gesamtstunden1

    # Csak a leírás-blokk javítása
    fix_beschreibung_block(ws)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"arbeitsnachweis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)

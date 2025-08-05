from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
import shutil
import os

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload_file(
    request: Request,
    file: UploadFile = File(...)
):
    try:
        contents = await file.read()
        filename = "GP-t_filled.xlsx"
        with open("GP-t.xlsx", "wb") as f:
            f.write(contents)

        wb = load_workbook("GP-t.xlsx")
        ws = wb.active

        # Példa: első mezőbe minta adat
        ws["B6"] = "2025.08.05"
        ws["D6"] = "M715"
        ws["G6"] = "John Doe"
        ws["J6"] = "8"
        ws["L6"] = "8"

        wb.save(filename)
        return FileResponse(filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=filename)
    except Exception as e:
        import traceback
        print("❌ Internal Server Error:", str(e))
        traceback.print_exc()
        return HTMLResponse(content="Internal Server Error: " + str(e), status_code=500)

from fastapi import FastAPI, Form
from fastapi.responses import FileResponse
from openpyxl import load_workbook
import shutil

app = FastAPI()

@app.post("/generate_excel/")
async def generate_excel(
    datum: str = Form(...),
    projekt: str = Form(...),
    geraet: str = Form(""),
    mitarbeiter1: str = Form(...),
    stunden1: float = Form(...)
):
    try:
        shutil.copyfile("GP-t.xlsx", "GP-t_filled.xlsx")
        wb = load_workbook("GP-t_filled.xlsx")
        ws = wb.active

        # Cellák kitöltése
        ws["B6"] = datum
        ws["D6"] = projekt
        ws["F6"] = geraet
        ws["G6"] = mitarbeiter1
        ws["J6"] = stunden1
        ws["L6"] = stunden1

        wb.save("GP-t_filled.xlsx")
        return FileResponse("GP-t_filled.xlsx", media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="GP-t_filled.xlsx")

    except Exception as e:
        return {"error": str(e)}

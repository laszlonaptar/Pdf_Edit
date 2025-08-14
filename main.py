from fastapi import FastAPI, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import io
import os
from datetime import datetime

app = FastAPI()

# Statikus mappák és sablonok
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_pdf")
async def generate_pdf(
    request: Request,
    datum: str = Form(...),
    projekt: str = Form(...),
    bf: str = Form(...),
    taetigkeiten: str = Form(...)
):
    buffer = io.BytesIO()

    # Oldal beállítása
    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4

    # Háttérkép betöltése
    img_path = os.path.join("static", "arbeitsnachweis_bg.jpg")
    img = ImageReader(img_path)

    img_width, img_height = img.getSize()

    # Arány megtartása
    aspect = img_height / float(img_width)
    new_width = page_width
    new_height = new_width * aspect

    # Kép pozíció (középen + 25px jobbra tolás)
    x_pos = 0
    y_pos = page_height - new_height
    c.drawImage(
        img,
        x_pos + 25,  # 25px jobbra tolás
        y_pos,
        width=new_width,
        height=new_height
    )

    # Betűtípus
    c.setFont("Helvetica", 10)

    # Dátum
    c.drawString(470, 550, datum)

    # Projekt
    c.drawString(100, 520, projekt)

    # BF
    c.drawString(470, 520, bf)

    # Tevékenység leírás (több sorba tördelve, ha kell)
    text_obj = c.beginText()
    text_obj.setTextOrigin(40, 480)
    text_obj.setFont("Helvetica", 10)
    for line in taetigkeiten.split("\n"):
        text_obj.textLine(line)
    c.drawText(text_obj)

    c.showPage()
    c.save()

    buffer.seek(0)

    filename = f"Arbeitsnachweis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    return FileResponse(
        buffer,
        media_type='application/pdf',
        filename=filename
    )

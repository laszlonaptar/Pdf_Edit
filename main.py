from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Betűtípus regisztrálása
pdfmetrics.registerFont(TTFont('Helvetica', 'Helvetica.ttf'))

# Állandó eltolás
OFFSET_X = 25  # px eltolás jobbra
OFFSET_Y = 0   # ha felfelé/lefelé kellene tolni, itt állítható

def generate_pdf(output_filename, background_image, text_data):
    c = canvas.Canvas(output_filename, pagesize=A4)
    width, height = A4

    # Háttérkép betöltése és rajzolása eltolással
    img = ImageReader(background_image)
    img_width, img_height = img.getSize()

    # A képet méretarányosan skálázzuk az A4 méretre
    scale_x = width / img_width
    scale_y = height / img_height
    scale = min(scale_x, scale_y)

    new_width = img_width * scale
    new_height = img_height * scale

    # Bal alsó sarok pozíciója eltolással
    x_pos = ((width - new_width) / 2) + OFFSET_X
    y_pos = ((height - new_height) / 2) + OFFSET_Y

    c.drawImage(img, x_pos, y_pos, width=new_width, height=new_height)

    # Szöveg rajzolása – minden X koordinátához hozzáadjuk az OFFSET_X-et
    c.setFont("Helvetica", 10)

    for item in text_data:
        text = item['text']
        x = item['x'] + OFFSET_X
        y = item['y'] + OFFSET_Y
        c.drawString(x, y, text)

    c.showPage()
    c.save()

# Tesztadatok
text_data = [
    {"text": "Bau und Ausführungsort: M715", "x": 100, "y": 700},
    {"text": "Datum: 14.08.2025", "x": 400, "y": 700},
    {"text": "Lorem ipsum dolor sit amet", "x": 100, "y": 650}
]

generate_pdf(
    output_filename="output.pdf",
    background_image="background.jpg",  # háttérkép elérési útja
    text_data=text_data
)

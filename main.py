from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

from openpyxl import load_workbook
from openpyxl.styles import Alignment

from datetime import datetime, time, timedelta
from io import BytesIO
from typing import Dict, Tuple, Optional

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


# --------- kis segédek ---------

def _norm(val: Optional[str]) -> str:
    if val is None:
        return ""
    return " ".join(str(val).strip().lower().split())


def is_in_range(r: int, c: int, rng) -> bool:
    return (rng.min_row <= r <= rng.max_row) and (rng.min_col <= c <= rng.max_col)


def merged_top_left(ws, r: int, c: int) -> Tuple[int, int]:
    """Ha (r, c) egy összevont cellában van, add vissza a blokk bal-felső sarkát, különben (r, c)."""
    for rng in ws.merged_cells.ranges:
        if is_in_range(r, c, rng):
            return rng.min_row, rng.min_col
    return r, c


def set_text(ws, r: int, c: int, text: str, wrap: bool = False, align_left: bool = True):
    r0, c0 = merged_top_left(ws, r, c)
    cell = ws.cell(r0, c0)
    cell.value = text
    cell.alignment = Alignment(
        wrap_text=wrap,
        horizontal="left" if align_left else cell.alignment.horizontal or "center",
        vertical="top"
    )


def merged_blocks_in_row(ws, row: int):
    return sorted(
        [rng for rng in ws.merged_cells.ranges if rng.min_row <= row <= rng.max_row],
        key=lambda r: r.min_col
    )


def right_block_top_left_after_label(ws, row: int, label_text: str) -> Optional[Tuple[int, int]]:
    """Megkeresi a soron belül azt az összevont blokkot, amelynek KÖVETKEZŐ blokkjába kell írni."""
    blocks = merged_blocks_in_row(ws, row)
    norm_label = _norm(label_text)
    for i, rng in enumerate(blocks):
        lbl = _norm(ws.cell(rng.min_row, rng.min_col).value)
        if lbl == norm_label:
            if i + 1 < len(blocks):
                nxt = blocks[i + 1]
                return nxt.min_row, nxt.min_col
            # ha nincs következő blokk, próbálkozzunk a label jobb melletti cellával
            return row, rng.max_col + 1
    return None


def right_value_cell_of_label(ws, label_text: str) -> Optional[Tuple[int, int]]:
    """Végigmegyünk a sorokon, és ahol megtaláljuk a labelt, visszaadjuk a tőle jobbra lévő blokk bal-felső celláját."""
    for r in range(1, ws.max_row + 1):
        res = right_block_top_left_after_label(ws, r, label_text)
        if res:
            return res
    return None


def normalize_header(val: Optional[str]) -> str:
    v = _norm(val)
    # pár tipikus header egységesítése
    replacements = {
        "einsatzzeit [hh:mm]": "einsatzzeit",
        "anzahl stunden (ohne pausen)": "anzahl stunden",
    }
    return replacements.get(v, v)


def find_table_headers(ws) -> Tuple[int, Dict[str, int]]:
    """
    Megkeresi a Name / Vorname / Ausweis / Beginn / Ende / Anzahl Stunden oszlopokat.
    A korábbi walrus-os megoldás helyett explicit if-eket használunk (Render barát).
    """
    for r in range(1, ws.max_row + 1):
        cols: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            val = normalize_header(ws.cell(r, c).value)
            if not val:
                continue
            if val == "name":
                cols["name"] = c
            if "vorname" in val:
                cols["vorname"] = c
            if "ausweis" in val or "kennzeichen" in val:
                cols["ausweis"] = c
            if "einsatzzeit" in val:
                # ez a blokk felett van a Beg/Ende feliratpár
                # a következő sorban lesz "Beginn" és "Ende"
                # de a tényleges oszlopokat lentebb fogjuk megtalálni külön
                pass
            if "beginn" in val:
                cols["beginn"] = c
            if "ende" in val:
                cols["ende"] = c
            if "anzahl stunden" in val:
                cols["stunden"] = c

        if {"name", "vorname", "ausweis", "beginn", "ende", "stunden"}.issubset(cols.keys()):
            return r, cols

    raise RuntimeError("Fejléc sor nem található a táblában.")


def parse_hhmm(s: str) -> time:
    s = s.strip().replace(" ", "")
    if not s:
        return time(0, 0)
    h, m = s.split(":")
    return time(int(h), int(m))


def td_overlap(a_start: time, a_end: time, b_start: time, b_end: time) -> timedelta:
    def to_minutes(t: time) -> int:
        return t.hour * 60 + t.minute
    a1, a2 = to_minutes(a_start), to_minutes(a_end)
    b1, b2 = to_minutes(b_start), to_minutes(b_end)
    inter = max(0, min(a2, b2) - max(a1, b1))
    return timedelta(minutes=inter)


def worked_hours_with_breaks(beg: time, end: time) -> float:
    """Számol: reggeli 09:00–09:15 (0,25h) + ebéd 12:00–12:45 (0,75h) automatikus levonás, ha belelóg."""
    m = (end.hour * 60 + end.minute) - (beg.hour * 60 + beg.minute)
    if m <= 0:
        return 0.0
    total = timedelta(minutes=m)

    total -= td_overlap(beg, end, time(9, 0), time(9, 15))
    total -= td_overlap(beg, end, time(12, 0), time(12, 45))

    return round(total.total_seconds() / 3600.0, 2)


# --------- útvonalak ---------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate_excel")
async def generate_excel(
    request: Request,
    datum: str = Form(...),
    bau: str = Form(...),
    basf_beauftragter: str = Form(""),
    geraet: str = Form(""),
    beschreibung: str = Form(""),
    vorname1: str = Form(""),
    nachname1: str = Form(""),
    ausweis1: str = Form(""),
    beginn1: str = Form(""),
    ende1: str = Form(""),
):
    """
    A form mezőit az Excel sablonba írja.
    A sablon fájl neve: GP-t.xlsx (a repo gyökerében).
    """
    # sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # 1) Fejléc mezők – mindig a label UTÁNI (jobbra lévő) blokkba írunk
    if datum:
        pos = right_value_cell_of_label(ws, "datum der leistungsausführung:")
        if pos:
            set_text(ws, pos[0], pos[1], datum, wrap=False)

    if bau:
        pos = right_value_cell_of_label(ws, "bau und ausführungsort:")
        if pos:
            set_text(ws, pos[0], pos[1], bau, wrap=False)

    if basf_beauftragter:
        # Ezt is a saját labelje UTÁN írjuk
        pos = right_value_cell_of_label(ws, "basf-beauftragter, org.-code:")
        if pos:
            set_text(ws, pos[0], pos[1], basf_beauftragter, wrap=False)

    # 2) Leírás – megkeressük a legszélesebb összevont sávot a címblokk alatt és oda írunk
    if beschreibung:
        best = None
        for rng in ws.merged_cells.ranges:
            width = rng.max_col - rng.min_col
            # heur: a címsorok alatt, elég széles sáv (jegyzet mező)
            if rng.min_row >= 5 and width >= 6:
                if best is None or width > (best.max_col - best.min_col):
                    best = rng
        if best:
            set_text(ws, best.min_row, best.min_col, beschreibung, wrap=True, align_left=True)
            # adunk egy kicsi magasságot az első sorának, hogy látszódjon a több sor
            ws.row_dimensions[best.min_row].height = 45

    # 3) Táblázat fejléc -> oszlopindexek
    try:
        header_row, cols = find_table_headers(ws)
    except Exception:
        # ha valamiért nem találjuk, hagyjuk ki a táblázat kitöltését (ne dobjuk el a generálást)
        cols = {}
        header_row = None

    total_hours = 0.0

    if cols:
        # az első adat sor rendszerint a header UTÁNI sor
        r = header_row + 1

        # 1. dolgozó (később bővíthető 2..5-ig ugyanígy)
        if (vorname1 or nachname1 or ausweis1 or beginn1 or ende1):
            if nachname1:
                set_text(ws, r, cols["name"], nachname1, wrap=False)
            if vorname1:
                set_text(ws, r, cols["vorname"], vorname1, wrap=False)
            if ausweis1:
                set_text(ws, r, cols["ausweis"], ausweis1, wrap=False)

            if beginn1:
                set_text(ws, r, cols["beginn"], beginn1, wrap=False)
            if ende1:
                set_text(ws, r, cols["ende"], ende1, wrap=False)

            try:
                bh = parse_hhmm(beginn1)
                eh = parse_hhmm(ende1)
                hrs = worked_hours_with_breaks(bh, eh)
            except Exception:
                hrs = 0.0

            total_hours += hrs
            if "stunden" in cols:
                set_text(ws, r, cols["stunden"], f"{hrs:.2f}", wrap=False)

        # összesített órák – megkeressük a „Gesamtstunden” felirat melletti cellát
        pos_total = right_value_cell_of_label(ws, "gesamtstunden")
        if pos_total:
            set_text(ws, pos_total[0], pos_total[1], f"{total_hours:.2f}", wrap=False)

    # 4) Mentés és válasz
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    headers = {
        "Content-Disposition": f'attachment; filename="leistungsnachweis_{datetime.utcnow().strftime("%Y%m%d_%H%M%S")}.xlsx"'
    }
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

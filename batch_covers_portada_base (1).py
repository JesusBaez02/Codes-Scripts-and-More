
#!/usr/bin/env python3
"""
Inserta portada base (PDF) con MES y TERRITORIO encima para cada PDF en `input/`.
Tamaño de página: LETTER (8.5 x 11 in).

Requisitos:
    pip install pypdf reportlab
"""
import re
import sys
from pathlib import Path
from typing import Optional, Dict, Tuple

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PORTADA_BASE = Path("portada_base.pdf")
MAPPING_CSV = Path("mapping.csv")  # opcional: nombre_sin_ext,Mes[,Territorio]

# --- Página en tamaño carta (LETTER) ---
PAGE_W, PAGE_H = letter  # (612.0, 792.0) puntos

# Posiciones de texto (en puntos). Ajusta a tu gusto.
MES_POS = (PAGE_W/2-150, PAGE_H - 7.4*72)      # centrado horizontal, a la altura 8 pulgadas desde el borde inferior
TERR_POS = (PAGE_W/2-160, PAGE_H - 7.8*72 - 40)

MES_FONT = ("Helvetica", 18)
TERR_FONT = ("Helvetica-Bold", 21)

NAME_TO_NUM = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"setiembre":9,"octubre":10,
    "noviembre":11,"diciembre":12,
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,
    "sep":9,"oct":10,"nov":11,"dic":12
}
NUM_TO_NAME = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}

MONTH_PATTERNS = [
    re.compile(r"(?:^|[_\-\s])(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre|ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)(?:[_\-\s]|$)", re.IGNORECASE),
    re.compile(r"(?:^|[_\-\s])(0?[1-9]|1[0-2])(?:[_\-\s]|$)"),
]

def parse_mapping_csv(path: Path) -> Dict[str, Tuple[Optional[str], Optional[str]]]:
    mapping = {}
    if not path.exists():
        return mapping
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) >= 2:
            name = parts[0]
            mes = parts[1] if parts[1] else None
            terr = parts[2] if len(parts) >= 3 and parts[2] else None
            mapping[name] = (mes, terr)
    return mapping

def detect_month_from_filename(filename: str) -> Optional[int]:
    stem = Path(filename).stem
    m = MONTH_PATTERNS[0].search(stem)
    if m:
        token = m.group(1).lower()
        return NAME_TO_NUM.get(token)
    m = MONTH_PATTERNS[1].search(stem)
    if m:
        val = int(m.group(1))
        return val if 1 <= val <= 12 else None
    return None

def detect_territory_from_filename(filename: str) -> Optional[str]:
    """
    Detecta territorio tomando el segmento antes del MES, usando '_' como separador principal.
    Ej.: ..._SUR_DICIEMBRE  -> 'SUR'
         ..._NORTE_12       -> 'NORTE'
    """
    stem = Path(filename).stem

    # 1) Divide por '_' primero (prioridad), y limpia espacios
    parts = [p.strip() for p in stem.split('_') if p.strip()]
    if len(parts) < 2:
        return None

    # 2) Localiza el índice del mes entre esas partes (por nombre o número)
    def is_month_token(tok: str) -> bool:
        t = tok.lower()
        if t.isdigit() and 1 <= int(t) <= 12:
            return True
        return t in NAME_TO_NUM  # nombres y abreviaturas (ene..dic)

    month_idx = None
    for i, p in enumerate(parts):
        if is_month_token(p):
            month_idx = i
            break

    # 3) Si encontró mes y hay un elemento previo, ese es el territorio
    if month_idx is not None and month_idx - 1 >= 0:
        return parts[month_idx - 1]

    # 4) Fallback: si no detectó mes, intenta que el último sea el mes y el anterior territorio
    if len(parts) >= 2 and is_month_token(parts[-1]):
        return parts[-2]

    return None


def normalize_month_token(token: Optional[str]) -> Optional[int]:
    if token is None:
        return None
    t = token.strip().lower()
    if t.isdigit():
        mm = int(t)
        return mm if 1 <= mm <= 12 else None
    return NAME_TO_NUM.get(t)

def month_text(month_num: int) -> str:
    return NUM_TO_NAME.get(month_num, "Mes")

def load_mapping_for(stem: str, mapping: Dict[str, Tuple[Optional[str], Optional[str]]]) -> Tuple[Optional[int], Optional[str]]:
    if stem not in mapping:
        return (None, None)
    mes_str, terr = mapping[stem]
    return (normalize_month_token(mes_str), terr)

def create_overlay(month_str: str, territory_str: Optional[str], overlay_path: Path):
    c = canvas.Canvas(str(overlay_path), pagesize=letter)
    c.setFont(MES_FONT[0], MES_FONT[1])
    c.drawCentredString(MES_POS[0], MES_POS[1], month_str)
    if territory_str:
        c.setFont(TERR_FONT[0], TERR_FONT[1])
        c.drawCentredString(TERR_POS[0], TERR_POS[1],f"{territory_str}")
    c.save()

def make_cover_with_overlay(base_pdf: Path, overlay_pdf: Path, out_pdf: Path):
    base_reader = PdfReader(str(base_pdf))
    overlay_reader = PdfReader(str(overlay_pdf))
    writer = PdfWriter()
    page = base_reader.pages[0]
    page.merge_page(overlay_reader.pages[0])
    writer.add_page(page)
    with out_pdf.open("wb") as f:
        writer.write(f)

def prepend_cover_to_pdf(cover_pdf: Path, src_pdf: Path, dest_pdf: Path):
    writer = PdfWriter()
    cover_reader = PdfReader(str(cover_pdf))
    for page in cover_reader.pages:
        writer.add_page(page)
    reader = PdfReader(str(src_pdf))
    for page in reader.pages:
        writer.add_page(page)
    with dest_pdf.open("wb") as f:
        writer.write(f)

def main():
    if not PORTADA_BASE.exists():
        print(f"[Error] No se encontró la portada base: {PORTADA_BASE}")
        sys.exit(1)
    mapping = parse_mapping_csv(MAPPING_CSV)
    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdfs:
        print(f"[Info] No hay PDFs en: {INPUT_DIR.resolve()}")
        sys.exit(0)
    for pdf in pdfs:
        stem = pdf.stem
        month_num = detect_month_from_filename(pdf.name)
        territory = detect_territory_from_filename(pdf.name)
        map_month, map_terr = load_mapping_for(stem, mapping)
        if month_num is None and map_month is not None:
            month_num = map_month
        if (territory is None or territory == "") and map_terr:
            territory = map_terr
        if month_num is None:
            print(f"[Aviso] No se detectó MES para '{pdf.name}'. Saltando...")
            continue
        month_str = month_text(month_num)
        overlay_path = OUTPUT_DIR / "__overlay.pdf"
        portada_temp = OUTPUT_DIR / "__portada_temp.pdf"
        create_overlay(month_str, territory, overlay_path)
        make_cover_with_overlay(PORTADA_BASE, overlay_path, portada_temp)
        out_pdf = OUTPUT_DIR / pdf.name
        prepend_cover_to_pdf(portada_temp, pdf, out_pdf)
        try:
            overlay_path.unlink()
            portada_temp.unlink()
        except:
            pass
        print(f"OK -> {out_pdf.name}")

if __name__ == "__main__":
    main()

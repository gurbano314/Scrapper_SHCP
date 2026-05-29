# ════════════════════════════════════════════════════════════════════
# CONCILIADOR DE COTIZACIONES PDF — v2.0
# Streamlit · Compatible Streamlit Cloud + Local
#
# Mejoras sobre v1:
#   • OCR solo en páginas que realmente lo necesitan (calidad de texto)
#   • Extracción de total con contexto (no greedy-max) → evita confundir
#     el presupuesto global del proyecto con el total de la cotización
#   • IVA incluido: detección de frases "precios incluyen IVA" + back-calc
#   • Integración API Banxico FIX/EUR para cotizaciones en moneda extranjera
#   • Salida Excel con openpyxl: formato exacto del template (colores,
#     anchos, fórmulas recalculables) → editable directo en Excel
#   • Session state desacoplado: sin reseteos fantasma al cambiar secciones
# ════════════════════════════════════════════════════════════════════

from __future__ import annotations

import datetime
import hashlib
import io
import math
import re
import os
import requests
from typing import Optional

import fitz                   # PyMuPDF
import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Side)
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliador de Cotizaciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { background: #1a2744; }
[data-testid="stSidebar"] * { color: #e8e8e8 !important; }
.hdr { background: linear-gradient(90deg,#1a2744,#2d4a8f);
       padding: 14px 22px; border-radius: 8px; margin-bottom: 18px; }
.hdr h1 { color: #fff; font-size: 1.4rem; margin: 0; font-weight: 700; }
.hdr p  { color: #a8c0ff; font-size: .87rem; margin: 4px 0 0; }
.ptitle { background: #2d4a8f; color: #fff !important; font-weight: 700;
          font-size: .88rem; padding: 7px 14px; border-radius: 6px 6px 0 0; margin-bottom: 4px; }
.kpi  { background: #f0f4ff; border: 1px solid #c5d3f5;
        border-radius: 8px; padding: 10px 12px; text-align: center; margin-bottom: 6px; }
.kpi .v { font-size: 1.2rem; font-weight: 700; color: #1a2744; }
.kpi .l { font-size: .72rem; color: #5566aa; }
.tc-box { background: #e8f4fd; border: 1px solid #90caf9; border-radius: 6px;
          padding: 8px 12px; font-size: .85rem; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────
COLS = [
    "Fecha", "Rubro", "QT", "T. Cambio", "(+ IVA)", "Cantidad",
    "Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%", "Total con IVA",
    "Diferencia final", "Monto en Anexo Escrito", "Observaciones",
]

COL_WIDTHS = [12, 51.7, 5.4, 10.3, 7.7, 6, 16.9, 18.6, 17.4, 16.4, 18, 22, 48.4]

MONEY_FMT  = '_-"$ "* #,##0.00_-;\\-"$ "* #,##0.00_-;_-"$ "* "-"??_-;_-@_-'
DATE_FMT   = "mm-dd-yy"

# Banxico SIE series for exchange rates (FIX)
BANXICO_SERIES: dict[str, str] = {
    "USD": "SF43718",
    "EUR": "SF46410",
    "CAD": "SF57771",
    "GBP": "SF46406",
    "JPY": "SF46407",
    "CHF": "SF46408",
}

_MESES: dict[str, int] = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    # English fallback (for OCR misreads)
    "january": 1, "february": 2, "march": 3, "april": 4,
    "june": 6, "july": 7, "august": 8, "september": 9,
    "october": 10, "november": 11, "december": 12,
    # Common abbrevs
    "ene": 1, "feb": 2, "mar": 3, "abr": 4,
    "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
}

# ──────────────────────────────────────────────────────────────
# REGEX PATTERNS
# ──────────────────────────────────────────────────────────────

# Money: captures $1,234.56 or 1,234,567.89 (requires comma OR $-prefix for safety)
_RE_MONEY = re.compile(
    r"\$\s*([\d]{1,3}(?:,\d{3})*(?:\.\d{1,2})?)"          # $1,234.56  ($ prefix)
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{0,2})?)"        # 1,234.56   (comma-grouped, no $ needed)
)

# Date: "26 de marzo de 2026", "2026-03-26", "26/03/2026", "26/03/26"
_RE_DATE = re.compile(
    r"(\d{1,2})\s*(?:de\s+)?([a-záéíóúüñ]{3,})\s*(?:de(?:l)?\s+)?(\d{4})"  # Spanish
    r"|(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})"                           # ISO
    r"|(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{2,4})",                        # DMY
    re.IGNORECASE,
)

# Section-total keywords — deliberately excludes "presupuesto total", "costo total del proyecto"
_RE_TOTAL = re.compile(
    r"\b(total|importe\s+total|gran\s+total)\b"
    r"(?!\s*(?:presupuest|del\s+proy|estimad|aprox))",
    re.IGNORECASE,
)

_RE_SUBTOTAL  = re.compile(r"\b(subtotal|sub[\s\-]?total|sin\s+iva|importe)\b", re.IGNORECASE)
_RE_IVA_LINE  = re.compile(r"\biva\b|\b16\s*%|impuesto\s+al\s+valor\s+agregado", re.IGNORECASE)
_RE_IVA_INC   = re.compile(
    r"(?:precios?|tarifas?|montos?)\s+(?:ya\s+)?inclu(?:ye[ns]?|ido[s]?)\s+(?:el\s+)?iva"
    r"|iva\s+inclu(?:ido|ye|yendo)",
    re.IGNORECASE,
)

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _f(v) -> Optional[float]:
    """Safe float conversion; returns None for None, nan, inf."""
    if v is None:
        return None
    try:
        x = float(str(v).replace(",", "").strip())
        return None if (math.isnan(x) or math.isinf(x)) else x
    except (ValueError, TypeError):
        return None


def _parse_money(text: str) -> list[float]:
    """Extract all positive monetary values from text, descending."""
    out: list[float] = []
    for m in _RE_MONEY.finditer(str(text)):
        raw = m.group(1) or m.group(2)
        if not raw:
            continue
        try:
            v = float(raw.replace(",", ""))
            if v > 0:
                out.append(v)
        except ValueError:
            pass
    return sorted(set(out), reverse=True)


def _parse_date(text: str) -> Optional[datetime.date]:
    """Extract first parseable date from text."""
    for m in _RE_DATE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]:                                   # Spanish: day month year
                mes = _resolve_month(g[1].strip())
                if mes:
                    return datetime.date(int(g[2]), mes, int(g[0]))
            elif g[3]:                                 # ISO: YYYY-MM-DD
                return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            elif g[6]:                                 # DMY
                yr = int(g[8])
                return datetime.date(yr + 2000 if yr < 100 else yr, int(g[7]), int(g[6]))
        except (ValueError, TypeError):
            pass
    return None


def _resolve_month(s: str) -> Optional[int]:
    s = s.lower().strip()
    if s in _MESES:
        return _MESES[s]
    # Partial / OCR-corrupted month names
    for k, v in _MESES.items():
        if len(s) >= 3 and (s.startswith(k[:3]) or k.startswith(s[:3])):
            return v
    return None


def _text_quality(text: str) -> float:
    """Ratio of alphanumeric chars (0–1). < 0.12 → likely scanned."""
    if not text:
        return 0.0
    return sum(1 for c in text if c.isalnum()) / len(text)


# ──────────────────────────────────────────────────────────────
# OCR ENGINE — lazy, multi-backend, cached
# ──────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def _ocr_engine():
    """Return (engine_name, engine) for the first available OCR backend."""
    # 1. RapidOCR (lightweight ONNX)
    try:
        from rapidocr_onnxruntime import RapidOCR
        return "rapidocr", RapidOCR()
    except Exception:
        pass
    # 2. EasyOCR
    try:
        import easyocr
        return "easyocr", easyocr.Reader(["es", "en"], gpu=False, verbose=False)
    except Exception:
        pass
    # 3. Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return "tesseract", pytesseract
    except Exception:
        pass
    return None, None


def _run_ocr(img_bgr: np.ndarray) -> str:
    """Run OCR on a BGR numpy array, returning extracted text."""
    name, engine = _ocr_engine()
    if name is None:
        return ""
    try:
        if name == "rapidocr":
            result, _ = engine(img_bgr)
            return "\n".join(r[1] for r in (result or []) if r and len(r) > 1)
        if name == "easyocr":
            return "\n".join(engine.readtext(img_bgr, detail=0, paragraph=True))
        if name == "tesseract":
            from PIL import Image
            pil = Image.fromarray(img_bgr[:, :, ::-1])
            return engine.image_to_string(pil, lang="spa+eng", config="--psm 3")
    except Exception:
        pass
    return ""


# ──────────────────────────────────────────────────────────────
# PDF EXTRACTION — per-page, cached, OCR-on-demand
# ──────────────────────────────────────────────────────────────

@st.cache_data(max_entries=200, show_spinner=False)
def _page_text(pdf_bytes: bytes, idx: int) -> tuple[str, bool]:
    """
    Return (text, used_ocr) for one page.
    OCR is applied ONLY when native text quality is below threshold.
    """
    native = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if idx < len(pdf.pages):
                native = pdf.pages[idx].extract_text() or ""
    except Exception:
        pass

    if _text_quality(native) > 0.12 and len(native.strip()) > 60:
        return native, False                    # native text is fine

    # Page is image-based → render and OCR
    ocr_text = ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            if idx < len(doc):
                pix = doc[idx].get_pixmap(
                    matrix=fitz.Matrix(2.0, 2.0),
                    colorspace=fitz.csRGB, alpha=False,
                )
                img = np.frombuffer(pix.samples, np.uint8).reshape(
                    pix.height, pix.width, 3
                )[:, :, ::-1]                  # RGB → BGR
                ocr_text = _run_ocr(img)
    except Exception:
        pass

    return (native + "\n" + ocr_text).strip(), True


@st.cache_data(max_entries=200, show_spinner=False)
def _page_tables(pdf_bytes: bytes, idx: int) -> list[list[list[str]]]:
    """Extract structured tables from a page via pdfplumber."""
    out: list[list[list[str]]] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if idx < len(pdf.pages):
                for tbl in pdf.pages[idx].extract_tables() or []:
                    cleaned = [
                        [str(c or "").strip() for c in row]
                        for row in tbl if row
                    ]
                    if cleaned:
                        out.append(cleaned)
    except Exception:
        pass
    return out


@st.cache_data(max_entries=50, show_spinner=False)
def _render_png(pdf_bytes: bytes, idx: int) -> bytes:
    """Render page to PNG bytes for the PDF viewer."""
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(
            matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False
        )
    return pix.tobytes("png")


# ──────────────────────────────────────────────────────────────
# DATA EXTRACTION — context-aware, multi-pass
# ──────────────────────────────────────────────────────────────

def _total_from_tables(tables: list[list[list[str]]]) -> Optional[float]:
    """
    Scan tables for total rows.
    KEY FIX vs v1: returns the MINIMUM candidate (most specific total),
    not the maximum (which greedily picks project-level budget figures).
    """
    candidates: list[float] = []
    for table in tables:
        for row in table:
            row_str = " ".join(row)
            if _RE_TOTAL.search(row_str):
                # Skip aggregate/project-level amounts
                if re.search(r"proyecto|presupuest|estimad", row_str, re.I):
                    continue
                vals = _parse_money(row_str)
                if vals:
                    candidates.append(vals[0])
    # Use the SMALLEST total found (line-item total < section total < project total)
    return min(candidates) if candidates else None


def _total_from_text(text: str) -> Optional[float]:
    """
    Scan free text for total lines.
    Looks for lines where the ONLY keyword is 'Total' (standalone).
    """
    candidates: list[float] = []
    for line in text.splitlines():
        if not _RE_TOTAL.search(line):
            continue
        if re.search(r"presupuest|del\s+proy|estimad|2[,.]000[,.]000", line, re.I):
            continue
        vals = _parse_money(line)
        if vals:
            candidates.append(vals[0])
    return min(candidates) if candidates else None


def _subtotal_iva_from_tables(
    tables: list[list[list[str]]],
) -> tuple[Optional[float], Optional[float]]:
    sub = iva = None
    for table in tables:
        for row in table:
            row_str = " ".join(row)
            row_low = row_str.lower()
            vals = _parse_money(row_str)
            if not vals:
                continue
            if _RE_SUBTOTAL.search(row_low) and "total" not in row_low.replace("subtotal", ""):
                if sub is None:
                    sub = vals[0]
            elif _RE_IVA_LINE.search(row_low) and not _RE_TOTAL.search(row_low):
                if iva is None:
                    iva = vals[0]
    return sub, iva


def _subtotal_iva_from_text(
    text: str,
) -> tuple[Optional[float], Optional[float]]:
    sub = iva = None
    for line in text.splitlines():
        vals = _parse_money(line)
        if not vals:
            continue
        low = line.lower()
        if _RE_SUBTOTAL.search(low) and "total" not in low.replace("subtotal", ""):
            if sub is None:
                sub = vals[0]
        elif _RE_IVA_LINE.search(low) and not _RE_TOTAL.search(low):
            if iva is None:
                iva = vals[0]
    return sub, iva


def _qty_price_from_tables(
    tables: list[list[list[str]]],
) -> tuple[int, Optional[float]]:
    """
    Detect quantity × unit_price from table rows via algebraic check:
    qty * pu ≈ row_total (within 1 peso tolerance).
    """
    for table in tables:
        for row in table:
            row_str = " ".join(row)
            nums = _parse_money(row_str)
            if len(nums) < 2:
                continue
            for pu in nums:
                for candidate_qty in nums:
                    if candidate_qty == pu or not (1 <= candidate_qty <= 50_000):
                        continue
                    if int(candidate_qty) != candidate_qty:
                        continue
                    if any(abs(candidate_qty * pu - t) < 1.5 for t in nums if t != pu and t != candidate_qty):
                        return int(candidate_qty), pu
    return 1, None


def _find_section_date(text: str) -> Optional[datetime.date]:
    """
    Prioritize lines that look like date labels before falling back
    to scanning the whole text.
    """
    for line in text.splitlines():
        low = line.lower()
        if re.search(r"^fecha|^date|potosí|potosi|s\.l\.p\.|a\s+\d{1,2}\s+de", low):
            d = _parse_date(line)
            if d:
                return d
    return _parse_date(text)


# ──────────────────────────────────────────────────────────────
# MAIN EXTRACTOR
# ──────────────────────────────────────────────────────────────

def extract_section(
    pdf_bytes: bytes,
    label: str,
    p0: int,
    p1: int,
    detect_iva: bool,
    calc_sub: bool,
) -> dict:
    """
    Extract one quotation row from pages p0..p1 (1-indexed, inclusive).
    Returns a dict keyed by COLS.
    """
    full_text = ""
    all_tables: list[list[list[str]]] = []
    used_ocr = False

    for idx in range(p0 - 1, p1):
        txt, ocr = _page_text(pdf_bytes, idx)
        full_text += "\n" + txt
        if ocr:
            used_ocr = True
        all_tables.extend(_page_tables(pdf_bytes, idx))

    obs_parts: list[str] = []
    if used_ocr:
        obs_parts.append("OCR aplicado")

    # ── DATE ──────────────────────────────────────────────────
    fecha = _find_section_date(full_text) or datetime.date.today()

    # ── IVA FLAGS ─────────────────────────────────────────────
    iva_included = bool(_RE_IVA_INC.search(full_text))
    iva_mentioned = detect_iva and bool(_RE_IVA_LINE.search(full_text))
    iva_flag = "Sí" if iva_mentioned else "N/M"

    # ── MONETARY EXTRACTION (tables first, then text) ─────────
    total = _total_from_tables(all_tables)
    if total is None:
        total = _total_from_text(full_text)

    sub, iva_val = _subtotal_iva_from_tables(all_tables)
    if sub is None or iva_val is None:
        s2, iv2 = _subtotal_iva_from_text(full_text)
        if sub is None:
            sub = s2
        if iva_val is None:
            iva_val = iv2

    # ── IVA RECONCILIATION ────────────────────────────────────
    if total is not None:
        if iva_included and sub is None and iva_val is None:
            sub = round(total / 1.16, 2)
            iva_val = round(total - sub, 2)
            obs_parts.append("IVA incluido → subtotal back-calculado")
        elif sub is not None and iva_val is None and iva_mentioned:
            iva_val = round(sub * 0.16, 2)
        elif sub is not None and iva_val is not None and total is None:
            total = round(sub + iva_val, 2)
    elif sub is not None and calc_sub and iva_mentioned:
        iva_val = round(sub * 0.16, 2)
        total = round(sub + iva_val, 2)
        obs_parts.append("Total calculado desde subtotal+IVA")

    # ── QUANTITY / UNIT PRICE ─────────────────────────────────
    qty, unit_price = _qty_price_from_tables(all_tables)
    if unit_price is None and sub is not None:
        unit_price = round(sub / qty, 2)

    return {
        "Fecha":                  fecha.isoformat(),
        "Rubro":                  label,
        "QT":                     "Sí",
        "T. Cambio":              "MXN",
        "(+ IVA)":                iva_flag,
        "Cantidad":               qty,
        "Precio Unitario":        unit_price,
        "Subtotal (Sin IVA)":     sub,
        "IVA 16%":                iva_val,
        "Total con IVA":          total,
        "Diferencia final":       None,
        "Monto en Anexo Escrito": None,
        "Observaciones":          " | ".join(obs_parts),
    }


# ──────────────────────────────────────────────────────────────
# BANXICO API
# ──────────────────────────────────────────────────────────────

@st.cache_data(ttl=3_600, show_spinner=False)
def banxico_rate(fecha_iso: str, currency: str, token: str) -> Optional[float]:
    """
    Fetch Banxico FIX exchange rate for the given date.
    Returns MXN per 1 unit of foreign currency, or None on failure.

    token: obtained free at https://www.banxico.org.mx/SieAPIRest/
    Falls back up to 5 business days back (weekends / holidays).
    """
    if not token or not token.strip() or currency == "MXN":
        return None
    serie = BANXICO_SERIES.get(currency.upper())
    if not serie:
        return None

    base = datetime.date.fromisoformat(fecha_iso)
    for offset in range(6):
        d = (base - datetime.timedelta(days=offset)).strftime("%Y-%m-%d")
        url = (
            f"https://www.banxico.org.mx/SieAPIRest/service/v1/"
            f"series/{serie}/datos/{d}/{d}"
        )
        try:
            resp = requests.get(url, headers={"Bmx-Token": token.strip()}, timeout=8)
            if resp.status_code != 200:
                break
            datos = (
                resp.json()
                .get("bmx", {})
                .get("series", [{}])[0]
                .get("datos", [])
            )
            if datos:
                raw = datos[0].get("dato", "N/E")
                if raw != "N/E":
                    return float(raw.replace(",", ""))
        except Exception:
            break
    return None


# ──────────────────────────────────────────────────────────────
# EXCEL BUILDER — exact template format via openpyxl
# ──────────────────────────────────────────────────────────────

def _side() -> Side:
    return Side(style="thin")

def _border() -> Border:
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex6: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex6)

def _font(bold=False, theme_white=False, size=11) -> Font:
    kwargs = dict(bold=bold, size=size, name="Calibri")
    if theme_white:
        from openpyxl.styles.colors import Color
        kwargs["color"] = Color(theme=0)     # white via theme (matches source file)
    return Font(**kwargs)

def _align(h="general", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)


def build_excel(
    df: pd.DataFrame,
    project_name: str = "VELOCIDAD ACTIVA",
    par_num: int = 9,
) -> bytes:
    """
    Build a .xlsx that exactly reproduces the Cotizaciones_EFIDEPORTE template:
      Row 1  — Project header  (maroon #6E152E)
      Row 2  — Column headers  (gold #D4C19C / beige #EBE2D1)
      Row 3+ — Data rows with Excel formulas (H=F*G, I=H*0.16, J=H+I, K=J-L)
      Last   — SUM row for J, K, L
    All monetary columns use the source file's accounting number format.
    Formulas are left as proper Excel formulas so the file is fully editable.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = f"PAR {str(par_num).zfill(3)}"

    n = len(df)
    D_START = 3          # first data row (1-indexed Excel)
    D_END   = D_START + n - 1
    TOT_ROW = D_END + 1

    HDR1_FILL = _fill("6E152E")
    HDR2_FILL = _fill("D4C19C")
    HDR2_FILL2 = _fill("EBE2D1")

    # ── ROW 1: project title ──────────────────────────────────
    c = ws.cell(1, 1, par_num)
    c.fill = HDR1_FILL
    c.font = _font(bold=True, theme_white=True, size=12)
    c.alignment = _align("center")
    c.number_format = "000"

    c2 = ws.cell(1, 2, project_name.upper())
    c2.fill = HDR1_FILL
    c2.font = _font(bold=True, theme_white=True, size=12)
    c2.alignment = _align("center")
    ws.merge_cells("B1:M1")
    ws.row_dimensions[1].height = 27.75

    # ── ROW 2: column headers ─────────────────────────────────
    headers = [
        "Fecha", "Rubro", "QT", "T. Cambio ", "(+ IVA)", "Cantidad",
        "Precio Unitario", "Subtotal (Sin IVA)", " IVA 16%", "Total con IVA",
        "Diferencia final ", "Monto en Anexo Escrito", "Observaciones",
    ]
    for ci, (hdr, width) in enumerate(zip(headers, COL_WIDTHS), start=1):
        c = ws.cell(2, ci, hdr)
        c.fill = HDR2_FILL2 if ci >= 12 else HDR2_FILL
        c.font = _font(bold=True)
        c.alignment = _align("center")
        c.border = _border()
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 21.75

    # ── DATA ROWS ─────────────────────────────────────────────
    CENTER_COLS = {3, 4, 5}   # QT, T. Cambio, (+IVA) → centered

    for i, (_, row) in enumerate(df.iterrows()):
        er = D_START + i       # Excel row (1-indexed)

        def wc(col: int, value=None, formula: str = None,
               fmt: str = None, bold: bool = False, wrap: bool = False):
            c = ws.cell(er, col, value if formula is None else formula)
            c.border = _border()
            c.alignment = _align("center" if col in CENTER_COLS else "left", wrap=wrap)
            if fmt:
                c.number_format = fmt
            if bold:
                c.font = _font(bold=True)
            return c

        # A — Fecha
        fv = row.get("Fecha")
        if isinstance(fv, str):
            try:
                fv = datetime.datetime.fromisoformat(fv)
            except ValueError:
                fv = datetime.datetime.today()
        elif isinstance(fv, datetime.date) and not isinstance(fv, datetime.datetime):
            fv = datetime.datetime.combine(fv, datetime.time.min)
        wc(1, fv, fmt=DATE_FMT)

        # B — Rubro
        wc(2, str(row.get("Rubro") or ""), wrap=True)

        # C — QT
        wc(3, str(row.get("QT") or "Sí"))

        # D — T. Cambio  (currency code or "MXN")
        tc = str(row.get("T. Cambio") or "MXN")
        wc(4, tc)

        # E — (+IVA)
        wc(5, str(row.get("(+ IVA)") or "N/M"))

        # F — Cantidad
        qty = int(_f(row.get("Cantidad")) or 1)
        wc(6, qty, fmt="0.00")

        # G — Precio Unitario
        g_val = _f(row.get("Precio Unitario"))
        wc(7, g_val, fmt=MONEY_FMT)

        # H — Subtotal (Sin IVA):  =F*G  or hardcoded when G is absent
        if g_val is not None:
            wc(8, formula=f"=F{er}*G{er}", fmt=MONEY_FMT)
        else:
            wc(8, _f(row.get("Subtotal (Sin IVA)")), fmt=MONEY_FMT)

        # I — IVA 16%: formula when (+IVA) is Sí or N/M; blank when No
        iva_flag = str(row.get("(+ IVA)") or "N/M")
        if iva_flag.upper() == "NO":
            wc(9, None, fmt=MONEY_FMT)
        else:
            wc(9, formula=f"=H{er}*0.16", fmt=MONEY_FMT)

        # J — Total con IVA = H + I
        wc(10, formula=f"=H{er}+I{er}", fmt=MONEY_FMT)

        # K — Diferencia final = J - L
        wc(11, formula=f"=J{er}-L{er}", fmt=MONEY_FMT)

        # L — Monto en Anexo Escrito (manual reference, bold)
        l_val = _f(row.get("Monto en Anexo Escrito"))
        wc(12, l_val, fmt=MONEY_FMT, bold=True)

        # M — Observaciones
        # Append T/C note if foreign currency
        obs = str(row.get("Observaciones") or "")
        tc_val = _f(row.get("_tc_rate"))  # internal field set in UI
        if tc_val is not None and tc != "MXN":
            tc_note = f"TC {tc}: {tc_val:,.4f} MXN (Banxico FIX {row.get('Fecha', '')})"
            obs = f"{obs} | {tc_note}".lstrip(" | ")
        wc(13, obs, wrap=True)

    # ── TOTAL ROW ─────────────────────────────────────────────
    for ci in (10, 11, 12):
        cl = get_column_letter(ci)
        c = ws.cell(TOT_ROW, ci, f"=SUM({cl}{D_START}:{cl}{D_END})")
        c.border = _border()
        c.number_format = MONEY_FMT
        c.font = _font(bold=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# SESSION STATE — initialise once, never reset accidentally
# ──────────────────────────────────────────────────────────────
_DEFAULTS: dict = {
    "pdf_bytes":      None,
    "pdf_hash":       None,
    "total_pages":    0,
    "current_page":   0,
    "num_sec":        1,
    "sec_cfg":        [],
    "df":             None,
    "extracted":      False,
    "banxico_token":  os.getenv("BANXICO_TOKEN", ""),
    "project_name":   "VELOCIDAD ACTIVA",
    "par_num":        9,
}
for _k, _v in _DEFAULTS.items():
    st.session_state.setdefault(_k, _v)

# ──────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Documento PDF")
    up = st.file_uploader("Sube el archivo PDF", type=["pdf"])

    if up:
        raw = up.read()
        h = _md5(raw)
        if h != st.session_state.pdf_hash:
            with fitz.open(stream=raw, filetype="pdf") as d:
                pages = len(d)
            st.session_state.update(
                pdf_bytes=raw, pdf_hash=h,
                total_pages=pages,
                current_page=0,
                extracted=False, df=None,
            )
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")

    st.markdown("---")

    # ── Project info ──────────────────────────────────────────
    st.markdown("### 📁 Proyecto")
    st.session_state.project_name = st.text_input(
        "Nombre del proyecto", value=st.session_state.project_name, key="inp_pname"
    )
    st.session_state.par_num = st.number_input(
        "Número PAR", min_value=1, max_value=999,
        value=st.session_state.par_num, step=1, key="inp_par"
    )

    st.markdown("---")

    # ── Banxico token ─────────────────────────────────────────
    st.markdown("### 💱 Tipo de cambio (Banxico)")
    st.session_state.banxico_token = st.text_input(
        "Token API Banxico",
        value=st.session_state.banxico_token,
        type="password",
        help="Obtén tu token gratuito en https://www.banxico.org.mx/SieAPIRest/",
        key="inp_token",
    )
    st.caption("Se usa cuando T. Cambio ≠ MXN para buscar el FIX de Banxico.")

    st.markdown("---")

    # ── Sections ──────────────────────────────────────────────
    st.markdown("### ⚙️ Secciones / Cotizaciones")
    tp = st.session_state.total_pages or 1
    n_sec = int(
        st.number_input(
            "Número de secciones", min_value=1, max_value=50,
            value=st.session_state.num_sec, step=1, key="inp_nsec",
        )
    )
    if n_sec != st.session_state.num_sec:
        st.session_state.num_sec = n_sec
        st.session_state.extracted = False
        st.session_state.df = None

    # Keep sec_cfg in sync with n_sec
    cfgs = st.session_state.sec_cfg
    while len(cfgs) < n_sec:
        i = len(cfgs) + 1
        cfgs.append({"label": f"Sección {i}", "p0": i, "p1": i,
                      "det_iva": True, "calc_sub": True})
    del cfgs[n_sec:]

    for i, cfg in enumerate(cfgs):
        with st.expander(f"📄 Sección {i + 1}", expanded=(n_sec <= 4)):
            cfg["label"] = st.text_input(
                "Rubro / Concepto", value=cfg["label"], key=f"lb_{i}"
            )
            col_a, col_b = st.columns(2)
            p0 = col_a.number_input(
                "Pág. inicio", min_value=1, max_value=tp,
                value=min(cfg["p0"], tp), key=f"p0_{i}",
            )
            p1 = col_b.number_input(
                "Pág. fin", min_value=p0, max_value=tp,
                value=max(min(cfg["p1"], tp), p0), key=f"p1_{i}",
            )
            cfg["p0"] = p0
            cfg["p1"] = p1
            cfg["det_iva"]  = st.checkbox("Detectar IVA",              value=cfg["det_iva"],  key=f"iv_{i}")
            cfg["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=cfg["calc_sub"], key=f"cs_{i}")

    st.markdown("---")
    btn_extract = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_bytes is None),
        use_container_width=True,
        type="primary",
    )

# ──────────────────────────────────────────────────────────────
# HEADER
# ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
  <h1>📋 Conciliador de Cotizaciones PDF</h1>
  <p>Extrae, edita y exporta montos desde documentos PDF — con tipo de cambio Banxico</p>
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# EXTRACTION
# ──────────────────────────────────────────────────────────────
if btn_extract and st.session_state.pdf_bytes:
    rows: list[dict] = []
    bar = st.progress(0, text="Iniciando extracción…")

    for i, cfg in enumerate(st.session_state.sec_cfg):
        bar.progress((i + 0.4) / n_sec, text=f"Extrayendo: {cfg['label']}")
        try:
            row = extract_section(
                st.session_state.pdf_bytes,
                cfg["label"], cfg["p0"], cfg["p1"],
                cfg["det_iva"], cfg["calc_sub"],
            )
        except Exception as exc:
            row = {k: None for k in COLS} | {
                "Rubro":    cfg["label"],
                "QT":       "Sí",
                "T. Cambio": "MXN",
                "Cantidad":  1,
                "Fecha":     datetime.date.today().isoformat(),
                "Observaciones": f"Error al procesar: {str(exc)[:80]}",
            }

        # Banxico lookup if foreign currency
        row["_tc_rate"] = None
        tc_currency = str(row.get("T. Cambio") or "MXN")
        if tc_currency != "MXN" and st.session_state.banxico_token:
            rate = banxico_rate(
                row["Fecha"], tc_currency, st.session_state.banxico_token
            )
            row["_tc_rate"] = rate

        rows.append(row)
        bar.progress((i + 1) / n_sec)

    bar.empty()

    df_new = pd.DataFrame(rows, columns=COLS + ["_tc_rate"])
    df_new["Diferencia final"] = (
        df_new["Total con IVA"].fillna(0) - df_new["Monto en Anexo Escrito"].fillna(0)
    ).where(df_new["Total con IVA"].notna() | df_new["Monto en Anexo Escrito"].notna())

    st.session_state.df = df_new
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()

# ──────────────────────────────────────────────────────────────
# MAIN LAYOUT
# ──────────────────────────────────────────────────────────────
left_col, right_col = st.columns(2, gap="medium")

# ── LEFT: PDF Viewer ──────────────────────────────────────────
with left_col:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)
    tp = st.session_state.total_pages
    cp = st.session_state.current_page

    nav1, nav2, nav3 = st.columns([1, 4, 1])
    if nav1.button("◀", key="btn_prev"):
        cp = max(0, cp - 1)
    cp = int(nav2.number_input("", 1, tp, cp + 1, label_visibility="collapsed", key="inp_pg")) - 1
    if nav3.button("▶", key="btn_next"):
        cp = min(tp - 1, cp + 1)
    st.session_state.current_page = cp
    st.caption(f"Página {cp + 1} de {tp}")

    # Highlight which section this page belongs to
    for cfg in st.session_state.sec_cfg:
        if cfg["p0"] <= cp + 1 <= cfg["p1"]:
            st.markdown(
                f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;'
                f'border-radius:4px;font-size:.8rem">📑 {cfg["label"]}</span>',
                unsafe_allow_html=True,
            )
            break

    with st.spinner("Cargando página…"):
        st.image(_render_png(st.session_state.pdf_bytes, cp), use_container_width=True)

    st.download_button(
        "📥 Descargar PDF original",
        data=st.session_state.pdf_bytes,
        file_name="documento.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

# ── RIGHT: Data Editor ────────────────────────────────────────
with right_col:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)

    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona **🔍 Extraer Montos**.")
    else:
        df: pd.DataFrame = st.session_state.df

        # ── Banxico T/C info panel ─────────────────────────────
        tc_rows = df[df["T. Cambio"].ne("MXN") & df["_tc_rate"].notna()]
        if not tc_rows.empty:
            for _, r in tc_rows.iterrows():
                st.markdown(
                    f'<div class="tc-box">💱 <b>{r["Rubro"]}</b>: '
                    f'1 {r["T. Cambio"]} = <b>{r["_tc_rate"]:,.4f} MXN</b> '
                    f'(Banxico FIX · {r["Fecha"]})</div>',
                    unsafe_allow_html=True,
                )

        # ── KPI placeholders ──────────────────────────────────
        kpi_area = st.container()

        # ── Data Editor (visible cols only — hide _tc_rate) ───
        display_cols = [c for c in df.columns if not c.startswith("_")]
        edited = st.data_editor(
            df[display_cols],
            key="editor_main",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha":                  st.column_config.TextColumn("Fecha (YYYY-MM-DD)", width="small"),
                "Rubro":                  st.column_config.TextColumn("Rubro", width="large"),
                "QT":                     st.column_config.SelectboxColumn("QT",       options=["Sí", "No"], width="small"),
                "T. Cambio":              st.column_config.SelectboxColumn("T. Cambio", options=["MXN", "USD", "EUR", "CAD", "GBP", "JPY"], width="small"),
                "(+ IVA)":               st.column_config.SelectboxColumn("(+ IVA)",  options=["Sí", "No", "N/M"], width="small"),
                "Cantidad":               st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario":        st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)":     st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%":               st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA":          st.column_config.NumberColumn("Total",  format="$%.2f", width="medium"),
                "Diferencia final":       st.column_config.NumberColumn("Dif.",   format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones":          st.column_config.TextColumn("Observaciones", width="large"),
            },
        )

        # Merge edits back (preserve _tc_rate column)
        if edited is not None:
            tc_col = df["_tc_rate"] if "_tc_rate" in df.columns else pd.Series([None]*len(edited))
            merged = edited.copy()
            merged["_tc_rate"] = tc_col.values[:len(merged)]
            st.session_state.df = merged
            df = st.session_state.df

        # ── Re-fetch Banxico for any row whose currency changed ─
        token = st.session_state.banxico_token
        if token:
            for idx in range(len(df)):
                cur = str(df.at[idx, "T. Cambio"] if "T. Cambio" in df.columns else "MXN")
                if cur != "MXN" and (
                    "_tc_rate" not in df.columns
                    or pd.isna(df.at[idx, "_tc_rate"])
                    or df.at[idx, "_tc_rate"] is None
                ):
                    fecha_str = str(df.at[idx, "Fecha"] or datetime.date.today())
                    rate = banxico_rate(fecha_str, cur, token)
                    df.at[idx, "_tc_rate"] = rate

        # ── KPIs ──────────────────────────────────────────────
        tot_sum = _f(df["Total con IVA"].sum(skipna=True)) or 0.0
        ref_sum = _f(df["Monto en Anexo Escrito"].sum(skipna=True)) or 0.0
        dif_sum = tot_sum - ref_sum

        with kpi_area:
            k1, k2, k3 = st.columns(3)
            for col, val, lbl in [
                (k1, tot_sum, "Total Extraído"),
                (k2, ref_sum, "Monto Referencia"),
                (k3, dif_sum, "Diferencia"),
            ]:
                color = (
                    "#c0392b" if (lbl == "Diferencia" and abs(dif_sum) > 0.01)
                    else ("#28a745" if lbl == "Diferencia" else "#1a2744")
                )
                col.markdown(
                    f'<div class="kpi"><div class="v" style="color:{color}">'
                    f'${val:,.2f}</div><div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

        st.markdown("---")

        # ── Export buttons ────────────────────────────────────
        col_xls, col_info = st.columns([3, 2])

        try:
            xlsx_bytes = build_excel(
                df,
                project_name=st.session_state.project_name,
                par_num=st.session_state.par_num,
            )
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            col_xls.download_button(
                "⬇️ Descargar Excel (editable)",
                data=xlsx_bytes,
                file_name=f"Cotizaciones_PAR{st.session_state.par_num:03d}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            col_info.info(
                "El Excel descargado contiene **fórmulas vivas** "
                "(Subtotal, IVA, Total, Diferencia). "
                "Puedes editar **Precio Unitario** y **Monto en Anexo Escrito** "
                "directamente en Excel sin necesidad de volver aquí."
            )
        except Exception as exc:
            st.error(f"Error generando Excel: {exc}")

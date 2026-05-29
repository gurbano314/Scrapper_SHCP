# ============================================================
# CONCILIADOR DE COTIZACIONES PDF — v2.1 (Corrección Total)
# Streamlit · Compatible Streamlit Cloud + Local
# ============================================================

from __future__ import annotations

import datetime
import hashlib
import io
import math
import re
import os
import requests
from typing import Optional

import fitz                    # PyMuPDF
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

BANXICO_SERIES: dict[str, str] = {
    "USD": "SF43718", "EUR": "SF46410", "CAD": "SF57771",
    "GBP": "SF46406", "JPY": "SF46407", "CHF": "SF46408",
}

_MESES: dict[str, int] = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
}

_RE_MONEY = re.compile(
    r"\$?\s*([\d]{1,3}(?:,\d{3})*(?:\.\d{1,2})?)"
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{0,2})?)"
)

_RE_DATE = re.compile(
    r"(\d{1,2})\s*(?:de\s+)?([a-záéíóúüñ]{3,})\s*(?:de(?:l)?\s+)?(\d{4})"
    r"|(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})"
    r"|(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{2,4})",
    re.IGNORECASE,
)

_RE_TOTAL = re.compile(r"\b(total|importe\s+total|gran\s+total)\b", re.IGNORECASE)
_RE_SUBTOTAL  = re.compile(r"\b(subtotal|sub[\s\-]?total|sin\s+iva|importe)\b", re.IGNORECASE)
_RE_IVA_LINE  = re.compile(r"\biva\b|\b16\s*%|impuesto\s+al\s+valor\s+agregado", re.IGNORECASE)
_RE_IVA_INC   = re.compile(r"(?:precios?|tarifas?|montos?).*inclu(?:ye[ns]?|ido[s]?).*iva|iva\s+inclu(?:ido|ye)", re.IGNORECASE)

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def _md5(b: bytes) -> str: return hashlib.md5(b).hexdigest()

def _f(v) -> Optional[float]:
    if v is None: return None
    try:
        x = float(str(v).replace(",", "").strip())
        return None if (math.isnan(x) or math.isinf(x)) else x
    except (ValueError, TypeError): return None

def _parse_money(text: str) -> list[float]:
    out: list[float] = []
    for m in _RE_MONEY.finditer(str(text)):
        raw = m.group(1) or m.group(2)
        if raw:
            try:
                v = float(raw.replace(",", ""))
                if v > 0: out.append(v)
            except ValueError: pass
    return sorted(set(out), reverse=True)

def _parse_date(text: str) -> Optional[datetime.date]:
    for m in _RE_DATE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]:
                mes = _resolve_month(g[1].strip())
                if mes: return datetime.date(int(g[2]), mes, int(g[0]))
            elif g[3]: return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            elif g[6]:
                yr = int(g[8])
                return datetime.date(yr + 2000 if yr < 100 else yr, int(g[7]), int(g[6]))
        except (ValueError, TypeError): pass
    return None

def _resolve_month(s: str) -> Optional[int]:
    s = s.lower().strip()
    if s in _MESES: return _MESES[s]
    for k, v in _MESES.items():
        if len(s) >= 3 and (s.startswith(k[:3]) or k.startswith(s[:3])): return v
    return None

# ──────────────────────────────────────────────────────────────
# OCR ENGINE
# ──────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR(det_model_dir=None, rec_model_dir=None, cls_model_dir=None)
        return "rapidocr", engine
    except Exception as e:
        st.error(f"⚠️ Error cargando motor OCR: {str(e)}")
        return None, None

def _run_ocr(img_bgr: np.ndarray) -> str:
    name, engine = _ocr_engine()
    if name is None: return ""
    try:
        if name == "rapidocr":
            result, _ = engine(img_bgr)
            return "\n".join(r[1] for r in (result or []) if r and len(r) > 1)
    except Exception as e:
        st.warning(f"⚠️ Error visual: {str(e)}")
    return ""

# ──────────────────────────────────────────────────────────────
# PDF EXTRACTION
# ──────────────────────────────────────────────────────────────
@st.cache_data(max_entries=200, show_spinner=False)
def _page_text_hybrid(pdf_bytes: bytes, idx: int) -> tuple[str, bool]:
    """Force OCR extraction concurrently to guarantee data rescue."""
    native = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if idx < len(pdf.pages): native = pdf.pages[idx].extract_text() or ""
    except Exception: pass

    ocr_text = ""
    used_ocr = False
    
    # We forcefully run OCR if native text is weak, or if requested silently
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            if idx < len(doc):
                pix = doc[idx].get_pixmap(matrix=fitz.Matrix(2.0, 2.0), colorspace=fitz.csRGB, alpha=False)
                img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)[:, :, ::-1]
                ocr_text = _run_ocr(img)
                if len(ocr_text) > 30: used_ocr = True
    except Exception: pass

    return (native + "\n" + ocr_text).strip(), used_ocr

@st.cache_data(max_entries=200, show_spinner=False)
def _page_tables(pdf_bytes: bytes, idx: int) -> list[list[list[str]]]:
    out: list[list[list[str]]] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if idx < len(pdf.pages):
                for tbl in pdf.pages[idx].extract_tables() or []:
                    cleaned = [[str(c or "").strip() for c in row] for row in tbl if row]
                    if cleaned: out.append(cleaned)
    except Exception: pass
    return out

@st.cache_data(max_entries=50, show_spinner=False)
def _render_png(pdf_bytes: bytes, idx: int) -> bytes:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False)
        return pix.tobytes("png")

# ──────────────────────────────────────────────────────────────
# EXTRACTION LOGIC
# ──────────────────────────────────────────────────────────────
def extract_section(pdf_bytes: bytes, label: str, p0: int, p1: int, detect_iva: bool, calc_sub: bool) -> dict:
    full_text = ""
    all_tables: list[list[list[str]]] = []
    used_ocr = False

    for idx in range(p0 - 1, p1):
        txt, ocr_flag = _page_text_hybrid(pdf_bytes, idx)
        full_text += "\n" + txt
        if ocr_flag: used_ocr = True
        all_tables.extend(_page_tables(pdf_bytes, idx))

    obs_parts: list[str] = []
    if used_ocr: obs_parts.append("OCR aplicado")

    fecha = _parse_date(full_text) or datetime.date.today()
    iva_included = bool(_RE_IVA_INC.search(full_text))
    iva_mentioned = detect_iva and bool(_RE_IVA_LINE.search(full_text))
    iva_flag = "Sí" if iva_mentioned else "N/M"

    # Extraction
    total = sub = iva_val = unit_price = None
    qty = 1

    # 1. Busca el total explicitamente filtrando los $2,000,000 de presupuesto global
    total_candidates = []
    for line in full_text.splitlines():
        if _RE_TOTAL.search(line) and not re.search(r"proyecto|presupuest", line, re.I):
            for v in _parse_money(line):
                if v != 2000000.0:  # Ignora presupuesto proyecto
                    total_candidates.append(v)
    if total_candidates:
        total = max(total_candidates) # Toma el valor maximo valido

    # 2. Busca Subtotal e IVA
    for line in full_text.splitlines():
        low = line.lower()
        if sub is None and _RE_SUBTOTAL.search(low) and "total" not in low.replace("subtotal", ""):
            v = _parse_money(line)
            if v: sub = v[0]
        if iva_val is None and _RE_IVA_LINE.search(low) and not _RE_TOTAL.search(low):
            v = _parse_money(line)
            if v: iva_val = v[0]

    # 3. RECONSTRUCCION DE LINEA (El motor que salva la cotización de Santa Úrsula y Desportik)
    valid_lines = []
    seen = set()
    for line in full_text.splitlines():
        nums = _parse_money(line)
        if len(nums) >= 2:
            for total_cand in nums:
                for pu_cand in nums:
                    for qty_cand in nums:
                        # Si (Cantidad * PU = Total) con un margen de 5 pesos
                        if total_cand != pu_cand and pu_cand != qty_cand and 1 <= qty_cand <= 50000:
                            if abs((qty_cand * pu_cand) - total_cand) < 5.0:
                                combo = (int(qty_cand), round(pu_cand, 2), round(total_cand, 2))
                                if combo not in seen:
                                    seen.add(combo)
                                    valid_lines.append((combo[0], combo[1], combo[2]))

    # Si hay lineas válidas y el total falló, usar la sumatoria reconstruida
    if valid_lines:
        if total is None:
            total = sum(item[2] for item in valid_lines)
            obs_parts.append("Total reconstruido por suma de ítems")
        # Si es un solo producto, heredar PU y Cantidad
        if len(valid_lines) == 1:
            qty = valid_lines[0][0]
            unit_price = valid_lines[0][1]

    # 4. Matemáticas Finales
    if total is not None:
        if iva_included and sub is None and iva_val is None:
            sub = round(total / 1.16, 2)
            iva_val = round(total - sub, 2)
            obs_parts.append("IVA incluido → subtotal calculado")
        elif sub is not None and iva_val is None and iva_mentioned:
            iva_val = round(sub * 0.16, 2)
        elif sub is not None and iva_val is not None and total is None:
            total = round(sub + iva_val, 2)
    elif sub is not None and calc_sub and iva_mentioned:
        iva_val = round(sub * 0.16, 2)
        total = round(sub + iva_val, 2)
        obs_parts.append("Total deducido desde subtotal")

    if unit_price is None and sub is not None:
        unit_price = round(sub / qty, 2)

    return {
        "Fecha":              fecha.isoformat(),
        "Rubro":              label,
        "QT":                 "Sí",
        "T. Cambio":          "MXN",
        "(+ IVA)":            iva_flag,
        "Cantidad":           qty,
        "Precio Unitario":    unit_price,
        "Subtotal (Sin IVA)": sub,
        "IVA 16%":            iva_val,
        "Total con IVA":      total,
        "Diferencia final":   None,
        "Monto en Anexo Escrito": None,
        "Observaciones":      " | ".join(obs_parts),
    }

# ──────────────────────────────────────────────────────────────
# BANXICO API
# ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=3_600, show_spinner=False)
def banxico_rate(fecha_iso: str, currency: str, token: str) -> Optional[float]:
    if not token or not token.strip() or currency == "MXN": return None
    serie = BANXICO_SERIES.get(currency.upper())
    if not serie: return None

    base = datetime.date.fromisoformat(fecha_iso)
    for offset in range(6):
        d = (base - datetime.timedelta(days=offset)).strftime("%Y-%m-%d")
        url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{serie}/datos/{d}/{d}"
        try:
            resp = requests.get(url, headers={"Bmx-Token": token.strip()}, timeout=8)
            if resp.status_code != 200: break
            datos = resp.json().get("bmx", {}).get("series", [{}])[0].get("datos", [])
            if datos and datos[0].get("dato", "N/E") != "N/E":
                return float(datos[0].get("dato").replace(",", ""))
        except Exception: break
    return None

# ──────────────────────────────────────────────────────────────
# EXCEL BUILDER 
# ──────────────────────────────────────────────────────────────
def _side() -> Side: return Side(style="thin")
def _border() -> Border: s = _side(); return Border(left=s, right=s, top=s, bottom=s)
def _fill(hex6: str) -> PatternFill: return PatternFill("solid", fgColor=hex6)
def _font(bold=False, theme_white=False, size=11) -> Font:
    kwargs = dict(bold=bold, size=size, name="Calibri")
    if theme_white:
        from openpyxl.styles.colors import Color
        kwargs["color"] = Color(theme=0)
    return Font(**kwargs)
def _align(h="general", wrap=False) -> Alignment: return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def build_excel(df: pd.DataFrame, project_name: str = " ", par_num: int = 9) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = f"PAR {str(par_num).zfill(3)}"

    n = len(df)
    D_START = 3; D_END = D_START + n - 1; TOT_ROW = D_END + 1

    c = ws.cell(1, 1, par_num)
    c.fill = _fill("6E152E"); c.font = _font(bold=True, theme_white=True, size=12); c.alignment = _align("center"); c.number_format = "000"
    
    c2 = ws.cell(1, 2, project_name.upper())
    c2.fill = _fill("6E152E"); c2.font = _font(bold=True, theme_white=True, size=12); c2.alignment = _align("center")
    ws.merge_cells("B1:M1")
    ws.row_dimensions[1].height = 27.75

    headers = ["Fecha", "Rubro", "QT", "T. Cambio ", "(+ IVA)", "Cantidad", "Precio Unitario", "Subtotal (Sin IVA)", " IVA 16%", "Total con IVA", "Diferencia final ", "Monto en Anexo Escrito", "Observaciones"]
    for ci, (hdr, width) in enumerate(zip(headers, COL_WIDTHS), start=1):
        c = ws.cell(2, ci, hdr)
        c.fill = _fill("EBE2D1") if ci >= 12 else _fill("D4C19C")
        c.font = _font(bold=True); c.alignment = _align("center"); c.border = _border()
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 21.75

    for i, (_, row) in enumerate(df.iterrows()):
        er = D_START + i
        def wc(col: int, value=None, formula: str = None, fmt: str = None, bold: bool = False, wrap: bool = False):
            c = ws.cell(er, col, value if formula is None else formula)
            c.border = _border(); c.alignment = _align("center" if col in {3, 4, 5} else "left", wrap=wrap)
            if fmt: c.number_format = fmt
            if bold: c.font = _font(bold=True)
            return c

        fv = row.get("Fecha")
        if isinstance(fv, str):
            try: fv = datetime.datetime.fromisoformat(fv)
            except ValueError: fv = datetime.datetime.today()
        wc(1, fv, fmt=DATE_FMT)
        wc(2, str(row.get("Rubro") or ""), wrap=True)
        wc(3, str(row.get("QT") or "Sí"))
        tc = str(row.get("T. Cambio") or "MXN")
        wc(4, tc)
        wc(5, str(row.get("(+ IVA)") or "N/M"))
        qty = int(_f(row.get("Cantidad")) or 1)
        wc(6, qty, fmt="0.00")
        
        g_val = _f(row.get("Precio Unitario"))
        wc(7, g_val, fmt=MONEY_FMT)
        if g_val is not None: wc(8, formula=f"=F{er}*G{er}", fmt=MONEY_FMT)
        else: wc(8, _f(row.get("Subtotal (Sin IVA)")), fmt=MONEY_FMT)
        
        if str(row.get("(+ IVA)") or "N/M").upper() == "NO": wc(9, None, fmt=MONEY_FMT)
        else: wc(9, formula=f"=H{er}*0.16", fmt=MONEY_FMT)
        
        wc(10, formula=f"=H{er}+I{er}", fmt=MONEY_FMT)
        wc(11, formula=f"=J{er}-L{er}", fmt=MONEY_FMT)
        wc(12, _f(row.get("Monto en Anexo Escrito")), fmt=MONEY_FMT, bold=True)
        
        obs = str(row.get("Observaciones") or "")
        tc_val = _f(row.get("_tc_rate"))
        if tc_val is not None and tc != "MXN":
            obs = f"{obs} | TC {tc}: {tc_val:,.4f} MXN".lstrip(" | ")
        wc(13, obs, wrap=True)

    for ci in (10, 11, 12):
        cl = get_column_letter(ci)
        c = ws.cell(TOT_ROW, ci, f"=SUM({cl}{D_START}:{cl}{D_END})")
        c.border = _border(); c.number_format = MONEY_FMT; c.font = _font(bold=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────
# SESSION STATE & SIDEBAR
# ──────────────────────────────────────────────────────────────
_DEFAULTS: dict = {"pdf_bytes": None, "pdf_hash": None, "total_pages": 0, "current_page": 0, "num_sec": 1, "sec_cfg": [], "df": None, "extracted": False, "banxico_token": os.getenv("BANXICO_TOKEN", ""), "project_name": " ", "par_num": 9}
for _k, _v in _DEFAULTS.items(): st.session_state.setdefault(_k, _v)

with st.sidebar:
    st.markdown("## 📂 Documento PDF")
    up = st.file_uploader("Sube el archivo PDF", type=["pdf"])
    if up:
        raw = up.read()
        h = _md5(raw)
        if h != st.session_state.pdf_hash:
            with fitz.open(stream=raw, filetype="pdf") as d: pages = len(d)
            st.session_state.update(pdf_bytes=raw, pdf_hash=h, total_pages=pages, current_page=0, extracted=False, df=None)
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")

    st.markdown("---")
    st.markdown("### 📁 Proyecto")
    st.session_state.project_name = st.text_input("Nombre del proyecto", value=st.session_state.project_name, key="inp_pname")
    st.session_state.par_num = st.number_input("Folio Proyecto", min_value=1, max_value=999, value=st.session_state.par_num, step=1, key="inp_par")

    st.markdown("---")
    st.markdown("### 💱 Tipo de cambio (Banxico)")
    st.session_state.banxico_token = st.text_input("Token API Banxico", value=st.session_state.banxico_token, type="password", key="inp_token")
    st.caption("Se usa cuando T. Cambio ≠ MXN para buscar el FIX de Banxico.")

    st.markdown("---")
    st.markdown("### ⚙️ Secciones / Cotizaciones")
    tp = st.session_state.total_pages or 1
    n_sec = int(st.number_input("Número de secciones", min_value=1, max_value=50, value=st.session_state.num_sec, step=1, key="inp_nsec"))
    if n_sec != st.session_state.num_sec:
        st.session_state.num_sec = n_sec; st.session_state.extracted = False; st.session_state.df = None

    cfgs = st.session_state.sec_cfg
    while len(cfgs) < n_sec:
        i = len(cfgs) + 1
        cfgs.append({"label": f"Sección {i}", "p0": i, "p1": i, "det_iva": True, "calc_sub": True})
    del cfgs[n_sec:]

    for i, cfg in enumerate(cfgs):
        with st.expander(f"📄 Sección {i + 1}", expanded=(n_sec <= 4)):
            cfg["label"] = st.text_input("Rubro / Concepto", value=cfg["label"], key=f"lb_{i}")
            col_a, col_b = st.columns(2)
            p0 = col_a.number_input("Pág. inicio", min_value=1, max_value=tp, value=min(cfg["p0"], tp), key=f"p0_{i}")
            p1 = col_b.number_input("Pág. fin", min_value=p0, max_value=tp, value=max(min(cfg["p1"], tp), p0), key=f"p1_{i}")
            cfg["p0"] = p0; cfg["p1"] = p1
            cfg["det_iva"]  = st.checkbox("Detectar IVA", value=cfg["det_iva"], key=f"iv_{i}")
            cfg["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=cfg["calc_sub"], key=f"cs_{i}")

    st.markdown("---")
    btn_extract = st.button("🔍 Extraer Montos", disabled=(st.session_state.pdf_bytes is None), use_container_width=True, type="primary")

# ──────────────────────────────────────────────────────────────
# HEADER & EXTRACTION
# ──────────────────────────────────────────────────────────────
st.markdown('<div class="hdr"><h1>📋 Conciliador de Cotizaciones PDF</h1><p>Extrae, edita y exporta montos desde documentos PDF — con tipo de cambio Banxico</p></div>', unsafe_allow_html=True)

if btn_extract and st.session_state.pdf_bytes:
    rows: list[dict] = []
    bar = st.progress(0, text="Iniciando extracción…")

    for i, cfg in enumerate(st.session_state.sec_cfg):
        bar.progress((i + 0.4) / n_sec, text=f"Extrayendo: {cfg['label']}")
        try:
            row = extract_section(st.session_state.pdf_bytes, cfg["label"], cfg["p0"], cfg["p1"], cfg["det_iva"], cfg["calc_sub"])
        except Exception as exc:
            row = {k: None for k in COLS} | {"Rubro": cfg["label"], "QT": "Sí", "T. Cambio": "MXN", "Cantidad": 1, "Fecha": datetime.date.today().isoformat(), "Observaciones": f"Error al procesar: {str(exc)[:80]}"}

        row["_tc_rate"] = None
        tc_currency = str(row.get("T. Cambio") or "MXN")
        if tc_currency != "MXN" and st.session_state.banxico_token:
            row["_tc_rate"] = banxico_rate(row["Fecha"], tc_currency, st.session_state.banxico_token)
        rows.append(row)
        bar.progress((i + 1) / n_sec)

    bar.empty()
    df_new = pd.DataFrame(rows, columns=COLS + ["_tc_rate"])
    df_new["Diferencia final"] = (df_new["Total con IVA"].fillna(0) - df_new["Monto en Anexo Escrito"].fillna(0)).where(df_new["Total con IVA"].notna() | df_new["Monto en Anexo Escrito"].notna())
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

with left_col:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)
    tp = st.session_state.total_pages; cp = st.session_state.current_page
    nav1, nav2, nav3 = st.columns([1, 4, 1])
    if nav1.button("◀", key="btn_prev"): cp = max(0, cp - 1)
    cp = int(nav2.number_input("", 1, tp, cp + 1, label_visibility="collapsed", key="inp_pg")) - 1
    if nav3.button("▶", key="btn_next"): cp = min(tp - 1, cp + 1)
    st.session_state.current_page = cp
    st.caption(f"Página {cp + 1} de {tp}")

    for cfg in st.session_state.sec_cfg:
        if cfg["p0"] <= cp + 1 <= cfg["p1"]:
            st.markdown(f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;border-radius:4px;font-size:.8rem">📑 {cfg["label"]}</span>', unsafe_allow_html=True)
            break
    with st.spinner("Cargando página…"): st.image(_render_png(st.session_state.pdf_bytes, cp), use_container_width=True)
    st.download_button("📥 Descargar PDF", data=st.session_state.pdf_bytes, file_name="documento.pdf", mime="application/pdf", use_container_width=True)

with right_col:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)
    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona **🔍 Extraer Montos**.")
    else:
        df: pd.DataFrame = st.session_state.df
        tc_rows = df[df["T. Cambio"].ne("MXN") & df["_tc_rate"].notna()]
        if not tc_rows.empty:
            for _, r in tc_rows.iterrows():
                st.markdown(f'<div class="tc-box">💱 <b>{r["Rubro"]}</b>: 1 {r["T. Cambio"]} = <b>{r["_tc_rate"]:,.4f} MXN</b> (Banxico FIX · {r["Fecha"]})</div>', unsafe_allow_html=True)

        kpi_area = st.container()
        display_cols = [c for c in df.columns if not c.startswith("_")]
        edited = st.data_editor(
            df[display_cols], key="editor_main", use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "Fecha": st.column_config.TextColumn("Fecha", width="small"),
                "Rubro": st.column_config.TextColumn("Rubro", width="large"),
                "QT": st.column_config.SelectboxColumn("QT", options=["Sí", "No"], width="small"),
                "T. Cambio": st.column_config.SelectboxColumn("T. Cambio", options=["MXN", "USD", "EUR", "CAD", "GBP", "JPY"], width="small"),
                "(+ IVA)": st.column_config.SelectboxColumn("(+ IVA)", options=["Sí", "No", "N/M"], width="small"),
                "Cantidad": st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario": st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)": st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%": st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA": st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Diferencia final": st.column_config.NumberColumn("Dif.", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
            }
        )

        if edited is not None:
            tc_col = df["_tc_rate"] if "_tc_rate" in df.columns else pd.Series([None]*len(edited))
            merged = edited.copy()
            merged["_tc_rate"] = tc_col.values[:len(merged)]
            st.session_state.df = merged
            df = st.session_state.df

        token = st.session_state.banxico_token
        if token:
            for idx in range(len(df)):
                cur = str(df.at[idx, "T. Cambio"] if "T. Cambio" in df.columns else "MXN")
                if cur != "MXN" and ("_tc_rate" not in df.columns or pd.isna(df.at[idx, "_tc_rate"]) or df.at[idx, "_tc_rate"] is None):
                    fecha_str = str(df.at[idx, "Fecha"] or datetime.date.today())
                    df.at[idx, "_tc_rate"] = banxico_rate(fecha_str, cur, token)

        tot_sum = _f(df["Total con IVA"].sum(skipna=True)) or 0.0
        ref_sum = _f(df["Monto en Anexo Escrito"].sum(skipna=True)) or 0.0
        dif_sum = tot_sum - ref_sum

        with kpi_area:
            k1, k2, k3 = st.columns(3)
            for col, val, lbl in [(k1, tot_sum, "Total Extraído"), (k2, ref_sum, "Monto Referencia"), (k3, dif_sum, "Diferencia")]:
                color = "#c0392b" if (lbl == "Diferencia" and abs(dif_sum) > 0.01) else ("#28a745" if lbl == "Diferencia" else "#1a2744")
                col.markdown(f'<div class="kpi"><div class="v" style="color:{color}">${val:,.2f}</div><div class="l">{lbl}</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        col_xls, col_info = st.columns([3, 2])
        try:
            xlsx_bytes = build_excel(df, project_name=st.session_state.project_name, par_num=st.session_state.par_num)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            col_xls.download_button(
                "⬇️ Descargar Excel (editable)", data=xlsx_bytes, file_name=f"Cotizaciones_PAR{st.session_state.par_num:03d}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
            )
            col_info.info("El Excel descargado contiene fórmulas vivas. Puedes editar montos directamente sin volver aquí.")
        except Exception as exc: st.error(f"Error generando Excel: {exc}")

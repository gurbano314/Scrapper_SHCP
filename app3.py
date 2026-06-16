# ============================================================
# CONCILIADOR DE COTIZACIONES PDF  |  app.py  v4.2
# ============================================================
# Novedades v4.2:
#   M1  — Carga múltiple de PDFs: file_uploader con
#         accept_multiple_files=True. Los PDFs se concatenan
#         internamente en un solo documento (PyMuPDF) para
#         paginación unificada. El visor muestra todos los PDFs
#         como un libro continuo, con indicador de archivo fuente.
#   M2  — Cada sección se vincula a un PDF específico o al
#         documento consolidado, con rangos de página globales.
#   M3  — Selector de PDF activo en el visor con indicador
#         visual del archivo fuente por página.
#   M4  — Descarga individual de cada PDF o del consolidado.
#
# Correcciones heredadas de v4.1:
#   P0  — to_excel: un solo Workbook (wb=Workbook(), ws=wb.active)
#   P1  — data_editor + st.rerun: guardia de reentrada
#   F1  — Detección cambios via hash MD5
#   F2  — Navegación: una sola fuente de verdad (on_click+args)
#   F3  — Caché de render: pdf_bytes fuera de la firma
#   F4  — OCR síncrono optimizado (sin ThreadPoolExecutor)
#   F5  — Regex grupo 3: solo decimales .XX
#   F6  — Fórmulas Excel con IF(ISNUMBER(…))
#   F7  — Guard explícito sobre df_cur antes de KPIs
#   F8  — Triangulación IVA con cota de cordura
#   F9  — cache_resource max_entries=1 para OCR
#   F10 — Caché negativo en API Banxico
#   F11 — Fila de totales Excel: rango correcto
#   F12 — getvalue() en file_uploader
#   F13 — División por cero protegida en P.U.
#   R1  — Reconstrucción qty × pu ≈ total
#   R2  — Fallback contexto monetario
#   R3  — Checkbox "Calcular subtotal si falta"
#   R4  — Plantilla vacía Excel
#   R5  — Botón "Descargar PDF" en visor
#   R6  — Alertas de extracción
#   R7  — Enlace documentación Banxico
# ============================================================

import datetime
import hashlib
import io
import re
from typing import Optional

import fitz  # PyMuPDF
import numpy as np
import openpyxl
import pandas as pd
import pdfplumber
import requests
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliador de Cotizaciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
[data-testid="stSidebar"] { background:#6E152E; }
[data-testid="stSidebar"] * { color:#f0e0e6 !important; }
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select,
[data-testid="stSidebar"] textarea {
    color: #1a1a1a !important;
    background-color: #f5e0e6 !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] input::placeholder,
[data-testid="stSidebar"] textarea::placeholder {
    color: #9a6070 !important;
    opacity: 1 !important;
}
.hdr {
    background:linear-gradient(90deg,#6E152E,#a02048);
    padding:14px 22px; border-radius:8px; margin-bottom:18px;
}
.hdr h1 { color:#fff; font-size:1.4rem; margin:0; font-weight:700; }
.hdr p  { color:#f8ccd6; font-size:.87rem; margin:4px 0 0; }
.ptitle {
    background:#6E152E; color:#fff !important; font-weight:700;
    font-size:.88rem; padding:7px 14px;
    border-radius:6px 6px 0 0; margin-bottom:4px;
}
.kpi {
    background:#fdf5f7; border:1px solid #e8b4c0;
    border-radius:8px; padding:10px 12px;
    text-align:center; margin-bottom:6px;
}
.kpi .v { font-size:1.2rem; font-weight:700; color:#6E152E; }
.kpi .l { font-size:.72rem; color:#9a4060; }
.tag-moneda {
    background:#6E152E; color:#fff !important; padding:2px 8px;
    border-radius:4px; font-size:.78rem; font-weight:600;
}
.tag-file {
    background:#2c3e50; color:#ecf0f1 !important; padding:2px 8px;
    border-radius:4px; font-size:.75rem; font-weight:600;
    margin-right:4px;
}
</style>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────
NATIVE_MIN_CHARS_PER_PAGE = 80
MAX_PLAUSIBLE_MXN = 50_000_000

_BANXICO_SERIES = {
    "USD": "SF43718",
    "EUR": "SF46410",
    "CAD": "SF60653",
}
_BANXICO_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1"

_COLS = [
    "Fecha",
    "Rubro",
    "QT",
    "T. Cambio",
    "(+ IVA)",
    "Cantidad",
    "Precio Unitario",
    "Subtotal (Sin IVA)",
    "IVA 16%",
    "Total con IVA",
    "Diferencia final",
    "Monto en Anexo Escrito",
    "Observaciones",
]
_WIDTHS_COL = [12, 52, 5, 10, 7, 8, 16, 18, 17, 16, 18, 22, 48]

_MONEY_RE = re.compile(
    r"\$\s*([\d,]+(?:\.\d{1,2})?)"
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)"
    r"|(?<!\d)(\d{1,9}\.\d{2})(?!\d)"
)
_MONETARY_CTX = re.compile(
    r"total|importe|monto|precio|valor|costo|cobro|cargo|pago"
    r"|subtotal|honorarios|tarifa|renta|fianza",
    re.IGNORECASE,
)
_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:de)?\s*"
    r"([a-záéíóúñA-ZÁÉÍÓÚÑ]+)[,\s]*(?:del?\s*)?(\d{4})"
    r"|(\d{4})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{1,2})"
    r"|(\d{1,2})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{2,4})"
)
_ITEM_RE = re.compile(
    r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)
_ES_DE_RE = re.compile(
    r"es\s+de\s+\$?\s*([\d,]+(?:\.\d{1,2})?)", re.IGNORECASE
)
_BUDGET_EXCL = re.compile(
    r"presupuestal|presupuesto\s+total", re.IGNORECASE
)
_MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    "ene": 1, "feb": 2, "mar": 3, "abr": 4,
    "may": 5, "jun": 6, "jul": 7, "ago": 8,
    "sep": 9, "oct": 10, "nov": 11, "dic": 12,
}

_FMT_MONEY = (
    '_-"$ "* #,##0.00_-;'
    '\\-"$ "* #,##0.00_-;'
    '_-"$ "* "-"??_-;_-@_-'
)
_FMT_DATE = "mm-dd-yy"


# ─────────────────────────────────────────────────────────────
# SESSION STATE — inicialización defensiva (M1: multi-PDF)
# ─────────────────────────────────────────────────────────────
_SS_DEFAULTS: dict = {
    # ── Multi-PDF (M1) ───────────────────────────────────────
    "pdf_files":       [],     # lista de dicts: {name, bytes, hash, pages}
    "pdf_combined":    None,   # bytes del PDF consolidado (todos concatenados)
    "pdf_combined_hash": None, # MD5 del consolidado para caché de render
    "total_pages":     0,      # total de páginas del consolidado
    "page_map":        [],     # mapeo: [{"file_idx":0,"local_page":0,"name":"a.pdf"}, ...]
    # ── Estado general ───────────────────────────────────────
    "current_page":    0,
    "num_sec":         1,
    "sec_cfg":         [],
    "df":              None,
    "df_hash":         "",
    "extracted":       False,
    "bx_cache":        {},
    "proyecto_nombre": "",
    "_rerun_guard":    False,
}
for _k, _v in _SS_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def _md5(b: bytes) -> str:
    """Calcula hash MD5 de un bloque de bytes."""
    return hashlib.md5(b).hexdigest()


def _df_hash(df: pd.DataFrame) -> str:
    """Hash determinista de un DataFrame para detección de cambios."""
    if df is None:
        return ""
    try:
        return hashlib.md5(
            pd.util.hash_pandas_object(df, index=False).values.tobytes()
        ).hexdigest()
    except Exception:
        return hashlib.md5(df.to_csv(index=False).encode()).hexdigest()


def _safe_f(v) -> Optional[float]:
    """Convierte un valor arbitrario a float. None si falla o NaN."""
    if v is None:
        return None
    try:
        f = float(str(v).replace(",", "").strip())
        return None if f != f else f
    except (ValueError, TypeError):
        return None


def _money(txt: str) -> Optional[float]:
    """Extrae el primer valor monetario positivo de una cadena."""
    for m in _MONEY_RE.finditer(str(txt)):
        raw = m.group(1) or m.group(2) or m.group(3)
        if raw:
            try:
                v = float(raw.replace(",", ""))
                if v > 0:
                    return v
            except ValueError:
                continue
    return None


def _date(text: str) -> Optional[datetime.date]:
    """Extrae la primera fecha válida de una cadena de texto."""
    for m in _DATE_RE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]:
                mes_str = g[1].lower().strip()
                mes_num = _MESES.get(mes_str) or _MESES.get(mes_str[:3])
                if not mes_num:
                    for k, v in _MESES.items():
                        if mes_str.startswith(k[:3]):
                            mes_num = v
                            break
                if mes_num:
                    return datetime.date(int(g[2]), mes_num, int(g[0]))
            if g[3]:
                return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            if g[6]:
                yr = int(g[8])
                return datetime.date(
                    yr + 2000 if yr < 100 else yr,
                    int(g[7]),
                    int(g[6]),
                )
        except (ValueError, KeyError):
            continue
    return None


# ─────────────────────────────────────────────────────────────
# CONSOLIDACIÓN DE PDFs (M1)
# ─────────────────────────────────────────────────────────────
def _build_combined_pdf(
    file_list: list[dict],
) -> tuple[bytes, str, int, list[dict]]:
    """Concatena múltiples PDFs en un solo documento PyMuPDF.

    Args:
        file_list: lista de dicts con keys 'name' y 'bytes'.

    Returns:
        (combined_bytes, combined_hash, total_pages, page_map)
        page_map es una lista de dicts, uno por página global:
          {"file_idx": int, "local_page": int, "name": str}
    """
    combined = fitz.open()
    page_map: list[dict] = []

    for file_idx, f_info in enumerate(file_list):
        src = fitz.open(stream=f_info["bytes"], filetype="pdf")
        n = len(src)
        combined.insert_pdf(src)
        for local_p in range(n):
            page_map.append(
                {
                    "file_idx": file_idx,
                    "local_page": local_p,
                    "name": f_info["name"],
                }
            )
        src.close()

    buf = io.BytesIO()
    combined.save(buf)
    combined.close()
    combined_bytes = buf.getvalue()
    combined_hash = _md5(combined_bytes)
    total_pages = len(page_map)

    return combined_bytes, combined_hash, total_pages, page_map


# ─────────────────────────────────────────────────────────────
# OCR — ejecución síncrona optimizada (F4, F9)
# ─────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False, max_entries=1)
def _get_ocr():
    """Carga el motor OCR RapidOCR una sola vez y lo cachea."""
    try:
        from rapidocr_onnxruntime import RapidOCR

        return RapidOCR(
            det_model_dir=None, rec_model_dir=None, cls_model_dir=None
        )
    except Exception as exc:
        st.warning(f"RapidOCR no disponible ({exc}). OCR desactivado.")
        return None


def _ocr_page(pdf_bytes: bytes, idx: int) -> str:
    """Ejecuta OCR sobre una página individual del PDF."""
    ocr = _get_ocr()
    if ocr is None:
        return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(1.0, 1.0),
                colorspace=fitz.csRGB,
                alpha=False,
            )
            img = np.frombuffer(pix.samples, np.uint8).reshape(
                pix.height, pix.width, 3
            )[:, :, ::-1]
            result, _ = ocr(img)
            if result:
                return "\n".join(
                    r[1] for r in result if r and len(r) > 1
                )
            return ""
    except Exception as exc:
        st.warning(f"Error OCR pág. {idx + 1}: {exc}")
        return ""


# ─────────────────────────────────────────────────────────────
# VISOR DE PDF — renderizado con caché (F3)
# ─────────────────────────────────────────────────────────────
@st.cache_data(max_entries=200, show_spinner=False)
def _render(pdf_hash: str, idx: int) -> bytes:
    """Renderiza una página del PDF consolidado como PNG.

    F3: pdf_bytes se obtiene de session_state; la invalidación
    se garantiza porque pdf_hash cambia cuando cambia el archivo.
    """
    pdf_bytes = st.session_state.pdf_combined
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(
            matrix=fitz.Matrix(1.5, 1.5),
            colorspace=fitz.csRGB,
            alpha=False,
        )
        return pix.tobytes("png")


# ─────────────────────────────────────────────────────────────
# API BANXICO SIE — con caché negativo (F10)
# ─────────────────────────────────────────────────────────────
def _banxico_tc(
    moneda: str, fecha: datetime.date, token: str
) -> Optional[float]:
    """Consulta el tipo de cambio FIX del Banco de México."""
    if moneda not in _BANXICO_SERIES or not token:
        return None

    cache_key = (moneda, fecha.isoformat())
    if cache_key in st.session_state.bx_cache:
        return st.session_state.bx_cache[cache_key]

    serie = _BANXICO_SERIES[moneda]
    f_ini = (fecha - datetime.timedelta(days=5)).isoformat()
    f_fin = fecha.isoformat()
    url = f"{_BANXICO_URL}/series/{serie}/datos/{f_ini}/{f_fin}"
    headers = {"Bmx-Token": token, "Accept": "application/json"}

    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        datos = resp.json()["bmx"]["series"][0]["datos"]
        if not datos:
            st.session_state.bx_cache[cache_key] = None
            return None
        tc = float(datos[-1]["dato"].replace(",", ""))
        st.session_state.bx_cache[cache_key] = tc
        return tc
    except Exception as exc:
        st.warning(f"Banxico ({moneda} · {fecha}): {exc}")
        st.session_state.bx_cache[cache_key] = None
        return None


# ─────────────────────────────────────────────────────────────
# RECALCULAR COLUMNAS DERIVADAS
# ─────────────────────────────────────────────────────────────
def recalc_derived(df: pd.DataFrame) -> pd.DataFrame:
    """Recalcula 'Diferencia final' = Total - Anexo."""
    if df is None or df.empty:
        return df
    df = df.copy()
    tot = pd.to_numeric(df["Total con IVA"], errors="coerce")
    anx = pd.to_numeric(df["Monto en Anexo Escrito"], errors="coerce")
    mask = tot.notna() | anx.notna()
    df.loc[mask, "Diferencia final"] = (
        tot.fillna(0) - anx.fillna(0)
    )[mask]
    df.loc[~mask, "Diferencia final"] = None
    return df


# ─────────────────────────────────────────────────────────────
# MOTOR DE EXTRACCIÓN
# ─────────────────────────────────────────────────────────────
def _parse_space_table(text: str) -> list[dict]:
    """Parsea líneas con formato tabular separado por espacios."""
    items = []
    for line in text.splitlines():
        m = _ITEM_RE.match(line.strip())
        if m:
            qty, desc, pu, total = m.groups()
            items.append(
                {
                    "qty": int(qty),
                    "desc": desc.strip(),
                    "pu": float(pu.replace(",", "")),
                    "total": float(total.replace(",", "")),
                }
            )
    return items


def extract(
    pdf_bytes: bytes,
    label: str,
    p0: int,
    p1: int,
    det_iva: bool,
    calc_sub: bool,
    moneda: str = "MXN",
    bx_token: str = "",
) -> dict:
    """Motor principal de extracción de montos desde PDF.

    Estrategia en 5 niveles:
      1. Tablas estructuradas (pdfplumber.extract_tables)
      2. Reconstrucción qty × pu ≈ total (R1: line totals)
      3. Texto libre con keyword matching
      4. Patrón "es de $X" (precio directo)
      5. Fallback: máximo valor en contexto monetario (R2)
    """
    native_text = ""
    table_rows: list[list[str]] = []
    ocr_used = False

    # ── Nivel 1 & 2: texto nativo + tablas ───────────────────
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        pr = range(max(0, p0 - 1), min(p1, n_pages))
        for i in pr:
            pg = pdf.pages[i]
            txt = pg.extract_text() or ""
            native_text += "\n" + txt
            for tbl in pg.extract_tables() or []:
                if tbl:
                    table_rows.extend(
                        [str(c or "").strip() for c in row]
                        for row in tbl
                        if row
                    )
            for item in _parse_space_table(txt):
                table_rows.append(
                    [
                        str(item["qty"]),
                        item["desc"],
                        f"${item['pu']:,.2f}",
                        f"${item['total']:,.2f}",
                    ]
                )

    # ── Nivel 3: OCR condicional ─────────────────────────────
    pages_count = max(len(pr), 1)
    if len(native_text.strip()) < NATIVE_MIN_CHARS_PER_PAGE * pages_count:
        ocr_used = True
        for i in pr:
            o_txt = _ocr_page(pdf_bytes, i)
            if o_txt:
                native_text += "\n" + o_txt
                for item in _parse_space_table(o_txt):
                    table_rows.append(
                        [
                            str(item["qty"]),
                            item["desc"],
                            f"${item['pu']:,.2f}",
                            f"${item['total']:,.2f}",
                        ]
                    )

    text = native_text
    tlow = text.lower()

    iva_f = (
        "Sí"
        if det_iva and re.search(r"\biva\b|16%|vat", tlow)
        else "N/M"
    )
    tot = iva = sub = pu = fecha = None
    qty = 1
    obs_parts: list[str] = []
    if ocr_used:
        obs_parts.append("OCR")

    # ── Paso 1: búsqueda en tablas estructuradas ─────────────
    for row in table_rows:
        j = "   ".join(row).lower()
        s = "   ".join(row)
        if re.search(r"\btotal\b", j) and not re.search(
            r"sub|parcial|acum", j
        ):
            v = _money(s)
            if v and (tot is None or v > tot):
                tot = v
        if re.search(r"\biva\b|16%|vat", j):
            v = _money(s)
            if v and (iva is None or v > iva):
                iva = v
        if re.search(r"subtotal|sin\s*iva|importe\s*neto", j):
            v = _money(s)
            if v and (sub is None or v > sub):
                sub = v

    # ── Paso 2 (R1): reconstrucción qty × pu ≈ total ────────
    valid_line_totals: list[float] = []
    seen_lines: set[tuple] = set()
    for row in table_rows:
        row_str = " ".join(str(c) for c in row)
        nums: list[float] = []
        for token in re.findall(r"[\d,]+(?:\.\d+)?", row_str):
            n = _safe_f(token.replace(",", ""))
            if n is not None and n > 0:
                nums.append(n)
        if len(nums) < 2:
            continue
        found = False
        for i_t, t_cand in enumerate(nums):
            if t_cand < 1.0 or found:
                continue
            for i_p, p_cand in enumerate(nums):
                if i_p == i_t or p_cand <= 0 or found:
                    continue
                for i_q, q_cand in enumerate(nums):
                    if i_q in (i_t, i_p):
                        continue
                    if not (
                        1 <= q_cand <= 9999
                        and q_cand == int(q_cand)
                    ):
                        continue
                    tol = max(0.5, t_cand * 0.01)
                    if abs(q_cand * p_cand - t_cand) <= tol:
                        key = (
                            int(q_cand),
                            round(p_cand, 2),
                            round(t_cand, 2),
                        )
                        if key not in seen_lines:
                            seen_lines.add(key)
                            valid_line_totals.append(t_cand)
                        found = True
                        break
                if found:
                    break

    if tot is None and valid_line_totals:
        tot = round(sum(valid_line_totals), 2)
        obs_parts.append("Total por líneas")

    # ── Paso 3: fecha y total por texto libre ────────────────
    for ln in text.splitlines():
        if fecha is None:
            d = _date(ln)
            if d:
                fecha = d

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b", ln, re.I) and not re.search(
                r"sub|parcial|acum", ln, re.I
            ):
                v = _money(ln)
                if v and (tot is None or v > tot):
                    tot = v

    # ── Paso 3.5: precio directo "es de $X" ──────────────────
    if tot is None:
        for ln in text.splitlines():
            m_ed = _ES_DE_RE.search(ln)
            if m_ed:
                v = _safe_f(m_ed.group(1).replace(",", ""))
                if v and 10.0 <= v <= MAX_PLAUSIBLE_MXN:
                    tot = v
                    obs_parts.append("Precio directo")
                    break

    # ── Paso 4 (R2): fallback — máximo con contexto monetario ─
    if tot is None:
        all_vals: list[float] = []
        for ln in text.splitlines():
            if _BUDGET_EXCL.search(ln):
                continue
            if not _MONETARY_CTX.search(ln):
                continue
            for m in _MONEY_RE.finditer(ln):
                raw = m.group(1) or m.group(2) or m.group(3)
                if raw:
                    try:
                        v = float(raw.replace(",", ""))
                        if 10.0 <= v <= MAX_PLAUSIBLE_MXN:
                            all_vals.append(v)
                    except ValueError:
                        pass
        if all_vals:
            tot = max(all_vals)
            obs_parts.append("Total inferido (máx.)")

    # ── F8: cota de cordura ──────────────────────────────────
    if tot is not None and tot > MAX_PLAUSIBLE_MXN:
        obs_parts.append(
            f"⚠ Valor sospechoso ({tot:,.2f}) — verificar"
        )
        tot = None

    # ── Subtotal por texto libre ─────────────────────────────
    if sub is None:
        for ln in text.splitlines():
            if re.search(
                r"subtotal|importe|sin\s*iva", ln, re.I
            ) and not re.search(
                r"\btotal\b", ln.replace("subtotal", ""), re.I
            ):
                v = _money(ln)
                if v and (tot is None or v <= tot):
                    sub = v
                    break

    # ── IVA por texto libre ──────────────────────────────────
    if iva is None and iva_f == "Sí":
        for ln in text.splitlines():
            if re.search(r"\biva\b|16%|vat|impuesto", ln, re.I):
                v = _money(ln)
                if v and (tot is None or v < tot):
                    iva = v
                    break

    # ── Cantidad y precio unitario desde tablas ──────────────
    for row in table_rows:
        row_str = " ".join(str(c) for c in row)
        nums_: list[float] = []
        for t in re.findall(r"[\d,]+(?:\.\d+)?", row_str):
            n_val = _safe_f(t)
            if n_val is not None:
                nums_.append(n_val)
        if len(nums_) >= 2 and 1 <= nums_[0] <= 9999:
            qty = int(nums_[0])
            pu = nums_[-2] if len(nums_) > 2 else nums_[-1]

    # ── Triangulación IVA / Subtotal / Total ─────────────────
    if tot and not sub and not iva and iva_f == "Sí":
        sub = round(tot / 1.16, 2)
        iva = round(tot - sub, 2)
        obs_parts.append("IVA desglosado")
    elif tot and iva and not sub:
        sub = round(tot - iva, 2)
    elif sub and iva and not tot:
        tot = round(sub + iva, 2)
    elif (
        calc_sub
        and sub
        and not iva
        and not tot
        and iva_f == "Sí"
    ):
        iva = round(sub * 0.16, 2)
        tot = round(sub + iva, 2)

    # F13: división segura por cero
    if pu is None and sub is not None and qty > 0:
        pu = round(sub / qty, 2)

    # ── Conversión de moneda con Banxico ─────────────────────
    tc: Optional[float] = None
    if moneda != "MXN" and bx_token:
        fecha_tc = fecha or datetime.date.today()
        tc = _banxico_tc(moneda, fecha_tc, bx_token)
        if tc and tot is not None:
            tot_orig = tot
            tot = round(tot * tc, 2)
            sub = round(sub * tc, 2) if sub else None
            iva = round(iva * tc, 2) if iva else None
            pu = round(pu * tc, 2) if pu else None
            obs_parts.append(f"1 {moneda} = ${tc:.4f} MXN")
            obs_parts.append(
                f"Total orig.: {moneda} {tot_orig:,.2f}"
            )

    tc_display = f"{moneda} ({tc:.4f})" if tc else moneda

    return {
        "Fecha": (
            fecha.isoformat()
            if fecha
            else datetime.date.today().isoformat()
        ),
        "Rubro": label,
        "QT": "Sí",
        "T. Cambio": tc_display,
        "(+ IVA)": iva_f,
        "Cantidad": qty,
        "Precio Unitario": pu,
        "Subtotal (Sin IVA)": sub,
        "IVA 16%": iva,
        "Total con IVA": tot,
        "Diferencia final": None,
        "Monto en Anexo Escrito": None,
        "Observaciones": (
            " | ".join(obs_parts) if obs_parts else ""
        ),
    }


# ─────────────────────────────────────────────────────────────
# EXPORTACIÓN EXCEL — P0 corregido, F6 fórmulas robustas
# ─────────────────────────────────────────────────────────────
def _side(style: str = "thin") -> Side:
    return Side(style=style, color="000000")


def _border(style: str = "thin") -> Border:
    s = _side(style)
    return Border(left=s, right=s, top=s, bottom=s)


def _fill(rgb: str) -> PatternFill:
    return PatternFill("solid", start_color=rgb, end_color=rgb)


def to_excel(
    df: pd.DataFrame, nombre: str = "", blank: bool = False
) -> bytes:
    """Genera archivo Excel con formato profesional.

    P0 FIX: un solo Workbook.
    F6: fórmulas con IF(ISNUMBER(…)) para tolerar celdas vacías.
    """
    buf = io.BytesIO()

    # P0 FIX: un solo Workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = nombre[:31] if nombre else "Conciliación"

    F_WHITE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    F_HDR = Font(name="Calibri", size=11, bold=True, color="000000")
    F_DATA = Font(name="Calibri", size=11, color="000000")
    F_TOT = Font(name="Calibri", size=11, bold=True, color="000000")
    F_BOLD = Font(name="Calibri", size=11, bold=True, color="000000")

    FILL_ROW1 = _fill("6E152E")
    FILL_HDR_A = _fill("D4C19C")
    FILL_HDR_B = _fill("EBE2D1")
    BD = _border("thin")
    BD_MED = _border("medium")
    AL_C = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    AL_L = Alignment(
        horizontal="left", vertical="center", wrap_text=True
    )
    AL_R = Alignment(horizontal="right", vertical="center")

    n_rows = len(df)

    for col_idx, width in enumerate(_WIDTHS_COL, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Fila 1: nombre del proyecto
    ws.row_dimensions[1].height = 27.75
    c = ws.cell(row=1, column=1, value=nombre or "")
    c.font = F_WHITE
    c.fill = FILL_ROW1
    c.border = BD
    c.alignment = AL_L
    for ci in range(2, len(_COLS) + 1):
        cx = ws.cell(row=1, column=ci)
        cx.fill = FILL_ROW1
        cx.border = BD

    # Fila 2: encabezados
    ws.row_dimensions[2].height = 21.75
    for ci, hdr in enumerate(_COLS, 1):
        c = ws.cell(row=2, column=ci, value=hdr)
        c.font = F_HDR
        c.fill = FILL_HDR_A if ci <= 11 else FILL_HDR_B
        c.border = BD
        c.alignment = AL_C

    DS = 3

    def _money_cell(ws_, r, ci, val=None):
        c_ = ws_.cell(row=r, column=ci, value=val)
        c_.number_format = _FMT_MONEY
        c_.font = F_DATA
        c_.border = BD
        c_.alignment = AL_R
        return c_

    for i, (_, row) in enumerate(df.iterrows()):
        r = DS + i
        ws.row_dimensions[r].height = 18

        # A: Fecha
        fv = row.get("Fecha")
        if isinstance(fv, str):
            try:
                fv = datetime.date.fromisoformat(fv)
            except Exception:
                fv = None
        if isinstance(fv, (datetime.date, datetime.datetime)):
            dt = (
                datetime.datetime.combine(fv, datetime.time())
                if isinstance(fv, datetime.date)
                else fv
            )
            c = ws.cell(row=r, column=1, value=dt)
            c.number_format = _FMT_DATE
        else:
            c = ws.cell(row=r, column=1, value=str(fv or ""))
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_C

        # B: Rubro
        c = ws.cell(
            row=r, column=2,
            value=str(row.get("Rubro", "") or ""),
        )
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_L

        # C: QT
        c = ws.cell(
            row=r, column=3, value=str(row.get("QT", "Sí"))
        )
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_C

        # D: T. Cambio
        c = ws.cell(
            row=r, column=4,
            value=str(row.get("T. Cambio", "MXN") or "MXN"),
        )
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_C

        # E: (+IVA)
        c = ws.cell(
            row=r, column=5,
            value=str(row.get("(+ IVA)", "") or ""),
        )
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_C

        # F: Cantidad
        q = _safe_f(row.get("Cantidad"))
        c = ws.cell(
            row=r, column=6,
            value=int(q) if q is not None else 1,
        )
        c.number_format = "0.00"
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_C

        # G: Precio Unitario (input)
        pu_val = (
            None if blank else _safe_f(row.get("Precio Unitario"))
        )
        _money_cell(ws, r, 7, pu_val)

        has_iva = (
            str(row.get("(+ IVA)", "N/M")).strip().lower() == "sí"
            and not blank
        )

        # H: Subtotal = F × G  (F6: ISNUMBER guard)
        c = ws.cell(
            row=r, column=8,
            value=f'=IF(ISNUMBER(G{r}),F{r}*G{r},"")',
        )
        c.number_format = _FMT_MONEY
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_R

        # I: IVA 16%
        if has_iva:
            c = ws.cell(
                row=r, column=9,
                value=f'=IF(ISNUMBER(H{r}),H{r}*0.16,"")',
            )
        else:
            c = ws.cell(row=r, column=9, value=None)
        c.number_format = _FMT_MONEY
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_R

        # J: Total con IVA
        if has_iva:
            c = ws.cell(
                row=r, column=10,
                value=(
                    f"=IF(ISNUMBER(H{r}),"
                    f"IF(ISNUMBER(I{r}),H{r}+I{r},H{r})"
                    f',"")'
                ),
            )
        else:
            c = ws.cell(
                row=r, column=10,
                value=f'=IF(ISNUMBER(H{r}),H{r},"")',
            )
        c.number_format = _FMT_MONEY
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_R

        # K: Diferencia = J - L
        c = ws.cell(
            row=r, column=11,
            value=(
                f"=IF(AND(ISNUMBER(J{r}),ISNUMBER(L{r})),"
                f'J{r}-L{r},"")'
            ),
        )
        c.number_format = _FMT_MONEY
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_R

        # L: Monto en Anexo Escrito (input)
        anx_val = (
            None
            if blank
            else _safe_f(row.get("Monto en Anexo Escrito"))
        )
        c = _money_cell(ws, r, 12, anx_val)
        c.font = F_BOLD

        # M: Observaciones
        c = ws.cell(
            row=r, column=13,
            value=(
                ""
                if blank
                else str(row.get("Observaciones", "") or "")
            ),
        )
        c.font = F_DATA
        c.border = BD
        c.alignment = AL_L

    # Fila de totales (F11)
    if n_rows > 0:
        tot_row = DS + n_rows
        s_xl = DS
        e_xl = DS + n_rows - 1
        ws.row_dimensions[tot_row].height = 18

        c = ws.cell(row=tot_row, column=1, value="TOTALES")
        c.font = F_TOT
        c.border = BD_MED

        for ci in (10, 11, 12):
            col_l = get_column_letter(ci)
            c = ws.cell(
                row=tot_row, column=ci,
                value=f"=SUM({col_l}{s_xl}:{col_l}{e_xl})",
            )
            c.number_format = _FMT_MONEY
            c.font = F_TOT
            c.fill = _fill("EBE2D1")
            c.border = BD_MED
            c.alignment = AL_R

    ws.freeze_panes = "A3"
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# CALLBACK NAVEGACIÓN (F2)
# ─────────────────────────────────────────────────────────────
def _go_to(target: int) -> None:
    """Mueve el visor a la página indicada, con clamping seguro."""
    tp = max(st.session_state.total_pages - 1, 0)
    st.session_state.current_page = max(0, min(tp, target))


# ─────────────────────────────────────────────────────────────
# SIDEBAR (M1: multi-PDF)
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Documentos PDF")
    uploaded = st.file_uploader(
        "Sube uno o varios PDFs de cotizaciones",
        type=["pdf"],
        accept_multiple_files=True,  # M1: multi-archivo
    )

    if uploaded:
        # F12: getvalue() — inmutable frente a rerenders
        new_files: list[dict] = []
        for f in uploaded:
            raw = f.getvalue()
            h = _md5(raw)
            with fitz.open(stream=raw, filetype="pdf") as doc:
                n = len(doc)
            new_files.append(
                {"name": f.name, "bytes": raw, "hash": h, "pages": n}
            )

        # Detectar cambios comparando hashes concatenados
        new_combo_hash = _md5(
            "".join(fi["hash"] for fi in new_files).encode()
        )
        if new_combo_hash != st.session_state.pdf_combined_hash:
            # Reconstruir PDF consolidado
            combined_bytes, combined_hash, total_pages, page_map = (
                _build_combined_pdf(new_files)
            )
            st.session_state.update(
                pdf_files=new_files,
                pdf_combined=combined_bytes,
                pdf_combined_hash=combined_hash,
                total_pages=total_pages,
                page_map=page_map,
                current_page=0,
                extracted=False,
                df=None,
                df_hash="",
            )

        # Resumen de archivos cargados
        n_files = len(st.session_state.pdf_files)
        tp_total = st.session_state.total_pages
        if n_files == 1:
            st.success(
                f"✅ 1 archivo · {tp_total} págs."
            )
        else:
            st.success(
                f"✅ {n_files} archivos · {tp_total} págs. totales"
            )
            with st.expander("📄 Detalle de archivos"):
                for idx, fi in enumerate(
                    st.session_state.pdf_files
                ):
                    st.caption(
                        f"**{idx + 1}.** {fi['name']}  —  "
                        f"{fi['pages']} pág(s)."
                    )

    st.markdown("---")
    st.markdown("### 🏷 Proyecto")
    st.session_state.proyecto_nombre = st.text_input(
        "Nombre del proyecto",
        value=st.session_state.proyecto_nombre,
    )

    # R7: enlace a documentación Banxico
    st.markdown("---")
    st.markdown("### 💱 Banxico – Tipo de Cambio")
    st.markdown(
        "Obtén tu token gratuito en "
        "[**SIE Banxico API** →]"
        "(https://www.banxico.org.mx/SieAPIRest/) "
        "*(solo necesario si hay cotizaciones en "
        "USD / EUR / CAD)*",
        unsafe_allow_html=False,
    )
    _token_default = (
        st.secrets.get("BANXICO_TOKEN", "")
        if hasattr(st, "secrets")
        else ""
    )
    bx_token = st.text_input(
        "Token Bmx-Token",
        value=_token_default,
        type="password",
        placeholder="Pega aquí tu token de Banxico",
    )
    if bx_token:
        st.caption("🔑 Token activo")

    st.markdown("---")
    st.markdown("### ⚙️ Secciones (cotizaciones)")
    n = int(
        st.number_input(
            "Número de secciones",
            min_value=1,
            max_value=50,
            value=st.session_state.num_sec,
            step=1,
        )
    )
    if n != st.session_state.num_sec:
        st.session_state.num_sec = n
        st.session_state.extracted = False
        st.session_state.df = None

    cfgs = list(st.session_state.sec_cfg)
    tp = max(st.session_state.total_pages, 1)

    # Generar opciones de PDF para el selectbox de cada sección
    pdf_names = [
        fi["name"] for fi in st.session_state.pdf_files
    ] if st.session_state.pdf_files else []

    while len(cfgs) < n:
        i = len(cfgs) + 1
        cfgs.append(
            {
                "label": f"Sección {i}",
                "p0": 1,
                "p1": 1,
                "det_iva": True,
                "calc_sub": True,
                "moneda": "MXN",
                "pdf_idx": 0,  # M2: índice del PDF fuente
            }
        )
    cfgs = cfgs[:n]

    for i, c in enumerate(cfgs):
        with st.expander(
            f"📄 Sección {i + 1}", expanded=(n <= 6)
        ):
            c["label"] = st.text_input(
                "Rubro / Concepto",
                value=c["label"],
                key=f"lb{i}",
            )

            # M2: selector de PDF fuente (solo si hay >1 archivo)
            if len(pdf_names) > 1:
                safe_idx = min(
                    c.get("pdf_idx", 0), len(pdf_names) - 1
                )
                c["pdf_idx"] = st.selectbox(
                    "PDF fuente",
                    range(len(pdf_names)),
                    index=safe_idx,
                    format_func=lambda x: pdf_names[x],
                    key=f"pdf{i}",
                )
            elif pdf_names:
                c["pdf_idx"] = 0
            else:
                c["pdf_idx"] = 0

            # Páginas: rango LOCAL dentro del PDF seleccionado
            if st.session_state.pdf_files:
                fi = st.session_state.pdf_files[c["pdf_idx"]]
                tp_local = max(fi["pages"], 1)
            else:
                tp_local = 1

            ca, cb = st.columns(2)
            c["p0"] = ca.number_input(
                "Pág. Inicio",
                1,
                tp_local,
                min(c["p0"], tp_local),
                key=f"p0{i}",
            )
            c["p1"] = cb.number_input(
                "Pág. Fin",
                c["p0"],
                tp_local,
                max(min(c["p1"], tp_local), c["p0"]),
                key=f"p1{i}",
            )
            c["moneda"] = st.selectbox(
                "Moneda de la cotización",
                ["MXN", "USD", "EUR", "CAD"],
                index=["MXN", "USD", "EUR", "CAD"].index(
                    c.get("moneda", "MXN")
                ),
                key=f"mon{i}",
            )
            c["det_iva"] = st.checkbox(
                "Detectar IVA",
                value=c["det_iva"],
                key=f"iv{i}",
            )
            # R3: restaurado
            c["calc_sub"] = st.checkbox(
                "Calcular subtotal si falta",
                value=c.get("calc_sub", True),
                key=f"cs{i}",
            )

    st.session_state.sec_cfg = cfgs

    st.markdown("---")
    run = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_combined is None),
        use_container_width=True,
        type="primary",
    )


# ─────────────────────────────────────────────────────────────
# CABECERA PRINCIPAL
# ─────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="hdr">
  <h1>📋 Conciliador de Cotizaciones</h1>
  <p>Extracción automática PDF · Carga múltiple ·
     OCR adaptativo · Tipo de cambio Banxico · v4.2</p>
</div>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────
# PROCESO DE EXTRACCIÓN (M2: usa PDF individual por sección)
# ─────────────────────────────────────────────────────────────
if run and st.session_state.pdf_combined:
    rows: list[dict] = []
    n_secs = st.session_state.num_sec
    bar = st.progress(0, text="Iniciando extracción…")

    for i, c in enumerate(st.session_state.sec_cfg):
        bar.progress(
            (i + 0.3) / n_secs, text=f"🔍 {c['label']}…"
        )

        # M2: seleccionar el PDF correcto para esta sección
        pdf_idx = c.get("pdf_idx", 0)
        if pdf_idx < len(st.session_state.pdf_files):
            section_pdf_bytes = st.session_state.pdf_files[pdf_idx]["bytes"]
        else:
            section_pdf_bytes = st.session_state.pdf_combined

        try:
            row = extract(
                section_pdf_bytes,
                c["label"],
                c["p0"],
                c["p1"],
                c["det_iva"],
                c.get("calc_sub", True),
                moneda=c.get("moneda", "MXN"),
                bx_token=bx_token,
            )
            rows.append(row)
        except Exception as exc:
            st.warning(
                f"⚠️ Sección {i + 1} «{c['label']}»: "
                f"{str(exc)[:120]}"
            )
            rows.append(
                {
                    **{k: None for k in _COLS},
                    "Rubro": c["label"],
                    "QT": "Sí",
                    "T. Cambio": c.get("moneda", "MXN"),
                    "Cantidad": 1,
                    "Fecha": datetime.date.today().isoformat(),
                    "Observaciones": f"Error: {str(exc)[:100]}",
                }
            )
        bar.progress((i + 1) / n_secs)

    bar.empty()

    df_new = pd.DataFrame(rows).reindex(columns=_COLS)
    df_new = recalc_derived(df_new)

    st.session_state.df = df_new
    st.session_state.df_hash = _df_hash(df_new)
    st.session_state.extracted = True
    st.session_state._rerun_guard = False
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")


if st.session_state.pdf_combined is None:
    st.info(
        "Sube uno o varios PDFs en la barra lateral para comenzar."
    )
    st.stop()


# ─────────────────────────────────────────────────────────────
# LAYOUT: VISOR + EDITOR
# ─────────────────────────────────────────────────────────────
col_L, col_R = st.columns(2, gap="medium")


# ── VISOR DE DOCUMENTO (M3: indicador de archivo fuente) ─────
with col_L:
    st.markdown(
        '<p class="ptitle">🔍 Visor de Documento</p>',
        unsafe_allow_html=True,
    )

    tp = st.session_state.total_pages

    nav_p, nav_c, nav_n = st.columns([1, 4, 1])
    nav_p.button(
        "◀",
        key="btn_prev",
        use_container_width=True,
        on_click=_go_to,
        args=(st.session_state.current_page - 1,),
    )
    nav_n.button(
        "▶",
        key="btn_next",
        use_container_width=True,
        on_click=_go_to,
        args=(st.session_state.current_page + 1,),
    )

    cp = st.session_state.current_page
    page_sel = nav_c.number_input(
        "Página",
        min_value=1,
        max_value=tp,
        value=cp + 1,
        step=1,
        label_visibility="collapsed",
        key="nav_page_input",
    )
    if page_sel - 1 != cp:
        st.session_state.current_page = page_sel - 1
        cp = page_sel - 1

    st.caption(f"Página {cp + 1} de {tp}")

    # M3: indicador de archivo fuente para la página actual
    pm = st.session_state.page_map
    if pm and cp < len(pm):
        pg_info = pm[cp]
        local_p = pg_info["local_page"] + 1
        fname = pg_info["name"]
        st.markdown(
            f'<span class="tag-file">📁 {fname}</span> '
            f'<span style="font-size:.8rem;color:#666">'
            f"pág. local {local_p}</span>",
            unsafe_allow_html=True,
        )

    # Indicador de sección activa (convertir a página global)
    for c_cfg in st.session_state.sec_cfg:
        # Calcular rango global de esta sección
        pdf_idx = c_cfg.get("pdf_idx", 0)
        # Offset global = suma de páginas de archivos anteriores
        global_offset = sum(
            fi["pages"]
            for fi in st.session_state.pdf_files[:pdf_idx]
        ) if st.session_state.pdf_files else 0
        g_p0 = global_offset + c_cfg["p0"]
        g_p1 = global_offset + c_cfg["p1"]
        if g_p0 <= cp + 1 <= g_p1:
            st.markdown(
                f'<span style="background:#6E152E;color:#fff;'
                f"padding:3px 10px;border-radius:4px;"
                f'font-size:.8rem">📑 {c_cfg["label"]}'
                f" · <span class=\"tag-moneda\">"
                f"{c_cfg.get('moneda', 'MXN')}</span></span>",
                unsafe_allow_html=True,
            )
            break

    with st.spinner("Cargando…"):
        st.image(
            _render(st.session_state.pdf_combined_hash, cp),
            use_container_width=True,
        )

    # R5: descarga de PDF(s)
    if len(st.session_state.pdf_files) == 1:
        st.download_button(
            "📥 Descargar PDF",
            data=st.session_state.pdf_files[0]["bytes"],
            file_name=st.session_state.pdf_files[0]["name"],
            mime="application/pdf",
            use_container_width=True,
        )
    elif len(st.session_state.pdf_files) > 1:
        # Descargar el consolidado
        st.download_button(
            "📥 Descargar PDF consolidado",
            data=st.session_state.pdf_combined,
            file_name="cotizaciones_consolidado.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        # Descargar individuales
        with st.expander("📄 Descargar archivos individuales"):
            for idx, fi in enumerate(st.session_state.pdf_files):
                st.download_button(
                    f"📥 {fi['name']}",
                    data=fi["bytes"],
                    file_name=fi["name"],
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"dl_pdf_{idx}",
                )


# ── EDITOR DE DATOS ──────────────────────────────────────────
with col_R:
    st.markdown(
        '<p class="ptitle">✏️ Editor de Datos</p>',
        unsafe_allow_html=True,
    )

    if not st.session_state.extracted or st.session_state.df is None:
        st.info(
            "Configura las secciones y presiona "
            "**🔍 Extraer Montos**."
        )

    else:
        df_cur = st.session_state.df
        if df_cur is None or df_cur.empty:
            st.warning(
                "Sin datos. Ejecuta nuevamente la extracción."
            )
            st.stop()

        kpi_slot = st.container()

        edited_df = st.data_editor(
            df_cur,
            key="data_editor_main",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha": st.column_config.TextColumn(
                    "Fecha", width="small"
                ),
                "Rubro": st.column_config.TextColumn(
                    "Rubro", width="large"
                ),
                "QT": st.column_config.SelectboxColumn(
                    "QT", options=["Sí", "No"], width="small"
                ),
                "T. Cambio": st.column_config.TextColumn(
                    "Moneda/TC", width="small"
                ),
                "(+ IVA)": st.column_config.SelectboxColumn(
                    "IVA",
                    options=["Sí", "No", "N/M"],
                    width="small",
                ),
                "Cantidad": st.column_config.NumberColumn(
                    "Cant.", format="%d", width="small"
                ),
                "Precio Unitario": st.column_config.NumberColumn(
                    "P. Unit.", format="$%.2f", width="medium"
                ),
                "Subtotal (Sin IVA)": st.column_config.NumberColumn(
                    "Subtotal", format="$%.2f", width="medium"
                ),
                "IVA 16%": st.column_config.NumberColumn(
                    "IVA 16%", format="$%.2f", width="medium"
                ),
                "Total con IVA": st.column_config.NumberColumn(
                    "Total c/IVA", format="$%.2f", width="medium"
                ),
                "Diferencia final": st.column_config.NumberColumn(
                    "Diferencia",
                    format="$%.2f",
                    width="medium",
                    disabled=True,
                ),
                "Monto en Anexo Escrito": (
                    st.column_config.NumberColumn(
                        "Anexo $", format="$%.2f", width="medium"
                    )
                ),
                "Observaciones": st.column_config.TextColumn(
                    "Observaciones", width="large"
                ),
            },
        )

        # P1 FIX: recalcular con guardia de reentrada
        updated_df = recalc_derived(edited_df)
        new_hash = _df_hash(updated_df)
        if new_hash != st.session_state.df_hash:
            st.session_state.df = updated_df
            st.session_state.df_hash = new_hash
            if not st.session_state.get("_rerun_guard"):
                st.session_state._rerun_guard = True
                st.rerun()
            else:
                st.session_state._rerun_guard = False
        else:
            st.session_state._rerun_guard = False

        df_cur = st.session_state.df

        # ── KPIs ─────────────────────────────────────────────
        ts_ = pd.to_numeric(
            df_cur["Total con IVA"], errors="coerce"
        ).sum()
        rs_ = pd.to_numeric(
            df_cur["Monto en Anexo Escrito"], errors="coerce"
        ).sum()
        dif = ts_ - rs_

        with kpi_slot:
            k1, k2, k3 = st.columns(3)
            for kol, val, lbl in [
                (k1, ts_, "Total Extraído"),
                (k2, rs_, "Monto Referencia"),
                (k3, dif, "Diferencia"),
            ]:
                color = (
                    "#6E152E"
                    if lbl != "Diferencia"
                    else (
                        "#c0392b" if abs(dif) > 0.01 else "#28a745"
                    )
                )
                kol.markdown(
                    f'<div class="kpi">'
                    f'<div class="v" style="color:{color}">'
                    f"${val:,.2f}</div>"
                    f'<div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

        # R6: alertas de extracción
        warn_rows = df_cur[
            df_cur["Observaciones"].str.contains(
                "⚠|OCR|inferido", na=False
            )
        ]
        if not warn_rows.empty:
            with st.expander(
                f"⚠ {len(warn_rows)} aviso(s) de extracción"
            ):
                for _, wr in warn_rows.iterrows():
                    st.caption(
                        f"• **{wr['Rubro']}**: "
                        f"{wr['Observaciones']}"
                    )

        st.markdown("---")

        # ── Descarga Excel ───────────────────────────────────
        ts_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        pname = (
            st.session_state.proyecto_nombre or "Cotizaciones"
        ).replace(" ", "_")
        nombre = st.session_state.proyecto_nombre

        try:
            xlsx_bytes = to_excel(df_cur, nombre=nombre)
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"Conciliacion_{pname}_{ts_str}.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True,
                type="primary",
            )
        except Exception as exc:
            st.error(f"Error generando Excel: {exc}")

        # R4: plantilla vacía
        with st.expander("📝 Edición manual directa en Excel"):
            st.info(
                "Descarga la **plantilla vacía**: Subtotal, "
                "IVA 16%, Total con IVA y Diferencia final se "
                "calculan automáticamente al escribir "
                "**Precio Unitario** (G) y "
                "**Monto en Anexo Escrito** (L)."
            )
            n_sec = st.session_state.num_sec
            df_blank = pd.DataFrame(
                {
                    "Fecha": [
                        datetime.date.today().isoformat()
                    ] * n_sec,
                    "Rubro": [
                        c["label"]
                        for c in st.session_state.sec_cfg
                    ],
                    "QT": ["Sí"] * n_sec,
                    "T. Cambio": [
                        c.get("moneda", "MXN")
                        for c in st.session_state.sec_cfg
                    ],
                    "(+ IVA)": ["Sí"] * n_sec,
                    "Cantidad": [1] * n_sec,
                    "Precio Unitario": [None] * n_sec,
                    "Subtotal (Sin IVA)": [None] * n_sec,
                    "IVA 16%": [None] * n_sec,
                    "Total con IVA": [None] * n_sec,
                    "Diferencia final": [None] * n_sec,
                    "Monto en Anexo Escrito": [None] * n_sec,
                    "Observaciones": [""] * n_sec,
                }
            )
            try:
                xlsx_blank = to_excel(
                    df_blank, nombre=nombre, blank=True
                )
                st.download_button(
                    "📄 Descargar plantilla vacía",
                    data=xlsx_blank,
                    file_name=f"Plantilla_{pname}.xlsx",
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"Error generando plantilla: {exc}")

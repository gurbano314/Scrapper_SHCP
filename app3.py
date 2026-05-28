import io
import re
import json
import hashlib
import datetime as dt
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from threading import Lock
from typing import Optional, Any

import fitz
import numpy as np
import pandas as pd
import pdfplumber
import requests
import streamlit as st
import xlsxwriter

# ============================================================
# SCRAPER HÍBRIDO DE COTIZACIONES PDF
# Texto nativo + OCR + Excel + Banxico
# Sin IA
# ============================================================

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN GENERAL
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
    [data-testid="stSidebar"] { background:#1a2744; }
    [data-testid="stSidebar"] * { color:#e8e8e8 !important; }
    .hdr {
        background:linear-gradient(90deg,#1a2744,#2d4a8f);
        padding:14px 22px; border-radius:8px; margin-bottom:18px;
    }
    .hdr h1 { color:#fff; font-size:1.4rem; margin:0; font-weight:700; }
    .hdr p  { color:#a8c0ff; font-size:.87rem; margin:4px 0 0; }
    .ptitle {
        background:#2d4a8f; color:#fff !important; font-weight:700;
        font-size:.88rem; padding:7px 14px; border-radius:6px 6px 0 0; margin-bottom:4px;
    }
    .kpi { background:#f0f4ff; border:1px solid #c5d3f5;
           border-radius:8px; padding:10px 12px; text-align:center; margin-bottom:6px; }
    .kpi .v { font-size:1.2rem; font-weight:700; color:#1a2744; }
    .kpi .l { font-size:.72rem; color:#5566aa; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────
NATIVE_MIN_CHARS_PER_PAGE = 80
OCR_SCALE = 2.0
BANXICO_TIMEOUT = 20

BANXICO_SERIES = {
    "USD": "SF43718",  # Peso por dólar FIX
    "EUR": "SF46410",
    "JPY": "SF46406",
    "GBP": "SF46407",
    "CAD": "SF60632",
}

SUPPORTED_CURRENCIES = ["MXN", "USD", "EUR", "JPY", "GBP", "CAD"]

_GEMINI_REPLACEMENT_NOTE = "Sin IA"

_MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

_COLS = [
    "Fecha",
    "Rubro",
    "Moneda",
    "Tipo Cambio",
    "QT",
    "T. Cambio",
    "(+ IVA)",
    "Cantidad",
    "Precio Unitario",
    "Subtotal (Sin IVA)",
    "IVA 16%",
    "Total con IVA",
    "Total MXN",
    "Diferencia final",
    "Monto en Anexo Escrito",
    "Observaciones",
]

_WIDTHS = [
    12, 42, 9, 12, 5, 10, 8, 8, 16, 18, 14, 16, 16, 18, 22, 50
]

MONEY_FMT = '_-"$"* #,##0.00_-;_-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'

_MONEY_RE = re.compile(
    r"\$\s*([-]?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)|"
    r"(?<!\d)([-]?\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?)"
)

_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:de)?\s*([a-záéíóúñA-ZÁÉÍÓÚÑ]+)[,\s]*(?:del?\s*)?(\d{4})|"
    r"(\d{4})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{1,2})|"
    r"(\d{1,2})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{2,4})"
)

_ITEM_RE = re.compile(
    r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
_SS_DEFAULTS: dict[str, Any] = {
    "pdf_bytes": None,
    "pdf_hash": None,
    "total_pages": 0,
    "current_page": 0,
    "num_sec": 1,
    "sec_cfg": [],
    "df": None,
    "extracted": False,
    "editor_df": None,
    "editor_initialized": False,
    "banxico_token": "",
    "banxico_cache": {},
}
for _k, _v in _SS_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)

# ─────────────────────────────────────────────────────────────
# UTILIDADES FINANCIERAS
# ─────────────────────────────────────────────────────────────
def D(x: Any) -> Optional[Decimal]:
    if x is None:
        return None
    try:
        s = str(x).strip().replace(",", "")
        if s == "" or s.lower() in {"nan", "none"}:
            return None
        return Decimal(s)
    except Exception:
        return None


def q2(x: Decimal) -> Decimal:
    return x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def md5_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def to_float_safe(v: Any) -> Optional[float]:
    if v is None:
        return None
    try:
        f = float(str(v).replace(",", "").strip())
        return None if f != f else f
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
# OCR ENGINE
# ─────────────────────────────────────────────────────────────
_OCR_LOCK = Lock()


@st.cache_resource(show_spinner=False)
def get_ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR(det_model_dir=None, rec_model_dir=None, cls_model_dir=None)
    except Exception as exc:
        st.warning(f"Motor OCR no disponible ({exc}). Solo se usará texto nativo.")
        return None


def ocr_page(pdf_bytes: bytes, idx: int) -> str:
    ocr = get_ocr_engine()
    if ocr is None:
        return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            idx = max(0, min(idx, len(doc) - 1))
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(OCR_SCALE, OCR_SCALE),
                colorspace=fitz.csRGB,
                alpha=False,
            )
            img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)[:, :, ::-1]
            with _OCR_LOCK:
                result, _ = ocr(img)
            return "\n".join(r[1] for r in result if r and len(r) > 1) if result else ""
    except Exception as exc:
        st.warning(f"OCR página {idx+1}: {exc}")
        return ""


# ─────────────────────────────────────────────────────────────
# HELPERS DE TEXTO / FECHA / MONEDA
# ─────────────────────────────────────────────────────────────
def extract_money_values(text: str) -> list[Decimal]:
    values: list[Decimal] = []
    for m in _MONEY_RE.finditer(str(text)):
        raw = m.group(1) or m.group(2)
        if not raw:
            continue
        try:
            values.append(Decimal(raw.replace(",", "")))
        except Exception:
            continue
    return values


def detect_currency(text: str) -> str:
    t = text.upper()
    if any(k in t for k in ["USD", "DÓLAR", "DOLAR", "DÓLARES", "DOLARES"]):
        return "USD"
    if any(k in t for k in ["EUR", "EURO", "EUROS"]):
        return "EUR"
    if any(k in t for k in ["JPY", "YEN"]):
        return "JPY"
    if any(k in t for k in ["GBP", "LIBRA", "STERLING"]):
        return "GBP"
    if any(k in t for k in ["CAD", "CANADIENSE", "CANADÁ", "CANADA"]):
        return "CAD"
    return "MXN"


def detect_date(text: str) -> Optional[dt.date]:
    low = text.lower()
    for m in _DATE_RE.finditer(low):
        g = m.groups()
        try:
            # 1) 26 de marzo del 2026
            if g[0] and g[1] and g[2]:
                day = int(g[0])
                month_txt = g[1].lower().strip(".,")
                year = int(g[2])
                month = _MONTHS.get(month_txt)
                if not month:
                    for k, v in _MONTHS.items():
                        if month_txt.startswith(k[:4]) or k.startswith(month_txt[:4]):
                            month = v
                            break
                if month:
                    return dt.date(year, month, day)
            # 2) 2026-03-26
            if g[3] and g[4] and g[5]:
                return dt.date(int(g[3]), int(g[4]), int(g[5]))
            # 3) 26/03/26
            if g[6] and g[7] and g[8]:
                year = int(g[8])
                year = 2000 + year if year < 100 else year
                return dt.date(year, int(g[7]), int(g[6]))
        except Exception:
            continue
    return None


def detect_textual_context_currency(text: str) -> str:
    t = text.upper()
    if "M.N." in t or "MONEDA NACIONAL" in t or "PESOS" in t:
        return "MXN"
    return detect_currency(text)


def parse_space_table(text: str) -> list[dict[str, Any]]:
    items = []
    for line in text.splitlines():
        m = _ITEM_RE.match(line.strip())
        if m:
            qty, desc, pu, total = m.groups()
            items.append(
                {
                    "qty": int(qty),
                    "desc": desc.strip(),
                    "pu": Decimal(pu.replace(",", "")),
                    "total": Decimal(total.replace(",", "")),
                }
            )
    return items


# ─────────────────────────────────────────────────────────────
# BANXICO
# ─────────────────────────────────────────────────────────────
class BanxicoService:
    def __init__(self, token: str):
        self.token = token.strip()

    def get_exchange_rate(self, currency: str, date_value: dt.date) -> Optional[Decimal]:
        if currency == "MXN":
            return Decimal("1")
        if not self.token:
            return None
        series = BANXICO_SERIES.get(currency)
        if not series:
            return None
        cache_key = (currency, date_value.isoformat())
        cached = st.session_state.banxico_cache.get(cache_key)
        if cached is not None:
            return cached

        # Banxico entrega series históricas por rango de fechas.
        # Tomamos la fecha exacta y, si no existe dato, buscamos hacia atrás un máximo de 7 días.
        headers = {"Bmx-Token": self.token}
        for offset in range(0, 8):
            d = date_value - dt.timedelta(days=offset)
            url = (
                "https://www.banxico.org.mx/SieAPIRest/service/v1/"
                f"series/{series}/datos/{d.isoformat()}/{d.isoformat()}"
            )
            try:
                resp = requests.get(url, headers=headers, timeout=BANXICO_TIMEOUT)
                resp.raise_for_status()
                data = resp.json()
                arr = data.get("bmx", {}).get("series", [])
                if arr and arr[0].get("datos"):
                    raw = arr[0]["datos"][0].get("dato")
                    rate = Decimal(str(raw).replace(",", ""))
                    st.session_state.banxico_cache[cache_key] = rate
                    return rate
            except Exception:
                continue
        return None


# ─────────────────────────────────────────────────────────────
# EXTRACCIÓN HÍBRIDA
# ─────────────────────────────────────────────────────────────
def extract_page_or_block(pdf_bytes: bytes, page_from: int, page_to: int) -> tuple[str, list[list[str]], bool]:
    native_text = ""
    table_rows: list[list[str]] = []
    ocr_used = False

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        start = max(0, page_from - 1)
        end = min(page_to, n_pages)
        page_indexes = range(start, end)
        for i in page_indexes:
            pg = pdf.pages[i]
            txt = pg.extract_text() or ""
            native_text += "\n" + txt
            for tbl in (pg.extract_tables() or []):
                if tbl:
                    table_rows.extend([ [str(c or "").strip() for c in row] for row in tbl if row ])
            for item in parse_space_table(txt):
                table_rows.append([
                    str(item["qty"]),
                    item["desc"],
                    f"${item['pu']:,}",
                    f"${item['total']:,}",
                ])

    pages_count = max(len(list(page_indexes)), 1)
    if len(native_text.strip()) < NATIVE_MIN_CHARS_PER_PAGE * pages_count:
        ocr_used = True
        for i in page_indexes:
            o_txt = ocr_page(pdf_bytes, i)
            if o_txt:
                native_text += "\n" + o_txt
                for item in parse_space_table(o_txt):
                    table_rows.append([
                        str(item["qty"]),
                        item["desc"],
                        f"${item['pu']:,}",
                        f"${item['total']:,}",
                    ])

    return native_text, table_rows, ocr_used


def recalc_derived(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in ["Total con IVA", "Monto en Anexo Escrito", "Subtotal (Sin IVA)", "IVA 16%", "Total MXN", "Precio Unitario"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    tot = df.get("Total con IVA")
    anex = df.get("Monto en Anexo Escrito")
    if tot is not None and anex is not None:
        mask = tot.notna() | anex.notna()
        df.loc[mask, "Diferencia final"] = (tot.fillna(0) - anex.fillna(0))[mask]
        df.loc[~mask, "Diferencia final"] = None
    else:
        df["Diferencia final"] = None
    return df


def build_record(
    text: str,
    table_rows: list[list[str]],
    label: str,
    currency: str,
    date_value: Optional[dt.date],
    banxico_rate: Optional[Decimal],
    det_iva: bool,
    calc_sub: bool,
    ocr_used: bool,
) -> dict[str, Any]:
    tlow = text.lower()
    iva_flag = "Sí" if det_iva and re.search(r"\biva\b|16%|vat|impuesto", tlow, re.I) else "N/M"

    tot: Optional[Decimal] = None
    iva: Optional[Decimal] = None
    sub: Optional[Decimal] = None
    pu: Optional[Decimal] = None
    qty: int = 1

    obs_parts: list[str] = []
    if ocr_used:
        obs_parts.append("OCR")
    if currency != "MXN":
        obs_parts.append(f"Moneda {currency}")
        if banxico_rate is not None:
            obs_parts.append(f"TC Banxico {banxico_rate}")
        else:
            obs_parts.append("Sin TC Banxico")

    # Paso 1: tablas estructuradas
    for row in table_rows:
        j = "   ".join(row).lower()
        s = "   ".join(row)
        money_vals = extract_money_values(s)

        if "total" in j and not re.search(r"sub|parcial|acum", j):
            if money_vals:
                candidate = max(money_vals)
                if tot is None or candidate > tot:
                    tot = candidate

        if re.search(r"\biva\b|16%|vat|impuesto", j):
            if money_vals:
                candidate = max(money_vals)
                if iva is None or candidate > iva:
                    iva = candidate

        if re.search(r"subtotal|sin\s*iva|importe\s*neto", j):
            if money_vals:
                candidate = max(money_vals)
                if sub is None or candidate > sub:
                    sub = candidate

    # Paso 2: líneas de producto
    valid_line_totals: list[Decimal] = []
    seen: set[tuple[int, Decimal, Decimal]] = set()
    for row in table_rows:
        row_str = " ".join(str(c) for c in row)
        nums: list[Decimal] = []
        for token in re.findall(r"[-]?[\d,]+(?:\.\d+)?", row_str):
            val = D(token)
            if val is not None and val > 0:
                nums.append(val)
        if len(nums) < 2:
            continue
        found = False
        for i_t, t_cand in enumerate(nums):
            if found or t_cand < 1:
                continue
            for i_p, p_cand in enumerate(nums):
                if found or i_p == i_t or p_cand <= 0:
                    continue
                for i_q, q_cand in enumerate(nums):
                    if i_q in (i_t, i_p):
                        continue
                    if not (1 <= q_cand <= 9999 and q_cand == int(q_cand)):
                        continue
                    tol = max(Decimal("0.50"), q2(t_cand * Decimal("0.01")))
                    if abs((q_cand * p_cand) - t_cand) <= tol:
                        key = (int(q_cand), q2(p_cand), q2(t_cand))
                        if key not in seen:
                            seen.add(key)
                            valid_line_totals.append(t_cand)
                        found = True
                        break
                if found:
                    break

    if tot is None and valid_line_totals:
        tot = q2(sum(valid_line_totals, Decimal("0")))
        obs_parts.append("Total por líneas")

    # Paso 3: texto libre
    if date_value is None:
        for ln in text.splitlines():
            d = detect_date(ln)
            if d:
                date_value = d
                break

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b", ln, re.I) and not re.search(r"sub|parcial|acum", ln, re.I):
                vals = extract_money_values(ln)
                if vals:
                    candidate = max(vals)
                    if tot is None or candidate > tot:
                        tot = candidate

    if tot is None:
        all_vals: list[Decimal] = []
        for row in table_rows:
            for cell in row:
                vals = extract_money_values(cell)
                for v in vals:
                    if v >= Decimal("10"):
                        all_vals.append(v)
        for ln in text.splitlines():
            for v in extract_money_values(ln):
                if v >= Decimal("10"):
                    all_vals.append(v)
        if all_vals:
            tot = max(all_vals)
            obs_parts.append("Total inferido (máx.)")

    if sub is None:
        for ln in text.splitlines():
            if re.search(r"subtotal|importe|sin\s*iva", ln, re.I) and not re.search(r"\btotal\b", ln.replace("subtotal", ""), re.I):
                vals = extract_money_values(ln)
                if vals:
                    candidate = max(vals)
                    if tot is None or candidate <= tot:
                        sub = candidate
                        break

    if iva is None and iva_flag == "Sí":
        for ln in text.splitlines():
            if re.search(r"\biva\b|16%|vat|impuesto", ln, re.I):
                vals = extract_money_values(ln)
                if vals:
                    candidate = max(vals)
                    if tot is None or candidate < tot:
                        iva = candidate
                        break

    # Cantidad y precio unitario
    for row in table_rows:
        row_str = " ".join(str(c) for c in row)
        nums_: list[Decimal] = []
        for token in re.findall(r"[-]?[\d,]+(?:\.\d+)?", row_str):
            val = D(token)
            if val is not None:
                nums_.append(val)
        if len(nums_) >= 2 and 1 <= nums_[0] <= 9999:
            qty = int(nums_[0])
            if len(nums_) >= 3:
                pu = nums_[-2]
            else:
                pu = nums_[-1]

    # Triangulación
    if tot is not None and sub is None and iva is None and iva_flag == "Sí":
        sub = q2(tot / Decimal("1.16"))
        iva = q2(tot - sub)
        obs_parts.append("IVA desglosado")
    elif tot is not None and iva is not None and sub is None:
        sub = q2(tot - iva)
    elif sub is not None and iva is not None and tot is None:
        tot = q2(sub + iva)
    elif calc_sub and sub is not None and iva is None and tot is None and iva_flag == "Sí":
        iva = q2(sub * Decimal("0.16"))
        tot = q2(sub + iva)

    if pu is None and sub is not None and qty:
        pu = q2(sub / Decimal(qty))

    # Conversión a MXN
    total_mxn: Optional[Decimal] = None
    if tot is not None:
        if currency == "MXN":
            total_mxn = tot
        elif banxico_rate is not None:
            total_mxn = q2(tot * banxico_rate)

    fecha_out = date_value.isoformat() if date_value else dt.date.today().isoformat()
    if not obs_parts:
        obs_parts.append("")

    return {
        "Fecha": fecha_out,
        "Rubro": label,
        "Moneda": currency,
        "Tipo Cambio": float(banxico_rate) if banxico_rate is not None else None,
        "QT": "Sí",
        "T. Cambio": currency,
        "(+ IVA)": iva_flag,
        "Cantidad": qty,
        "Precio Unitario": float(pu) if pu is not None else None,
        "Subtotal (Sin IVA)": float(sub) if sub is not None else None,
        "IVA 16%": float(iva) if iva is not None else None,
        "Total con IVA": float(tot) if tot is not None else None,
        "Total MXN": float(total_mxn) if total_mxn is not None else None,
        "Diferencia final": None,
        "Monto en Anexo Escrito": None,
        "Observaciones": " | ".join([x for x in obs_parts if x]),
    }


def extract_block(
    pdf_bytes: bytes,
    label: str,
    p0: int,
    p1: int,
    det_iva: bool,
    calc_sub: bool,
    banxico: Optional[BanxicoService] = None,
) -> dict[str, Any]:
    text, table_rows, ocr_used = extract_page_or_block(pdf_bytes, p0, p1)
    currency = detect_textual_context_currency(text)
    date_value = None
    for ln in text.splitlines():
        d = detect_date(ln)
        if d:
            date_value = d
            break
    rate = None
    if banxico is not None and currency != "MXN" and date_value is not None:
        rate = banxico.get_exchange_rate(currency, date_value)
    return build_record(text, table_rows, label, currency, date_value, rate, det_iva, calc_sub, ocr_used)


# ─────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("Conciliación")
    review = wb.add_worksheet("Revisión")
    audit = wb.add_worksheet("Auditoría")

    B = {"font_name": "Calibri", "font_size": 11, "border": 1}
    hdr = wb.add_format({**B, "bold": True, "bg_color": "#D4C19C", "align": "center", "valign": "vcenter"})
    hdr2 = wb.add_format({**B, "bold": True, "bg_color": "#EBE2D1", "align": "center", "valign": "vcenter"})
    txt = wb.add_format({**B, "text_wrap": True})
    num = wb.add_format({**B, "num_format": MONEY_FMT})
    intf = wb.add_format({**B, "num_format": "#,##0"})
    datef = wb.add_format({**B, "num_format": "dd/mm/yyyy"})
    bold_num = wb.add_format({**B, "bold": True, "num_format": MONEY_FMT})
    note = wb.add_format({**B, "text_wrap": True, "bg_color": "#F7F7F7"})
    red = wb.add_format({**B, "bg_color": "#FDEDEC"})

    for sheet in (ws, review):
        for c, h in enumerate(_COLS):
            sheet.write(0, c, h, hdr2 if c >= 14 else hdr)
        for i, w in enumerate(_WIDTHS):
            sheet.set_column(i, i, w)
        sheet.set_row(0, 28)

    # Hoja Conciliación: con fórmulas
    DROW = 1
    n_rows = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        r0 = DROW + i
        r1 = r0 + 1

        # Fecha
        fecha_val = row.get("Fecha")
        if isinstance(fecha_val, str):
            try:
                fecha_val = dt.date.fromisoformat(fecha_val)
            except Exception:
                fecha_val = None
        if isinstance(fecha_val, (dt.date, dt.datetime)):
            dtv = dt.datetime.combine(fecha_val, dt.time()) if isinstance(fecha_val, dt.date) and not isinstance(fecha_val, dt.datetime) else fecha_val
            ws.write_datetime(r0, 0, dtv, datef)
            review.write_datetime(r0, 0, dtv, datef)
        else:
            ws.write(r0, 0, str(fecha_val or ""), txt)
            review.write(r0, 0, str(fecha_val or ""), txt)

        for col_idx, col_name in enumerate(_COLS[1:], start=1):
            val = row.get(col_name)
            if col_name in {"Rubro", "Moneda", "QT", "T. Cambio", "(+ IVA)", "Observaciones"}:
                ws.write(r0, col_idx, str(val or ""), note if col_name == "Observaciones" else txt)
                review.write(r0, col_idx, str(val or ""), note if col_name == "Observaciones" else txt)
            elif col_name in {"Cantidad"}:
                q = to_float_safe(val)
                if q is not None:
                    ws.write_number(r0, col_idx, int(q), intf)
                    review.write_number(r0, col_idx, int(q), intf)
                else:
                    ws.write_blank(r0, col_idx, None, intf)
                    review.write_blank(r0, col_idx, None, intf)
            elif col_name in {"Tipo Cambio", "Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%", "Total con IVA", "Total MXN", "Diferencia final", "Monto en Anexo Escrito"}:
                v = to_float_safe(val)
                if v is not None:
                    ws.write_number(r0, col_idx, v, num)
                    review.write_number(r0, col_idx, v, num)
                else:
                    ws.write_blank(r0, col_idx, None, num)
                    review.write_blank(r0, col_idx, None, num)
            else:
                ws.write(r0, col_idx, "" if val is None else str(val), txt)
                review.write(r0, col_idx, "" if val is None else str(val), txt)

        # Fórmula de diferencia final
        ws.write_formula(r0, 13, f"=L{r1}-O{r1}", num)
        review.write_formula(r0, 13, f"=L{r1}-O{r1}", num)

    # Totales
    total_row = DROW + n_rows
    s_xl = DROW + 1
    e_xl = DROW + n_rows
    for ci in (9, 10, 11, 12, 13, 14):
        cl = chr(ord("A") + ci)
        ws.write_formula(total_row, ci, f"=SUM({cl}{s_xl}:{cl}{e_xl})", bold_num)

    # Hoja Auditoría
    audit.write(0, 0, "Regla", hdr)
    audit.write(0, 1, "Resultado", hdr)
    audit.write(0, 2, "Detalle", hdr)
    audit.set_column(0, 0, 26)
    audit.set_column(1, 1, 18)
    audit.set_column(2, 2, 80)
    audit.write(1, 0, "Origen", txt)
    audit.write(1, 1, "Automático", txt)
    audit.write(1, 2, "Extracción híbrida: texto nativo + OCR condicional.", txt)
    audit.write(2, 0, "Validación", txt)
    audit.write(2, 1, "Determinística", txt)
    audit.write(2, 2, "Cálculo de diferencias y consistencia basado en reglas financieras.", txt)

    wb.close()
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# NAVEGACIÓN
# ─────────────────────────────────────────────────────────────
def nav_prev():
    st.session_state.current_page = max(0, st.session_state.current_page - 1)


def nav_next():
    tp = max(st.session_state.total_pages - 1, 0)
    st.session_state.current_page = min(tp, st.session_state.current_page + 1)


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Documento PDF")
    up = st.file_uploader("Sube el archivo PDF", type=["pdf"])

    if up:
        raw = up.read()
        h = md5_bytes(raw)
        if h != st.session_state.pdf_hash:
            st.session_state.update(
                pdf_bytes=raw,
                pdf_hash=h,
                extracted=False,
                df=None,
                editor_df=None,
                editor_initialized=False,
                current_page=0,
            )
            with fitz.open(stream=raw, filetype="pdf") as d:
                st.session_state.total_pages = len(d)
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")

    st.markdown("---")
    st.markdown("### ⚙️ Secciones")
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
        st.session_state.editor_df = None
        st.session_state.editor_initialized = False

    cfgs = list(st.session_state.sec_cfg)
    tp = max(st.session_state.total_pages, 1)
    while len(cfgs) < n:
        i = len(cfgs) + 1
        cfgs.append({"label": f"Sección {i}", "p0": 1, "p1": 1, "det_iva": True, "calc_sub": True})
    del cfgs[n:]
    st.session_state.sec_cfg = cfgs

    for i, c in enumerate(st.session_state.sec_cfg):
        with st.expander(f"📄 Sección {i+1}", expanded=(n <= 5)):
            c["label"] = st.text_input("Rubro/Concepto", value=c["label"], key=f"lb{i}")
            col_a, col_b = st.columns(2)
            c["p0"] = col_a.number_input("Pág. Inicio", 1, tp, min(c["p0"], tp), key=f"p0{i}")
            c["p1"] = col_b.number_input("Pág. Fin", c["p0"], tp, max(min(c["p1"], tp), c["p0"]), key=f"p1{i}")
            c["det_iva"] = st.checkbox("Detectar IVA", value=c["det_iva"], key=f"iv{i}")
            c["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=c["calc_sub"], key=f"cs{i}")

    st.markdown("---")
    st.markdown("### 💱 Banxico")
    st.text_input(
        "Token Banxico",
        value=st.session_state.banxico_token,
        type="password",
        key="banxico_token_input",
        help="Token del API SIE de Banxico.",
    )
    st.session_state.banxico_token = st.session_state.banxico_token_input.strip()

    run = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_bytes is None),
        use_container_width=True,
        type="primary",
    )


# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="hdr">
      <h1>📋 Conciliador de Cotizaciones PDF</h1>
      <p>Extrae, valida y exporta montos desde documentos PDF · OCR automático · Banxico · Excel editable</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────
# EXTRACCIÓN
# ─────────────────────────────────────────────────────────────
if run and st.session_state.pdf_bytes:
    banxico = BanxicoService(st.session_state.banxico_token)
    rows: list[dict[str, Any]] = []
    n_secs = st.session_state.num_sec
    bar = st.progress(0, text="Iniciando extracción…")

    for i, c in enumerate(st.session_state.sec_cfg):
        bar.progress((i + 0.3) / max(n_secs, 1), text=f"🔍 {c['label']}…")
        try:
            row = extract_block(
                st.session_state.pdf_bytes,
                c["label"],
                int(c["p0"]),
                int(c["p1"]),
                bool(c["det_iva"]),
                bool(c["calc_sub"]),
                banxico=banxico,
            )
            rows.append(row)
        except Exception as exc:
            rows.append(
                {
                    "Fecha": dt.date.today().isoformat(),
                    "Rubro": c["label"],
                    "Moneda": "MXN",
                    "Tipo Cambio": 1.0,
                    "QT": "Sí",
                    "T. Cambio": "MXN",
                    "(+ IVA)": "N/M",
                    "Cantidad": 1,
                    "Precio Unitario": None,
                    "Subtotal (Sin IVA)": None,
                    "IVA 16%": None,
                    "Total con IVA": None,
                    "Total MXN": None,
                    "Diferencia final": None,
                    "Monto en Anexo Escrito": None,
                    "Observaciones": f"Error: {str(exc)[:150]}",
                }
            )
            st.warning(f"⚠️ Sección {i+1} «{c['label']}»: {str(exc)[:120]}")
        bar.progress((i + 1) / max(n_secs, 1))

    bar.empty()
    df_new = pd.DataFrame(rows)
    for col in _COLS:
        if col not in df_new.columns:
            df_new[col] = None
    df_new = df_new[_COLS]
    df_new = recalc_derived(df_new)

    st.session_state.df = df_new.copy()
    st.session_state.editor_df = df_new.copy()
    st.session_state.editor_initialized = True
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()


# ─────────────────────────────────────────────────────────────
# LAYOUT PRINCIPAL
# ─────────────────────────────────────────────────────────────
col_L, col_R = st.columns(2, gap="medium")

# ── VISOR ────────────────────────────────────────────────────
with col_L:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)

    tp = st.session_state.total_pages
    cp = st.session_state.current_page

    nav_p, nav_c, nav_n = st.columns([1, 4, 1])
    nav_p.button("◀", key="btn_prev", use_container_width=True, on_click=nav_prev)
    nav_n.button("▶", key="btn_next", use_container_width=True, on_click=nav_next)

    cp = st.session_state.current_page
    new_page = nav_c.number_input(
        "Página",
        min_value=1,
        max_value=max(tp, 1),
        value=cp + 1,
        step=1,
        label_visibility="collapsed",
        key="nav_page_input",
    )
    if new_page - 1 != cp:
        st.session_state.current_page = new_page - 1
        cp = new_page - 1

    st.caption(f"Página {cp + 1} de {tp}")

    for c in st.session_state.sec_cfg:
        if c["p0"] <= cp + 1 <= c["p1"]:
            st.markdown(
                f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;'
                f'border-radius:4px;font-size:.8rem">📑 {c["label"]}</span>',
                unsafe_allow_html=True,
            )
            break

    @st.cache_data(max_entries=80, show_spinner=False)
    def render_page(pdf_hash: str, idx: int) -> bytes:
        with fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf") as doc:
            idx = max(0, min(idx, len(doc) - 1))
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(1.5, 1.5),
                colorspace=fitz.csRGB,
                alpha=False,
            )
            return pix.tobytes("png")

    with st.spinner("Cargando…"):
        st.image(
            render_page(st.session_state.pdf_hash or "", cp),
            use_container_width=True,
        )

    st.download_button(
        "📥 Descargar PDF",
        data=st.session_state.pdf_bytes,
        file_name="documento.pdf",
        mime="application/pdf",
        use_container_width=True,
    )


# ── EDITOR ───────────────────────────────────────────────────
with col_R:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)

    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona 🔍 Extraer Montos.")
    else:
        kpi_slot = st.container()

        if not st.session_state.editor_initialized or st.session_state.editor_df is None:
            st.session_state.editor_df = st.session_state.df.copy()
            st.session_state.editor_initialized = True

        edited_df = st.data_editor(
            st.session_state.editor_df,
            key="data_editor_main",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha": st.column_config.TextColumn("Fecha (YYYY-MM-DD)", width="small"),
                "Rubro": st.column_config.TextColumn("Rubro", width="large"),
                "Moneda": st.column_config.SelectboxColumn("Moneda", options=SUPPORTED_CURRENCIES, width="small"),
                "Tipo Cambio": st.column_config.NumberColumn("TC Banxico", format="%.6f", width="small"),
                "QT": st.column_config.SelectboxColumn("QT", options=["Sí", "No"], width="small"),
                "T. Cambio": st.column_config.SelectboxColumn("T. Cambio", options=SUPPORTED_CURRENCIES, width="small"),
                "(+ IVA)": st.column_config.SelectboxColumn("(+ IVA)", options=["Sí", "No", "N/M"], width="small"),
                "Cantidad": st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario": st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)": st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%": st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA": st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Total MXN": st.column_config.NumberColumn("Total MXN", format="$%.2f", width="medium"),
                "Diferencia final": st.column_config.NumberColumn("Diferencia", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
            },
        )

        if edited_df is not None:
            recalculated = recalc_derived(edited_df)
            if not recalculated.equals(st.session_state.editor_df):
                st.session_state.editor_df = recalculated.copy()
                st.session_state.df = recalculated.copy()

        df_cur = st.session_state.df
        ts_ = pd.to_numeric(df_cur["Total con IVA"], errors="coerce").sum()
        rs_ = pd.to_numeric(df_cur["Monto en Anexo Escrito"], errors="coerce").sum()
        dif = ts_ - rs_

        with kpi_slot:
            k1, k2, k3 = st.columns(3)
            for col, val, lbl in [
                (k1, ts_, "Total Extraído"),
                (k2, rs_, "Monto Referencia"),
                (k3, dif, "Diferencia"),
            ]:
                color = "#1a2744" if lbl != "Diferencia" else ("#c0392b" if abs(dif) > 0.01 else "#28a745")
                col.markdown(
                    f'<div class="kpi"><div class="v" style="color:{color}">${val:,.2f}</div><div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

        st.markdown("---")

        try:
            xlsx_bytes = to_excel(st.session_state.df)
            ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"Conciliacion_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Error generando Excel: {exc}")

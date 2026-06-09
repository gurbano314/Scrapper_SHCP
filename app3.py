# ============================================================
# CONCILIADOR DE COTIZACIONES PDF  |  app.py  v3.0
# Luis Gustavo Urbano Jiménez · SHCP / EFIDEPORTE
# ============================================================
# Correcciones sobre v2.1:
#   F1  — Race condition data_editor: detección real de cambios via hash
#   F2  — Navegación: una sola fuente de verdad (on_click con args)
#   F3  — Caché de render: pdf_bytes fuera de la firma → sin fuga de memoria
#   F4  — OCR con timeout + ThreadPoolExecutor
#   F5  — Regex grupo 3: solo captura números con decimales .XX
#   F6  — Fórmulas Excel siempre escritas (nunca sobreescritas por valores)
#   F7  — Guard explícito sobre df_cur antes de KPIs
#   F8  — Triangulación IVA protegida con cota de cordura
#   F9  — cache_resource con max_entries=1
#   F10 — Parser JSON de IA robusto con re.search()
#   F11 — Fila de totales Excel: rango s_xl..e_xl siempre < tot_row
# Nuevas funciones:
#   ● API Banxico SIE: tipo de cambio USD/EUR/CAD por fecha de cotización
#   ● Plantilla Excel fiel al formato PAR (fila 1 burdeos, fórmulas en H/I/J/K)
#   ● Opción "plantilla vacía" para edición manual directa en Excel
#   ● Moneda por sección + conversión automática a MXN al exportar
# ============================================================

import concurrent.futures
import contextlib
import datetime
import hashlib
import io
import json
import re
from typing import Optional

import fitz
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

st.markdown("""
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
    font-size:.88rem; padding:7px 14px; border-radius:6px 6px 0 0; margin-bottom:4px;
}
.kpi { background:#fdf5f7; border:1px solid #e8b4c0;
       border-radius:8px; padding:10px 12px; text-align:center; margin-bottom:6px; }
.kpi .v { font-size:1.2rem; font-weight:700; color:#6E152E; }
.kpi .l { font-size:.72rem; color:#9a4060; }
.tag-moneda {
    background:#6E152E; color:#fff !important; padding:2px 8px;
    border-radius:4px; font-size:.78rem; font-weight:600;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────
NATIVE_MIN_CHARS_PER_PAGE = 80
OCR_TIMEOUT_S             = 25
MAX_PLAUSIBLE_MXN         = 50_000_000          # $50 MDP – cota de cordura (F8)

# Banxico SIE – series de tipo de cambio FIX
_BANXICO_SERIES = {
    "USD": "SF43718",
    "EUR": "SF46410",
    "CAD": "SF60653",
}
_BANXICO_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1"

# Columnas del modelo de datos (= plantilla Excel)
_COLS = [
    "Fecha", "Rubro", "QT", "T. Cambio", "(+ IVA)", "Cantidad",
    "Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%", "Total con IVA",
    "Diferencia final", "Monto en Anexo Escrito", "Observaciones",
]
_WIDTHS_COL = [12, 52, 5, 10, 7, 8, 16, 18, 17, 16, 18, 22, 48]

# F5 CORREGIDO: grupo 3 exige decimales .XX → no captura folios/teléfonos
_MONEY_RE = re.compile(
    r"\$\s*([\d,]+(?:\.\d{1,2})?)"                      # con signo $
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)"    # con separador de miles
    r"|(?<!\d)(\d{1,9}\.\d{2})(?!\d)"                   # solo si tiene .XX
)
# Palabras que dan contexto monetario a una línea
_MONETARY_CTX = re.compile(
    r"total|importe|monto|precio|valor|costo|cobro|cargo|pago"
    r"|subtotal|honorarios|tarifa|renta|fianza",
    re.IGNORECASE,
)
_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:de)?\s*([a-záéíóúñA-ZÁÉÍÓÚÑ]+)[,\s]*(?:del?\s*)?(\d{4})"
    r"|(\d{4})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{1,2})"
    r"|(\d{1,2})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{2,4})"
)
_ITEM_RE = re.compile(
    r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)
# Fix Dorama: detecta precio directo "es de $X" antes del fallback máximo
_ES_DE_RE = re.compile(
    r"es\s+de\s+\$?\s*([\d,]+(?:\.\d{1,2})?)", re.IGNORECASE
)
# Fix Dorama: excluye líneas que mencionan el presupuesto total del proyecto
_BUDGET_EXCL = re.compile(
    r"presupuestal|presupuesto\s+total", re.IGNORECASE
)
_MESES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12,
}

# Formato numérico de dinero (idéntico al de la plantilla)
_FMT_MONEY = '_-"$ "* #,##0.00_-;\\-"$ "* #,##0.00_-;_-"$ "* "-"??_-;_-@_-'
_FMT_DATE  = "mm-dd-yy"


# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
_SS_DEFAULTS: dict = {
    "pdf_bytes":      None,
    "pdf_hash":       None,
    "total_pages":    0,
    "current_page":   0,
    "num_sec":        1,
    "sec_cfg":        [],
    "df":             None,
    "df_hash":        "",          # F1: hash del df para detectar cambios reales
    "extracted":      False,
    "bx_cache":       {},          # {(moneda, fecha_iso): float}
    "proyecto_nombre":"",
}
for _k, _v in _SS_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)


# ─────────────────────────────────────────────────────────────
# HELPERS GENERALES
# ─────────────────────────────────────────────────────────────
def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _df_hash(df: pd.DataFrame) -> str:
    """Hash rápido del contenido del DataFrame (F1: detecta cambios reales)."""
    try:
        return hashlib.md5(
            pd.util.hash_pandas_object(df, index=False).values.tobytes()
        ).hexdigest()
    except Exception:
        return hashlib.md5(df.to_csv(index=False).encode()).hexdigest()


def _safe_f(v) -> Optional[float]:
    if v is None:
        return None
    try:
        f = float(str(v).replace(",", "").strip())
        return None if f != f else f    # NaN check
    except (ValueError, TypeError):
        return None


def _money(txt: str, need_ctx: bool = False) -> Optional[float]:
    """Extrae el primer importe válido del texto. F5: grupo 3 solo con .XX"""
    if need_ctx and not _MONETARY_CTX.search(str(txt)):
        return None
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
                return datetime.date(yr + 2000 if yr < 100 else yr, int(g[7]), int(g[6]))
        except (ValueError, KeyError):
            continue
    return None


# ─────────────────────────────────────────────────────────────
# OCR  (F4: timeout · F9: max_entries=1)
# ─────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False, max_entries=1)   # F9
def _get_ocr():
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR(det_model_dir=None, rec_model_dir=None, cls_model_dir=None)
    except Exception as exc:
        st.warning(f"RapidOCR no disponible ({exc}). OCR desactivado.")
        return None


def _ocr_page(pdf_bytes: bytes, idx: int) -> str:
    """Extrae texto de una página imagen con timeout (F4)."""
    ocr = _get_ocr()
    if ocr is None:
        return ""

    def _run() -> str:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False
            )
            img = np.frombuffer(pix.samples, np.uint8).reshape(
                pix.height, pix.width, 3
            )[:, :, ::-1]
            result, _ = ocr(img)
            return "\n".join(r[1] for r in result if r and len(r) > 1) if result else ""

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
        fut = ex.submit(_run)
        try:
            return fut.result(timeout=OCR_TIMEOUT_S)
        except concurrent.futures.TimeoutError:
            st.warning(f"⏱ OCR pág. {idx + 1} excedió {OCR_TIMEOUT_S}s — omitida.")
            return ""
        except Exception as exc:
            st.warning(f"OCR pág. {idx + 1}: {exc}")
            return ""


# ─────────────────────────────────────────────────────────────
# VISOR  (F3: pdf_bytes fuera de la firma de caché)
# ─────────────────────────────────────────────────────────────
@st.cache_data(max_entries=120, show_spinner=False)
def _render(pdf_hash: str, idx: int) -> bytes:          # F3: clave = (hash, idx)
    pdf_bytes = st.session_state.pdf_bytes              # bytes en runtime, no en clave
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(
            matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False
        )
        return pix.tobytes("png")


# ─────────────────────────────────────────────────────────────
# API BANXICO SIE
# ─────────────────────────────────────────────────────────────
def _banxico_tc(moneda: str, fecha: datetime.date, token: str) -> Optional[float]:
    """
    Devuelve el tipo de cambio Fix (MXN por unidad) para `moneda` en `fecha`.
    Busca en un rango de 5 días para cubrir fines de semana y festivos.
    Usa caché de sesión para no repetir llamadas.
    """
    if moneda not in _BANXICO_SERIES or not token:
        return None

    cache_key = (moneda, fecha.isoformat())
    if cache_key in st.session_state.bx_cache:
        return st.session_state.bx_cache[cache_key]

    serie   = _BANXICO_SERIES[moneda]
    f_ini   = (fecha - datetime.timedelta(days=5)).isoformat()
    f_fin   = fecha.isoformat()
    url     = f"{_BANXICO_URL}/series/{serie}/datos/{f_ini}/{f_fin}"
    headers = {"Bmx-Token": token, "Accept": "application/json"}

    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        datos = resp.json()["bmx"]["series"][0]["datos"]
        if not datos:
            return None
        # Tomar el último dato disponible (el más cercano a la fecha solicitada)
        tc = float(datos[-1]["dato"].replace(",", ""))
        st.session_state.bx_cache[cache_key] = tc
        return tc
    except Exception as exc:
        st.warning(f"Banxico ({moneda} · {fecha}): {exc}")
        return None


# ─────────────────────────────────────────────────────────────
# RECALCULAR DIFERENCIA FINAL
# ─────────────────────────────────────────────────────────────
def recalc_derived(df: pd.DataFrame) -> pd.DataFrame:
    df  = df.copy()
    tot = pd.to_numeric(df["Total con IVA"],          errors="coerce")
    anx = pd.to_numeric(df["Monto en Anexo Escrito"], errors="coerce")
    mask = tot.notna() | anx.notna()
    df.loc[mask,  "Diferencia final"] = (tot.fillna(0) - anx.fillna(0))[mask]
    df.loc[~mask, "Diferencia final"] = None
    return df


# ─────────────────────────────────────────────────────────────
# MOTOR DE EXTRACCIÓN  (F5, F8 corregidos)
# ─────────────────────────────────────────────────────────────
def _parse_space_table(text: str) -> list[dict]:
    items = []
    for line in text.splitlines():
        m = _ITEM_RE.match(line.strip())
        if m:
            qty, desc, pu, total = m.groups()
            items.append({
                "qty": int(qty), "desc": desc.strip(),
                "pu": float(pu.replace(",", "")),
                "total": float(total.replace(",", "")),
            })
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
    """
    Extrae datos financieros de las páginas p0..p1 del PDF.
    Si moneda != MXN y se dispone de token Banxico, convierte a MXN.
    Devuelve un dict con todas las columnas de _COLS + '_tc' para uso interno.
    """
    native_text = ""
    table_rows: list[list[str]] = []
    ocr_used = False

    # ── Nivel 1 y 2: texto nativo + tablas ───────────────────
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        pr = range(max(0, p0 - 1), min(p1, n_pages))
        for i in pr:
            pg  = pdf.pages[i]
            txt = pg.extract_text() or ""
            native_text += "\n" + txt
            for tbl in (pg.extract_tables() or []):
                if tbl:
                    table_rows.extend(
                        [str(c or "").strip() for c in row] for row in tbl if row
                    )
            for item in _parse_space_table(txt):
                table_rows.append([
                    str(item["qty"]), item["desc"],
                    f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                ])

    # ── Nivel 3: OCR condicional ──────────────────────────────
    pages_count = max(len(pr), 1)
    if len(native_text.strip()) < NATIVE_MIN_CHARS_PER_PAGE * pages_count:
        ocr_used = True
        for i in pr:
            o_txt = _ocr_page(pdf_bytes, i)
            if o_txt:
                native_text += "\n" + o_txt
                for item in _parse_space_table(o_txt):
                    table_rows.append([
                        str(item["qty"]), item["desc"],
                        f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                    ])

    text = native_text
    tlow = text.lower()
    iva_f = "Sí" if det_iva and re.search(r"\biva\b|16%|vat", tlow) else "N/M"

    tot = iva = sub = pu = fecha = None
    qty = 1
    obs_parts: list[str] = []
    if ocr_used:
        obs_parts.append("OCR")

    # ── Paso 1: tablas estructuradas ──────────────────────────
    for row in table_rows:
        j = "   ".join(row).lower()
        s = "   ".join(row)
        if re.search(r"\btotal\b", j) and not re.search(r"sub|parcial|acum", j):
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

    # ── Paso 2: reconstrucción cantidad × precio unitario ─────
    valid_line_totals: list[float] = []
    seen: set[tuple] = set()
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
                    if not (1 <= q_cand <= 9999 and q_cand == int(q_cand)):
                        continue
                    tol = max(0.5, t_cand * 0.01)
                    if abs(q_cand * p_cand - t_cand) <= tol:
                        key = (int(q_cand), round(p_cand, 2), round(t_cand, 2))
                        if key not in seen:
                            seen.add(key)
                            valid_line_totals.append(t_cand)
                        found = True
                        break
                if found:
                    break

    if tot is None and valid_line_totals:
        tot = round(sum(valid_line_totals), 2)
        obs_parts.append("Total por líneas")

    # ── Paso 3: fecha y total por texto libre ─────────────────
    for ln in text.splitlines():
        if fecha is None:
            d = _date(ln)
            if d:
                fecha = d

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b", ln, re.I) and not re.search(r"sub|parcial|acum", ln, re.I):
                v = _money(ln)
                if v and (tot is None or v > tot):
                    tot = v

    # ── Paso 3.5: precio directo "es de $X" ──────────────────────
    # Cubre fianzas, honorarios y servicios expresados en prosa sin
    # etiqueta "Total" explícita (p. ej. Dorama: "...es de $31,520.00").
    if tot is None:
        for ln in text.splitlines():
            m_ed = _ES_DE_RE.search(ln)
            if m_ed:
                v = _safe_f(m_ed.group(1).replace(",", ""))
                if v and 10.0 <= v <= MAX_PLAUSIBLE_MXN:
                    tot = v
                    obs_parts.append("Precio directo")
                    break

    # ── Fallback: máximo con contexto monetario (F5 corregido) ─
    # Excluye líneas de presupuesto del proyecto (_BUDGET_EXCL).
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

    # F8: cota de cordura – valores absurdos (folios, teléfonos)
    if tot is not None and tot > MAX_PLAUSIBLE_MXN:
        obs_parts.append(f"⚠ Valor sospechoso ({tot:,.2f}) — verificar")
        tot = None

    # ── Subtotal e IVA por texto libre ────────────────────────
    if sub is None:
        for ln in text.splitlines():
            if re.search(r"subtotal|importe|sin\s*iva", ln, re.I) and \
               not re.search(r"\btotal\b", ln.replace("subtotal", ""), re.I):
                v = _money(ln)
                if v and (tot is None or v <= tot):
                    sub = v
                    break

    if iva is None and iva_f == "Sí":
        for ln in text.splitlines():
            if re.search(r"\biva\b|16%|vat|impuesto", ln, re.I):
                v = _money(ln)
                if v and (tot is None or v < tot):
                    iva = v
                    break

    # ── Cantidad y precio unitario ────────────────────────────
    for row in table_rows:
        row_str = " ".join(str(c) for c in row)
        nums_: list[float] = []
        for t in re.findall(r"[\d,]+(?:\.\d+)?", row_str):
            try:
                nums_.append(float(t.replace(",", "")))
            except ValueError:
                pass
        if len(nums_) >= 2 and 1 <= nums_[0] <= 9999:
            qty = int(nums_[0])
            pu  = nums_[-2] if len(nums_) > 2 else nums_[-1]

    # ── Triangulación IVA / Subtotal / Total (F8: protegida) ─
    if tot and not sub and not iva and iva_f == "Sí":
        sub = round(tot / 1.16, 2)
        iva = round(tot - sub, 2)
        obs_parts.append("IVA desglosado")
    elif tot and iva and not sub:
        sub = round(tot - iva, 2)
    elif sub and iva and not tot:
        tot = round(sub + iva, 2)
    elif calc_sub and sub and not iva and not tot and iva_f == "Sí":
        iva = round(sub * 0.16, 2)
        tot = round(sub + iva, 2)

    if pu is None and sub is not None and qty:
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
            pu  = round(pu  * tc, 2) if pu  else None
            obs_parts.append(f"1 {moneda} = ${tc:.4f} MXN")
            obs_parts.append(f"Total orig.: {moneda} {tot_orig:,.2f}")

    tc_display = f"{moneda} ({tc:.4f})" if tc else moneda

    return {
        "Fecha":                  (fecha.isoformat() if fecha else datetime.date.today().isoformat()),
        "Rubro":                  label,
        "QT":                     "Sí",
        "T. Cambio":              tc_display,
        "(+ IVA)":                iva_f,
        "Cantidad":               qty,
        "Precio Unitario":        pu,
        "Subtotal (Sin IVA)":     sub,
        "IVA 16%":                iva,
        "Total con IVA":          tot,
        "Diferencia final":       None,
        "Monto en Anexo Escrito": None,
        "Observaciones":          " | ".join(obs_parts) if obs_parts else "",
        "_tc":                    tc,    # columna auxiliar, se elimina antes del df
    }


# ─────────────────────────────────────────────────────────────
# EXPORTACIÓN EXCEL  (F6, F11 corregidos + formato exacto plantilla)
# ─────────────────────────────────────────────────────────────
def _side(style: str = "thin") -> Side:
    return Side(style=style, color="000000")


def _border(style: str = "thin") -> Border:
    s = _side(style)
    return Border(left=s, right=s, top=s, bottom=s)


def _fill(rgb: str) -> PatternFill:
    return PatternFill("solid", start_color=rgb, end_color=rgb)


def to_excel(
    df: pd.DataFrame,
    nombre: str = "",
    blank: bool = False,
) -> bytes:
    """
    Genera el archivo Excel en el formato exacto de la plantilla PAR.
    Si blank=True genera una plantilla vacía con fórmulas para edición manual.

    F6: H, I, J, K siempre usan fórmulas (nunca sobrescritos con valores).
    F11: fila de totales en rango s_xl..e_xl, siempre menor que tot_row.
    """
    buf = io.BytesIO()
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = (nombre[:31] if nombre else "Conciliación")

    # ── Estilos base ──────────────────────────────────────────
    FN = {"name": "Calibri", "size": 11}
    F_WHITE = Font(**FN, bold=True, color="FFFFFF")
    F_HDR   = Font(**FN, bold=True, color="000000")
    F_DATA  = Font(**FN, color="000000")
    F_TOT   = Font(**FN, bold=True, color="000000")

    FILL_ROW1  = _fill("6E152E")    # burdeos (fila 1)
    FILL_HDR_A = _fill("D4C19C")    # arena oscuro (cols A-K)
    FILL_HDR_B = _fill("EBE2D1")    # arena claro  (cols L-M)

    BD = _border("thin")
    BD_MED = _border("medium")

    AL_C  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    AL_R  = Alignment(horizontal="right",  vertical="center")

    n_rows = len(df)

    # ── Anchos de columna ─────────────────────────────────────
    for col_idx, width in enumerate(_WIDTHS_COL, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Fila 1: nombre del proyecto ──────────────────────────────
    ws.row_dimensions[1].height = 27.75
    c = ws.cell(row=1, column=1, value=nombre or "")
    c.font      = F_WHITE
    c.fill      = FILL_ROW1
    c.border    = BD
    c.alignment = AL_L
    # Aplicar relleno burdeos a todas las columnas de la fila 1
    for ci in range(2, len(_COLS) + 1):
        cx = ws.cell(row=1, column=ci)
        cx.fill   = FILL_ROW1
        cx.border = BD

    # ── Fila 2: encabezados ───────────────────────────────────
    ws.row_dimensions[2].height = 21.75
    for ci, hdr in enumerate(_COLS, 1):
        fill = FILL_HDR_A if ci <= 11 else FILL_HDR_B
        c = ws.cell(row=2, column=ci, value=hdr)
        c.font   = F_HDR
        c.fill   = fill
        c.border = BD
        c.alignment = AL_C

    # ── Filas de datos (DATA_START = 3) ───────────────────────
    DS = 3  # DATA_START

    def _money_cell(ws, r, ci, val=None):
        c = ws.cell(row=r, column=ci, value=val)
        c.number_format = _FMT_MONEY
        c.font   = F_DATA
        c.border = BD
        c.alignment = AL_R
        return c

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
            dt = datetime.datetime.combine(fv, datetime.time()) if isinstance(fv, datetime.date) else fv
            c = ws.cell(row=r, column=1, value=dt)
            c.number_format = _FMT_DATE
        else:
            c = ws.cell(row=r, column=1, value=str(fv or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # B: Rubro
        c = ws.cell(row=r, column=2, value=str(row.get("Rubro","") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_L

        # C: QT
        c = ws.cell(row=r, column=3, value=str(row.get("QT","Sí")))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # D: T. Cambio
        c = ws.cell(row=r, column=4, value=str(row.get("T. Cambio","MXN") or "MXN"))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # E: (+IVA)
        c = ws.cell(row=r, column=5, value=str(row.get("(+ IVA)","") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # F: Cantidad  (numfmt 0.00 = plantilla original)
        q = _safe_f(row.get("Cantidad"))
        c = ws.cell(row=r, column=6, value=int(q) if q is not None else 1)
        c.number_format = "0.00"; c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # G: Precio Unitario (input del usuario)
        pu = None if blank else _safe_f(row.get("Precio Unitario"))
        _money_cell(ws, r, 7, pu)

        # ─── F6: columnas calculadas SIEMPRE con fórmula ────────
        has_iva = str(row.get("(+ IVA)", "N/M")).strip().lower() == "sí" and not blank

        # H: Subtotal = F * G
        ws.cell(row=r, column=8, value=f"=F{r}*G{r}").number_format = _FMT_MONEY
        ws.cell(row=r, column=8).font = F_DATA
        ws.cell(row=r, column=8).border = BD
        ws.cell(row=r, column=8).alignment = AL_R

        # I: IVA 16% = H * 0.16  (solo si tiene IVA; si no, fórmula igual a 0)
        i_formula = f"=H{r}*0.16" if has_iva else ""
        ws.cell(row=r, column=9, value=i_formula or None).number_format = _FMT_MONEY
        ws.cell(row=r, column=9).font = F_DATA
        ws.cell(row=r, column=9).border = BD
        ws.cell(row=r, column=9).alignment = AL_R

        # J: Total con IVA = H + I
        j_formula = f"=H{r}+I{r}" if has_iva else f"=H{r}"
        ws.cell(row=r, column=10, value=j_formula).number_format = _FMT_MONEY
        ws.cell(row=r, column=10).font = F_DATA
        ws.cell(row=r, column=10).border = BD
        ws.cell(row=r, column=10).alignment = AL_R

        # K: Diferencia final = J - L
        ws.cell(row=r, column=11, value=f"=J{r}-L{r}").number_format = _FMT_MONEY
        ws.cell(row=r, column=11).font = F_DATA
        ws.cell(row=r, column=11).border = BD
        ws.cell(row=r, column=11).alignment = AL_R

        # L: Monto en Anexo Escrito (input del usuario)
        anx = None if blank else _safe_f(row.get("Monto en Anexo Escrito"))
        c = _money_cell(ws, r, 12, anx)
        c.font = Font(name="Calibri", size=11, bold=True, color="000000")

        # M: Observaciones
        c = ws.cell(row=r, column=13, value="" if blank else str(row.get("Observaciones","") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_L

    # ── Fila de TOTALES  (F11 corregido) ─────────────────────
    if n_rows > 0:
        tot_row = DS + n_rows           # siempre DESPUÉS de la última fila de datos
        s_xl    = DS                    # primera fila de datos (1-based Excel)
        e_xl    = DS + n_rows - 1       # última  fila de datos → siempre < tot_row ✓
        ws.row_dimensions[tot_row].height = 18

        c = ws.cell(row=tot_row, column=1, value="TOTALES")
        c.font = F_TOT; c.border = BD_MED

        for ci in (10, 11, 12):     # J, K, L
            col_l = get_column_letter(ci)
            c = ws.cell(
                row=tot_row, column=ci,
                value=f"=SUM({col_l}{s_xl}:{col_l}{e_xl})"
            )
            c.number_format = _FMT_MONEY
            c.font   = F_TOT
            c.fill   = _fill("EBE2D1")
            c.border = BD_MED
            c.alignment = AL_R

    # Congelar fila de encabezados
    ws.freeze_panes = "A3"

    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# CALLBACKS DE NAVEGACIÓN  (F2 corregido)
# ─────────────────────────────────────────────────────────────
def _go_to(target: int) -> None:
    """Único punto de escritura para current_page."""
    tp = max(st.session_state.total_pages - 1, 0)
    st.session_state.current_page = max(0, min(tp, target))


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Documento PDF")
    up = st.file_uploader("Sube el PDF de cotizaciones", type=["pdf"])
    if up:
        raw = up.read()
        h   = _md5(raw)
        if h != st.session_state.pdf_hash:
            st.session_state.update(
                pdf_bytes=raw, pdf_hash=h,
                extracted=False, df=None, df_hash="", current_page=0,
            )
            with fitz.open(stream=raw, filetype="pdf") as d:
                st.session_state.total_pages = len(d)
        st.success(f"✅ {st.session_state.total_pages} págs. cargadas")

    st.markdown("---")
    st.markdown("### 🏷 Proyecto")
    st.session_state.proyecto_nombre = st.text_input(
        "Nombre del proyecto",
        value=st.session_state.proyecto_nombre,
          )

    st.markdown("---")
    st.markdown("### 💱 Banxico – Tipo de Cambio")
    st.markdown(
        "Obtén tu token gratuito en "
        "[**SIE Banxico API** →](https://www.banxico.org.mx/SieAPIRest/) "
        "*(solo necesario si hay cotizaciones en USD / EUR / CAD)*",
        unsafe_allow_html=False,
    )
    # Prioridad: secrets → sidebar
    _token_default = st.secrets.get("BANXICO_TOKEN", "") if hasattr(st, "secrets") else ""
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
    n = int(st.number_input(
        "Número de secciones", min_value=1, max_value=50,
        value=st.session_state.num_sec, step=1,
    ))
    if n != st.session_state.num_sec:
        st.session_state.num_sec  = n
        st.session_state.extracted = False
        st.session_state.df       = None

    cfgs = list(st.session_state.sec_cfg)
    tp   = max(st.session_state.total_pages, 1)

    while len(cfgs) < n:
        i = len(cfgs) + 1
        cfgs.append({
            "label": f"Sección {i}", "p0": 1, "p1": 1,
            "det_iva": True, "calc_sub": True, "moneda": "MXN",
        })
    cfgs = cfgs[:n]

    for i, c in enumerate(cfgs):
        with st.expander(f"📄 Sección {i + 1}", expanded=(n <= 6)):
            c["label"]  = st.text_input("Rubro / Concepto", value=c["label"], key=f"lb{i}")
            ca, cb      = st.columns(2)
            c["p0"]     = ca.number_input("Pág. Inicio", 1, tp, min(c["p0"], tp), key=f"p0{i}")
            c["p1"]     = cb.number_input("Pág. Fin", c["p0"], tp,
                                          max(min(c["p1"], tp), c["p0"]), key=f"p1{i}")
            c["moneda"] = st.selectbox(
                "Moneda de la cotización",
                ["MXN", "USD", "EUR", "CAD"],
                index=["MXN", "USD", "EUR", "CAD"].index(c.get("moneda", "MXN")),
                key=f"mon{i}",
            )
            c["det_iva"]  = st.checkbox("Detectar IVA", value=c["det_iva"],  key=f"iv{i}")
            c["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=c["calc_sub"], key=f"cs{i}")

    st.session_state.sec_cfg = cfgs

    st.markdown("---")
    run = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_bytes is None),
        use_container_width=True,
        type="primary",
    )


# ─────────────────────────────────────────────────────────────
# CABECERA PRINCIPAL
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
  <h1>📋 Conciliador de Cotizaciones</h1>
  <p>Extracción automática PDF · OCR adaptativo · Tipo de cambio Banxico </p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# PROCESO DE EXTRACCIÓN
# ─────────────────────────────────────────────────────────────
if run and st.session_state.pdf_bytes:
    rows: list[dict] = []
    n_secs = st.session_state.num_sec
    bar    = st.progress(0, text="Iniciando extracción…")

    for i, c in enumerate(st.session_state.sec_cfg):
        bar.progress((i + 0.3) / n_secs, text=f"🔍 {c['label']}…")
        try:
            row = extract(
                st.session_state.pdf_bytes,
                c["label"], c["p0"], c["p1"],
                c["det_iva"], c["calc_sub"],
                moneda=c.get("moneda", "MXN"),
                bx_token=bx_token,
            )
            rows.append(row)
        except Exception as exc:
            st.warning(f"⚠️ Sección {i + 1} «{c['label']}»: {str(exc)[:120]}")
            rows.append({
                **{k: None for k in _COLS},
                "Rubro": c["label"], "QT": "Sí",
                "T. Cambio": c.get("moneda", "MXN"),
                "Cantidad": 1,
                "Fecha": datetime.date.today().isoformat(),
                "Observaciones": f"Error: {str(exc)[:100]}",
                "_tc": None,
            })
        bar.progress((i + 1) / n_secs)

    bar.empty()

    # Quitar columna auxiliar _tc antes de guardar en session_state
    df_new = pd.DataFrame(rows).reindex(columns=_COLS + ["_tc"])
    df_new = df_new.drop(columns=["_tc"], errors="ignore").reindex(columns=_COLS)
    df_new = recalc_derived(df_new)

    st.session_state.df        = df_new
    st.session_state.df_hash   = _df_hash(df_new)
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")


if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()


# ─────────────────────────────────────────────────────────────
# LAYOUT: VISOR  +  EDITOR
# ─────────────────────────────────────────────────────────────
col_L, col_R = st.columns(2, gap="medium")


# ── VISOR DE DOCUMENTO ───────────────────────────────────────
with col_L:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)

    tp = st.session_state.total_pages

    # F2: botones usan on_click con args → no hay desincronización de estado
    nav_p, nav_c, nav_n = st.columns([1, 4, 1])
    nav_p.button("◀", key="btn_prev", use_container_width=True,
                 on_click=_go_to, args=(st.session_state.current_page - 1,))
    nav_n.button("▶", key="btn_next", use_container_width=True,
                 on_click=_go_to, args=(st.session_state.current_page + 1,))

    # F2: number_input como fuente de verdad única fuera de callbacks
    cp = st.session_state.current_page
    page_sel = nav_c.number_input(
        "Página", min_value=1, max_value=tp,
        value=cp + 1, step=1,
        label_visibility="collapsed", key="nav_page_input",
    )
    if page_sel - 1 != cp:
        st.session_state.current_page = page_sel - 1
        cp = page_sel - 1

    st.caption(f"Página {cp + 1} de {tp}")

    # Badge de sección activa
    for c_cfg in st.session_state.sec_cfg:
        if c_cfg["p0"] <= cp + 1 <= c_cfg["p1"]:
            st.markdown(
                f'<span style="background:#6E152E;color:#fff;padding:3px 10px;'
                f'border-radius:4px;font-size:.8rem">📑 {c_cfg["label"]}'
                f' · <span class="tag-moneda">{c_cfg.get("moneda","MXN")}</span></span>',
                unsafe_allow_html=True,
            )
            break

    with st.spinner("Cargando…"):
        st.image(
            _render(st.session_state.pdf_hash, cp),
            use_container_width=True,
        )

    st.download_button(
        "📥 Descargar PDF", data=st.session_state.pdf_bytes,
        file_name="cotizaciones.pdf", mime="application/pdf",
        use_container_width=True,
    )


# ── EDITOR DE DATOS ───────────────────────────────────────────
with col_R:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)

    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona **🔍 Extraer Montos**.")

    else:
        # F7: guard explícito — nunca operar sobre df_cur si es None
        df_cur = st.session_state.df
        if df_cur is None or df_cur.empty:
            st.warning("Sin datos. Ejecuta nuevamente la extracción.")
            st.stop()

        kpi_slot = st.container()

        # ── Editor interactivo (F1: hash real para detectar cambios) ──
        edited_df = st.data_editor(
            df_cur,
            key="data_editor_main",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha":                  st.column_config.TextColumn("Fecha",         width="small"),
                "Rubro":                  st.column_config.TextColumn("Rubro",         width="large"),
                "QT":                     st.column_config.SelectboxColumn("QT",       options=["Sí","No"], width="small"),
                "T. Cambio":              st.column_config.TextColumn("Moneda/TC",     width="small"),
                "(+ IVA)":                st.column_config.SelectboxColumn("IVA",      options=["Sí","No","N/M"], width="small"),
                "Cantidad":               st.column_config.NumberColumn("Cant.",       format="%d",    width="small"),
                "Precio Unitario":        st.column_config.NumberColumn("P. Unit.",    format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)":     st.column_config.NumberColumn("Subtotal",    format="$%.2f", width="medium"),
                "IVA 16%":                st.column_config.NumberColumn("IVA 16%",     format="$%.2f", width="medium"),
                "Total con IVA":          st.column_config.NumberColumn("Total c/IVA", format="$%.2f", width="medium"),
                "Diferencia final":       st.column_config.NumberColumn("Diferencia",  format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Anexo $",     format="$%.2f", width="medium"),
                "Observaciones":          st.column_config.TextColumn("Observaciones", width="large"),
            },
        )

        # F1: solo recalcular y guardar si el usuario REALMENTE editó algo
        new_hash = _df_hash(edited_df)
        if new_hash != st.session_state.df_hash:
            st.session_state.df      = recalc_derived(edited_df)
            st.session_state.df_hash = new_hash

        df_cur = st.session_state.df

        # ── KPIs ─────────────────────────────────────────────
        ts_ = pd.to_numeric(df_cur["Total con IVA"],          errors="coerce").sum()
        rs_ = pd.to_numeric(df_cur["Monto en Anexo Escrito"], errors="coerce").sum()
        dif = ts_ - rs_
        with kpi_slot:
            k1, k2, k3 = st.columns(3)
            for kol, val, lbl in [
                (k1, ts_, "Total Extraído"),
                (k2, rs_, "Monto Referencia"),
                (k3, dif, "Diferencia"),
            ]:
                color = "#6E152E" if lbl != "Diferencia" else (
                    "#c0392b" if abs(dif) > 0.01 else "#28a745"
                )
                kol.markdown(
                    f'<div class="kpi">'
                    f'<div class="v" style="color:{color}">${val:,.2f}</div>'
                    f'<div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

        # Alertas de observaciones
        warn_rows = df_cur[df_cur["Observaciones"].str.contains("⚠|OCR|inferido", na=False)]
        if not warn_rows.empty:
            with st.expander(f"⚠ {len(warn_rows)} aviso(s) de extracción"):
                for _, wr in warn_rows.iterrows():
                    st.caption(f"• **{wr['Rubro']}**: {wr['Observaciones']}")

        st.markdown("---")

        # ── Descarga Excel con formato de plantilla PAR ───────
        ts_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        pname  = (st.session_state.proyecto_nombre or "Cotizaciones").replace(" ", "_")
        nombre = st.session_state.proyecto_nombre

        try:
            xlsx_bytes = to_excel(df_cur, nombre=nombre)
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"Conciliacion_{pname}_{ts_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        except Exception as exc:
            st.error(f"Error generando Excel: {exc}")

        # ── Plantilla vacía para edición manual directa ───────
        with st.expander("📝 Edición manual directa en Excel"):
            st.info(
                "Descarga la **plantilla vacía**: las columnas Subtotal, IVA 16%, "
                "Total con IVA y Diferencia final se calculan automáticamente con "
                "fórmulas al escribir **Precio Unitario** (G) y **Monto en Anexo Escrito** (L). "
                "El número de filas coincide con las secciones configuradas."
            )
            n_sec = st.session_state.num_sec
            df_blank = pd.DataFrame({
                "Fecha":                  [datetime.date.today().isoformat()] * n_sec,
                "Rubro":                  [c["label"]          for c in st.session_state.sec_cfg],
                "QT":                     ["Sí"]                * n_sec,
                "T. Cambio":              [c.get("moneda","MXN") for c in st.session_state.sec_cfg],
                "(+ IVA)":                ["Sí"]                * n_sec,
                "Cantidad":               [1]                   * n_sec,
                "Precio Unitario":        [None]                * n_sec,
                "Subtotal (Sin IVA)":     [None]                * n_sec,
                "IVA 16%":                [None]                * n_sec,
                "Total con IVA":          [None]                * n_sec,
                "Diferencia final":       [None]                * n_sec,
                "Monto en Anexo Escrito": [None]                * n_sec,
                "Observaciones":          [""]                  * n_sec,
            })
            try:
                xlsx_blank = to_excel(df_blank, nombre=nombre, blank=True)
                st.download_button(
                    "📄 Descargar plantilla vacía",
                    data=xlsx_blank,
                    file_name=f"Plantilla_{pname}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"Error generando plantilla: {exc}")

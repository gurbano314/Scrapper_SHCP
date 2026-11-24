# ============================================================
# CONCILIADOR DE COTIZACIONES  |  engine.py  v6.6
# Motor de extracción puro – sin dependencias de GUI.
# Compatible con PyQt6 y cualquier otro frontend.
# ============================================================
from __future__ import annotations

import datetime
import hashlib
import io
import math
import os
import re
import sys
from typing import Optional

# ── Tesseract path for Windows & PyInstaller ─────────────────
if hasattr(sys, '_MEIPASS'):
    _base_dir = sys._MEIPASS
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))

_tess_candidates = [
    os.path.join(_base_dir, 'bin', 'tesseract', 'tesseract.exe'),
    os.path.join(os.path.dirname(_base_dir), 'bin', 'tesseract', 'tesseract.exe'),
]
_env_tess = os.getenv('TESSERACT_PATH')
if _env_tess:
    _tess_candidates.append(_env_tess)
_tess_candidates.append(r'C:\Program Files\Tesseract-OCR\tesseract.exe')

TESSERACT_PATH = next(
    (p for p in _tess_candidates if os.path.exists(p)),
    _tess_candidates[-1],
)

if os.path.exists(TESSERACT_PATH):
    _tess_dir = os.path.dirname(TESSERACT_PATH)
    if _tess_dir not in os.environ.get('PATH', ''):
        os.environ['PATH'] = os.environ.get('PATH', '') + os.pathsep + _tess_dir
    _tessdata = os.path.join(_tess_dir, 'tessdata')
    if os.path.isdir(_tessdata):
        os.environ['TESSDATA_PREFIX'] = _tessdata


def _resolve_tess_lang(preferred: str = "spa+eng") -> str:
    """Devuelve 'spa+eng' si ambos .traineddata están disponibles; si no, 'spa'."""
    td_dir = os.environ.get('TESSDATA_PREFIX', '')
    if not td_dir or not os.path.isdir(td_dir):
        if os.path.exists(TESSERACT_PATH):
            td_dir = os.path.join(os.path.dirname(TESSERACT_PATH), 'tessdata')
    if td_dir and os.path.isdir(td_dir):
        has_spa = os.path.exists(os.path.join(td_dir, 'spa.traineddata'))
        has_eng = os.path.exists(os.path.join(td_dir, 'eng.traineddata'))
        if '+' in preferred and has_spa and has_eng:
            return preferred
        if has_spa:
            return "spa"
        if has_eng:
            return "eng"
    return "spa"

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    import fitz
import numpy as np
import openpyxl
import pandas as pd
import pdfplumber
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────
# CONSTANTES Y REGEX
# ─────────────────────────────────────────────────────────────
NATIVE_MIN_CHARS_PER_PAGE = 80
MAX_PLAUSIBLE_MXN = 500_000_000  # 500 M – presupuestos de obra pueden superar 50 M
OCR_SCALE = 3.0
RAPIDOCR_SCALE = 2.5
OCR_MAX_CACHE = 256

_BANXICO_SERIES = {"USD": "SF43718", "EUR": "SF46410", "CAD": "SF60653"}
_BANXICO_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1"

_COLS = [
    "Nº Sec", "Sección", "Tipo",
    "Fecha", "Rubro", "QT", "T. Cambio", "(+ IVA)", "Cantidad",
    "Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%", "Total con IVA",
    "Diferencia final", "Monto en Anexo Escrito", "Observaciones",
]
_WIDTHS_COL = [8, 12, 20, 12, 52, 5, 10, 7, 8, 16, 18, 17, 16, 18, 22, 48]

_MONEY_RE = re.compile(
    r"(?:MX\$|US\$|CA\$|\$|€)\s*(\d{1,3}(?:[.,\s]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)(?!\d|[.,]\d)"  # g1: prefijo $, MX$, US$, CA$, €
    r"|(?<![\d.,])(\d{1,3}(?:[.,\s]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2}))\s*€"                         # g2: sufijo €
    r"|(?<![\d.,])(\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})?)(?!\d|[.,]\d)"                                       # g3: agrupado por miles
    r"|(?<![\d.,])(\d{1,9}\.\d{2})(?!\d|[.,]\d)"                                                            # g4: decimal suelto
)
_BARE_INT_MONEY_RE = re.compile(
    r"(?:^|(?<=[:\s]))([1-9]\d{3,6})(?=\s*(?:mxn|usd|eur|cad|pesos|m\.?n\.?)?\s*$)",
    re.IGNORECASE,
)
_MONETARY_CTX = re.compile(
    r"total|importe|monto|precio|valor|costo|cobro|cargo|pago"
    r"|subtotal|honorarios|tarifa|renta|fianza",
    re.IGNORECASE,
)
_DATE_RE = re.compile(
    r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})"
    r"|\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\b"
    r"|\b(\d{1,2})\s*(?:de|[/\-.\s])?\s*(?!del\b|al\b|el\b)([a-záéíóúñ]{3,12})(?:\s*,?\s*(?:del?|[/-])?\s*(\d{2,4}))(?!\s*[./\-]\d)"
    r"|\b([a-záéíóúñ]{3,12})\s+(\d{1,2})\s*,?\s*(\d{4})\b",
    re.IGNORECASE,
)
_DATE_EXCL_LINE_RE = re.compile(
    r"vigencia|periodo|per[ií]odo|temporada|service\s*period|fecha\s+estimada"
    r"|entrega|comprendid|comprende|conclusi[oó]n|constituit|empezando|campamento"
    r"|concentraci[oó]n|competencia|pretemporada|durante\s+el"
    r"|20\d{2}\s*[—–]\s*\d{1,2}|\b(?:al|hasta)\s+\d{1,2}\s+de\b"
    r"|\b(?:al|a)\s+\d{1,2}[./\-]\d{1,2}[./\-]\d{4}\b",
    re.IGNORECASE,
)
_DATE_ISSUE_RE = re.compile(
    r"fecha(?:\s+de\s+emisi[oó]n)?\b|issue\s*date"
    r"|data(?:\s+(?:del\s*tasso|di\s*riferimento))?\b"
    r"|documento\s*nr|cdmx|ciudad\s+de\s+m[eé]xico",
    re.IGNORECASE,
)
_ITEM_RE = re.compile(
    r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)
_GHOST_MONEY_RE = re.compile(
    r"(?:^|\s)[$€]?\s*(\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?)"
    r"|(?:^|\s)[$€]?\s*(\d{1,9}\.\d{2})(?!\d)",
    re.IGNORECASE,
)
_IVA_INCLUDED_RE = re.compile(
    r"iva\s+inclu[iy]|inclu[iy]\w*\s*(?:el\s*)?(?:impuesto\s+al\s+valor\s+agregado|iva)"
    r"|ya\s+inclu|tarifa\s+inclu|con\s+iva\s+inclu|precio\s+(?:total|final)\s+con\s+iva",
    re.IGNORECASE,
)
_IVA_EXEMPT_RE = re.compile(
    r"exento|tasa\s+0|no\s+aplica\s+iva|sin\s+iva|iva\s*:\s*0|libre\s+de\s+iva"
    r"|iva\s+no\s+aplica|no\s+causa\s+iva",
    re.IGNORECASE,
)
_IVA_ZERO_EXPLICIT_RE = re.compile(
    r"(?:\biva\b|impuesto(?:\s*\(\s*16\s*%\s*\))?)\s*[:\-]?\s*\$\s*0+(?:\.0{1,2})?\b",
    re.IGNORECASE,
)
_ES_DE_RE = re.compile(
    r"(?:es|ser[aá]|era|monto)\s+de\s+([\$€]\s*\d[\d.,]*|\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|\d{3,9}\.\d{2})"
    r"(?!\s*(?:d[ií]as|meses|a[ñn]os|horas|personas|atletas|ciclistas))",
    re.IGNORECASE,
)
_TBL_TOTAL_RE  = re.compile(
    r"\btot(?:ales?|ale?|[ae][1!i]|ts[!]|e[!])(?!\w)|total\s+documento|totale?\s+articoli|costo\s+total",
    re.IGNORECASE,
)
_TBL_SUBTOT_RE = re.compile(
    r"(?:sub|sut)\s*-?\s*(?:total|tota[lt1!]|iota[lt]?|ouat|oral)|sin\s*iva|importe\s*neto",
    re.IGNORECASE,
)
_TBL_IVA_RE    = re.compile(r"\biva\b|i\.?\s*v\.?\s*a\.?|16\s*%|vat|\bimpuestos?\b", re.IGNORECASE)
_TBL_EXCL_RE   = re.compile(r"sub|sut|parcial|acum|period", re.IGNORECASE)
_TBL_HEADER_RE = re.compile(
    r"(?:c[oó]digo|descripci[oó]n|concepto|descrizione|rubro).*(?:unidad|cantidad|precio|total|totale|importe)"
    r"|(?:unidad|cantidad|precio\s*unit|prezzo|tarifa).*(?:total|totale|importe)",
    re.IGNORECASE,
)
_RET_IVA_RE = re.compile(r"ret\.?\s*(?:de\s*)?(?:iva|isr)|retenci[oó]n", re.IGNORECASE)
_EXCH_RATE_LINE_RE = re.compile(
    r"tasso\s*d[il]\s*cambio|data\s+del\s+tasso|exchange\s*rate|tipo\s*de\s*cambio|1\s*(?:eur|usd|cad)\s*[=>\-]",
    re.IGNORECASE,
)
_MXN_EQUIV_PATTERNS = [
    re.compile(r"c[o0][st]{1,3}[o0]\s+[il1]n\s+p[eo]s[o0]s\s+m[eo]ss?[il1]c\w*[^\d\n]{0,20}([\d.,\s]+)", re.IGNORECASE),
    re.compile(r"\btotale?s?\s*[:\-]?\s*\$?\s*([\d.,]+)\s*m\.?x\.?n\.?\b", re.IGNORECASE),
    re.compile(r"(?:^|\s)\$\s*([\d.,]+)\s*m\.?x\.?n\.?\b", re.IGNORECASE),
]
_FOREIGN_DOC_RE = re.compile(
    r"(?:total\s+documento|totale?\s+articoli)\s*[:!\-]?\s*[€$]?\s*([\d.,]+)\s*(€|eur)?",
    re.IGNORECASE,
)
_FOREIGN_USD_RE = re.compile(
    r"\$\s*([\d.,]+)\s*(usd|eur|cad)\b",
    re.IGNORECASE,
)
_GLOBAL_TOTAL_RE = re.compile(
    r"presupuesto\s+total|costo\s+total\s+del?\s+(?:p\w{1,4}(?:\s+[o0]\s+par)?|par|proyecto)|^\s*totales\b",
    re.IGNORECASE,
)
_BUDGET_EXCL   = re.compile(
    r"presupuestal|presupuesto\s+total"
    r"|costo\s+(?:total\s+)?del?\s+proyecto"
    r"|presupuesto\s+del?\s+proyecto"
    r"|inversi[óo]n\s+total\s+del\s+(?:proyecto|programa)"  # NO excluir "del servicio"
    r"|anexo\s*\d|aportante",
    re.IGNORECASE,
)
# Líneas de plan de pagos – no confundir con el total global
_PAYMENT_SCHED_RE = re.compile(
    r"pago\s+mensual|mensualidad|por\s+mes\s*\("
    r"|cantidad\s+de\s+pagos|n[uú]mero\s+de\s+pagos"
    r"|cuota\s+mensual|abono\s+mensual",
    re.IGNORECASE,
)

_MESES = {
    # Español
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
    # Inglés
    "january": 1, "jan": 1, "february": 2, "march": 3, "april": 4, "apr": 4,
    "june": 6, "july": 7, "august": 8, "aug": 8, "september": 9,
    "october": 10, "november": 11, "december": 12, "dec": 12,
    # Italiano
    "gennaio": 1, "gen": 1, "febbraio": 2, "aprile": 4, "maggio": 5, "mag": 5,
    "giugno": 6, "giu": 6, "luglio": 7, "lug": 7, "settembre": 9, "set": 9,
    "ottobre": 10, "ott": 10, "novembre": 11, "dicembre": 12,
}

_FMT_MONEY = (
    '_-"$ "* #,##0.00_-;'
    '\\-"$ "* #,##0.00_-;'
    '_-"$ "* "-"??_-;_-@_-'
)
_FMT_DATE = "mm-dd-yy"

# Caché Banxico y OCR a nivel de módulo (compartido entre sesiones del mismo proceso)
_BX_CACHE: dict[tuple, Optional[float]] = {}
_OCR_CACHE: dict[tuple[str, int, str], str] = {}

def _ocr_cache_set(key: tuple[str, int, str], val: str) -> None:
    if key in _OCR_CACHE:
        _OCR_CACHE[key] = val
        return
    if len(_OCR_CACHE) >= OCR_MAX_CACHE:
        oldest = next(iter(_OCR_CACHE))
        del _OCR_CACHE[oldest]
    _OCR_CACHE[key] = val

# OCR – singleton a nivel de módulo para evitar recargar el modelo
_ocr_engine = None
_ocr_available: Optional[bool] = None   # None = no determinado aún

# ─────────────────────────────────────────────────────────────
# VALIDACIÓN NORMATIVA — Lineamientos EFIDEPORTE
# ─────────────────────────────────────────────────────────────
MAX_VIGENCIA_DIAS = 90  # Cotizaciones vigentes ≤ 3 meses
TOPE_ADMIN_PCT    = 0.05  # Gastos administrativos ≤ 5%

_ADMIN_KEYWORDS_RE = re.compile(
    r"papeler[ií]a|renta\s+de\s+oficina|servicios?\s+administrat"
    r"|gasto\s+administrat|compra\s+de\s+papeler"
    r"|art[ií]culos?\s+de\s+oficina|material\s+de\s+oficina"
    r"|honorarios?\s+del?\s+(?:contador|profesional|auditor)"
    r"|servicio\s+de\s+limpieza|agua\s+potable|luz\s+el[eé]ctrica"
    r"|tel[eé]fono|internet\s+oficina|mensajer[ií]a",
    re.IGNORECASE,
)
_RUBROS_PROHIBIDOS_RE = re.compile(
    r"otros\s+gastos|imprevistos|gastos\s+varios"
    r"|pr[eé]stamo|anticipo\s+de\s+n[oó]mina"
    r"|gesti[oó]n\s+de\s+recursos|honorarios?\s+por\s+gesti[oó]n"
    r"|ganancia\s+propia|utilidad\s+del\s+proyecto"
    r"|servicios?\s+imprecis|concepto\s+gen[eé]rico",
    re.IGNORECASE,
)
_SEGURO_RE = re.compile(
    r"seguro|p[oó]liza\s+de\s+(?:seguro|cobertura|responsabilidad)"
    r"|cobertura\s+(?:total|contra\s+siniestro|del?\s+proyecto)"
    r"|prima\s+de\s+seguro",
    re.IGNORECASE,
)
_EQUIPO_RE = re.compile(
    r"equipo|aparato|maquinaria|instrumento|implemento"
    r"|bicicleta|bal[oó]n|raqueta|arco|rifle|pistola|kayak"
    r"|canoa|tabla|vela|esquí|trampolín",
    re.IGNORECASE,
)


def validate_normative(
    df: pd.DataFrame,
    fecha_solicitud: Optional[datetime.date] = None,
    pii_solo_equipo: bool = False,
) -> list[str]:
    """
    Ejecuta las 7 validaciones normativas de EFIDEPORTE sobre el DataFrame
    de resultados. Retorna una lista de advertencias textuales.

    Validaciones:
    1. Vigencia de cotizaciones (≤ 3 meses)
    2. Tope gastos administrativos ≤ 5%
    3. Detección de rubros prohibidos
    4. Detección de duplicados
    5. Verificación de seguro obligatorio
    6. PII solo equipo → sin gastos administrativos
    7. Aportación en especie sin CFDI (informativa)
    """
    if df is None or df.empty:
        return []

    warnings: list[str] = []
    df_norm = _normalize_editor_df(df)
    if df_norm.empty:
        return []

    # Separar cotizaciones individuales del Presupuesto Global para no duplicar totales
    tipo_col = df_norm["Tipo"].astype(str).str.strip()
    df_prov = df_norm[tipo_col != "Presupuesto Global"]
    df_presup = df_norm[tipo_col == "Presupuesto Global"]

    # ── 1. Vigencia de cotizaciones ──────────────────────────
    if fecha_solicitud is not None:
        for idx, row in df_norm.iterrows():
            tipo = str(row.get("Tipo", "")).strip()
            if tipo != "Cotización Proveedor":
                continue
            fecha_str = str(row.get("Fecha", "")).strip()
            if not fecha_str:
                continue
            try:
                fecha_cot = datetime.date.fromisoformat(fecha_str)
                delta = (fecha_solicitud - fecha_cot).days
                if delta > MAX_VIGENCIA_DIAS:
                    rubro = str(row.get("Rubro", f"Fila {idx+1}"))
                    warnings.append(
                        f"📅 «{rubro}»: Cotización con fecha {fecha_str} "
                        f"excede los 3 meses de vigencia ({delta} días). "
                        f"— Lineamiento 2, pág. 10"
                    )
            except (ValueError, TypeError):
                pass

    # ── 2. Tope gastos administrativos ≤ 5% ──────────────────
    total_prov = float(pd.to_numeric(df_prov["Total con IVA"], errors="coerce").sum())
    total_presup = float(pd.to_numeric(df_presup["Total con IVA"], errors="coerce").sum())
    total_global = total_presup if total_presup > 0 else total_prov

    admin_total = 0.0
    admin_rubros: list[str] = []
    for idx, row in df_prov.iterrows():
        rubro = str(row.get("Rubro", ""))
        if _ADMIN_KEYWORDS_RE.search(rubro):
            monto = _safe_f(row.get("Total con IVA"))
            if monto:
                admin_total += monto
                admin_rubros.append(rubro)

    if total_global > 0 and admin_total > 0:
        pct = admin_total / total_global
        if pct > TOPE_ADMIN_PCT:
            warnings.append(
                f"🚫 Gastos administrativos ({pct:.1%}) exceden el tope "
                f"del 5% del presupuesto total. "
                f"Monto admin: ${admin_total:,.2f} / Total: ${total_global:,.2f}. "
                f"Rubros: {', '.join(admin_rubros[:5])}. "
                f"— Lineamiento 3, pág. 11"
            )

    # ── 3. Detección de rubros prohibidos ────────────────────
    for idx, row in df_norm.iterrows():
        if str(row.get("Tipo", "")).strip() == "Presupuesto Global":
            continue
        rubro = str(row.get("Rubro", ""))
        if _RUBROS_PROHIBIDOS_RE.search(rubro):
            warnings.append(
                f"🚫 «{rubro}»: Rubro potencialmente no elegible para "
                f"estímulo fiscal. Verificar que no sea impreciso, duplicado "
                f"o de gestión de recursos. — Lineamiento 10, pág. 11-12"
            )

    # ── 4. Detección de duplicados ───────────────────────────
    rubros_vistos: list[tuple[str, float, int]] = []  # (rubro_lower, total, idx)
    for idx, row in df_norm.iterrows():
        if str(row.get("Tipo", "")).strip() == "Presupuesto Global":
            continue
        rubro = str(row.get("Rubro", "")).strip().lower()
        total = _safe_f(row.get("Total con IVA"))
        if not rubro or total is None:
            continue
        for prev_rubro, prev_total, prev_idx in rubros_vistos:
            # Comparar similitud simple
            if prev_rubro == rubro and abs(total - prev_total) / max(total, 1) < 0.05:
                warnings.append(
                    f"⚠ Posible duplicado: «{row.get('Rubro', '')}» "
                    f"(fila {idx+1}) tiene el mismo rubro y monto similar "
                    f"a la fila {prev_idx+1} (${total:,.2f}). "
                    f"— Lineamiento 10c, pág. 11"
                )
        rubros_vistos.append((rubro, total, idx))

    # ── 5. Verificación de seguro obligatorio ────────────────
    tiene_seguro = False
    for idx, row in df_prov.iterrows():
        rubro = str(row.get("Rubro", ""))
        if _SEGURO_RE.search(rubro):
            tiene_seguro = True
            break
    if not tiene_seguro and len(df_prov) >= 3:
        warnings.append(
            "📋 No se detectó un rubro de seguro en el presupuesto. "
            "Los lineamientos exigen un seguro con cobertura total del "
            "valor del PII o PAR contra cualquier siniestro. "
            "— Lineamiento 1, pág. 10"
        )

    # ── 6. PII solo equipo → sin gastos administrativos ──────
    if pii_solo_equipo and admin_rubros:
        warnings.append(
            f"🚫 PII exclusivamente de equipo: No se permiten gastos "
            f"administrativos. Se detectaron: {', '.join(admin_rubros[:5])}. "
            f"— Lineamiento 3, pág. 11"
        )

    # ── 7. Info: aportación en especie (recordatorio CFDI) ───
    for idx, row in df_norm.iterrows():
        rubro = str(row.get("Rubro", ""))
        if _EQUIPO_RE.search(rubro):
            tipo = str(row.get("Tipo", "")).strip()
            if tipo == "Presupuesto Global":
                continue
            warnings.append(
                f"📎 «{rubro}»: Si es aportación en especie (equipo), "
                f"verificar que se cuente con el CFDI correspondiente "
                f"y cotizaciones de valor comercial. — Lineamiento e), pág. 9"
            )

    return warnings


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _is_blank_value(v) -> bool:
    if v is None:
        return True
    try:
        if bool(pd.isna(v)):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(v, str) and not v.strip()


def _normalize_editor_df(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=_COLS)
    clean = pd.DataFrame(df).copy()
    private_cols = [c for c in clean.columns if str(c).startswith("_")]
    clean = clean.drop(columns=private_cols, errors="ignore").reindex(columns=_COLS)
    if clean.empty:
        return clean.reset_index(drop=True)
    editable_cols = [c for c in _COLS if c != "Diferencia final"]
    blank_rows = clean[editable_cols].apply(
        lambda row: all(_is_blank_value(v) for v in row), axis=1
    )
    return clean.loc[~blank_rows].reset_index(drop=True)


def _df_hash(df: pd.DataFrame) -> str:
    if df is None:
        return ""
    df = _normalize_editor_df(df)
    hash_df = df.astype(object).where(pd.notna(df), "")
    try:
        return hashlib.md5(
            pd.util.hash_pandas_object(hash_df, index=False).values.tobytes()
        ).hexdigest()
    except Exception:
        return hashlib.md5(hash_df.to_csv(index=False).encode()).hexdigest()


def _safe_f(v) -> Optional[float]:
    if v is None:
        return None
    try:
        f = float(str(v).replace(",", "").strip())
        return None if math.isnan(f) or math.isinf(f) else f
    except (ValueError, TypeError):
        return None


def _norm_ocr_currency_chars(txt: str) -> str:
    """Normaliza símbolos OCR comunes confundidos con '$' antes de dígitos."""
    return re.sub(r"§(?=\s*\d)", "$", str(txt))


def _parse_money_token(raw: str) -> Optional[float]:
    """Convierte un token numérico monetario (MX/US/EU o artefacto OCR mixto) a float."""
    if not raw:
        return None
    s = str(raw).strip().rstrip(" .,\t")
    # Quitar espacios internos usados como separador de miles (ej. "75 844,312.00")
    s = re.sub(r"(?<=\d)\s+(?=\d{3}(?:[.,\s]|$))", "", s)
    s = s.replace(" ", "")
    if not s:
        return None
    if "." not in s and "," not in s:
        try:
            v = float(s)
            return v if v > 0 else None
        except ValueError:
            return None
    parts = re.split(r"[.,]", s)
    if any(not p.isdigit() for p in parts):
        return None
    if len(parts) == 2:
        int_p, tail_p = parts
        if len(tail_p) in (1, 2):
            val_s = f"{int_p}.{tail_p}"
        elif len(tail_p) == 3 and 1 <= len(int_p) <= 3:
            val_s = f"{int_p}{tail_p}"
        else:
            return None
    else:
        if not (1 <= len(parts[0]) <= 3):
            return None
        last = parts[-1]
        mids = parts[1:-1]
        if any(len(m) != 3 for m in mids):
            return None
        if len(last) == 2:
            val_s = "".join(parts[:-1]) + "." + last
        elif len(last) == 3:
            val_s = "".join(parts)
        else:
            return None
    try:
        v = float(val_s)
        return v if v > 0 else None
    except ValueError:
        return None


def _money(txt: str, allow_bare_int: bool = False) -> Optional[float]:
    """Devuelve el primer monto positivo encontrado en txt."""
    norm = _norm_ocr_currency_chars(txt)
    for m in _MONEY_RE.finditer(norm):
        raw = m.group(1) or m.group(2) or m.group(3) or m.group(4)
        v = _parse_money_token(raw)
        if v is not None:
            return v
    if allow_bare_int:
        for m in _BARE_INT_MONEY_RE.finditer(norm):
            v = _parse_money_token(m.group(1))
            if v is not None and not (1990 <= v <= 2040):
                return v
    return None


def _money_max(txt: str, allow_bare_int: bool = False) -> Optional[float]:
    """Devuelve el MAYOR monto positivo encontrado en txt (útil para líneas multivalor)."""
    norm = _norm_ocr_currency_chars(txt)
    best: Optional[float] = None
    for m in _MONEY_RE.finditer(norm):
        raw = m.group(1) or m.group(2) or m.group(3) or m.group(4)
        v = _parse_money_token(raw)
        if v is not None and (best is None or v > best):
            best = v
    if best is None and allow_bare_int:
        for m in _BARE_INT_MONEY_RE.finditer(norm):
            v = _parse_money_token(m.group(1))
            if v is not None and not (1990 <= v <= 2040) and (best is None or v > best):
                best = v
    return best


def _money_after_label(
    ln: str,
    label_re: re.Pattern,
    use_max: bool = False,
    allow_bare_int: bool = False,
) -> Optional[float]:
    """Extrae el monto preferentemente después de la etiqueta; fallback a la línea completa."""
    fn = _money_max if use_max else _money
    m = label_re.search(ln)
    if m:
        tail = ln[m.end():]
        v = fn(tail, allow_bare_int=allow_bare_int)
        if v is not None:
            return v
    if not _TBL_HEADER_RE.search(ln):
        return fn(ln, allow_bare_int=allow_bare_int)
    return None


_FULL_MONTH_NAMES = {k for k in _MESES if len(k) >= 6}


def _match_month(raw_mes: str) -> Optional[int]:
    m = raw_mes.lower().strip(" .,:;")
    if m in _MESES:
        return _MESES[m]
    if len(m) >= 5:
        for full_m in _FULL_MONTH_NAMES:
            if full_m[:4] == m[:4]:
                return _MESES[full_m]
    return None


def _date(text: str) -> Optional[datetime.date]:
    norm = re.sub(r"\b[Oo](\d)\b", r"0\1", text)
    for m in _DATE_RE.finditer(norm.lower()):
        g = m.groups()
        try:
            if g[0]:
                yr = int(g[0])
                if 2015 <= yr <= 2035:
                    return datetime.date(yr, int(g[1]), int(g[2]))
            elif g[3]:
                d_val, m_val, yr = int(g[3]), int(g[4]), int(g[5])
                yr = yr + 2000 if yr < 100 else yr
                if 2015 <= yr <= 2035:
                    if 1 <= m_val <= 12 and 1 <= d_val <= 31:
                        return datetime.date(yr, m_val, d_val)
                    if 1 <= d_val <= 12 and 1 <= m_val <= 31:
                        return datetime.date(yr, d_val, m_val)
            elif g[6]:
                mes_num = _match_month(g[7])
                if mes_num:
                    yr = int(g[8])
                    yr = yr + 2000 if yr < 100 else yr
                    if 2015 <= yr <= 2035:
                        return datetime.date(yr, mes_num, int(g[6]))
            elif g[9]:
                mes_num = _match_month(g[9])
                if mes_num:
                    yr = int(g[11])
                    if 2015 <= yr <= 2035:
                        return datetime.date(yr, mes_num, int(g[10]))
        except (ValueError, KeyError):
            continue
    return None


def _extract_best_date(text: str) -> Optional[datetime.date]:
    """Prioriza la fecha de emisión/cotización sobre fechas de periodo o vigencia."""
    if not text:
        return None
    max_plausible_yr = datetime.date.today().year + 2
    lines = text.splitlines()
    p1_good: list[datetime.date] = []
    p1_late: list[datetime.date] = []
    p2_good: list[datetime.date] = []
    p2_late: list[datetime.date] = []
    p3_all: list[datetime.date] = []

    standalone_date_hdr = re.compile(r"^\s*(?:fecha|issue\s*date|data)\s*[:.]?\s*$", re.IGNORECASE)

    for idx_ln, raw_ln in enumerate(lines):
        ln = raw_ln.strip()
        if not ln:
            continue
        d = _date(ln)
        if not d:
            continue
        p3_all.append(d)
        is_excl = bool(_DATE_EXCL_LINE_RE.search(ln))
        if not is_excl:
            prev_ln = lines[idx_ln - 1].strip() if idx_ln > 0 else ""
            is_issue = bool(_DATE_ISSUE_RE.search(ln) or standalone_date_hdr.search(prev_ln))
            if is_issue:
                (p1_good if d.year <= max_plausible_yr else p1_late).append(d)
            else:
                (p2_good if d.year <= max_plausible_yr else p2_late).append(d)

    for bucket in (p1_good, p2_good, p1_late, p2_late, p3_all):
        if bucket:
            return bucket[0]
    return None


_SPANISH_WORDS_MAP = {
    "cero": 0, "un": 1, "uno": 1, "una": 1,
    "dos": 2, "tres": 3, "cuatro": 4, "cuaatro": 4, "cinco": 5,
    "seis": 6, "sais": 6, "siete": 7, "ocho": 8, "nueve": 9,
    "diez": 10, "once": 11, "doce": 12, "trece": 13, "catorce": 14, "quince": 15,
    "dieciseis": 16, "diecisiete": 17, "dieciocho": 18, "diecinueve": 19,
    "veinte": 20, "veite": 20, "veinie": 20,
    "veintiun": 21, "veintiuno": 21, "veintiuna": 21,
    "veintidos": 22, "veintitres": 23, "veinticuatro": 24, "veinticinco": 25,
    "veintiseis": 26, "veintisiete": 27, "veintiocho": 28, "veintinueve": 29,
    "treinta": 30, "cuarenta": 40, "cincuenta": 50, "sesenta": 60,
    "setenta": 70, "serenta": 70, "ochenta": 80, "noventa": 90,
    "cien": 100, "clen": 100, "clan": 100, "cian": 100, "ciento": 100,
    "doscientos": 200, "docientos": 200,
    "trescientos": 300, "cuatrocientos": 400, "quinientos": 500,
    "seiscientos": 600, "setecientos": 700,
    "ochocientos": 800, "ochoscientos": 800, "novecientos": 900,
}
_AMOUNT_WORDS_LINE_RE = re.compile(
    r"([a-záéíóúñ\s]{6,}?\b(?:pesos(?:\s+pesos)?(?:\s+m\.?x\.?n\.?)?|mil(?:\s+m\.?x\.?n\.?)?))\s*(\d{1,2})\s*/\s*(?:100|00|700)\b",
    re.IGNORECASE,
)


def _norm_ascii(s: str) -> str:
    rep = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u"}
    out = s.lower()
    for k, v in rep.items():
        out = out.replace(k, v)
    return out


def _parse_spanish_amount_words(text: str) -> Optional[float]:
    """Convierte montos escritos en letra en español (ej. 'Cien mil pesos 00/100 M.N.') a float."""
    if not text:
        return None
    norm_text = _norm_ascii(text)
    norm_text = re.sub(r"\bdoscemil\b", "doce mil", norm_text)
    for m in _AMOUNT_WORDS_LINE_RE.finditer(norm_text):
        phrase = re.sub(r"\b(?:pesos|mxn|m\.?n\.?|moneda|nacional)\b", " ", m.group(1))
        cents = int(m.group(2))
        tokens = [t for t in re.split(r"[^a-z]+", phrase) if t]
        # Recortar palabras iniciales no numéricas (ej. 'son ciento...', 'de ...')
        start_idx = 0
        while start_idx < len(tokens) and tokens[start_idx] not in _SPANISH_WORDS_MAP and tokens[start_idx] not in ("mil", "millon", "millones"):
            start_idx += 1
        tokens = tokens[start_idx:]
        if not tokens:
            continue
        total_val = 0
        current = 0
        valid = True
        for tok in tokens:
            if tok == "y":
                continue
            if tok in _SPANISH_WORDS_MAP:
                current += _SPANISH_WORDS_MAP[tok]
            elif tok == "mil":
                if current == 0:
                    current = 1
                total_val += current * 1000
                current = 0
            elif tok in ("millon", "millones"):
                if current == 0:
                    current = 1
                total_val += current * 1_000_000
                current = 0
            else:
                valid = False
                break
        if valid:
            val = float(total_val + current) + (cents / 100.0)
            if val >= 100.0:
                return round(val, 2)
    return None


# ─────────────────────────────────────────────────────────────
# CONSOLIDACIÓN DE PDFs
# ─────────────────────────────────────────────────────────────
def build_combined_pdf(file_list: list[dict]) -> tuple[bytes, str, int, list[dict]]:
    """Combina varios PDFs en uno solo y devuelve (bytes, hash, total_pages, page_map)."""
    combined = fitz.open()
    page_map: list[dict] = []
    for file_idx, f_info in enumerate(file_list):
        src = fitz.open(stream=f_info["bytes"], filetype="pdf")
        n = len(src)
        combined.insert_pdf(src)
        for local_p in range(n):
            page_map.append({
                "file_idx":   file_idx,
                "local_page": local_p,
                "name":       f_info["name"],
            })
        src.close()
    buf = io.BytesIO()
    combined.save(buf)
    combined.close()
    combined_bytes = buf.getvalue()
    return combined_bytes, _md5(combined_bytes), len(page_map), page_map


# ─────────────────────────────────────────────────────────────
# RENDER DE PÁGINA (sin caché – la gestiona el caller)
# ─────────────────────────────────────────────────────────────
def render_page(pdf_bytes: bytes, idx: int, scale: float = 1.5) -> bytes:
    """Renderiza la página idx del PDF a PNG. Devuelve b'' en caso de error."""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            if idx < 0 or idx >= len(doc):
                return b""
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False
            )
            return pix.tobytes("png")
    except Exception:
        return b""


# ─────────────────────────────────────────────────────────────
# OCR — Motor mejorado con preprocesamiento inteligente y 3 niveles
# 1. PyTesseract (spa+eng, 3.0x, inversión adaptativa + Max RGB)
# 2. RapidOCR (ONNX, reconstrucción espacial de filas)
# 3. PyMuPDF OCR integrado como último recurso
# ─────────────────────────────────────────────────────────────

def _page_has_invisible_ocr(doc: fitz.Document, idx: int) -> bool:
    """Detecta páginas escaneadas con capa OCR invisible incrustada ('3 Tr', ej. iLovePDF)."""
    try:
        page = doc[idx]
        p_area = max(page.rect.width * page.rect.height, 1.0)
        img_cov = 0.0
        for im in page.get_images(full=True):
            for r in page.get_image_rects(im[0]):
                img_cov = max(img_cov, (r.width * r.height) / p_area)
        if img_cov < 0.75:
            return False
        xrefs = page.get_contents() or []
        stream = b"".join(doc.xref_stream(x) or b"" for x in xrefs)
        return (b"3 Tr" in stream) or bool(re.search(rb"\b3\s+Tr\b", stream))
    except Exception:
        return False


def _preprocess_for_ocr(rgb: np.ndarray):
    """Preprocesamiento adaptativo PIL+numpy para fondos oscuros, resaltadores y escaneos."""
    from PIL import Image, ImageFilter

    gray = (0.299 * rgb[:, :, 0] + 0.587 * rgb[:, :, 1] + 0.114 * rgb[:, :, 2]).astype(np.uint8)
    dark_frac = float((gray < 105).mean())
    if dark_frac > 0.30:
        h, w = gray.shape
        sw, sh = max(w // 8, 1), max(h // 8, 1)
        small_dark = Image.fromarray(((gray < 115).astype(np.uint8) * 255)).resize(
            (sw, sh), Image.Resampling.NEAREST
        )
        bg = small_dark.filter(ImageFilter.MaxFilter(11)).filter(ImageFilter.MinFilter(11))
        bg = bg.filter(ImageFilter.MinFilter(11)).filter(ImageFilter.MaxFilter(11))
        bg_full = np.array(bg.resize((w, h), Image.Resampling.NEAREST)) > 127
        inv = 255 - gray
        if bg_full.any():
            p90 = float(np.percentile(inv[bg_full], 90))
            if 140 < p90 < 248:
                inv = np.clip(inv.astype(np.float32) * (255.0 / p90), 0, 255).astype(np.uint8)
        out = gray.copy()
        out[bg_full] = inv[bg_full]
        img = Image.fromarray(out)
    else:
        max_rgb = np.max(rgb, axis=2)
        img = Image.fromarray(max_rgb)

    return img.filter(ImageFilter.UnsharpMask(radius=1.2, percent=130, threshold=3))


def _rapidocr_to_lines(result) -> str:
    """Reconstruye líneas horizontales a partir de las cajas espaciales de RapidOCR."""
    if not result:
        return ""
    boxes: list[tuple[float, float, float, str]] = []
    fallback_texts: list[str] = []
    for r in result:
        if not r or len(r) < 2 or not r[1]:
            continue
        txt = str(r[1]).strip()
        if not txt:
            continue
        pts = r[0]
        try:
            xs = [float(p[0]) for p in pts]
            ys = [float(p[1]) for p in pts]
            x_left = min(xs)
            y_min, y_max = min(ys), max(ys)
            y_center = 0.5 * (y_min + y_max)
            height = max(y_max - y_min, 1.0)
            boxes.append((y_center, x_left, height, txt))
        except Exception:
            fallback_texts.append(txt)

    if not boxes:
        return "\n".join(fallback_texts)

    boxes.sort(key=lambda b: b[0])
    rows: list[list] = []  # [row_y, row_h, [(x_left, txt), ...]]
    for y_c, x_l, h, txt in boxes:
        if rows:
            row_y, row_h, items = rows[-1]
            if abs(y_c - row_y) <= max(8.0, 0.55 * min(h, row_h)):
                n = len(items)
                rows[-1][0] = (row_y * n + y_c) / (n + 1)
                rows[-1][1] = max(row_h, h)
                items.append((x_l, txt))
                continue
        rows.append([y_c, h, [(x_l, txt)]])

    lines_out: list[str] = []
    for _, _, items in rows:
        items.sort(key=lambda it: it[0])
        lines_out.append("   ".join(it[1] for it in items))
    if fallback_texts:
        lines_out.extend(fallback_texts)
    return "\n".join(lines_out)


def _try_pytesseract_page(
    pdf_bytes: bytes,
    idx: int,
    lang: Optional[str] = None,
    scale: float = OCR_SCALE,
    psm: int = 6,
) -> str:
    """OCR con pytesseract + preprocesamiento inteligente (3.0x por defecto)."""
    try:
        import pytesseract

        if os.path.exists(TESSERACT_PATH):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False
            )
            rgb = np.frombuffer(pix.samples, np.uint8).reshape(
                pix.height, pix.width, 3
            ).copy()

        img_proc = _preprocess_for_ocr(rgb)
        lang_use = lang or _resolve_tess_lang("spa+eng")
        cfg = f"--psm {psm}"
        try:
            text = pytesseract.image_to_string(img_proc, lang=lang_use, config=cfg)
        except Exception:
            text = pytesseract.image_to_string(img_proc, lang="spa", config=cfg)
        return text.strip()
    except Exception:
        return ""


def _get_ocr():
    """Devuelve el motor RapidOCR (singleton). None si no está disponible.

    Soporta dos generaciones de rapidocr-onnxruntime:
    - v1.3.x+ (API nueva): RapidOCR(det_model_dir=..., rec_model_dir=..., cls_model_dir=...)
    - v1.2.x  (API vieja, Python 3.14): RapidOCR() sin argumentos de directorio
    """
    global _ocr_engine, _ocr_available
    if _ocr_available is False:
        return None
    if _ocr_engine is not None:
        return _ocr_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
        try:
            _ocr_engine = RapidOCR()
        except Exception:
            _ocr_engine = RapidOCR(
                det_model_dir=None, rec_model_dir=None, cls_model_dir=None
            )
        _ocr_available = True
        return _ocr_engine
    except Exception as exc:
        print(f"[engine] RapidOCR no disponible: {exc}", file=sys.stderr)
        _ocr_available = False
        return None


def _try_rapidocr_page(pdf_bytes: bytes, idx: int, scale: float = RAPIDOCR_SCALE) -> str:
    """OCR con RapidOCR (ONNX) + preprocesamiento inteligente y reconstrucción de filas."""
    ocr = _get_ocr()
    if ocr is None:
        return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False
            )
            rgb = np.frombuffer(pix.samples, np.uint8).reshape(
                pix.height, pix.width, 3
            ).copy()
        img_proc = _preprocess_for_ocr(rgb)
        arr = np.array(img_proc)
        bgr = np.ascontiguousarray(np.stack([arr, arr, arr], axis=-1))
        result, _ = ocr(bgr)
        return _rapidocr_to_lines(result)
    except Exception:
        return ""


def _try_pymupdf_ocr(pdf_bytes: bytes, idx: int, lang: str = "spa") -> str:
    """Último recurso: OCR integrado de PyMuPDF."""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            page = doc[idx]
            tp = page.get_textpage_ocr(language=lang, dpi=300, full=True)
            return page.get_text(textpage=tp).strip()
    except Exception:
        return ""


def ocr_page(
    pdf_bytes: bytes,
    idx: int,
    on_warning: Optional[callable] = None,
    mode: str = "fast",
    psm: int = 6,
    pdf_hash: str = "",
) -> str:
    """OCR de una página con preprocesamiento inteligente, caché y 3 niveles de fallback.

    mode='fast'      : PyTesseract 3.0x -> RapidOCR 2.5x -> PyMuPDF OCR
    mode='precision' : RapidOCR 2.5x -> PyTesseract 4.0x -> PyMuPDF OCR
    """
    cache_key = (pdf_hash, idx, f"{mode}_{psm}") if pdf_hash else None
    if cache_key and cache_key in _OCR_CACHE:
        return _OCR_CACHE[cache_key]

    if mode == "precision":
        text = _try_rapidocr_page(pdf_bytes, idx, scale=RAPIDOCR_SCALE)
        if not (text and len(text) > 20):
            text = _try_pytesseract_page(pdf_bytes, idx, scale=4.0, psm=6)
        if not (text and len(text) > 20):
            text = _try_pymupdf_ocr(pdf_bytes, idx)
    else:
        text = _try_pytesseract_page(pdf_bytes, idx, scale=OCR_SCALE, psm=psm)
        if not (text and len(text) > 20):
            text = _try_rapidocr_page(pdf_bytes, idx, scale=RAPIDOCR_SCALE)
        if not (text and len(text) > 20):
            text = _try_pymupdf_ocr(pdf_bytes, idx)

    if text and len(text) > 20:
        if cache_key:
            _ocr_cache_set(cache_key, text)
        return text

    if on_warning:
        on_warning(f"OCR pág. {idx + 1}: No se pudo extraer texto con ningún motor.")
    return text or ""


# ─────────────────────────────────────────────────────────────
# API BANXICO
# ─────────────────────────────────────────────────────────────
def banxico_tc(
    moneda: str,
    fecha: datetime.date,
    token: str,
    on_warning: Optional[callable] = None,
) -> Optional[float]:
    """Tipo de cambio Banxico. Usa caché de módulo. Emite on_warning ante errores."""
    if moneda not in _BANXICO_SERIES or not token:
        return None
    cache_key = (moneda, fecha.isoformat())
    if cache_key in _BX_CACHE and _BX_CACHE[cache_key] is not None:
        return _BX_CACHE[cache_key]
    serie = _BANXICO_SERIES[moneda]
    f_ini = (fecha - datetime.timedelta(days=5)).isoformat()
    url = f"{_BANXICO_URL}/series/{serie}/datos/{f_ini}/{fecha.isoformat()}"
    try:
        resp = requests.get(
            url,
            headers={"Bmx-Token": token, "Accept": "application/json"},
            timeout=10,
        )
        resp.raise_for_status()
        datos = resp.json()["bmx"]["series"][0]["datos"]
        if not datos:
            return None
        tc = float(datos[-1]["dato"].replace(",", ""))
        _BX_CACHE[cache_key] = tc
        return tc
    except Exception as exc:
        msg = f"Banxico ({moneda} · {fecha}): {exc}"
        if on_warning:
            on_warning(msg)
        else:
            print(f"[engine] {msg}", file=sys.stderr)
        return None


# ─────────────────────────────────────────────────────────────
# RECALCULAR COLUMNAS DERIVADAS
# ─────────────────────────────────────────────────────────────
def recalc_derived(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recalcula columnas derivadas conservadoramente:
    solo rellena campos vacíos (None/NaN), nunca sobreescribe
    valores que el usuario ya proporcionó o editó manualmente.
    """
    df = _normalize_editor_df(df)
    for idx, row in df.iterrows():
        qty = _safe_f(row.get("Cantidad"))
        qty = qty if qty and qty > 0 else 1

        unit = _safe_f(row.get("Precio Unitario"))
        sub  = _safe_f(row.get("Subtotal (Sin IVA)"))
        iva  = _safe_f(row.get("IVA 16%"))
        tot  = _safe_f(row.get("Total con IVA"))
        anx  = _safe_f(row.get("Monto en Anexo Escrito"))

        iva_flag = (
            str(row.get("(+ IVA)", "N/M") or "N/M").strip()
            .replace("\u00ed", "i").replace("\u00c3\u00ad", "i")
        ).lower()
        has_iva    = iva_flag in {"si", "yes", "true", "1", "sí"}
        iva_incl   = iva_flag == "incluido"
        iva_exento = iva_flag in {"exento", "no", "n/m", "nm", "n.a.", "na", ""}

        if sub is None and tot is not None and iva is not None:
            sub = round(tot - iva, 2)
        if unit is None:
            if sub is not None and qty > 0:
                unit = round(sub / qty, 2)
            elif tot is not None and qty > 0:
                base = (tot / 1.16) if (has_iva or iva_incl) else tot
                unit = round(base / qty, 2)
        if sub is None and unit is not None:
            sub = round(qty * unit, 2)
        if sub is not None:
            if iva is None:
                if has_iva or iva_incl:
                    iva = round(sub * 0.16, 2)
                elif iva_exento:
                    iva = 0.0
            if tot is None:
                tot = round(sub + (iva or 0), 2)

        df.at[idx, "Cantidad"]           = int(qty) if float(qty).is_integer() else qty
        df.at[idx, "Precio Unitario"]    = unit
        df.at[idx, "Subtotal (Sin IVA)"] = sub
        df.at[idx, "IVA 16%"]            = iva
        df.at[idx, "Total con IVA"]      = tot
        df.at[idx, "Diferencia final"]   = (
            round(float(tot) - float(anx), 2)
            if tot is not None and anx is not None else None
        )
    return df


# ─────────────────────────────────────────────────────────────
# GHOST GRID PARSER
# ─────────────────────────────────────────────────────────────
_SPACE_TBL_META_RE = re.compile(
    r"\b(?:vigencia|tel[eé]fono|c[oó]digo\s*postal|c\.?\s*p\.?|rfc|iban|clabe"
    r"|correo|e-?mail|issue\s*date|documento\s*nr|data\s*di\s*riferimento)\b"
    r"|\bdescuentos?\b",
    re.IGNORECASE,
)


def _parse_space_table(text: str, is_global: bool = False) -> list[dict]:
    items = []
    MIN_DESC = 4
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if not is_global and _BUDGET_EXCL.search(stripped):
            continue
        if (
            _EXCH_RATE_LINE_RE.search(stripped)
            or _TBL_HEADER_RE.search(stripped)
            or _RET_IVA_RE.search(stripped)
            or _PAYMENT_SCHED_RE.search(stripped)
            or _SPACE_TBL_META_RE.search(stripped)
            or _TBL_SUBTOT_RE.search(stripped)
            or _TBL_IVA_RE.search(stripped)
            or (_TBL_TOTAL_RE.search(stripped) and not _TBL_EXCL_RE.search(stripped))
        ):
            continue
        m = _ITEM_RE.match(stripped)
        if m:
            qty_s, desc, pu_s, total_s = m.groups()
            items.append({
                "qty": int(qty_s), "desc": desc.strip(),
                "pu":  float(pu_s.replace(",", "")),
                "total": float(total_s.replace(",", "")),
            })
            continue
        norm_ln = _norm_ocr_currency_chars(stripped)
        montos = []
        for m2 in _GHOST_MONEY_RE.finditer(norm_ln):
            raw = m2.group(1) or m2.group(2)
            v = _parse_money_token(raw)
            if v and v >= 1.0:
                montos.append(v)
        if not montos:
            continue
        first_match = _GHOST_MONEY_RE.search(norm_ln)
        desc_part = stripped[:first_match.start()].strip().rstrip(":-")
        if len(desc_part) < MIN_DESC:
            continue
        qty_d = 1
        desc_clean = desc_part
        qty_m = re.match(r"^(\d{1,4})\s+(.+)", desc_part)
        if qty_m and int(qty_m.group(1)) <= 9999:
            qty_d = int(qty_m.group(1))
            desc_clean = qty_m.group(2).strip()
        if len(montos) >= 2:
            montos_sorted = sorted(montos)
            pu_v  = montos_sorted[0]
            tot_v = montos_sorted[-1]
            tol = max(0.5, tot_v * 0.05)
            if abs(qty_d * pu_v - tot_v) <= tol or len(montos) == 2:
                items.append({"qty": qty_d, "desc": desc_clean, "pu": pu_v, "total": tot_v})
        elif len(montos) == 1:
            items.append({"qty": qty_d, "desc": desc_clean,
                          "pu": montos[0] / qty_d, "total": montos[0]})
    return items


# ─────────────────────────────────────────────────────────────
# MOTOR DE EXTRACCIÓN PRINCIPAL (HELPERS + EXTRACT)
# ─────────────────────────────────────────────────────────────
def _extract_supp_rate(text: str) -> Optional[float]:
    """Extrae tipo de cambio indicado en la cotización (ej. 21.0016, 18.56, 19.485)."""
    for ln in text.splitlines():
        if not _EXCH_RATE_LINE_RE.search(ln):
            continue
        if re.search(r"data\s+del\s+tasso|fecha", ln, re.IGNORECASE) and not re.search(
            r"tasso\s*d[il]\s*cambio|exchange\s*rate|=", ln, re.IGNORECASE
        ):
            continue
        for m in re.finditer(r"(\d{1,3}[.,]\d{2,4})", ln):
            raw = m.group(1)
            if "," in raw and "." not in raw:
                int_p, tail_p = raw.split(",", 1)
                if len(tail_p) == 3 and int(int_p) < 200:
                    val = float(f"{int_p}.{tail_p}")
                else:
                    val = _safe_f(raw)
            else:
                val = _safe_f(raw.replace(",", ""))
            if val is not None and 1.2 < val < 200.0:
                return val
    return None


def _extract_foreign_total(text: str, tlow: str) -> tuple[Optional[float], Optional[str]]:
    """Detecta total explícito en divisa extranjera (Total documento EUR, Total Period USD)."""
    doc_vals: list[float] = []
    for m in _FOREIGN_DOC_RE.finditer(text):
        v = _parse_money_token(m.group(1))
        if v is not None and v >= 1.0:
            doc_vals.append(v)
    if doc_vals:
        curr = "EUR"
        if re.search(r"\b(usd|dollars?|d[oó]lares)\b|us\$", tlow) and "€" not in text and "eur" not in tlow:
            curr = "USD"
        return max(doc_vals), curr

    usd_vals: list[tuple[float, str]] = []
    for m in _FOREIGN_USD_RE.finditer(text):
        v = _parse_money_token(m.group(1))
        if v is not None and v >= 1.0:
            usd_vals.append((v, m.group(2).upper()))
    if usd_vals:
        return max(usd_vals, key=lambda x: x[0])

    imp_vals: list[float] = []
    for m in re.finditer(r"importo\s+totale\s*[:!\-]?\s*[€$]?\s*([\d.,]+)\s*(?:€|eur)?", text, re.IGNORECASE):
        v = _parse_money_token(m.group(1))
        if v is not None and v >= 1.0:
            imp_vals.append(v)
    if imp_vals:
        return max(imp_vals), "EUR"

    return None, None


def _extract_mxn_equiv(text: str, tlow: str, allow_bare_mxn: bool) -> Optional[float]:
    """Extrae equivalente explícito en MXN dentro de cotizaciones en divisa extranjera."""
    cands: list[float] = []
    for idx_p, pat in enumerate(_MXN_EQUIV_PATTERNS):
        if idx_p == 2 and not allow_bare_mxn:
            continue
        for m in pat.finditer(text):
            v = _parse_money_token(m.group(1))
            if v is not None and 50.0 <= v <= MAX_PLAUSIBLE_MXN:
                cands.append(v)
    return max(cands) if cands else None


def _scan_table_rows(
    rows: list[list[str]],
    tot_in: Optional[float],
    iva_in: Optional[float],
    sub_in: Optional[float],
    is_global: bool = False,
) -> tuple[list[float], Optional[float], Optional[float], Optional[float], int]:
    """Escanea filas de tabla buscando Total / IVA / Subtotal."""
    candidates: list[float] = []
    sub_count = 0
    for row in rows:
        j = "   ".join(row).lower()
        if not is_global and _BUDGET_EXCL.search(j):
            continue
        if _TBL_HEADER_RE.search(j):
            continue
        if _TBL_TOTAL_RE.search(j) and not _TBL_EXCL_RE.search(j):
            cell_moneys = [_money(c) for c in row if c]
            cell_moneys = [v for v in cell_moneys if v]
            v = max(cell_moneys) if cell_moneys else _money_after_label(
                "   ".join(row), _TBL_TOTAL_RE, use_max=True
            )
            if v:
                candidates.append(v)
                if tot_in is None or v > tot_in:
                    tot_in = v
        if _TBL_IVA_RE.search(j) and not _RET_IVA_RE.search(j):
            v = _money_after_label("   ".join(row), _TBL_IVA_RE)
            if v and (iva_in is None or v > iva_in):
                iva_in = v
        if _TBL_SUBTOT_RE.search(j) and not _RET_IVA_RE.search(j):
            v = _money_after_label("   ".join(row), _TBL_SUBTOT_RE)
            if v:
                if sub_in is None or v > sub_in:
                    sub_in = v
                sub_count += 1
    return candidates, tot_in, iva_in, sub_in, sub_count


def _parse_extraction(
    text: str,
    table_rows: list[list[str]],
    tipo: str,
    moneda_in: str,
    det_iva: bool,
    calc_sub: bool,
) -> dict:
    """Ejecuta toda la lógica de interpretación (moneda, fecha, IVA y montos) sobre texto + tablas."""
    tlow = text.lower()
    is_global = (tipo == "Presupuesto Global")

    # ── G1.5: Moneda y Bimoneda ──────────────────────────────
    moneda = moneda_in
    found_currency = True
    if moneda_in == "AUTO":
        if re.search(r"\b(usd|dl[ls]s?|dollars?|dolares|dólares)\b|us\$", tlow):
            moneda = "USD"
        elif re.search(r"\b(eur|euros?)\b|€", tlow):
            moneda = "EUR"
        elif re.search(
            r"ca\$|\bcanadian\b|\bcad\s*\$|\$\s*[\d.,]+\s*cad\b|\b\d[\d.,]*\s+cad\b|moneda\s*[:\-]?\s*cad\b",
            tlow,
        ):
            moneda = "CAD"
        elif re.search(r"\b(mxn|pesos?|m\.n\.|mn)\b", tlow):
            moneda = "MXN"
        else:
            moneda = "MXN"
            found_currency = False

    supp_rate = _extract_supp_rate(text)
    foreign_val, foreign_curr = _extract_foreign_total(text, tlow)
    has_fx_ctx = bool(
        foreign_val is not None
        or supp_rate is not None
        or re.search(r"\b(usd|eur|cad|euro|dolar)\b|€", tlow)
    )
    mxn_equiv = _extract_mxn_equiv(text, tlow, allow_bare_mxn=has_fx_ctx)

    # Validar coherencia de mxn_equiv cuando tenemos total extranjero + tipo de cambio
    if mxn_equiv is not None and foreign_val is not None and supp_rate is not None:
        expected_mxn = foreign_val * supp_rate
        ratio = mxn_equiv / max(expected_mxn, 1.0)
        if not (0.70 <= ratio <= 1.30):
            mxn_equiv = None

    # ── G2: Tipo de IVA ──────────────────────────────────────
    iva_mention = re.search(
        r"\biva\b|i\.?\s*v\.?\s*a\.?|16\s*%|vat|impuesto|exento|tasa\s*0|facturaci[oó]n\s*e?\s*iva|eiva\b",
        tlow,
    )
    iva_zero_explicit = bool(_IVA_ZERO_EXPLICIT_RE.search(text))
    _exempt_hit = bool(_IVA_EXEMPT_RE.search(text))
    if _exempt_hit and (is_global or re.search(r"\bcon\s+iva\b", text, re.IGNORECASE)) and not re.search(
        r"exento|tasa\s*0|no\s+aplica\s+iva|iva\s*:\s*0|libre\s+de\s+iva|no\s+causa\s+iva",
        text, re.IGNORECASE,
    ):
        _exempt_hit = False
    if not det_iva or not iva_mention:
        iva_f = "N/M"
    elif _exempt_hit:
        iva_f = "Exento"
    elif _IVA_INCLUDED_RE.search(text):
        iva_f = "Incluido"
    else:
        iva_f = "Sí"

    tot = iva = sub = pu = None
    qty = 1
    obs_parts: list[str] = []
    if iva_f in ("Exento", "Incluido"):
        obs_parts.append(f"IVA {iva_f}")

    fecha = _extract_best_date(text)
    has_retenciones = bool(_RET_IVA_RE.search(text))
    converted_mxn_in_doc = False

    if moneda_in == "AUTO":
        if mxn_equiv is not None:
            moneda = "MXN"
            tot = mxn_equiv
            converted_mxn_in_doc = True
            found_currency = True
            if foreign_val is not None and foreign_curr:
                obs_parts.append(f"Equiv. MXN en cotización ({foreign_curr} {foreign_val:,.2f})")
            else:
                obs_parts.append("Equiv. MXN en cotización")
        elif foreign_val is not None and foreign_curr:
            moneda = foreign_curr
            tot = foreign_val
            found_currency = True
            obs_parts.append(f"Total en {foreign_curr}")

    lines_clean = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # ── Prioridad Presupuesto Global ─────────────────────────
    if is_global and tot is None:
        global_vals: list[float] = []
        for idx_ln, ln in enumerate(lines_clean):
            if _GLOBAL_TOTAL_RE.search(ln):
                v = _money_max(ln)
                if v is None and idx_ln + 1 < len(lines_clean):
                    v = _money_max(lines_clean[idx_ln + 1])
                if v and v >= 100.0:
                    global_vals.append(v)
        if global_vals:
            tot = max(global_vals)
            obs_parts.append("Presupuesto global")

    # ── Estrategia 1: tablas estructuradas ───────────────────
    tot_candidates, tot_tbl, iva, sub, sub_count = _scan_table_rows(
        table_rows, None, iva, sub, is_global=is_global
    )
    if tot is None:
        tot = tot_tbl
        if sub is not None and len(tot_candidates) > 1:
            factor = 1.0 if iva_f in ("Exento", "N/M") else 1.16
            target = sub * factor
            best = min(tot_candidates, key=lambda x: abs(x - target))
            if abs(best - target) / max(target, 1.0) <= 0.20:
                tot = best
                obs_parts.append("Total seleccionado por coherencia")

    # ── Estrategia 3: totales etiquetados en texto libre ─────
    if tot is None:
        for ln in lines_clean:
            if not is_global and _BUDGET_EXCL.search(ln):
                continue
            if _PAYMENT_SCHED_RE.search(ln) or _TBL_HEADER_RE.search(ln) or _EXCH_RATE_LINE_RE.search(ln):
                continue
            if _RET_IVA_RE.search(ln):
                continue
            if _TBL_TOTAL_RE.search(ln) and not _TBL_EXCL_RE.search(ln):
                v = _money_after_label(ln, _TBL_TOTAL_RE, use_max=True, allow_bare_int=True)
                if v and (tot is None or v > tot):
                    tot = v

    # ── G3: Pares label:valor en líneas adyacentes ─────────────
    _g3_hdr_guard = re.compile(
        r"c[oó]digo|descripci[oó]n|unidad|cantidad|precio\s*unit|concepto|descrizione|prezzo",
        re.IGNORECASE,
    )
    if tot is None or sub is None:
        for i_ln, ln_a in enumerate(lines_clean[:-1]):
            ln_b = lines_clean[i_ln + 1]
            if not is_global and _BUDGET_EXCL.search(ln_a):
                continue
            if _RET_IVA_RE.search(ln_a) or _EXCH_RATE_LINE_RE.search(ln_a):
                continue
            if _money(ln_a) is not None:
                continue
            prev_win = lines_clean[max(0, i_ln - 3):i_ln + 1]
            if any(_g3_hdr_guard.search(p) for p in prev_win):
                continue
            v = _money(ln_b, allow_bare_int=False)
            if v is None:
                continue
            ln_a_low = ln_a.lower()
            if _TBL_TOTAL_RE.search(ln_a_low) and not _TBL_EXCL_RE.search(ln_a_low) and (tot is None or v > tot):
                tot = v
                obs_parts.append("Total (línea adyacente)")
            elif sub is None and _TBL_SUBTOT_RE.search(ln_a_low):
                sub = v
                obs_parts.append("Subtotal (línea adyacente)")

    # ── Estrategia 3.5: "es de $X" / "será de $X" ────────────
    if tot is None:
        for ln in lines_clean:
            if not is_global and _BUDGET_EXCL.search(ln):
                continue
            m_ed = _ES_DE_RE.search(ln)
            if m_ed:
                v = _parse_money_token(re.sub(r"[^0-9.,]", "", m_ed.group(1)))
                if v and 10.0 <= v <= MAX_PLAUSIBLE_MXN:
                    tot = v
                    obs_parts.append("Precio directo")
                    break

    # ── Estrategia 2 (R1): triangulación qty × pu ≈ total ────
    valid_line_totals: list[float] = []
    seen_lines: set[tuple] = set()
    _non_item_row_re = re.compile(
        r"tasso|exchange|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre"
        r"|january|february|march|april|june|july|august|september|october|november|december",
        re.IGNORECASE,
    )
    if tot is None:
        for row in table_rows:
            row_str = " ".join(str(c) for c in row)
            if (
                _EXCH_RATE_LINE_RE.search(row_str)
                or _RET_IVA_RE.search(row_str)
                or _PAYMENT_SCHED_RE.search(row_str)
                or _TBL_HEADER_RE.search(row_str)
                or _non_item_row_re.search(row_str)
            ):
                continue
            nums: list[float] = []
            for token in re.findall(r"\d+(?:[.,]\d+)*", row_str):
                # Excluir códigos SKU de 7-8 dígitos sin decimales/comas y años 2020-2035
                if token.isdigit() and (len(token) >= 7 or 2020 <= int(token) <= 2035):
                    continue
                n = _parse_money_token(token) if ("," in token or "." in token) else _safe_f(token)
                if n is not None and n > 0:
                    nums.append(n)
            if len(nums) < 3:
                continue
            found = False
            for i_t, t_cand in enumerate(nums):
                if t_cand < 10.0 or found:
                    continue
                for i_p, p_cand in enumerate(nums):
                    if i_p == i_t or p_cand < 10.0 or found:
                        continue
                    if p_cand == int(p_cand) and 2020 <= int(p_cand) <= 2035:
                        continue
                    for i_q, q_cand in enumerate(nums):
                        if i_q in (i_t, i_p):
                            continue
                        if not (1 <= q_cand <= 9999 and q_cand == int(q_cand)):
                            continue
                        tol = max(0.5, t_cand * 0.01)
                        if abs(q_cand * p_cand - t_cand) <= tol:
                            key = (int(q_cand), round(p_cand, 2), round(t_cand, 2))
                            if key not in seen_lines:
                                seen_lines.add(key)
                                valid_line_totals.append(t_cand)
                            found = True
                            break
                    if found:
                        break

        if valid_line_totals:
            tot = round(sum(valid_line_totals), 2)
            obs_parts.append("Total por líneas")

    # ── Subtotal e IVA por texto libre ───────────────────────
    if sub is None:
        for ln in lines_clean:
            if _RET_IVA_RE.search(ln) or _TBL_HEADER_RE.search(ln):
                continue
            if _TBL_SUBTOT_RE.search(ln):
                v = _money_after_label(ln, _TBL_SUBTOT_RE)
                if v and (tot is None or has_retenciones or v <= tot * 1.05):
                    sub = v
                    break
        if sub is None:
            _imp_sub_re = re.compile(
                r"^\s*importe\s*[:$]|\bimporte\s+(?:neto|total|sin\s*iva)\s*[:$]",
                re.IGNORECASE,
            )
            for ln in lines_clean:
                if (
                    _RET_IVA_RE.search(ln)
                    or _TBL_HEADER_RE.search(ln)
                    or _PAYMENT_SCHED_RE.search(ln)
                    or re.search(r"honorarios?\s+mensual|dicho\s+importe", ln, re.IGNORECASE)
                ):
                    continue
                if _imp_sub_re.search(ln) and not _TBL_TOTAL_RE.search(ln):
                    v = _money_after_label(ln, _imp_sub_re)
                    if v and (tot is None or has_retenciones or v <= tot * 1.05):
                        sub = v
                        break

    if iva is None and iva_f in ("Sí", "Incluido"):
        _strict_iva_re = re.compile(r"\biva\b|i\.?\s*v\.?\s*a\.?|16\s*%|vat|\bimpuestos?\b", re.IGNORECASE)
        for ln in lines_clean:
            if _RET_IVA_RE.search(ln) or _TBL_HEADER_RE.search(ln):
                continue
            if _strict_iva_re.search(ln) and not (_TBL_SUBTOT_RE.search(ln) and "iva" not in ln.lower()):
                v = _money_after_label(ln, _strict_iva_re)
                if v and (tot is None or v < tot):
                    iva = v
                    break
        # Tolerancia OCR en etiqueta IVA ('wa', '(VA)', '1A') solo validado con 16% del subtotal
        if iva is None and sub is not None and sub > 0:
            _ocr_iva_re = re.compile(r"(?<!\w)(?:\(?[i1l]?va\)?|wa|1a)(?!\w)", re.IGNORECASE)
            for ln in lines_clean:
                if _RET_IVA_RE.search(ln):
                    continue
                if _ocr_iva_re.search(ln):
                    v = _money_after_label(ln, _ocr_iva_re)
                    if v and abs(v - sub * 0.16) / max(v, 1.0) < 0.08:
                        iva = v
                        break

    if iva is None and iva_zero_explicit:
        iva = 0.0
        if iva_f == "Sí":
            iva_f = "Exento"
            if "IVA Exento" not in obs_parts:
                obs_parts.append("IVA Exento")

    # ── V1+O3: Cota de Cordura Proporcional ──────────────────
    incoherent = False
    if tot is not None:
        _iva_coherent = False
        if sub is not None and iva is not None and tot > 0:
            if abs(sub + iva - tot) / tot < 0.02:
                _iva_coherent = True
        _multi_section = sub_count >= 2
        if not has_retenciones and not _iva_coherent and not _multi_section:
            if sub is not None and tot > sub * 1.5:
                obs_parts.append("⚠ Total descartado (desproporcionado vs subtotal)")
                tot = None
                incoherent = True

    # ── Estrategia 4 (R2): fallback – máximo en contexto monetario
    if tot is None and not incoherent:
        all_vals: list[float] = []
        for ln in lines_clean:
            if not is_global and _BUDGET_EXCL.search(ln):
                continue
            if (
                _PAYMENT_SCHED_RE.search(ln)
                or _EXCH_RATE_LINE_RE.search(ln)
                or _RET_IVA_RE.search(ln)
                or _TBL_SUBTOT_RE.search(ln)
                or _TBL_IVA_RE.search(ln)
            ):
                continue
            if not _MONETARY_CTX.search(ln):
                continue
            for m in _MONEY_RE.finditer(_norm_ocr_currency_chars(ln)):
                raw = m.group(1) or m.group(2) or m.group(3) or m.group(4)
                v = _parse_money_token(raw)
                if v and 10.0 <= v <= MAX_PLAUSIBLE_MXN:
                    all_vals.append(v)
        if all_vals:
            if sub is not None:
                all_vals = [v for v in all_vals if v <= sub * 1.5]
            elif valid_line_totals:
                sum_lines = sum(valid_line_totals)
                all_vals = [v for v in all_vals if v <= sum_lines * 1.5]
            if all_vals:
                cand_tot = max(all_vals)
                if not (sub is not None and iva_f == "Sí" and abs(cand_tot - sub) <= 0.02):
                    tot = cand_tot
                    obs_parts.append("Total inferido (máx. validado)")

    # Evitar contaminación de sub/iva extranjeros en cotizaciones con equivalente MXN explícito
    if converted_mxn_in_doc and not has_retenciones:
        if sub is not None and tot is not None and abs(tot - sub * 1.16) / max(tot, 1.0) > 0.03:
            sub = None
        if iva is not None and tot is not None:
            expected_iva = tot - (tot / 1.16)
            if abs(iva - expected_iva) / max(expected_iva, 1.0) > 0.05:
                iva = None

    # Inferencia matemática de IVA 16% cuando Subtotal y Total están en proporción exacta 1.16
    if (
        det_iva
        and not has_retenciones
        and sub is not None
        and tot is not None
        and sub > 0
        and tot > sub
        and iva_f not in ("Exento",)
        and abs(tot - sub * 1.16) / max(tot, 1.0) <= 0.003
    ):
        if iva_f == "N/M":
            iva_f = "Sí"
        if iva is None or abs(sub + iva - tot) > 1.0:
            iva = round(tot - sub, 2)

    # Validar consistencia de IVA con Subtotal cuando no hay retenciones
    if not has_retenciones and sub is not None and sub > 0 and iva is not None and iva > 0:
        if iva_f == "Sí" and tot is None:
            if not (0.13 <= (iva / sub) <= 0.19):
                iva = None
        elif iva_f == "Sí" and tot is not None and tot > sub:
            if abs(sub + iva - tot) / max(tot, 1.0) > 0.05 and abs(sub * 1.16 - tot) / max(tot, 1.0) < 0.02:
                iva = round(tot - sub, 2)

    # Cantidad y P.U. (solo de filas MXN coherentes)
    if not converted_mxn_in_doc:
        best_row_total = 0.0
        for row in table_rows:
            row_str = " ".join(str(c) for c in row)
            if _EXCH_RATE_LINE_RE.search(row_str) or _RET_IVA_RE.search(row_str):
                continue
            nums_: list[float] = []
            for t in re.findall(r"\d+(?:[.,]\d+)*", row_str):
                if t.isdigit() and len(t) >= 7:
                    continue
                n_val = _parse_money_token(t) if ("," in t or "." in t) else _safe_f(t)
                if n_val is not None:
                    nums_.append(n_val)
            if len(nums_) >= 3 and 1 <= nums_[0] <= 9999:
                row_tot = nums_[-1]
                if row_tot > best_row_total:
                    best_row_total = row_tot
                    qty = int(nums_[0])
                    pu  = nums_[-2] if len(nums_) > 2 else nums_[-1]

    return {
        "moneda": moneda,
        "found_currency": found_currency,
        "iva_f": iva_f,
        "tot": tot,
        "iva": iva,
        "sub": sub,
        "pu": pu,
        "qty": qty,
        "fecha": fecha,
        "obs_parts": obs_parts,
        "incoherent": incoherent,
        "has_retenciones": has_retenciones,
        "converted_mxn_in_doc": converted_mxn_in_doc,
    }


def _needs_precision_ocr(parsed: dict) -> bool:
    """Determina si conviene escalar a Pass 2 (RapidOCR de alta precisión)."""
    tot = parsed["tot"]
    sub = parsed["sub"]
    iva_f = parsed["iva_f"]
    if tot is None or parsed["incoherent"]:
        return True
    if (
        tot is not None
        and sub is not None
        and abs(tot - sub) / max(tot, 1.0) <= 0.02
        and iva_f in ("Sí", "Incluido")
        and "Total inferido (máx. validado)" in parsed["obs_parts"]
    ):
        return True
    if (
        sub is not None
        and tot is not None
        and iva_f in ("Sí", "Incluido")
        and not parsed["has_retenciones"]
        and abs(tot - sub * 1.16) / max(tot, 1.0) > 0.08
        and abs(tot - sub) / max(tot, 1.0) > 0.02
    ):
        return True
    return False


def _score_parsed(parsed: dict) -> int:
    """Puntúa la calidad y coherencia de una extracción para elegir entre Pass 1 y Pass 2."""
    score = 0
    tot = parsed["tot"]
    sub = parsed["sub"]
    iva_f = parsed["iva_f"]
    if tot is not None and not parsed["incoherent"]:
        score += 100
        if sub is not None:
            if parsed["has_retenciones"] or iva_f in ("N/M", "Exento"):
                score += 40
            elif abs(tot - sub * 1.16) / max(tot, 1.0) <= 0.03:
                score += 50
            elif abs(tot - sub) / max(tot, 1.0) <= 0.02:
                score += 25
        else:
            score += 20
    if sub is not None:
        score += 20
    if parsed["fecha"] is not None:
        score += 10
    return score


def extract(
    pdf_bytes: bytes,
    label: str,
    p0: int,
    p1: int,
    det_iva: bool,
    calc_sub: bool = True,
    tipo: str = "Cotización Proveedor",
    moneda: str = "AUTO",
    bx_token: str = "",
    sec_num: int = 1,
    on_progress: Optional[callable] = None,
    on_warning:  Optional[callable] = None,
    force_ocr:   bool = False,
) -> dict:
    """
    Motor de extracción de montos (texto nativo + OCR inteligente 2 pasadas + 5 estrategias).
    on_progress(msg: str) – llamado al avanzar fases.
    on_warning(msg: str)  – llamado ante condiciones no fatales.
    """
    is_global = (tipo == "Presupuesto Global")
    moneda_in = moneda
    pdf_hash = _md5(pdf_bytes)
    native_text_parts: list[str] = []
    table_rows: list[list[str]] = []
    ocr_pages: list[int] = []
    img_pages: list[int] = []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as fitz_doc, \
         pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        pr = range(max(0, p0 - 1), min(p1, n_pages))
        for i in pr:
            page_f = fitz_doc[i]
            p_area = max(page_f.rect.width * page_f.rect.height, 1.0)
            has_big_img = any(
                (r.width * r.height) / p_area >= 0.75
                for im in page_f.get_images(full=True)
                for r in page_f.get_image_rects(im[0])
            )
            if has_big_img:
                img_pages.append(i)

            if force_ocr or _page_has_invisible_ocr(fitz_doc, i):
                ocr_pages.append(i)
                continue

            pg = pdf.pages[i]
            txt = pg.extract_text() or ""
            if len(txt.strip()) < NATIVE_MIN_CHARS_PER_PAGE:
                ocr_pages.append(i)
                continue

            native_text_parts.append(txt)
            for tbl in pg.extract_tables() or []:
                if tbl:
                    table_rows.extend(
                        [str(c or "").replace("\n", " ").strip() for c in row]
                        for row in tbl if row
                    )
            for item in _parse_space_table(txt, is_global=is_global):
                table_rows.append([
                    str(item["qty"]), item["desc"],
                    f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                ])

    native_text = "\n".join(native_text_parts)
    ocr_used = bool(ocr_pages)
    ocr_mode_default = "precision" if force_ocr else "fast"
    ocr_psm_default = 3 if is_global else 6

    ocr_text_p1 = ""
    ocr_rows_p1: list[list[str]] = []
    global_fast_solved = False

    if ocr_pages:
        if on_progress:
            on_progress("OCR de alta precisión…" if force_ocr else "OCR de páginas…")

        # Optimización para Presupuestos Globales multipágina: priorizar portada y hoja final
        if is_global and len(ocr_pages) > 6 and not force_ocr:
            prio_pages = list(dict.fromkeys([
                ocr_pages[0], ocr_pages[-1], ocr_pages[1], ocr_pages[-2]
            ]))
            prio_parts: list[str] = []
            for idx_p in prio_pages:
                o_txt = ocr_page(pdf_bytes, idx_p, on_warning=on_warning, mode="fast", psm=3, pdf_hash=pdf_hash)
                if o_txt:
                    prio_parts.append(o_txt)
                    for item in _parse_space_table(o_txt, is_global=True):
                        ocr_rows_p1.append([
                            str(item["qty"]), item["desc"],
                            f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                        ])
            prio_text_joined = "\n".join([native_text, *prio_parts])
            prio_parsed = _parse_extraction(
                prio_text_joined,
                table_rows + ocr_rows_p1, tipo, moneda_in, det_iva, calc_sub
            )
            has_explicit_global_lbl = any(
                re.search(r"presupuesto\s+total|costo\s+total", l, re.IGNORECASE) and _money_max(l) is not None
                for l in prio_text_joined.splitlines()
            )
            if prio_parsed["tot"] is not None and has_explicit_global_lbl:
                ocr_text_p1 = "\n".join(prio_parts)
                global_fast_solved = True
            else:
                if on_progress:
                    on_progress("OCR de alta precisión (Presupuesto Global)…")
                edge_pages = list(dict.fromkeys([ocr_pages[0], ocr_pages[-1]]))
                edge_parts: list[str] = []
                edge_rows: list[list[str]] = []
                for idx_p in edge_pages:
                    o_txt = ocr_page(pdf_bytes, idx_p, on_warning=on_warning, mode="precision", psm=3, pdf_hash=pdf_hash)
                    if o_txt:
                        edge_parts.append(o_txt)
                        for item in _parse_space_table(o_txt, is_global=True):
                            edge_rows.append([
                                str(item["qty"]), item["desc"],
                                f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                            ])
                edge_parsed = _parse_extraction(
                    "\n".join([native_text, *edge_parts]),
                    table_rows + edge_rows, tipo, moneda_in, det_iva, calc_sub
                )
                if edge_parsed["tot"] is not None:
                    ocr_text_p1 = "\n".join(edge_parts)
                    ocr_rows_p1 = edge_rows
                    global_fast_solved = True

        if not global_fast_solved:
            parts_p1: list[str] = []
            for idx_p in ocr_pages:
                o_txt = ocr_page(
                    pdf_bytes, idx_p, on_warning=on_warning,
                    mode=ocr_mode_default, psm=ocr_psm_default, pdf_hash=pdf_hash,
                )
                if o_txt:
                    parts_p1.append(o_txt)
                    for item in _parse_space_table(o_txt, is_global=is_global):
                        ocr_rows_p1.append([
                            str(item["qty"]), item["desc"],
                            f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                        ])
            ocr_text_p1 = "\n".join(parts_p1)

    text_p1 = "\n".join(p for p in (native_text, ocr_text_p1) if p)
    rows_p1 = table_rows + ocr_rows_p1
    parsed = _parse_extraction(text_p1, rows_p1, tipo, moneda, det_iva, calc_sub)
    text_best = text_p1
    all_texts_for_words = [text_p1]

    # ── G5: OCR selectivo para PDFs mixtos ────────────────────
    if not ocr_used and not table_rows and parsed["tot"] is None and parsed["sub"] is None:
        ocr_used = True
        if on_progress:
            on_progress("OCR selectivo (tabla imagen)…")
        sel_parts: list[str] = []
        sel_rows: list[list[str]] = []
        for idx_p in pr:
            o_txt = ocr_page(
                pdf_bytes, idx_p, on_warning=on_warning,
                mode=ocr_mode_default, psm=ocr_psm_default, pdf_hash=pdf_hash,
            )
            if o_txt:
                sel_parts.append(o_txt)
                for item in _parse_space_table(o_txt, is_global=is_global):
                    sel_rows.append([
                        str(item["qty"]), item["desc"],
                        f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                    ])
        if sel_parts:
            text_p1 = "\n".join(p for p in (text_p1, *sel_parts) if p)
            rows_p1 = rows_p1 + sel_rows
            parsed = _parse_extraction(text_p1, rows_p1, tipo, moneda, det_iva, calc_sub)
            parsed["obs_parts"].append("OCR-Selectivo (tabla imagen)")
            text_best = text_p1
            all_texts_for_words.append(text_p1)

    # ── Pass 2: Escalado automático a RapidOCR ───────────────
    if not force_ocr and not global_fast_solved and (ocr_used or img_pages) and _needs_precision_ocr(parsed):
        target_p2 = ocr_pages if ocr_pages else list(pr)
        if is_global and len(target_p2) > 6:
            target_p2 = list(dict.fromkeys([target_p2[0], target_p2[-1]]))
        if on_progress:
            on_progress("OCR de alta precisión (RapidOCR)…")
        p2_parts: list[str] = []
        p2_rows: list[list[str]] = []
        for idx_p in target_p2:
            o_txt = ocr_page(pdf_bytes, idx_p, on_warning=on_warning, mode="precision", psm=ocr_psm_default, pdf_hash=pdf_hash)
            if o_txt:
                p2_parts.append(o_txt)
                for item in _parse_space_table(o_txt, is_global=is_global):
                    p2_rows.append([
                        str(item["qty"]), item["desc"],
                        f"${item['pu']:,.2f}", f"${item['total']:,.2f}",
                    ])
        if p2_parts:
            text_p2 = "\n".join(p for p in (native_text, *p2_parts) if p)
            all_texts_for_words.append(text_p2)
            rows_p2 = table_rows + p2_rows
            parsed2 = _parse_extraction(text_p2, rows_p2, tipo, moneda_in, det_iva, calc_sub)
            if _score_parsed(parsed2) >= _score_parsed(parsed):
                if parsed2["fecha"] is None and parsed["fecha"] is not None:
                    parsed2["fecha"] = parsed["fecha"]
                if parsed2["iva_f"] == "N/M" and parsed["iva_f"] == "Sí":
                    parsed2["iva_f"] = "Sí"
                parsed = parsed2
                text_best = text_p2
                parsed["obs_parts"].append("OCR-Precisión")

    moneda = parsed["moneda"]
    found_currency = parsed["found_currency"]
    iva_f = parsed["iva_f"]
    tot = parsed["tot"]
    iva = parsed["iva"]
    sub = parsed["sub"]
    pu = parsed["pu"]
    qty = parsed["qty"]
    fecha = parsed["fecha"]
    has_retenciones = parsed["has_retenciones"]
    obs_parts = list(parsed["obs_parts"])
    if ocr_used and "OCR" not in obs_parts:
        obs_parts.insert(0, "OCR")

    if not found_currency and moneda_in == "AUTO" and on_warning:
        on_warning(f"Sección '{label}': Moneda no explícita en texto. Se asume MXN. Por favor corrobore.")

    # ── Auto-reparación con montos en letra ──────────────────
    if tot is None:
        words_val = None
        for wt in all_texts_for_words:
            words_val = _parse_spanish_amount_words(wt)
            if words_val is not None:
                break
        if words_val is not None:
            if sub is not None and sub > 0:
                if abs(words_val - sub * 1.16) / max(words_val, 1.0) <= 0.02:
                    tot = words_val
                    iva = round(tot - sub, 2)
                    if iva_f == "N/M":
                        iva_f = "Sí"
                    obs_parts.append("Total validado (subtotal + letra)")
                elif abs(words_val - sub) / max(words_val, 1.0) <= 0.02:
                    tot = words_val
                    obs_parts.append("Total validado (subtotal + letra)")
            else:
                tot = words_val
                obs_parts.append("Total por texto en letra — verificar")

    if tot is not None and tot > MAX_PLAUSIBLE_MXN:
        obs_parts.append(f"⚠ Valor muy alto ({tot:,.2f}) — verificar")

    # ── Triangulación final (G2: 4 estados de IVA) ───────────
    if iva_f == "Incluido":
        if tot and not sub:
            sub = round(tot / 1.16, 2)
            iva = round(tot - sub, 2)
            obs_parts.append("IVA descompuesto del total")
        elif tot and sub and not iva and tot > sub:
            iva = round(tot - sub, 2)
        elif sub and not tot:
            tot = sub
            sub = round(tot / 1.16, 2)
            iva = round(tot - sub, 2)
    elif iva_f == "Exento":
        if tot and not sub:
            sub = tot
        elif sub and not tot:
            tot = sub
        iva = 0.0
    elif iva_f == "Sí":
        if tot and sub and abs(tot - sub) < 0.02 and iva == 0.0:
            pass
        elif tot and not sub and not iva:
            sub = round(tot / 1.16, 2)
            iva = round(tot - sub, 2)
            obs_parts.append("IVA desglosado")
        elif tot and iva and not sub:
            sub = round(tot - iva, 2)
        elif tot and sub and not iva:
            iva = round(tot - sub, 2)
            if iva < 0:
                iva = None
        elif sub and iva and not tot:
            tot = round(sub + iva, 2)
        elif calc_sub and sub and not iva and not tot:
            iva = round(sub * 0.16, 2)
            tot = round(sub + iva, 2)
    else:  # N/M
        if tot and iva and not sub:
            sub = round(tot - iva, 2)
        elif sub and iva and not tot:
            tot = round(sub + iva, 2)

    if pu is None and qty > 0:
        if sub is not None:
            pu = round(sub / qty, 2)
        elif tot is not None:
            if iva_f in ("Sí", "Incluido"):
                pu = round((tot / 1.16) / qty, 2)
            else:
                pu = round(tot / qty, 2)

    # Conversión Banxico
    tc: Optional[float] = None
    if moneda != "MXN" and bx_token:
        fecha_tc = fecha or datetime.date.today()
        tc = banxico_tc(moneda, fecha_tc, bx_token, on_warning=on_warning)
        if tc and tot is not None:
            tot_orig = tot
            tot = round(tot * tc, 2)
            sub = round(sub * tc, 2) if sub else None
            iva = round(iva * tc, 2) if iva else None
            pu  = round(pu  * tc, 2) if pu  else None
            obs_parts.append(f"1 {moneda} = ${tc:.4f} MXN")
            obs_parts.append(f"Total orig.: {moneda} {tot_orig:,.2f}")

    if tot is None and on_warning:
        on_warning(f"Sección '{label}': No se detectó un Total claro. Corrobora los montos manualmente.")

    return {
        "Nº Sec":                  sec_num,
        "Sección":                 label,
        "Tipo":                    tipo,
        "Fecha":                   fecha.isoformat() if fecha else datetime.date.today().isoformat(),
        "Rubro":                   label,
        "QT":                      "Sí",
        "T. Cambio":               moneda,
        "(+ IVA)":                 iva_f,
        "Cantidad":                qty,
        "Precio Unitario":         pu,
        "Subtotal (Sin IVA)":      sub,
        "IVA 16%":                 iva,
        "Total con IVA":           tot,
        "Diferencia final":        None,
        "Monto en Anexo Escrito":  None,
        "Observaciones":           " | ".join(obs_parts) if obs_parts else "",
    }


# ─────────────────────────────────────────────────────────────
# AUTO-DETECCIÓN DE SECCIONES (PDF ÚNICO)
# ─────────────────────────────────────────────────────────────
from concurrent.futures import ThreadPoolExecutor as _AutoThreadPool, as_completed as _auto_as_completed

_AUTOSEC_TOTAL_RE = re.compile(
    r"\btot(?:ales?|ale?)\b|\btora[lt1!]\b|\btot(?=\s*[:$€\d])|importe\s+total|monto\s+total|gran\s+total"
    r"|total\s+documento|totale?\s+articoli|importe\s+con\s+letra"
    r"|\b\d{1,2}\s*/\s*100\s*m\.?\s*n\.?",
    re.IGNORECASE,
)
_AUTOSEC_CLOSING_RE = re.compile(
    r"atentamente|representante\s+legal|importe\s+con\s+letra"
    r"|vigencia\s+del?\s+(?:la\s+)?presente|\b\d{1,2}\s*/\s*100\s*m\.?\s*n\.?|\bfirma\b",
    re.IGNORECASE,
)
_AUTOSEC_OPENING_RE = re.compile(
    r"por\s+medio\s+de\s+la\s+presente\s+se\s+env[ií]a\s+propuesta"
    r"|hago\s+de\s+su\s+conocimiento\s+la\s+cotizaci[oó]n"
    r"|le\s+hago\s+llegar\s+la\s+siguiente\s+cotizaci[oó]n"
    r"|presupuesto\s+de\s+servicios\s+profesionales"
    r"|^\s*presupuesto\s*:"
    r"|^\s*(?:cotizaci[oó]n|quotation|preventivo(?:\s*/\s*cotizaci[oó]n)?)\b"
    r"|propuesta\s+(?:comercial|econ[oó]mica|de\s+servicios)"
    r"|atenci[oó]n\s+a\s*:"
    r"|documento\s*nr\s*:?\s*\d+",
    re.IGNORECASE | re.MULTILINE,
)
_AUTOSEC_CAPTION_RE = re.compile(
    r"^cotizaci[oó]n\s+correspondiente\b",
    re.IGNORECASE,
)
_AUTOSEC_SUPPORT_RE = re.compile(
    r"c[eé]dula\s+profesional|fo-of-cprs|constancia\s+de\s+situaci[oó]n\s+fiscal"
    r"|grados\s+acj|registro\s+federal\s+de\s+contribuyentes.*acuse",
    re.IGNORECASE,
)
_AUTOSEC_GENERIC_CAPS_RE = re.compile(
    r"^(?:press?upuesto\s+(?:desglosado|global|general)|presupuesto:\s*inmueble|giant\s+bicycle"
    r"|training\s*peaks?|constructora\s+cicomex|aremcox|meza\b|fianza\s+de\s+obra"
    r"|datos\s+del\s+cliente|cotizaci[oó]n|quotation|preventivo|condiciones"
    r"|anexos?|efideporte|director\s+responsable|gerardo\s+reyes|domain\s+resources"
    r"|grados\s+acj|c[eé]dula\s+profesional|secretar[ií]a\s+de\s+educaci[oó]n)\b",
    re.IGNORECASE,
)
_AUTOSEC_MANIFEST_TOPICS: tuple[tuple[re.Pattern, str], ...] = (
    (re.compile(r"\bgastos\s+administrativos\b", re.I), "Carta: Gastos administrativos"),
    (re.compile(r"\bpropuesta\s+del?\s+personal\b", re.I), "Carta: Propuesta del personal"),
    (re.compile(r"\bmantenimiento\s+(?:posterior|al\s+proyecto|del\s+proyecto)\b", re.I), "Carta: Mantenimiento del proyecto"),
    (re.compile(r"\bgastos\s+de\s+capacitaci[oó]n\b", re.I), "Carta: Gastos de capacitación"),
    (re.compile(r"\badquisici[oó]n\s+de\s+bienes\b", re.I), "Carta: Adquisición de bienes"),
    (re.compile(r"\brecursos?\s+de\s+ganancias\s+propias\b", re.I), "Carta: Ganancias propias"),
    (re.compile(r"\brubros?\s+no\s+permitidos\b", re.I), "Carta: Rubros no permitidos"),
    (re.compile(r"\bmanifiest[oa]\w*\s+(?:de\s+)?inter[eé]s\b", re.I), "Carta: Manifiesto de interés"),
)
_AUTOSEC_MANIFEST_CTX_RE = re.compile(
    r"efideporte|representante\s+legal|en\s+mi\s+car[aá]cter|bajo\s+protesta|manifiest",
    re.IGNORECASE,
)
_AUTOSEC_VENDOR_KW_RE = re.compile(
    r"\b(?:zurich|cicomex|aremcox|sofimex|sinergia|giant|training\s*peaks?"
    r"|logistica\s+perfect|ar\.?\s*performance|medical\s+premium"
    r"|domain\s+resources|intelicorp|tecno\s*office|office\s*depot|mybike"
    r"|meza\b|alons[oa]\s+abrego|rodr[ií]guez\s+moedano|aranzabal|espinosa\s+espinosa"
    r"|gerardo\s+reyes|petrarq|deportan)\b",
    re.IGNORECASE,
)
_AUTOSEC_EMAIL_RE = re.compile(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", re.IGNORECASE)
_AUTOSEC_URL_RE   = re.compile(r"(?:https?://|www\.)([a-z0-9.\-]+\.[a-z]{2,})", re.IGNORECASE)
_AUTOSEC_RFC_RE   = re.compile(r"\b[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{2,3}\b")
_AUTOSEC_PHONE_RE = re.compile(r"(?:\+?52[\s\-.]*)?(?:\(?\d{2,3}\)?[\s\-.]*){2,4}\d{2,4}")
_AUTOSEC_TAIL_STAGE_RE = re.compile(
    r"temporada|pretemporada|competencias|campeonato|\(\s*\d{1,2}\s+de\s+[a-z]",
    re.IGNORECASE,
)
_AUTOSEC_TAIL_EXCL_RE = re.compile(
    r"[$€]|\b(?:pesos|mxn|m\.?n\.?|00/100|total|subtotal|iva|s\.?a\.?|s\.?r\.?l\.?|c\.?v\.?"
    r"|rfc|atentamente|pieza|pza|iaaf|certificado|medida)\b",
    re.IGNORECASE,
)
_AUTOSEC_TOC_RANGE_RE = re.compile(r"\(\s*p\.?\s*(\d+)\s*[-–]\s*(\d+)\s*\)", re.IGNORECASE)


def _same_page_geometry(
    a: tuple[int, int],
    b: tuple[int, int],
    ignore_rotation: bool = True,
    tol: int = 12,
) -> bool:
    if ignore_rotation:
        a0, a1 = sorted(a)
        b0, b1 = sorted(b)
    else:
        a0, a1 = a
        b0, b1 = b
    return abs(a0 - b0) <= tol and abs(a1 - b1) <= tol


def _clean_section_label(raw: str, max_len: int = 95) -> str:
    s = re.sub(r"\s+", " ", (raw or "")).strip(" |-–—:;,.")
    if len(s) > max_len:
        s = s[: max_len - 1].rstrip() + "…"
    return s


def _strip_ocr_noise(lines: list[str]) -> list[str]:
    cleaned = []
    for ln in lines:
        s = re.sub(r"\s+", " ", ln).strip()
        if len(s) < 3:
            continue
        alnum = sum(c.isalnum() for c in s)
        if alnum < 3 or alnum / max(len(s), 1) < 0.45:
            continue
        cleaned.append(s)
    return cleaned


def _extract_caps_title(lines: list[str]) -> str:
    """Detecta un título de rubro en MAYÚSCULAS en las primeras 1-3 líneas."""
    cand_lines: list[str] = []
    for ln in lines[:4]:
        s = re.sub(r"\s+", " ", ln).strip()
        if not s:
            continue
        if len(s) <= 4 and cand_lines:
            break
        letters = [c for c in s if c.isalpha()]
        if not letters:
            if cand_lines:
                break
            continue
        up_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
        if up_ratio >= 0.78:
            if not cand_lines and _AUTOSEC_GENERIC_CAPS_RE.search(s):
                return ""
            cand_lines.append(s)
            if len(cand_lines) >= 3:
                break
        elif cand_lines:
            break
        else:
            return ""
    joined = " ".join(cand_lines).strip()
    if len(joined) < 18 or len(joined.split()) < 3:
        return ""
    if _AUTOSEC_GENERIC_CAPS_RE.search(joined):
        return ""
    if any(ch in joined for ch in ("$", "€")):
        return ""
    return _clean_section_label(joined)


def _extract_tail_caps_title(lines: list[str]) -> str:
    """Detecta un título de rubro desbordado en las últimas 1-3 líneas de la página."""
    tail_src = lines[-4:] if len(lines) >= 4 else lines
    cand: list[str] = []
    for ln in reversed(tail_src):
        s = re.sub(r"\s+", " ", ln).strip()
        if not s:
            if cand:
                break
            continue
        if _AUTOSEC_TAIL_EXCL_RE.search(s):
            break
        letters = [c for c in s if c.isalpha()]
        if not letters:
            break
        up_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
        if up_ratio >= 0.78:
            cand.insert(0, s)
            if len(cand) >= 3:
                break
        else:
            break
    joined = " ".join(cand).strip()
    if len(joined) < 26 or len(joined.split()) < 4:
        return ""
    if _AUTOSEC_GENERIC_CAPS_RE.search(joined) or _AUTOSEC_TAIL_EXCL_RE.search(joined):
        return ""
    if not _AUTOSEC_TAIL_STAGE_RE.search(joined):
        return ""
    return _clean_section_label(joined)


def _extract_vendor_tokens(head: list[str], tail: list[str], text: str) -> set[str]:
    tokens: set[str] = set()
    edge_txt = "\n".join(head + tail)
    for m in _AUTOSEC_EMAIL_RE.finditer(edge_txt):
        em = m.group(0).lower()
        dom = em.split("@", 1)[-1]
        if dom in ("gmail.com", "hotmail.com", "yahoo.com", "outlook.com", "live.com"):
            tokens.add(f"email:{em}")
        else:
            tokens.add(f"dom:{dom}")
    for m in _AUTOSEC_URL_RE.finditer(edge_txt):
        dom = m.group(1).lower().lstrip("www.")
        if dom and not dom.endswith("gob.mx"):
            tokens.add(f"dom:{dom}")
    for m in _AUTOSEC_RFC_RE.finditer(edge_txt):
        tokens.add(f"rfc:{m.group(0).upper()}")
    for m in _AUTOSEC_PHONE_RE.finditer(edge_txt):
        digits = re.sub(r"\D", "", m.group(0))
        if digits.startswith("52") and len(digits) >= 12:
            digits = digits[2:]
        if len(digits) == 10:
            tokens.add(f"tel:{digits[-8:]}")
    for m in _AUTOSEC_VENDOR_KW_RE.finditer(edge_txt + "\n" + text[:500]):
        kw = re.sub(r"[^a-z]", "", _norm_ascii(m.group(0)))
        if kw:
            tokens.add(f"kw:{kw}")
    return tokens


def _classify_non_quote(head: list[str], text: str, has_money: bool, has_total: bool, has_opening: bool) -> Optional[tuple[str, str, str]]:
    head_txt = "\n".join(head)
    # 1. Resumen financiero / de cuentas
    if re.search(r"resumen\s+(?:financiero|de\s+cuentas)", head_txt, re.I):
        return ("Resumen financiero", "Resumen financiero (no duplicar)", "Resumen financiero")
    # 2. Croquis / plano
    if re.search(r"\bcroquis\b", head_txt, re.I) and not has_total:
        lbl = _clean_section_label(head[0]) if head else "Croquis / plano"
        return ("Anexo / Carta", "Croquis / plano de ubicación", lbl)
    # 3. Constancia sin montos
    if re.search(r"se\s+hace\s+constar\s+que\s+la\s+obra", text, re.I) and not has_money and not has_total:
        return ("Anexo / Carta", "Carta constancia", "Constancia de obra")
    # 4. Cartas manifiesto EFIDEPORTE (sin tabla cotizada)
    has_price_table = bool(_TBL_HEADER_RE.search(text) and has_money)
    if _AUTOSEC_MANIFEST_CTX_RE.search(head_txt + "\n" + text[-500:]) and not has_price_table and not (has_total and has_money):
        for pat, lbl in _AUTOSEC_MANIFEST_TOPICS:
            if pat.search(head_txt + "\n" + text[:600]):
                return ("Anexo / Carta", "Carta manifiesto EFIDEPORTE", lbl)
    # 5. Oficio gubernamental / municipal sin montos ni apertura de cotización
    if re.search(r"ayuntamiento|\.gob\.mx|presidencia\s+municipal", text, re.I) and not has_opening and not has_money and not has_total:
        return ("Anexo / Carta", "Oficio oficial", "Oficio oficial")
    return None


def _build_page_fp(
    idx: int,
    text: str,
    w: int,
    h: int,
    size_key: tuple[int, int],
    is_empty_nat: bool,
    is_inv_ocr: bool,
    native_len: int,
    has_big_img: bool,
) -> dict:
    raw_lines = [l.strip() for l in text.splitlines() if l.strip()]
    lines = _strip_ocr_noise(raw_lines)
    head = lines[:6]
    tail = lines[-5:] if len(lines) > 5 else lines
    n_txt = len(text)
    upper_half = text[: max(1, (n_txt * 3) // 5)]
    lower_half = text[max(0, (n_txt * 3) // 5) :]

    has_money = bool(_MONEY_RE.search(_norm_ocr_currency_chars(text)))
    has_total = bool(_AUTOSEC_TOTAL_RE.search(text))
    has_closing = bool(_AUTOSEC_CLOSING_RE.search(lower_half or text))
    has_opening = bool(_AUTOSEC_OPENING_RE.search(upper_half or text))

    caption_title = ""
    if lines and _AUTOSEC_CAPTION_RE.match(lines[0]):
        cap_parts = [lines[0]]
        for nxt in lines[1:3]:
            if len(cap_parts) >= 3:
                break
            if re.search(r"temporada|altitud|ruta|monta[ñn]a|202\d|\)|etapa|cdmx|europa", nxt, re.I) or len(cap_parts[-1]) > 55:
                cap_parts.append(nxt)
            else:
                break
        caption_title = _clean_section_label(" ".join(cap_parts))

    caps_title = _extract_caps_title(lines) if not caption_title else ""
    tail_caps_title = _extract_tail_caps_title(lines)
    non_quote_info = _classify_non_quote(head, text, has_money, has_total, has_opening)
    is_support_doc = (
        bool(_AUTOSEC_SUPPORT_RE.search(text))
        and not has_opening
        and not has_money
        and not caption_title
        and non_quote_info is None
    )
    vendor_tokens = _extract_vendor_tokens(head, tail, text)

    return {
        "idx": idx,
        "text": text,
        "lines": lines,
        "head": head,
        "tail": tail,
        "w": w,
        "h": h,
        "size_key": size_key,
        "is_empty_nat": is_empty_nat,
        "is_inv_ocr": is_inv_ocr,
        "native_len": native_len,
        "has_big_img": has_big_img,
        "has_money": has_money,
        "has_total": has_total,
        "has_closing": has_closing,
        "has_opening": has_opening,
        "caption_title": caption_title,
        "caps_title": caps_title,
        "tail_caps_title": tail_caps_title,
        "is_support_doc": is_support_doc,
        "non_quote_info": non_quote_info,
        "vendor_tokens": vendor_tokens,
    }


def _extract_page_fingerprints(
    pdf_bytes: bytes,
    on_progress: Optional[callable] = None,
) -> list[dict]:
    """Extrae huellas rápidas de todas las páginas; ejecuta OCR ligero solo en páginas sin texto."""
    from PIL import Image
    import pytesseract

    if os.path.exists(TESSERACT_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    fps_raw: list[dict] = []
    ocr_jobs: dict[int, Image.Image] = {}

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        n = len(doc)
        if on_progress:
            on_progress(0, max(n, 1), "Analizando estructura del PDF...")
        for idx in range(n):
            pg = doc[idx]
            r = pg.rect
            w, h = int(round(r.width)), int(round(r.height))
            size_key = (int(round(w / 10.0) * 10), int(round(h / 10.0) * 10))
            raw_text = (pg.get_text("text") or "").strip()
            native_len = len(raw_text)
            is_empty_nat = native_len < 25
            is_inv_ocr = False
            has_big_img = False
            if 25 <= native_len < 350:
                p_area = max(r.width * r.height, 1.0)
                try:
                    has_big_img = any(
                        (ir.width * ir.height) / p_area >= 0.70
                        for im in pg.get_images(full=True)
                        for ir in pg.get_image_rects(im[0])
                    )
                except Exception:
                    pass
            if is_empty_nat:
                try:
                    pix = pg.get_pixmap(
                        matrix=fitz.Matrix(1.5, 1.5),
                        colorspace=fitz.csRGB,
                        alpha=False,
                    )
                    pil_im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    ocr_jobs[idx] = pil_im
                except Exception:
                    pass
            fps_raw.append({
                "idx": idx,
                "raw_text": raw_text,
                "w": w,
                "h": h,
                "size_key": size_key,
                "is_empty_nat": is_empty_nat,
                "is_inv_ocr": is_inv_ocr,
                "native_len": native_len,
                "has_big_img": has_big_img,
            })

    if ocr_jobs:
        lang_use = _resolve_tess_lang("spa")
        total_ocr = len(ocr_jobs)

        def _run_one(item: tuple[int, Image.Image]) -> tuple[int, str]:
            idx_p, im = item
            try:
                t = pytesseract.image_to_string(im, lang=lang_use, config="--psm 6")
                return idx_p, (t or "").strip()
            except Exception:
                return idx_p, ""

        done = 0
        workers = min(4, max(1, os.cpu_count() or 2), total_ocr)
        with _AutoThreadPool(max_workers=workers) as pool:
            futs = [pool.submit(_run_one, it) for it in ocr_jobs.items()]
            for fut in _auto_as_completed(futs):
                idx_p, ocr_txt = fut.result()
                if ocr_txt:
                    fps_raw[idx_p]["raw_text"] = ocr_txt
                done += 1
                if on_progress:
                    on_progress(done, total_ocr, f"OCR rápido pág. {idx_p + 1}/{len(fps_raw)}…")

    if on_progress:
        on_progress(len(fps_raw), max(len(fps_raw), 1), "Clasificando secciones…")

    return [
        _build_page_fp(
            item["idx"],
            item["raw_text"],
            item["w"],
            item["h"],
            item["size_key"],
            item["is_empty_nat"],
            item["is_inv_ocr"],
            item["native_len"],
            item["has_big_img"],
        )
        for item in fps_raw
    ]


def _detect_global_budget_end(fps: list[dict], filename: str = "") -> int:
    """Si la pág. 0 inicia un Presupuesto Global, devuelve el índice final (0-based); -1 si no."""
    if not fps:
        return -1
    fp0 = fps[0]
    t0 = fp0["text"]
    is_budget_cover = bool(
        re.search(
            r"press?upuesto\s+(?:desglosado|global|general|del\s+proyecto|de\s+inversi[oó]n)"
            r"|sistema\s+integral\s+de\s+detecci[oó]n|anexos?\s*\n\s*press?upuesto",
            t0,
            re.IGNORECASE,
        )
        or (
            re.search(r"\bconcepto\b", t0, re.I)
            and re.search(r"\bimporte\b|\bsin\s*iva\b|\bcon\s*iva\b|\bunidad\b", t0, re.I)
        )
        or (
            "presupuesto" in (filename or "").lower()
            and not fp0["has_opening"]
            and not fp0["caption_title"]
        )
    )
    if not is_budget_cover:
        return -1

    end_idx = 0
    g_w, g_h = fp0["w"], fp0["h"]
    for k in range(1, len(fps)):
        cur = fps[k]
        if cur["caption_title"]:
            break
        if cur["non_quote_info"] is not None:
            break
        if not _same_page_geometry((cur["w"], cur["h"]), (g_w, g_h), ignore_rotation=True):
            break
        if cur["has_opening"] and not re.search(r"press?upuesto\s+desglosado", cur["text"][:400], re.I):
            break
        if cur["caps_title"] and (
            cur["has_opening"]
            or bool(_AUTOSEC_EMAIL_RE.search(cur["text"]))
            or bool(_AUTOSEC_RFC_RE.search(cur["text"]))
            or bool(re.search(r"\b\d{1,2}\s*/\s*100\s*m\.?\s*n\.?", cur["text"], re.I))
        ):
            break
        end_idx = k
    return end_idx


def _strip_caption_prefix(caption: str) -> str:
    s = re.sub(r"^cotizaci[oó]n\s+correspondiente\s+(?:a\s+)?", "", caption, flags=re.I).strip()
    if s:
        s = s[0].upper() + s[1:]
    return _clean_section_label(s or caption)


def _extract_provider_hint(fp: dict) -> str:
    for ln in fp["head"][:5]:
        if re.search(r"s\.?\s*a\.?\s*de\s*c\.?\s*v|s\.?\s*r\.?\s*l|s\.?\s*c\.?\b", ln, re.I) or _AUTOSEC_VENDOR_KW_RE.search(ln):
            return _clean_section_label(ln, 35)
    for ln in fp["head"][:4]:
        if re.search(r"cotizaci[oó]n|preventivo|quotation|documento\s*nr|^(?:via\b|calle\b|av\.|blvd\.|col\.)", ln, re.I):
            continue
        if len(ln) >= 4:
            return _clean_section_label(ln, 35)
    return ""


def _build_section_label(
    fps_slice: list[dict],
    p0: int,
    p1: int,
    inherited_title: str = "",
    inherited_caption: str = "",
    fallback_title: str = "",
) -> str:
    fp0 = fps_slice[0]
    if inherited_title:
        return _clean_section_label(inherited_title)
    if inherited_caption:
        base_cap = _strip_caption_prefix(inherited_caption)
        prov = _extract_provider_hint(fp0)
        if prov and prov.lower() not in base_cap.lower():
            return _clean_section_label(f"{base_cap[:60]} ({prov})")
        return _clean_section_label(f"{base_cap} (adicional)")
    if fp0["caption_title"]:
        return _strip_caption_prefix(fp0["caption_title"])
    if fp0["caps_title"]:
        return _clean_section_label(fp0["caps_title"])
    if fp0["non_quote_info"] is not None:
        return _clean_section_label(fp0["non_quote_info"][2])
    txt0 = fp0["text"]
    if re.search(r"propuesta\s+del\s+costo\s+de\s+la\s+fianza", txt0, re.I):
        if "sofimex" in txt0.lower():
            return "Cotización Fianza (Sofimex)"
        return "Cotización Fianza"
    if re.search(r"monto\s+afianzado", txt0, re.I) and "fianza" in txt0.lower():
        return "Cotización Fianza de Obra"
    if re.search(r"\bDRO[-\s]*\d+", txt0):
        return "Presupuesto Servicios DRO"
    for ln in fp0["head"]:
        compact = re.sub(r"\s+", "", _norm_ascii(ln))
        if compact == "presente":
            continue
        if "@" in ln or "www." in ln.lower():
            continue
        if re.match(
            r"^(?:ciudad|guadalajara|tijuana|fecha|a\s+qui[eé]n\s+corresponda|presente\b"
            r"|blvd\.|poniente|calle|av\.|presupuesto\s*:|inmueble\s*:|ubicaci[oó]n\s*:"
            r"|exts?\.|tel[eé]?f?|cel\.|correo|e-?mail|cunda)",
            ln,
            re.I,
        ):
            continue
        if len(ln) >= 5:
            return _clean_section_label(ln)
    if fallback_title:
        return _clean_section_label(fallback_title)
    return f"Sección p.{p0}" if p0 == p1 else f"Sección p.{p0}-{p1}"


def _segment_range_heuristic(
    fps: list[dict],
    start_idx: int,
    end_idx: int,
    fallback_title: str = "",
) -> list[dict]:
    """Segmenta un rango continuo de páginas [start_idx..end_idx] en secciones candidatas."""
    if start_idx > end_idx:
        return []

    groups: list[dict] = [{
        "s": start_idx,
        "e": start_idx,
        "inherited_title": "",
        "inherited_caption": "",
        "reason": "Inicio de bloque",
    }]
    last_caption_seen = fps[start_idx]["caption_title"]

    for i in range(start_idx + 1, end_idx + 1):
        prev = fps[i - 1]
        cur = fps[i]
        g = groups[-1]
        split_reason = ""
        inh_title = ""
        inh_cap = ""

        # Regla 1: Anexo soporte (cédula / SAT) pegado a la cotización previa
        if cur["is_support_doc"]:
            split_reason = ""
        # Regla 3: Encabezado explícito "Cotización correspondiente..." (antes de portada corta)
        elif cur["caption_title"]:
            split_reason = "Título de cotización"
            last_caption_seen = cur["caption_title"]
        # Regla 2: Portada corta sin montos ni imagen de página completa + cuerpo en misma geometría
        elif (
            len(prev["text"]) < 150
            and not prev["has_money"]
            and not prev["caption_title"]
            and not prev["has_big_img"]
            and g["s"] == i - 1
            and _same_page_geometry((prev["w"], prev["h"]), (cur["w"], cur["h"]), ignore_rotation=False)
            and cur["non_quote_info"] is None
            and prev["non_quote_info"] is None
        ):
            split_reason = ""
        # Regla 6: Transición con carta manifiesto / croquis / oficio (antes de caps_title)
        elif (
            (cur["non_quote_info"] is not None) != (prev["non_quote_info"] is not None)
            or (
                cur["non_quote_info"] is not None
                and prev["non_quote_info"] is not None
                and cur["non_quote_info"][2] != prev["non_quote_info"][2]
            )
        ):
            split_reason = cur["non_quote_info"][1] if cur["non_quote_info"] else "Fin de anexo / carta"
            last_caption_seen = ""
        # Regla 4: Encabezado de rubro en MAYÚSCULAS
        elif cur["caps_title"]:
            split_reason = "Título de rubro"
            last_caption_seen = ""
        # Regla 5: Título desbordado al pie de página previa
        elif prev["tail_caps_title"] and (
            cur["has_money"] or cur["has_total"] or cur["has_opening"] or bool(cur["vendor_tokens"])
        ):
            split_reason = "Título en pie previo"
            inh_title = prev["tail_caps_title"]
            last_caption_seen = ""
        # Regla 7: Cotización escaneada autocontenida tras título nativo breve o cotización completa
        elif cur["is_empty_nat"] and (cur["has_opening"] or cur["has_total"]) and (
            prev["caption_title"]
            or prev["has_total"]
            or (prev["has_big_img"] and prev["native_len"] < 350 and not cur["vendor_tokens"].intersection(prev["vendor_tokens"]))
        ):
            if cur["vendor_tokens"] and prev["vendor_tokens"] and cur["vendor_tokens"] == prev["vendor_tokens"] and not prev["has_closing"] and not cur["has_opening"]:
                split_reason = ""
            else:
                split_reason = "Cotización independiente (OCR)"
                if last_caption_seen and (prev["caption_title"] or prev["is_empty_nat"]):
                    inh_cap = last_caption_seen
        # Regla 8: Cambio de tamaño/papel (no rotación 90°)
        elif not _same_page_geometry((cur["w"], cur["h"]), (prev["w"], prev["h"]), ignore_rotation=True):
            split_reason = "Cambio de formato de página"
        # Regla 9: Cambio nativo <-> escaneado
        elif cur["is_empty_nat"] != prev["is_empty_nat"] and not (
            cur["vendor_tokens"] and prev["vendor_tokens"] and cur["vendor_tokens"].intersection(prev["vendor_tokens"])
        ):
            split_reason = "Cambio de documento (nativo/escaneado)"
        # Regla 10: Nueva apertura de cotización tras cierre, total o monto de otro proveedor
        elif cur["has_opening"] and (
            prev["has_closing"]
            or prev["has_total"]
            or (prev["has_money"] and not cur["vendor_tokens"].intersection(prev["vendor_tokens"]))
        ):
            split_reason = "Nueva cotización detectada"
            if last_caption_seen and prev["is_empty_nat"] and cur["is_empty_nat"]:
                inh_cap = last_caption_seen
        # Regla 11: Cambio claro de proveedor (sin intersección de tokens de identidad)
        elif (
            cur["vendor_tokens"]
            and prev["vendor_tokens"]
            and not cur["vendor_tokens"].intersection(prev["vendor_tokens"])
            and not (cur["head"] and re.match(r"^(?:condiciones\b|[•\-*]|\d+\.)", cur["head"][0], re.I))
        ):
            split_reason = "Cambio de proveedor"

        if split_reason:
            groups.append({
                "s": i,
                "e": i,
                "inherited_title": inh_title,
                "inherited_caption": inh_cap,
                "reason": split_reason,
            })
        else:
            g["e"] = i

    out: list[dict] = []
    for g in groups:
        s_i, e_i = g["s"], g["e"]
        p0, p1 = s_i + 1, e_i + 1
        fp0 = fps[s_i]
        nq = fp0["non_quote_info"]
        if nq is not None and not any(fps[k]["caption_title"] for k in range(s_i, e_i + 1)):
            category, reason, label_nq = nq
            tipo = "Cotización Proveedor"
            selected = False
            label = label_nq
        else:
            category = "Cotización"
            tipo = "Cotización Proveedor"
            selected = True
            reason = g["reason"]
            if reason == "Inicio de bloque":
                if fp0["caption_title"] or fp0["caps_title"]:
                    reason = "Título de cotización"
                elif fp0["has_opening"]:
                    reason = "Encabezado de cotización"
                else:
                    reason = "Cotización detectada"
            label = _build_section_label(
                fps[s_i : e_i + 1],
                p0,
                p1,
                inherited_title=g["inherited_title"],
                inherited_caption=g["inherited_caption"],
                fallback_title=fallback_title,
            )
        out.append({
            "label": label,
            "tipo": tipo,
            "p0": p0,
            "p1": p1,
            "selected": selected,
            "category": category,
            "reason": reason,
            "pdf_idx": 0,
        })
    return out


def _clean_toc_label(title: str) -> str:
    s = re.sub(r"\(\s*p\.?\s*\d+(?:\s*[-–]\s*\d+)?\s*\)\s*$", "", title or "", flags=re.I).strip()
    s = re.sub(r"\.pdf\s*$", "", s, flags=re.I).strip()
    s = re.sub(r"^(?:\d{1,3}[.\-_ ]+|[a-z]{1,2}\.)", "", s, flags=re.I).strip()
    s = re.sub(r"\s+", " ", s).strip()
    return _clean_section_label(s or title)


def _classify_toc_entry(clean_label: str, raw_title: str, fps_slice: list[dict]) -> tuple[str, str, bool, str]:
    """Devuelve (tipo, category, selected, reason) para una entrada de marcador PDF."""
    low = _norm_ascii(f"{clean_label} {raw_title}")
    if re.search(r"resumen\s+(?:financiero|de\s+cuentas)", low):
        return ("Cotización Proveedor", "Resumen financiero", False, "Resumen financiero (marcador)")
    if re.search(r"press?upuesto\s+(?:desglosado|global|general)", low):
        return ("Presupuesto Global", "Presupuesto Global", True, "Presupuesto desglosado (marcador)")
    if re.search(
        r"manifiest[oa]\w*\s+de\s+inter[eé]s|gastos\s+administrativos|propuesta\s+del?\s+personal"
        r"|mantenimiento\s+(?:al|del|posterior)|gastos\s+de\s+capacitaci[oó]n"
        r"|adquisici[oó]n\s+de\s+bienes|recurs\w*\s+de\s+ganancias|rubros\s+no\s+permitidos"
        r"|croquis|constancia",
        low,
    ) or re.match(r"^\s*(?:29|34|35|36|37|38|39|40)\b", raw_title.strip()):
        return ("Cotización Proveedor", "Anexo / Carta", False, "Carta manifiesto EFIDEPORTE (marcador)")
    if fps_slice and fps_slice[0]["non_quote_info"] is not None and not re.search(r"cotiz|seguro|fianza|dro|contador", low):
        nq = fps_slice[0]["non_quote_info"]
        return ("Cotización Proveedor", nq[0], False, nq[1])
    return ("Cotización Proveedor", "Cotización", True, "Marcador PDF")


def _toc_to_leaf_entries(raw_toc: list[list], n_pages: int) -> list[tuple[str, int]]:
    """Convierte get_toc() jerárquico en entradas hoja ordenadas (raw_title, page_1based)."""
    valid: list[tuple[int, str, int]] = []
    for item in raw_toc:
        if not item or len(item) < 3:
            continue
        lvl = int(item[0])
        title = str(item[1] or "").strip()
        pg = int(item[2])
        if not (1 <= pg <= n_pages):
            continue
        if not title or re.match(r"^(?:sin\s*t[ií]tulo|untitled|bookmark)$", title, re.I):
            continue
        valid.append((lvl, title, pg))
    if not valid:
        return []

    # Construir nodos para decidir si un padre se expande en sus hijos
    nodes: list[dict] = []
    stack: list[dict] = []
    for lvl, title, pg in valid:
        node = {"lvl": lvl, "title": title, "pg": pg, "children": []}
        while stack and stack[-1]["lvl"] >= lvl:
            stack.pop()
        if stack:
            stack[-1]["children"].append(node)
        else:
            nodes.append(node)
        stack.append(node)

    leaves: list[tuple[str, int]] = []

    def _collect(node: dict):
        ch = node["children"]
        distinct_child_pages = {c["pg"] for c in ch}
        if ch and len(distinct_child_pages) >= 2:
            for c in ch:
                _collect(c)
        else:
            leaves.append((node["title"], node["pg"]))

    for nd in nodes:
        _collect(nd)

    # Desduplicar por página inicial conservando el título más específico
    dedup: dict[int, str] = {}
    for title, pg in leaves:
        dedup[pg] = title
    return sorted(((t, pg) for pg, t in dedup.items()), key=lambda x: x[1])


def auto_detect_sections(
    pdf_bytes: bytes,
    filename: str = "",
    on_progress: Optional[callable] = None,
) -> list[dict]:
    """Analiza un PDF único y devuelve secciones candidatas para revisión del usuario."""
    if not pdf_bytes:
        return []

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        n_pages = len(doc)
        raw_toc = doc.get_toc(simple=True) or []

    if n_pages == 0:
        return []

    fps = _extract_page_fingerprints(pdf_bytes, on_progress=on_progress)

    # Caso base: PDF de 1 página
    if n_pages == 1:
        fp0 = fps[0]
        base_name = re.sub(r"\.pdf$", "", os.path.basename(filename or ""), flags=re.I).strip()
        is_presup = (
            "presupuesto" in (filename or "").lower()
            or bool(re.search(r"press?upuesto\s+(?:desglosado|global|general)|costo\s+total\s+del?\s+(?:proyecto|p\w+)", fp0["text"], re.I))
        ) and not fp0["caption_title"] and not re.search(r"presupuesto\s+de\s+servicios\s+profesionales", fp0["text"], re.I)
        if is_presup:
            return [{
                "label": _clean_section_label(base_name or "Presupuesto Global"),
                "tipo": "Presupuesto Global",
                "p0": 1,
                "p1": 1,
                "selected": True,
                "category": "Presupuesto Global",
                "reason": "Presupuesto desglosado (1 pág.)",
                "pdf_idx": 0,
            }]
        lbl = _build_section_label(fps, 1, 1, fallback_title=base_name)
        return [{
            "label": lbl,
            "tipo": "Cotización Proveedor",
            "p0": 1,
            "p1": 1,
            "selected": True,
            "category": "Cotización",
            "reason": "Documento único (1 pág.)",
            "pdf_idx": 0,
        }]

    # Capa 1: Marcadores PDF (TOC) con refinamiento anti-huecos
    leaf_toc = _toc_to_leaf_entries(raw_toc, n_pages)
    if len(leaf_toc) >= 2:
        results: list[dict] = []
        if leaf_toc[0][1] > 1:
            results.extend(_segment_range_heuristic(fps, 0, leaf_toc[0][1] - 2))

        for k, (raw_title, start_p) in enumerate(leaf_toc):
            end_p = (leaf_toc[k + 1][1] - 1) if (k + 1 < len(leaf_toc)) else n_pages
            if end_p < start_p:
                continue
            s_i, e_i = start_p - 1, end_p - 1
            fps_slice = fps[s_i : e_i + 1]
            clean_lbl = _clean_toc_label(raw_title)
            tipo, category, selected, reason = _classify_toc_entry(clean_lbl, raw_title, fps_slice)

            # Detectar bloques TOC heterogéneos que ocultan cotizaciones insertadas (ej. 003 p.12-26)
            span_len = end_p - start_p + 1
            m_rng = _AUTOSEC_TOC_RANGE_RE.search(raw_title)
            rng_mismatch = False
            if m_rng:
                claimed_len = int(m_rng.group(2)) - int(m_rng.group(1)) + 1
                if claimed_len > 0 and claimed_len < span_len - 1:
                    rng_mismatch = True

            has_nat_scan_mix = (
                any(fp["is_empty_nat"] for fp in fps_slice)
                and any(not fp["is_empty_nat"] for fp in fps_slice)
            )
            has_geom_change = any(
                not _same_page_geometry((fp["w"], fp["h"]), (fps_slice[0]["w"], fps_slice[0]["h"]), ignore_rotation=True)
                for fp in fps_slice[1:]
            )

            if category == "Cotización" and span_len >= 4 and (rng_mismatch or has_nat_scan_mix or has_geom_change):
                sub_secs = _segment_range_heuristic(fps, s_i, e_i, fallback_title=clean_lbl)
                # Conservar el nombre del marcador en el primer bloque si corresponde al mismo rubro principal
                if sub_secs and clean_lbl and not fps[s_i]["caps_title"] and not fps[s_i]["caption_title"]:
                    if sub_secs[0]["p1"] - sub_secs[0]["p0"] >= 1 or sub_secs[0]["label"].startswith("Sección"):
                        sub_secs[0]["label"] = clean_lbl
                results.extend(sub_secs)
            else:
                results.append({
                    "label": clean_lbl or f"Sección p.{start_p}-{end_p}",
                    "tipo": tipo,
                    "p0": start_p,
                    "p1": end_p,
                    "selected": selected,
                    "category": category,
                    "reason": reason,
                    "pdf_idx": 0,
                })
        return results

    # Capa 2: Segmentación heurística sin marcadores
    out: list[dict] = []
    gb_end = _detect_global_budget_end(fps, filename=filename)
    start_after = 0
    if gb_end >= 0:
        out.append({
            "label": "Presupuesto Global",
            "tipo": "Presupuesto Global",
            "p0": 1,
            "p1": gb_end + 1,
            "selected": True,
            "category": "Presupuesto Global",
            "reason": "Presupuesto desglosado",
            "pdf_idx": 0,
        })
        start_after = gb_end + 1

    if start_after < n_pages:
        out.extend(_segment_range_heuristic(fps, start_after, n_pages - 1))
    return out


# ─────────────────────────────────────────────────────────────
# EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────────────────────
def _side(style: str = "thin") -> Side:
    return Side(style=style, color="000000")

def _border(style: str = "thin") -> Border:
    s = _side(style)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(rgb: str) -> PatternFill:
    return PatternFill("solid", start_color=rgb, end_color=rgb)


def to_excel(df: pd.DataFrame, nombre: str = "", blank: bool = False) -> bytes:
    """
    Genera Excel con formato profesional alineado a las 16 columnas de _COLS.
    blank=True produce plantilla vacía con fórmulas vivas.
    """
    buf = io.BytesIO()
    wb  = openpyxl.Workbook()
    ws  = wb.active
    clean_title = re.sub(r'[\\/*?:\[\]]+', ' ', (nombre or '').strip()).strip().strip("'")
    ws.title = clean_title[:31] if clean_title else "Conciliación"

    F_WHITE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    F_HDR   = Font(name="Calibri", size=11, bold=True, color="000000")
    F_DATA  = Font(name="Calibri", size=11, color="000000")
    F_TOT   = Font(name="Calibri", size=11, bold=True, color="000000")
    F_BOLD  = Font(name="Calibri", size=11, bold=True, color="000000")

    FILL_ROW1  = _fill("6E152E")
    FILL_HDR_A = _fill("D4C19C")
    FILL_HDR_B = _fill("EBE2D1")
    BD     = _border("thin")
    BD_MED = _border("medium")
    AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    AL_R = Alignment(horizontal="right",  vertical="center")

    # Configuraciones de impresión – Carta Horizontal
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.paperSize        = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation      = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage        = True
    ws.page_setup.fitToWidth       = 1   # encajar todas las columnas en 1 página de ancho
    ws.page_setup.fitToHeight      = 0   # sin límite de páginas en alto
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75, header=0.3, footer=0.3
    )
    ws.print_options.horizontalCentered = True

    df_clean = _normalize_editor_df(df)
    n_rows = len(df_clean)

    for col_idx, width in enumerate(_WIDTHS_COL, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 27.75
    c = ws.cell(row=1, column=1, value=nombre or "")
    c.font = F_WHITE; c.fill = FILL_ROW1; c.border = BD; c.alignment = AL_L
    for ci in range(2, len(_COLS) + 1):
        cx = ws.cell(row=1, column=ci)
        cx.fill = FILL_ROW1; cx.border = BD

    ws.row_dimensions[2].height = 21.75
    for ci, hdr in enumerate(_COLS, 1):
        c = ws.cell(row=2, column=ci, value=hdr)
        c.font = F_HDR
        c.fill = FILL_HDR_A if ci <= 13 else FILL_HDR_B
        c.border = BD; c.alignment = AL_C

    DS = 3

    def _money_cell(ws_, r, ci, val=None):
        c_ = ws_.cell(row=r, column=ci, value=val)
        c_.number_format = _FMT_MONEY; c_.font = F_DATA
        c_.border = BD; c_.alignment = AL_R
        return c_

    for i, row_dict in enumerate(df_clean.to_dict('records'), DS):
        r = i
        ws.row_dimensions[r].height = 16.5
        blank_row = blank

        # 1 (A): Nº Sec
        n_sec = _safe_f(row_dict.get("Nº Sec"))
        c = ws.cell(row=r, column=1, value=int(n_sec) if n_sec is not None else (r - DS + 1))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 2 (B): Sección
        c = ws.cell(row=r, column=2, value=str(row_dict.get("Sección", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_L

        # 3 (C): Tipo
        c = ws.cell(row=r, column=3, value=str(row_dict.get("Tipo", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 4 (D): Fecha
        c = ws.cell(row=r, column=4, value=str(row_dict.get("Fecha", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 5 (E): Rubro
        c = ws.cell(row=r, column=5, value=str(row_dict.get("Rubro", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_L

        # 6 (F): QT
        c = ws.cell(row=r, column=6, value=str(row_dict.get("QT", "Sí") or "Sí"))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 7 (G): T. Cambio
        c = ws.cell(row=r, column=7, value=str(row_dict.get("T. Cambio", "MXN") or "MXN"))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 8 (H): (+ IVA)
        c = ws.cell(row=r, column=8, value=str(row_dict.get("(+ IVA)", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 9 (I): Cantidad (soporta decimales sin truncar)
        q = _safe_f(row_dict.get("Cantidad"))
        q_val = int(q) if (q is not None and float(q).is_integer()) else (q if q is not None else 1)
        c = ws.cell(row=r, column=9, value=q_val)
        c.number_format = "0.##"; c.font = F_DATA; c.border = BD; c.alignment = AL_C

        # 10 (J): Precio Unitario
        pu_val = None if blank_row else _safe_f(row_dict.get("Precio Unitario"))
        _money_cell(ws, r, 10, pu_val)

        val_iva = str(row_dict.get("(+ IVA)", "N/M")).strip().lower().replace("\u00ed", "i").replace("\u00c3\u00ad", "i")
        has_iva = val_iva in ["si", "incluido"]

        # 11 (K): Subtotal (Sin IVA) = I * J
        c = ws.cell(row=r, column=11, value=f'=IF(ISNUMBER(J{r}),ROUND(I{r}*J{r},2),"")')
        c.number_format = _FMT_MONEY; c.font = F_DATA; c.border = BD; c.alignment = AL_R

        # 12 (L): IVA 16% = K * 0.16
        iva_formula = f'=IF(ISNUMBER(K{r}),ROUND(K{r}*0.16,2),"")' if has_iva else None
        c = ws.cell(row=r, column=12, value=iva_formula)
        c.number_format = _FMT_MONEY; c.font = F_DATA; c.border = BD; c.alignment = AL_R

        # 13 (M): Total con IVA = K + L
        if has_iva:
            tot_formula = f'=IF(ISNUMBER(K{r}),IF(ISNUMBER(L{r}),K{r}+L{r},K{r}),"")'
        else:
            tot_formula = f'=IF(ISNUMBER(K{r}),K{r},"")'
        c = ws.cell(row=r, column=13, value=tot_formula)
        c.number_format = _FMT_MONEY; c.font = F_DATA; c.border = BD; c.alignment = AL_R

        # 14 (N): Diferencia final = M - O
        c = ws.cell(row=r, column=14, value=f'=IF(AND(ISNUMBER(M{r}),ISNUMBER(O{r})),M{r}-O{r},"")')
        c.number_format = _FMT_MONEY; c.font = F_DATA; c.border = BD; c.alignment = AL_R

        # 15 (O): Monto en Anexo Escrito
        anx_val = None if blank_row else _safe_f(row_dict.get("Monto en Anexo Escrito"))
        c = _money_cell(ws, r, 15, anx_val); c.font = F_BOLD

        # 16 (P): Observaciones
        c = ws.cell(row=r, column=16, value="" if blank_row else str(row_dict.get("Observaciones", "") or ""))
        c.font = F_DATA; c.border = BD; c.alignment = AL_L

    if n_rows > 0:
        tot_row = DS + n_rows
        s_xl, e_xl = DS, DS + n_rows - 1
        ws.row_dimensions[tot_row].height = 18
        c = ws.cell(row=tot_row, column=1, value="TOTALES")
        c.font = F_TOT; c.border = BD_MED
        for ci in range(2, len(_COLS) + 1):
            cx = ws.cell(row=tot_row, column=ci)
            cx.border = BD_MED
        # SUMIF excluyendo Presupuesto Global para Subtotal (11/K), IVA (12/L) y Total (13/M)
        for ci in (11, 12, 13):
            col_l = get_column_letter(ci)
            formula = f'=SUMIF(C{s_xl}:C{e_xl},"<>Presupuesto Global",{col_l}{s_xl}:{col_l}{e_xl})'
            c = ws.cell(row=tot_row, column=ci, value=formula)
            c.number_format = _FMT_MONEY; c.font = F_TOT
            c.fill = _fill("EBE2D1"); c.border = BD_MED; c.alignment = AL_R
        # Suma de Monto en Anexo Escrito (15/O)
        c = ws.cell(row=tot_row, column=15, value=f"=SUM(O{s_xl}:O{e_xl})")
        c.number_format = _FMT_MONEY; c.font = F_TOT
        c.fill = _fill("EBE2D1"); c.border = BD_MED; c.alignment = AL_R

    ws.freeze_panes = "A3"

    # Definir área de impresión explícita
    last_col = get_column_letter(len(_COLS))
    last_row = DS + n_rows if n_rows > 0 else DS
    ws.print_area = f"A1:{last_col}{last_row}"

    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ============================================================
# EXTRACTOR DE COTIZACIONES PDF | app.py  v2.0
# ============================================================

import hashlib, io, re, datetime, json
from typing import Optional

import fitz
import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st
import xlsxwriter
import google.generativeai as genai


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
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────
# [OPT-1] Umbral: si texto nativo < este valor × n_páginas → activa OCR
NATIVE_MIN_CHARS_PER_PAGE = 80

# [BF-4] Modelos Gemini en orden de preferencia (sin list_models())
_GEMINI_PREFERRED = [
    "gemini-2.0-flash",
    "gemini-1.5-flash-latest",
    "gemini-1.5-flash",
    "gemini-1.5-pro-latest",
    "gemini-1.5-pro",
    "gemini-pro",
]

# [BF-2] Regex de importes:
#   Grupo 1 (con $): acepta cualquier cantidad de dígitos → $12345.67 ✓
#   Grupo 2 (sin $): exige separador de miles → 1,234.56 ✓ | 1234 ✗
_MONEY_RE = re.compile(
    r"\$\s*([\d,]+(?:\.\d{1,2})?)"
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)"
)

_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:de)?\s*([a-záéíóúñA-ZÁÉÍÓÚÑ]+)[,\s]*(?:del?\s*)?(\d{4})|"
    r"(\d{4})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{1,2})|"
    r"(\d{1,2})[\s.\-/]+(\d{1,2})[\s.\-/]+(\d{2,4})"
)

_ITEM_RE = re.compile(
    r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)

_MESES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
}

_COLS = [
    "Fecha","Rubro","QT","T. Cambio","(+ IVA)","Cantidad",
    "Precio Unitario","Subtotal (Sin IVA)","IVA 16%","Total con IVA",
    "Diferencia final","Monto en Anexo Escrito","Observaciones",
]
_WIDTHS = [12,52,5,10,7,8,16,18,17,16,18,22,48]


# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
_SS_DEFAULTS: dict = {
    "pdf_bytes": None, "pdf_hash": None, "total_pages": 0,
    "current_page": 0, "num_sec": 1, "sec_cfg": [],
    "df": None, "extracted": False,
}
for _k, _v in _SS_DEFAULTS.items():
    st.session_state.setdefault(_k, _v)


# ─────────────────────────────────────────────────────────────
# MOTOR OCR  (singleton por worker, cargado una sola vez)
# ─────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _get_ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR(det_model_dir=None, rec_model_dir=None, cls_model_dir=None)
    except Exception as exc:
        st.warning(f"Motor OCR no disponible ({exc}). Solo se usará texto nativo.")
        return None


def _ocr_page(pdf_bytes: bytes, idx: int) -> str:
    """Extrae texto de una página vía OCR (fallback cuando nativo es escaso)."""
    ocr = _get_ocr_engine()
    if ocr is None:
        return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(
                matrix=fitz.Matrix(2.0, 2.0), colorspace=fitz.csRGB, alpha=False
            )
            img = np.frombuffer(pix.samples, np.uint8).reshape(
                pix.height, pix.width, 3
            )[:, :, ::-1]
            result, _ = ocr(img)
            return "\n".join(r[1] for r in result if r and len(r) > 1) if result else ""
    except Exception as exc:
        st.warning(f"OCR página {idx+1}: {exc}")
        return ""


# ─────────────────────────────────────────────────────────────
# HELPERS DE CONVERSIÓN
# ─────────────────────────────────────────────────────────────
def _safe_f(v) -> Optional[float]:
    """Convierte v a float; devuelve None para NaN, None o no numérico."""
    if v is None:
        return None
    try:
        # Eliminar comas de miles antes de convertir
        f = float(str(v).replace(",", "").strip())
        return None if (f != f) else f   # f != f es True solo para NaN
    except (ValueError, TypeError):
        return None


def _money(txt: str) -> Optional[float]:
    """Devuelve el primer importe monetario válido encontrado en txt."""
    for m in _MONEY_RE.finditer(str(txt)):
        raw = m.group(1) or m.group(2)
        if raw:
            try:
                val = float(raw.replace(",", ""))
                if val > 0:
                    return val
            except ValueError:
                continue
    return None


def _date(text: str) -> Optional[datetime.date]:
    """Extrae la primera fecha válida de un texto."""
    for m in _DATE_RE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]:
                mes_str = g[1].lower()
                mes_num = _MESES.get(mes_str)
                if not mes_num:
                    # Coincidencia parcial (p.ej. "sep" → "septiembre")
                    for k, v in _MESES.items():
                        if k.startswith(mes_str[:4]) or mes_str.startswith(k[:4]):
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


def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


# [OPT-2] Clave de caché = hash (str) + idx, no los bytes completos
@st.cache_data(max_entries=80, show_spinner=False)
def _render(pdf_hash: str, pdf_bytes: bytes, idx: int) -> bytes:
    """Renderiza la página idx del PDF a PNG para el visor."""
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(
            matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False
        )
        return pix.tobytes("png")


def _parse_space_table(text: str) -> list[dict]:
    """Parsea tablas de cotización con alineación por espacios."""
    items = []
    for line in text.splitlines():
        m = _ITEM_RE.match(line.strip())
        if m:
            qty, desc, pu, total = m.groups()
            items.append({
                "qty": int(qty),
                "desc": desc.strip(),
                "pu": float(pu.replace(",", "")),
                "total": float(total.replace(",", "")),
            })
    return items


# ─────────────────────────────────────────────────────────────
# [BF-1] RECALCULAR CAMPOS DERIVADOS
# ─────────────────────────────────────────────────────────────
def recalc_derived(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recalcula 'Diferencia final' = Total con IVA − Monto en Anexo Escrito.
    Se llama después de cada cambio en el data_editor.
    """
    df = df.copy()
    tot  = pd.to_numeric(df["Total con IVA"],          errors="coerce")
    anex = pd.to_numeric(df["Monto en Anexo Escrito"], errors="coerce")
    mask = tot.notna() | anex.notna()
    df.loc[mask,  "Diferencia final"] = (tot.fillna(0) - anex.fillna(0))[mask]
    df.loc[~mask, "Diferencia final"] = None
    return df


# ─────────────────────────────────────────────────────────────
# MOTOR DE EXTRACCIÓN HÍBRIDO
# ─────────────────────────────────────────────────────────────
def extract(
    pdf_bytes: bytes, label: str,
    p0: int, p1: int,
    det_iva: bool, calc_sub: bool,
) -> dict:
    """
    Extrae datos financieros de p0..p1 usando tres niveles:
      Nivel 1 — texto nativo (pdfplumber).
      Nivel 2 — tablas estructuradas.
      Nivel 3 — OCR (solo si el texto nativo es insuficiente).  [OPT-1]
    """
    native_text = ""
    table_rows: list[list[str]] = []
    ocr_used = False

    # ── Niveles 1 y 2: nativo ────────────────────────────────
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
        if "total" in j and not re.search(r"sub|parcial|acum", j):
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

    # ── Paso 2: reconstrucción por líneas de producto ─────────
    # [BF-6] qty debe ser entero exacto; tolerancia relativa 1%
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
                    # [BF-6] qty: entero exacto entre 1 y 9 999
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

    # ── Paso 3: escaneo de texto libre ───────────────────────
    for ln in text.splitlines():
        if fecha is None:
            d = _date(ln)
            if d:
                fecha = d

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b", ln, re.I) and \
               not re.search(r"sub|parcial|acum", ln, re.I):
                v = _money(ln)
                if v and (tot is None or v > tot):
                    tot = v

    # ── Fallback absoluto: importe más alto del documento ────
    # [BF-3] Solo si todo lo anterior falló; filtra valores < 10
    if tot is None:
        all_vals: list[float] = []
        for row in table_rows:
            for cell in row:
                v = _money(cell)
                if v and v >= 10.0:
                    all_vals.append(v)
        for ln in text.splitlines():
            for m in _MONEY_RE.finditer(ln):
                raw = m.group(1) or m.group(2)
                if raw:
                    try:
                        v = float(raw.replace(",", ""))
                        if v >= 10.0:
                            all_vals.append(v)
                    except ValueError:
                        pass
        if all_vals:
            tot = max(all_vals)
            obs_parts.append("Total inferido (máx.)")

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

    # ── Cantidad y precio unitario desde tablas ───────────────
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

    # ── [BF-5] Triangulación IVA / Subtotal / Total ───────────
    if tot and not sub and not iva and iva_f == "Sí":
        # Tot ya incluye IVA
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

    return {
        "Fecha":                  (fecha.isoformat() if fecha else datetime.date.today().isoformat()),
        "Rubro":                  label,
        "QT":                     "Sí",
        "T. Cambio":              "MXN",
        "(+ IVA)":                iva_f,
        "Cantidad":               qty,
        "Precio Unitario":        pu,
        "Subtotal (Sin IVA)":     sub,
        "IVA 16%":                iva,
        "Total con IVA":          tot,
        "Diferencia final":       None,   # se calcula con recalc_derived()
        "Monto en Anexo Escrito": None,
        "Observaciones":          " | ".join(obs_parts) if obs_parts else "",
    }


# ─────────────────────────────────────────────────────────────
# EXPORTACIÓN A EXCEL
# ─────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws  = wb.add_worksheet("Conciliación")

    MF  = '_-"$ "* #,##0.00_-;\\-"$ "* #,##0.00_-;_-"$ "* "-"??_-;_-@_-'
    B   = {"font_name": "Calibri", "font_size": 11, "border": 1}
    hf  = wb.add_format({**B, "bold": True, "bg_color": "#D4C19C", "align": "center", "valign": "vcenter"})
    h2f = wb.add_format({**B, "bold": True, "bg_color": "#EBE2D1", "align": "center", "valign": "vcenter"})
    df_ = wb.add_format({**B, "num_format": "dd/mm/yy"})
    mf  = wb.add_format({**B, "num_format": MF})
    tf  = wb.add_format({**B})
    nf  = wb.add_format({**B, "num_format": "#,##0"})
    totf = wb.add_format({**B, "bold": True, "num_format": MF})
    of  = wb.add_format({**B, "text_wrap": True})

    # Encabezados
    for c, h in enumerate(_COLS):
        ws.write(0, c, h, h2f if c >= 11 else hf)
    for i, w in enumerate(_WIDTHS):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)

    # Columnas con fórmulas (clave = nombre columna, valor = (col_idx_0based, fórmula))
    # {r} se sustituye por el número de fila Excel (1-based)
    MC = {
        "Precio Unitario":        (6,  None),           # G = entrada del usuario
        "Subtotal (Sin IVA)":     (7,  "=F{r}*G{r}"),  # H = Cantidad × PU
        "IVA 16%":                (8,  "=H{r}*0.16"),  # I = Subtotal × 16%
        "Total con IVA":          (9,  "=H{r}+I{r}"),  # J = Subtotal + IVA
        "Monto en Anexo Escrito": (11, None),            # L = entrada del usuario
    }

    # D = primera fila de datos (0-based xlsxwriter)
    # La fila Excel (1-based) equivalente es D + 1 = 2
    D      = 1
    n_rows = len(df)

    for i, (_, row) in enumerate(df.iterrows()):
        r0 = D + i          # índice 0-based para xlsxwriter
        r1 = r0 + 1         # número de fila Excel (1-based) = r0 + 1

        # Fecha
        fecha_val = row.get("Fecha")
        if isinstance(fecha_val, str):
            try:
                fecha_val = datetime.date.fromisoformat(fecha_val)
            except Exception:
                fecha_val = None
        if isinstance(fecha_val, (datetime.date, datetime.datetime)):
            dt = (datetime.datetime.combine(fecha_val, datetime.time())
                  if isinstance(fecha_val, datetime.date) else fecha_val)
            ws.write_datetime(r0, 0, dt, df_)
        else:
            ws.write(r0, 0, str(fecha_val or ""), tf)

        ws.write(r0, 1, str(row.get("Rubro", "") or ""), of)
        ws.write(r0, 2, str(row.get("QT", "Sí")), tf)
        ws.write(r0, 3, str(row.get("T. Cambio", "MXN")), tf)
        ws.write(r0, 4, str(row.get("(+ IVA)", "") or ""), tf)

        q = _safe_f(row.get("Cantidad"))
        ws.write(r0, 5, int(q) if q is not None else 1, nf)

        for cn, (xi, fm) in MC.items():
            v = _safe_f(row.get(cn))
            if v is not None:
                ws.write_number(r0, xi, v, mf)
            elif fm:
                ws.write_formula(r0, xi, fm.format(r=r1), mf)
            else:
                ws.write_blank(r0, xi, None, mf)

        # Col K (10) = Diferencia final = Total − Monto Anexo
        ws.write_formula(r0, 10, f"=J{r1}-L{r1}", mf)
        # Col M (12) = Observaciones
        ws.write(r0, 12, str(row.get("Observaciones", "") or ""), of)

    # [BF-7] Fila de totales
    # total_row (0-based) = D + n_rows  →  Excel row D + n_rows + 1
    # s_xl (Excel 1-based) = D + 1 = 2  →  primera fila de datos
    # e_xl (Excel 1-based) = D + n_rows →  última fila de datos
    # (D + n_rows como Excel row = última fila dato porque Excel rows son D+1..D+n_rows)
    total_row = D + n_rows      # 0-based: donde se escribe el total
    s_xl      = D + 1           # 1-based: primera fila de datos en fórmula SUM
    e_xl      = D + n_rows      # 1-based: última  fila de datos en fórmula SUM
    for ci in (9, 10, 11):
        cl = chr(ord("A") + ci)
        ws.write_formula(total_row, ci, f"=SUM({cl}{s_xl}:{cl}{e_xl})", totf)

    wb.close()
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# CALLBACKS DE NAVEGACIÓN  [UI-1]
# ─────────────────────────────────────────────────────────────
def _nav_prev():
    st.session_state.current_page = max(0, st.session_state.current_page - 1)

def _nav_next():
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
        h = _md5(raw)
        if h != st.session_state.pdf_hash:
            st.session_state.update(
                pdf_bytes=raw, pdf_hash=h,
                extracted=False, df=None, current_page=0,
            )
            with fitz.open(stream=raw, filetype="pdf") as d:
                st.session_state.total_pages = len(d)
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")

    st.markdown("---")
    st.markdown("### ⚙️ Secciones")
    n = int(st.number_input(
        "Número de secciones", min_value=1, max_value=50,
        value=st.session_state.num_sec, step=1,
    ))

    if n != st.session_state.num_sec:
        st.session_state.num_sec = n
        st.session_state.extracted = False
        st.session_state.df = None

    cfgs = st.session_state.sec_cfg
    # [UI-4] tp mínimo 1 para evitar error en number_input cuando no hay PDF
    tp = max(st.session_state.total_pages, 1)

    while len(cfgs) < n:
        i = len(cfgs) + 1
        cfgs.append({"label": f"Sección {i}", "p0": 1, "p1": 1,
                     "det_iva": True, "calc_sub": True})
    del cfgs[n:]

    for i, c in enumerate(cfgs):
        with st.expander(f"📄 Sección {i+1}", expanded=(n <= 5)):
            c["label"]    = st.text_input("Rubro/Concepto", value=c["label"], key=f"lb{i}")
            col_a, col_b  = st.columns(2)
            c["p0"]       = col_a.number_input("Pág. Inicio", 1, tp, min(c["p0"], tp), key=f"p0{i}")
            c["p1"]       = col_b.number_input(
                "Pág. Fin", c["p0"], tp,
                max(min(c["p1"], tp), c["p0"]),
                key=f"p1{i}",
            )
            c["det_iva"]  = st.checkbox("Detectar IVA", value=c["det_iva"],  key=f"iv{i}")
            c["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=c["calc_sub"], key=f"cs{i}")

    st.markdown("---")
    run = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_bytes is None),
        use_container_width=True, type="primary",
    )


# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
  <h1>📋 Conciliador de Cotizaciones PDF</h1>
  <p>Extrae, valida y exporta montos desde documentos PDF · OCR automático · Auditoría IA</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# EXTRACCIÓN
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
            )
            rows.append(row)
        except Exception as exc:
            st.warning(f"⚠️ Sección {i+1} «{c['label']}»: {str(exc)[:120]}")
            rows.append({
                **{k: None for k in _COLS},
                "Rubro": c["label"], "QT": "Sí", "T. Cambio": "MXN",
                "Cantidad": 1, "Fecha": datetime.date.today().isoformat(),
                "Observaciones": f"Error: {str(exc)[:100]}",
            })
        bar.progress((i + 1) / n_secs)

    bar.empty()
    df_new = pd.DataFrame(rows, columns=_COLS)
    df_new = recalc_derived(df_new)   # [BF-1]

    st.session_state.df        = df_new
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()


# ─────────────────────────────────────────────────────────────
# LAYOUT PRINCIPAL
# ─────────────────────────────────────────────────────────────
col_L, col_R = st.columns(2, gap="medium")


# ── VISOR DE DOCUMENTO ────────────────────────────────────────
with col_L:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)

    tp = st.session_state.total_pages
    cp = st.session_state.current_page

    # [UI-1] Callbacks garantizan que session_state esté actualizado
    #        antes del re-run, eliminando el desync entre botón e imagen.
    nav_p, nav_c, nav_n = st.columns([1, 4, 1])
    nav_p.button("◀", key="btn_prev", use_container_width=True, on_click=_nav_prev)
    nav_n.button("▶", key="btn_next", use_container_width=True, on_click=_nav_next)

    # El number_input lee cp actualizado tras el callback
    cp = st.session_state.current_page
    new_page = nav_c.number_input(
        "Página", min_value=1, max_value=tp,
        value=cp + 1, step=1,
        label_visibility="collapsed", key="nav_page_input",
    )
    if new_page - 1 != cp:
        st.session_state.current_page = new_page - 1
        cp = new_page - 1

    st.caption(f"Página {cp + 1} de {tp}")

    # Badge de sección activa
    for c in st.session_state.sec_cfg:
        if c["p0"] <= cp + 1 <= c["p1"]:
            st.markdown(
                f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;'
                f'border-radius:4px;font-size:.8rem">📑 {c["label"]}</span>',
                unsafe_allow_html=True,
            )
            break

    with st.spinner("Cargando…"):
        st.image(
            _render(st.session_state.pdf_hash, st.session_state.pdf_bytes, cp),
            use_container_width=True,
        )

    st.download_button(
        "📥 Descargar PDF", data=st.session_state.pdf_bytes,
        file_name="documento.pdf", mime="application/pdf",
        use_container_width=True,
    )


# ── EDITOR DE DATOS ───────────────────────────────────────────
with col_R:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)

    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona 🔍 Extraer Montos.")
    else:
        # KPIs arriba — se rellenan después del editor con datos frescos
        kpi_slot = st.container()

        # ── Editor interactivo ────────────────────────────────
        edited_df = st.data_editor(
            st.session_state.df,
            key="data_editor_main",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha":                  st.column_config.TextColumn("Fecha (YYYY-MM-DD)", width="small"),
                "Rubro":                  st.column_config.TextColumn("Rubro", width="large"),
                "QT":                     st.column_config.SelectboxColumn("QT", options=["Sí","No"], width="small"),
                "T. Cambio":              st.column_config.SelectboxColumn("T. Cambio", options=["MXN","USD","EUR"], width="small"),
                "(+ IVA)":                st.column_config.SelectboxColumn("(+ IVA)", options=["Sí","No","N/M"], width="small"),
                "Cantidad":               st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario":        st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)":     st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%":                st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA":          st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Diferencia final":       st.column_config.NumberColumn("Diferencia", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones":          st.column_config.TextColumn("Observaciones", width="large"),
            },
        )

        # [BF-1] Guardar con Diferencia final recalculada en cada edición
        if edited_df is not None:
            st.session_state.df = recalc_derived(edited_df)

        # ── KPIs con datos frescos  [OPT-3] ──────────────────
        df_cur = st.session_state.df
        ts_  = pd.to_numeric(df_cur["Total con IVA"],          errors="coerce").sum()
        rs_  = pd.to_numeric(df_cur["Monto en Anexo Escrito"], errors="coerce").sum()
        dif  = ts_ - rs_

        with kpi_slot:
            k1, k2, k3 = st.columns(3)
            for col, val, lbl in [
                (k1, ts_,  "Total Extraído"),
                (k2, rs_,  "Monto Referencia"),
                (k3, dif,  "Diferencia"),
            ]:
                color = "#1a2744" if lbl != "Diferencia" else (
                    "#c0392b" if abs(dif) > 0.01 else "#28a745"
                )
                col.markdown(
                    f'<div class="kpi">'
                    f'<div class="v" style="color:{color}">${val:,.2f}</div>'
                    f'<div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

        # ── Auditoría con IA  [BF-4, UI-3] ───────────────────
        if st.button("🤖 Auditar montos con IA", type="secondary", use_container_width=True):
            if "GEMINI_API_KEY" not in st.secrets:
                st.error("🔑 Agrega GEMINI_API_KEY en los Secrets de la app para usar esta función.")
            else:
                with st.spinner("Analizando con Gemini…"):
                    gemini_resp_text = ""
                    try:
                        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

                        # [BF-4] Intentar modelos en orden, sin list_models()
                        model_obj  = None
                        used_model = ""
                        for m_name in _GEMINI_PREFERRED:
                            try:
                                candidate = genai.GenerativeModel(m_name)
                                candidate.count_tokens("ping")   # valida disponibilidad
                                model_obj  = candidate
                                used_model = m_name
                                break
                            except Exception:
                                continue

                        if model_obj is None:
                            st.error("No se encontró ningún modelo Gemini disponible. "
                                     "Verifica que tu API key tenga acceso.")
                        else:
                            datos  = df_cur.to_dict(orient="records")
                            prompt = f"""Eres un auditor financiero experto en cotizaciones mexicanas.
Revisa los siguientes registros extraídos de un PDF de cotización:

{json.dumps(datos, ensure_ascii=False, indent=2)}

Verifica EXCLUSIVAMENTE la consistencia matemática:
1. ¿Cantidad × Precio Unitario ≈ Subtotal (Sin IVA)?
2. ¿Subtotal × 0.16 ≈ IVA 16%?
3. ¿Subtotal + IVA ≈ Total con IVA?

Si hay discrepancias, indica el rubro afectado, el campo incorrecto,
el valor extraído y el valor matemáticamente correcto.

Responde ÚNICAMENTE con JSON válido sin texto adicional ni backticks:
{{
  "ok": true,
  "observacion": "Resumen conciso de máximo dos oraciones.",
  "discrepancias": [
    {{"rubro": "...", "campo": "...", "valor_extraido": 0.0, "valor_correcto": 0.0}}
  ]
}}"""

                            response = model_obj.generate_content(prompt)
                            gemini_resp_text = response.text

                            clean = (gemini_resp_text.strip()
                                     .removeprefix("```json").removeprefix("```")
                                     .removesuffix("```").strip())
                            result = json.loads(clean)

                            if result.get("ok", False):
                                st.success(f"✅ Matemáticas correctas · modelo: {used_model}")
                            else:
                                st.warning(f"⚠️ Inconsistencias detectadas · {used_model}")

                            st.info(f"**Veredicto:** {result.get('observacion','Sin comentarios.')}")

                            # [UI-3] Tabla de discrepancias estructurada
                            discs = result.get("discrepancias", [])
                            if discs:
                                st.dataframe(
                                    pd.DataFrame(discs),
                                    hide_index=True,
                                    use_container_width=True,
                                )

                    except json.JSONDecodeError:
                        st.warning("La IA respondió en formato inesperado. Respuesta cruda:")
                        st.code(gemini_resp_text[:600] if gemini_resp_text else "(vacía)")
                    except Exception as exc:
                        st.error(f"Error en auditoría IA: {exc}")

        st.markdown("---")

        # ── Descarga Excel ────────────────────────────────────
        try:
            xlsx_bytes = to_excel(st.session_state.df)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"Conciliacion_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Error generando Excel: {exc}")

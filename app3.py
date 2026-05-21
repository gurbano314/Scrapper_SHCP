# ============================================================
# EFIDEPORTE – Extractor Genérico de Cotizaciones  |  app.py
# ============================================================

# ── 0. AUTO-INSTALADOR ──────────────────────────────────────
import importlib, subprocess, sys

_PKGS = {
    "streamlit": "streamlit", "pdfplumber": "pdfplumber", "fitz": "PyMuPDF",
    "rapidocr_onnxruntime": "rapidocr-onnxruntime", "pandas": "pandas",
    "xlsxwriter": "xlsxwriter", "numpy": "numpy", "PIL": "pillow",
}
_missing = [pkg for mod, pkg in _PKGS.items() if importlib.util.find_spec(mod) is None]
if _missing:
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages"] + _missing
    )

# ── 1. IMPORTS ───────────────────────────────────────────────
import io, re, datetime

import fitz
import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st
import xlsxwriter

# ── 2. CONFIGURACIÓN ─────────────────────────────────────────
st.set_page_config(
    page_title="Extractor de Cotizaciones",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { background:#1a2744; }
[data-testid="stSidebar"] * { color:#e8e8e8 !important; }
.main-header {
    background:linear-gradient(90deg,#1a2744,#2d4a8f);
    padding:14px 20px; border-radius:8px; margin-bottom:16px;
}
.main-header h1 { color:#fff; font-size:1.5rem; margin:0; }
.main-header span { color:#a8c0ff; font-size:.9rem; }
.panel-title {
    background:#2d4a8f; color:white !important;
    padding:8px 14px; border-radius:6px 6px 0 0;
    font-weight:700; font-size:.9rem; margin-bottom:0;
}
.metric-card {
    background:#f0f4ff; border:1px solid #c5d3f5;
    border-radius:8px; padding:10px 14px; text-align:center;
}
.metric-card .val { font-size:1.3rem; font-weight:700; color:#1a2744; }
.metric-card .lbl { font-size:.75rem; color:#5566aa; }
.stDownloadButton>button {
    background:#1a6b35; color:white; font-weight:600;
    border-radius:6px; border:none; padding:8px 20px; width:100%;
}
</style>
""", unsafe_allow_html=True)

# ── 3. SESSION STATE ─────────────────────────────────────────
_DEFAULTS: dict = {
    "pdf_bytes": None,
    "total_pages": 0,
    "current_page": 0,
    "num_sections": 1,
    "sections_config": [],
    "tabla_datos": None,
    "extracted": False,
}
for _k, _v in _DEFAULTS.items():
    st.session_state.setdefault(_k, _v)

# ── 4. CONSTANTES DE EXTRACCIÓN ──────────────────────────────
# FIX #1: regex requiere símbolo $ explícito O coma-decimal para evitar
#         capturar conteos de unidades como montos.
_MONEY_RE = re.compile(r"\$\s*([\d,]+(?:\.\d{1,2})?)|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)")
_DATE_RE  = re.compile(
    r"(\d{1,2})\s+de\s+(\w+)[,\s]+(?:del?\s+)?(\d{4})"   # 26 de marzo de 2026
    r"|(\d{4})[/-](\d{1,2})[/-](\d{1,2})"                 # 2026-03-26
    r"|(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"               # 26/03/26
)
_MESES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
}
_HEADERS = [
    "Fecha","Rubro","QT","T. Cambio","(+ IVA)","Cantidad",
    "Precio Unitario","Subtotal (Sin IVA)","IVA 16%","Total con IVA",
    "Diferencia final","Monto en Anexo Escrito","Observaciones",
]
_COL_WIDTHS = [12, 52, 5, 10, 7, 8, 16, 18, 17, 16, 18, 22, 48]

# ── 5. SINGLETON OCR ─────────────────────────────────────────
# FIX #2: el modelo ONNX (~200 MB) se carga UNA SOLA VEZ por sesión del servidor.
@st.cache_resource(show_spinner=False)
def _get_ocr():
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()

# ── 6. HELPERS ───────────────────────────────────────────────
def _parse_money(txt: str) -> float | None:
    """Extrae el primer monto válido. Requiere $ o formato de miles con comas."""
    for m in _MONEY_RE.finditer(str(txt)):
        raw = m.group(1) or m.group(2)
        if raw:
            try:
                return float(raw.replace(",", ""))
            except ValueError:
                continue
    return None

def _parse_date(text: str) -> datetime.date | None:
    for m in _DATE_RE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]:
                return datetime.date(int(g[2]), _MESES.get(g[1], 1), int(g[0]))
            if g[3]:
                return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            if g[6]:
                yr = int(g[8])
                return datetime.date(yr + 2000 if yr < 100 else yr, int(g[7]), int(g[6]))
        except (ValueError, KeyError):
            continue
    return None

# FIX #4: render_page cachea la imagen PNG, no re-abre el PDF en cada rerun.
@st.cache_data(max_entries=60, show_spinner=False)
def render_page(pdf_bytes: bytes, page_idx: int, zoom: float = 1.6) -> bytes:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:   # FIX #10: with-statement
        pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB)
    return pix.tobytes("png")

def _ocr_page(pdf_bytes: bytes, page_idx: int) -> str:
    """OCR de una página usando el singleton RapidOCR."""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:   # FIX #10
            pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(2.5, 2.5), colorspace=fitz.csRGB)
        img_bgr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)[:, :, ::-1]
        result, _ = _get_ocr()(img_bgr)
        if result:
            return "\n".join(r[1] for r in result if r and len(r) > 1)
    except Exception:
        pass
    return ""

# ── 7. MOTOR DE EXTRACCIÓN ───────────────────────────────────
def extract_section(pdf_bytes: bytes, label: str, p_ini: int, p_fin: int,
                    detect_iva: bool = True, calc_subtotal: bool = True) -> dict:
    """
    3 niveles: tablas nativas → regex en texto → OCR.
    Retorna un dict con las 13 columnas de la plantilla.
    """
    all_text, all_rows = "", []

    # FIX #7: pdfplumber se abre UNA SOLA VEZ para todas las páginas del rango.
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        for pg in range(p_ini - 1, min(p_fin, n_pages)):
            page = pdf.pages[pg]
            all_text += "\n" + (page.extract_text() or "")
            for tbl in page.extract_tables():
                if tbl:
                    all_rows.extend([str(c or "").strip() for c in row] for row in tbl if row)

    # Fallback OCR si el texto extraído es insignificante
    if len(all_text.strip()) < 30:
        for pg in range(p_ini - 1, min(p_fin, n_pages)):
            all_text += "\n" + _ocr_page(pdf_bytes, pg)

    fecha_val = _parse_date(all_text)
    text_low  = all_text.lower()

    iva_flag = "N/M"
    if detect_iva and re.search(r"\biva\b|16%|vat", text_low):
        iva_flag = "Sí"

    # ── Extraer montos de tablas ─────────────────────────────
    total_val = iva_val = sub_val = pu_val = None
    qty = 1

    for row in all_rows:
        joined = "  ".join(row).lower()
        row_str = "  ".join(row)
        if "total" in joined and not re.search(r"sub|parcial", joined):
            v = _parse_money(row_str)
            if v and (total_val is None or v > total_val):
                total_val = v
        if re.search(r"\biva\b|16%|vat", joined):
            v = _parse_money(row_str)
            if v:
                iva_val = v
        if re.search(r"subtotal|sin\s*iva", joined):
            v = _parse_money(row_str)
            if v:
                sub_val = v

    # Fallback: buscar montos en líneas de texto libre
    if total_val is None:
        for line in all_text.splitlines():
            if re.search(r"\btotal\b", line, re.I) and not re.search(r"sub|parcial", line, re.I):
                v = _parse_money(line)
                if v and (total_val is None or v > total_val):
                    total_val = v
    if iva_val is None and iva_flag == "Sí":
        for line in all_text.splitlines():
            if re.search(r"\biva\b|16%|vat", line, re.I):
                v = _parse_money(line)
                if v:
                    iva_val = v
                    break

    # FIX #6: extracción de cantidad y precio unitario con conversión segura
    for row in all_rows:
        nums_f = []
        for tok in re.findall(r"[\d,]+(?:\.\d+)?", "  ".join(row)):
            try:
                nums_f.append(float(tok.replace(",", "")))
            except ValueError:
                continue
        if len(nums_f) >= 2 and 1 <= nums_f[0] <= 999 and nums_f[-1] > 0:
            qty   = int(nums_f[0])
            pu_val = nums_f[-2] if len(nums_f) > 2 else nums_f[-1]

    # ── Lógica de montos ─────────────────────────────────────
    obs = ""
    if total_val and not sub_val and not iva_val and iva_flag == "Sí":
        sub_val = round(total_val / 1.16, 2)
        iva_val = round(total_val - sub_val, 2)
        obs = "Precios ya incluyen impuestos"
    elif total_val and iva_val and not sub_val:
        sub_val = round(total_val - iva_val, 2)
    elif sub_val and iva_val and not total_val:
        total_val = round(sub_val + iva_val, 2)
    elif calc_subtotal and sub_val and not iva_val and not total_val and iva_flag == "Sí":
        iva_val   = round(sub_val * 0.16, 2)
        total_val = round(sub_val + iva_val, 2)

    if pu_val is None and sub_val:
        pu_val = round(sub_val / qty, 2) if qty else sub_val

    # FIX #9: Monto en Anexo Escrito se inicializa en None para que
    #         el usuario lo capture manualmente y se detecte discrepancia real.
    return {
        "Fecha":                  fecha_val or datetime.date.today(),
        "Rubro":                  label,
        "QT":                     "Sí",
        "T. Cambio":              "MXN",
        "(+ IVA)":                iva_flag,
        "Cantidad":               qty,
        "Precio Unitario":        pu_val,
        "Subtotal (Sin IVA)":     sub_val,
        "IVA 16%":                iva_val,
        "Total con IVA":          total_val,
        "Diferencia final":       None,
        "Monto en Anexo Escrito": None,   # ← intencionalmente vacío
        "Observaciones":          obs,
    }

# ── 8. EXPORTACIÓN EXCEL ─────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws  = wb.add_worksheet("Conciliación")

    MONEY_FMT = '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'
    base    = {"font_name": "Calibri", "font_size": 11, "border": 1}
    h_fmt   = wb.add_format({**base, "bold": True, "bg_color": "#D4C19C", "align": "center", "valign": "vcenter"})
    h2_fmt  = wb.add_format({**base, "bold": True, "bg_color": "#EBE2D1", "align": "center", "valign": "vcenter"})
    date_f  = wb.add_format({**base, "num_format": "dd/mm/yy"})
    money_f = wb.add_format({**base, "num_format": MONEY_FMT})
    text_f  = wb.add_format({**base})
    num_f   = wb.add_format({**base, "num_format": "#,##0"})
    total_f = wb.add_format({**base, "bold": True, "num_format": MONEY_FMT})
    obs_f   = wb.add_format({**base, "text_wrap": True})

    # Encabezados
    for c, h in enumerate(_HEADERS):
        ws.write(0, c, h, h2_fmt if c >= 11 else h_fmt)
    for i, w in enumerate(_COL_WIDTHS):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)

    # Mapeo columna → (índice_xl, formula_si_nulo)
    MONEY_COLS = {
        "Precio Unitario":    (6,  None),
        "Subtotal (Sin IVA)": (7,  "=F{r}*G{r}"),
        "IVA 16%":            (8,  "=H{r}*0.16"),
        "Total con IVA":      (9,  "=H{r}+I{r}"),
        "Monto en Anexo Escrito": (11, None),
    }

    D = 1  # primera fila de datos (0-indexed)
    n = len(df)

    for i, (_, row) in enumerate(df.iterrows()):
        r0 = D + i          # 0-indexed row
        r1 = r0 + 1         # 1-indexed (para fórmulas)

        # A: Fecha
        fecha = row.get("Fecha")
        if isinstance(fecha, (datetime.date, datetime.datetime)):
            dt = datetime.datetime.combine(fecha, datetime.time()) if isinstance(fecha, datetime.date) else fecha
            ws.write_datetime(r0, 0, dt, date_f)
        else:
            ws.write(r0, 0, str(fecha) if fecha else "", date_f)

        ws.write(r0, 1, str(row.get("Rubro",    "")),    obs_f)
        ws.write(r0, 2, str(row.get("QT",       "Sí")), text_f)
        ws.write(r0, 3, str(row.get("T. Cambio","MXN")),text_f)
        ws.write(r0, 4, str(row.get("(+ IVA)",  "")),   text_f)

        qty = row.get("Cantidad")
        ws.write(r0, 5, int(qty) if pd.notna(qty) else 1, num_f)

        for col_name, (xl_c, formula) in MONEY_COLS.items():
            val = row.get(col_name)
            if pd.notna(val):
                ws.write_number(r0, xl_c, float(val), money_f)
            elif formula:
                ws.write_formula(r0, xl_c, formula.format(r=r1), money_f)
            else:
                ws.write(r0, xl_c, "", money_f)

        # K: Diferencia final (siempre fórmula)
        ws.write_formula(r0, 10, f"=J{r1}-L{r1}", money_f)
        ws.write(r0, 12, str(row.get("Observaciones", "") or ""), obs_f)

    # Fila de totales
    s1, e1 = D + 1, D + n
    ws.write_formula(D + n, 9,  f"=SUM(J{s1}:J{e1})", total_f)
    ws.write_formula(D + n, 10, f"=SUM(K{s1}:K{e1})", total_f)
    ws.write_formula(D + n, 11, f"=SUM(L{s1}:L{e1})", total_f)

    wb.close()
    buf.seek(0)
    return buf.read()

# ── 9. SIDEBAR ───────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Cargar Documento")
    uploaded = st.file_uploader("Sube un archivo PDF", type=["pdf"])

    if uploaded:
        pdf_bytes = uploaded.read()
        if pdf_bytes != st.session_state.pdf_bytes:
            st.session_state.pdf_bytes    = pdf_bytes
            st.session_state.extracted    = False
            st.session_state.tabla_datos  = None
            st.session_state.current_page = 0
            with fitz.open(stream=pdf_bytes, filetype="pdf") as _d:   # FIX #10
                st.session_state.total_pages = len(_d)
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")

    st.markdown("---")
    st.markdown("### ⚙️ Configuración de Secciones")

    # FIX #8: n_sec se guarda en session_state para que sea accesible en todo el script
    n_sec = st.number_input(
        "Número de secciones/rubros", min_value=1, max_value=50,
        value=st.session_state.num_sections, step=1,
    )
    if n_sec != st.session_state.num_sections:
        st.session_state.num_sections = n_sec
        st.session_state.extracted    = False
        st.session_state.tabla_datos  = None

    cfg_list = st.session_state.sections_config
    while len(cfg_list) < n_sec:
        idx = len(cfg_list) + 1
        cfg_list.append({"label": f"Sección {idx}", "p_ini": idx, "p_fin": idx,
                         "detect_iva": True, "calc_subtotal": True})
    del cfg_list[n_sec:]

    total_pgs = st.session_state.total_pages or 1
    for i in range(n_sec):
        cfg = cfg_list[i]
        with st.expander(f"📄 Sección {i+1}", expanded=(n_sec <= 5)):
            cfg["label"]        = st.text_input("Rubro/Concepto", value=cfg["label"], key=f"lbl_{i}")
            c1, c2              = st.columns(2)
            cfg["p_ini"]        = c1.number_input("Pág. Inicio", min_value=1, max_value=total_pgs, value=min(cfg["p_ini"], total_pgs), key=f"pi_{i}")
            cfg["p_fin"]        = c2.number_input("Pág. Fin",    min_value=cfg["p_ini"], max_value=total_pgs, value=max(min(cfg["p_fin"], total_pgs), cfg["p_ini"]), key=f"pf_{i}")
            cfg["detect_iva"]   = st.checkbox("Detectar IVA automáticamente", value=cfg["detect_iva"],   key=f"iva_{i}")
            cfg["calc_subtotal"]= st.checkbox("Calcular subtotal si falta",   value=cfg["calc_subtotal"],key=f"cs_{i}")

    st.markdown("---")
    extract_btn = st.button(
        "🔍 Extraer Montos",
        disabled=(st.session_state.pdf_bytes is None),
        use_container_width=True, type="primary",
    )

# ── 10. HEADER ───────────────────────────────────────────────
st.markdown(
    '<div class="main-header"><h1>📊 Extractor Genérico de Cotizaciones</h1>'
    '<span>Extrae, edita y concilia montos desde cualquier PDF</span></div>',
    unsafe_allow_html=True,
)

# ── 11. EXTRACCIÓN ───────────────────────────────────────────
if extract_btn and st.session_state.pdf_bytes:
    rows = []
    prog = st.progress(0, text="Procesando…")
    cfgs = st.session_state.sections_config

    for i, cfg in enumerate(cfgs):
        prog.progress((i + 0.5) / n_sec, text=f"Extrayendo: {cfg['label']}")
        try:
            rows.append(extract_section(
                st.session_state.pdf_bytes, cfg["label"],
                cfg["p_ini"], cfg["p_fin"],
                cfg["detect_iva"], cfg["calc_subtotal"],
            ))
        except Exception as e:
            st.warning(f"⚠️ Error en sección {i+1}: {str(e)[:80]}")
            rows.append({k: None for k in _HEADERS} | {
                "Rubro": cfg["label"], "QT": "Sí", "T. Cambio": "MXN",
                "Cantidad": 1, "Fecha": datetime.date.today(),
                "Observaciones": f"Error: {str(e)[:60]}",
            })
        prog.progress((i + 1) / n_sec)

    prog.empty()
    df_new = pd.DataFrame(rows, columns=_HEADERS)
    # Diferencia final = Total - Monto Anexo (ambos pueden ser None)
    df_new["Diferencia final"] = (
        df_new["Total con IVA"].fillna(0) - df_new["Monto en Anexo Escrito"].fillna(0)
    ).where(df_new["Total con IVA"].notna() | df_new["Monto en Anexo Escrito"].notna())

    st.session_state.tabla_datos = df_new
    st.session_state.extracted   = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()

# ── 12. LAYOUT ───────────────────────────────────────────────
col_left, col_right = st.columns(2, gap="medium")

# ══ VISOR PDF ════════════════════════════════════════════════
with col_left:
    st.markdown('<p class="panel-title">🔍 Visor de Documento</p>', unsafe_allow_html=True)

    tp = st.session_state.total_pages
    cp = st.session_state.current_page

    nav1, nav2, nav3, nav4 = st.columns([1, 1, 3, 1])
    if nav1.button("◀", key="prev"): cp = max(0, cp - 1)
    pg_sel = nav3.number_input("", min_value=1, max_value=tp, value=cp + 1,
                               label_visibility="collapsed", key="pg_input")
    cp = pg_sel - 1
    if nav4.button("▶", key="next"): cp = min(tp - 1, cp + 1)
    st.session_state.current_page = cp

    st.caption(f"Página **{cp + 1}** de **{tp}**")

    # Etiqueta de sección activa
    for cfg in st.session_state.sections_config:
        if cfg["p_ini"] <= cp + 1 <= cfg["p_fin"]:
            st.markdown(
                f'<span style="background:#3b5998;color:#fff;padding:3px 10px;'
                f'border-radius:4px;font-size:.8rem">📑 {cfg["label"]}</span>',
                unsafe_allow_html=True,
            )
            break

    with st.spinner("Renderizando…"):
        st.image(render_page(st.session_state.pdf_bytes, cp), use_container_width=True)  # FIX #3

    st.download_button(
        "📥 Descargar PDF original",
        data=st.session_state.pdf_bytes,
        file_name="documento.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

# ══ EDITOR DE DATOS ══════════════════════════════════════════
with col_right:
    st.markdown('<p class="panel-title">✏️ Editor de Datos Extraídos</p>', unsafe_allow_html=True)

    if not st.session_state.extracted or st.session_state.tabla_datos is None:
        st.info("Configura las secciones y presiona **🔍 Extraer Montos**.")
    else:
        df = st.session_state.tabla_datos

        # Métricas
        total_sum = df["Total con IVA"].sum(skipna=True)
        ref_sum   = df["Monto en Anexo Escrito"].sum(skipna=True)
        diff      = total_sum - ref_sum
        mc1, mc2, mc3 = st.columns(3)
        for col, val, lbl in [(mc1, total_sum, "Total Extraído"),
                              (mc2, ref_sum,   "Monto Referencia"),
                              (mc3, diff,      "Diferencia")]:
            color = "#1a2744" if lbl != "Diferencia" else ("#c0392b" if abs(diff) > 0.01 else "#28a745")
            col.markdown(
                f'<div class="metric-card">'
                f'<div class="val" style="color:{color}">${val:,.2f}</div>'
                f'<div class="lbl">{lbl}</div></div>',
                unsafe_allow_html=True,
            )

        # FIX #5: sin df.copy() — session_state es la única fuente de verdad.
        edited = st.data_editor(
            st.session_state.tabla_datos,   # ← referencia directa, sin copia
            key="editor_datos",             # ← clave estable: sobrevive reruns
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha":                  st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", width="small"),
                "Rubro":                  st.column_config.TextColumn("Rubro", width="large"),
                "QT":                     st.column_config.SelectboxColumn("QT", options=["Sí","No"], width="small"),
                "T. Cambio":              st.column_config.SelectboxColumn("T. Cambio", options=["MXN","USD","EUR"], width="small"),
                "(+ IVA)":               st.column_config.SelectboxColumn("(+ IVA)", options=["Sí","No","N/M"], width="small"),
                "Cantidad":               st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario":        st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)":     st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%":               st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA":          st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Diferencia final":       st.column_config.NumberColumn("Diff.", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones":          st.column_config.TextColumn("Observaciones", width="large"),
            },
        )
        if edited is not None:
            st.session_state.tabla_datos = edited

        st.markdown("---")
        st.markdown("### 📥 Exportar Resultados")
        try:
            xlsx = build_excel(st.session_state.tabla_datos)
            ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "⬇️ Descargar Excel Formateado",
                data=xlsx,
                file_name=f"Conciliacion_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Error generando Excel: {e}")

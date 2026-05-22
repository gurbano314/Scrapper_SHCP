# ============================================================
# EXTRACTOR DE COTIZACIONES PDF | app.py
# Compatible con Streamlit Community Cloud + OCR Robusto + IA Auditora
# ============================================================

import hashlib, io, re, os, yaml, pathlib, tempfile, datetime, sys, json
from collections import defaultdict
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
.hdr p { color:#a8c0ff; font-size:.87rem; margin:4px 0 0; }
.ptitle {
    background:#2d4a8f; color:#fff !important; font-weight:700;
    font-size:.88rem; padding:7px 14px; border-radius:6px 6px 0 0;
    margin-bottom:4px;
}
.kpi {
    background:#f0f4ff; border:1px solid #c5d3f5;
    border-radius:8px; padding:10px 12px; text-align:center; margin-bottom:6px;
}
.kpi .v { font-size:1.2rem; font-weight:700; color:#1a2744; }
.kpi .l { font-size:.72rem; color:#5566aa; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# CONSTANTES Y REGEX TOLERANTES A OCR
# ─────────────────────────────────────────────────────────────
_MONEY_RE = re.compile(r"\$?\s*([\d,]+(?:\.\d{1,2})?)|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)")
_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:de)?\s*([a-zA-Z]+)[,\s]*(?:del?\s*)?(\d{4})|"
    r"(\d{4})[\s\.\-\/]+(\d{1,2})[\s\.\-\/]+(\d{1,2})|"
    r"(\d{1,2})[\s\.\-\/]+(\d{1,2})[\s\.\-\/]+(\d{2,4})"
)
_ITEM_RE = re.compile(r"^\d+\s+(\d+)\s+\w+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$")
_MESES = {
    "enero":1, "febrero":2, "marzo":3, "abril":4, "mayo":5, "junio":6,
    "julio":7, "agosto":8, "septiembre":9, "octubre":10, "noviembre":11, "diciembre":12,
}
_COLS = [
    "Fecha", "Rubro", "QT", "T. Cambio", "(+ IVA)", "Cantidad",
    "Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%", "Total con IVA",
    "Diferencia final", "Monto en Anexo Escrito", "Observaciones",
]
_WIDTHS = [12,52,5,10,7,8,16,18,17,16,18,22,48]

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
for k, v in {
    "pdf_bytes": None, "pdf_hash": None, "total_pages": 0,
    "current_page": 0, "num_sec": 1, "sec_cfg": [],
    "df": None, "extracted": False,
}.items():
    st.session_state.setdefault(k, v)

# ─────────────────────────────────────────────────────────────
# OCR ROBUSTO PARA STREAMLIT CLOUD
# ─────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _get_ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
        ocr = RapidOCR(det_model_dir=None, rec_model_dir=None, cls_model_dir=None)
        return ocr
    except Exception as e:
        st.error(f"⚠️ Error cargando motor OCR: {str(e)}")
        return None

def _ocr_page(pdf_bytes: bytes, idx: int) -> str:
    ocr = _get_ocr_engine()
    if ocr is None: return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(matrix=fitz.Matrix(2.0, 2.0), colorspace=fitz.csRGB, alpha=False)
            img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)[:, :, ::-1]
            result, _ = ocr(img)
            return "\n".join(r[1] for r in result if r and len(r) > 1) if result else ""
    except Exception as e:
        st.warning(f"⚠️ Error procesando OCR en página {idx+1}: {str(e)}")
        return ""

# ─────────────────────────────────────────────────────────────
# HELPERS DE CONVERSIÓN
# ─────────────────────────────────────────────────────────────
def _safe_f(v) -> Optional[float]:
    try:
        f = float(v)
        return None if f != f else f
    except Exception:
        return None

def _money(txt: str) -> Optional[float]:
    for m in _MONEY_RE.finditer(str(txt)):
        raw = m.group(1) or m.group(2)
        if raw:
            try: return float(raw.replace(",", ""))
            except ValueError: continue
    return None

def _date(text: str) -> Optional[datetime.date]:
    for m in _DATE_RE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]: 
                mes_str = g[1].lower()
                mes_num = _MESES.get(mes_str)
                if not mes_num:
                    for k, v in _MESES.items():
                        if k in mes_str or mes_str in k:
                            mes_num = v
                            break
                if mes_num: return datetime.date(int(g[2]), mes_num, int(g[0]))
            if g[3]: return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            if g[6]:
                yr = int(g[8])
                return datetime.date(yr + 2000 if yr < 100 else yr, int(g[7]), int(g[6]))
        except (ValueError, KeyError): continue
    return None

def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

@st.cache_data(max_entries=50, show_spinner=False)
def _render(pdf_bytes: bytes, idx: int) -> bytes:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), colorspace=fitz.csRGB, alpha=False)
        return pix.tobytes("png")

def _parse_space_table(text: str) -> list[dict]:
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
# MOTOR DE EXTRACCIÓN HÍBRIDO AUTOMÁTICO
# ─────────────────────────────────────────────────────────────
def extract(pdf_bytes: bytes, label: str, p0: int, p1: int, det_iva: bool, calc_sub: bool) -> dict:
    native_text, ocr_text, table_rows = "", "", []
    
    # Procesamiento Nativo (Nivel 1 y 2)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n = len(pdf.pages)
        pr = range(p0 - 1, min(p1, n))
        for i in pr:
            pg = pdf.pages[i]
            txt = pg.extract_text() or ""
            native_text += "\n" + txt
            
            for tbl in pg.extract_tables():
                if tbl:
                    table_rows.extend([str(c or "").strip() for c in r] for r in tbl if r)
            
            for item in _parse_space_table(txt):
                table_rows.append([str(item["qty"]), item["desc"], f"${item['pu']:,.2f}", f"${item['total']:,.2f}"])
    
    # Procesamiento OCR Automático de Respaldo Integrado (Nivel 3)
    for i in pr:
        o_txt = _ocr_page(pdf_bytes, i)
        if o_txt:
            ocr_text += "\n" + o_txt
            for item in _parse_space_table(o_txt):
                table_rows.append([str(item["qty"]), item["desc"], f"${item['pu']:,.2f}", f"${item['total']:,.2f}"])
    
    # Unión total de capas de texto
    text = native_text + "\n" + ocr_text
    tlow = text.lower()
    iva_f = "Sí" if det_iva and re.search(r"\biva\b|16%|vat", tlow) else "N/M"
    
    tot = iva = sub = pu = fecha = None
    qty = 1
    obs = ""
    
    # 1. Búsqueda de Totales Explícitos en Tablas estructuradas
    for row in table_rows:
        j, s = "   ".join(row).lower(), "   ".join(row)
        if "total" in j and not re.search(r"sub|parcial", j):
            v = _money(s)
            if v and (tot is None or v > tot): tot = v
        if re.search(r"\biva\b|16%|vat", j):
            v = _money(s)
            if v: iva = v
        if re.search(r"subtotal|sin\s*iva|importe", j):
            v = _money(s)
            if v: sub = v

    # 2. ALGORITMO DE RECONSTRUCCIÓN: Sumar subtotales de productos si falta la fila Total
    valid_line_totals = []
    seen_combinations = set()
    
    for row in table_rows:
        row_str = " ".join([str(c) for c in row])
        numbers = []
        for token in re.findall(r"[\d,]+(?:\.\d+)?", row_str):
            num = _safe_f(token.replace(",", ""))
            if num is not None and num > 0: numbers.append(num)
            
        if len(numbers) >= 2:
            found_item = False
            for idx_tot, total_cand in enumerate(numbers):
                for idx_pu, pu_cand in enumerate(numbers):
                    for idx_qty, qty_cand in enumerate(numbers):
                        if idx_tot != idx_pu and idx_tot != idx_qty and idx_pu != idx_qty:
                            if 1 <= qty_cand <= 50000 and abs(qty_cand * pu_cand - total_cand) < 5.0:
                                combo_key = (int(qty_cand), round(pu_cand, 2), round(total_cand, 2))
                                if combo_key not in seen_combinations:
                                    seen_combinations.add(combo_key)
                                    valid_line_totals.append(total_cand)
                                found_item = True
                                break
                    if found_item: break
                if found_item: break

    if tot is None and valid_line_totals:
        tot = round(sum(valid_line_totals), 2)
        obs = "Total calculado sumando líneas de productos"

    # 3. Escaneo de Texto Libre Completo (Global Fallback)
    for ln in text.splitlines():
        if fecha is None:
            m_date = _date(ln)
            if m_date: fecha = m_date

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b", ln, re.I) and not re.search(r"sub|parcial", ln, re.I):
                v = _money(ln)
                if v and (tot is None or v > tot): tot = v

    if tot is None:
        posibles_montos = []
        for row in table_rows:
            for cell in row:
                val = _money(cell)
                if val is not None and val > 0: posibles_montos.append(val)
        for ln in text.splitlines():
            for m in _MONEY_RE.finditer(ln):
                raw = m.group(1) or m.group(2)
                if raw:
                    try:
                        val = float(raw.replace(",", ""))
                        if val > 0: posibles_montos.append(val)
                    except ValueError: continue
        if posibles_montos:
            tot = max(posibles_montos)
            obs = "Total inferido (monto máximo)"

    if sub is None:
        for ln in text.splitlines():
            if re.search(r"sub\s*total|importe|sin\s*iva", ln, re.I) and not re.search(r"total\b", ln.replace("subtotal", ""), re.I):
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

    for row in table_rows:
        row_str = " ".join([str(c) for c in row])
        nf = []
        for t in re.findall(r"[\d,]+(?:\.\d+)?", row_str):
            try: nf.append(float(t.replace(",", "")))
            except ValueError: continue
        if len(nf) >= 2 and 1 <= nf[0] <= 999 and nf[-1] > 0:
            qty = int(nf[0])
            pu = nf[-2] if len(nf) > 2 else nf[-1]

    if tot and not sub and not iva and iva_f == "Sí":
        sub = round(tot / 1.16, 2)
        iva = round(tot - sub, 2)
        obs = obs + " | IVA incluido" if obs else "IVA incluido"
    elif tot and iva and not sub: sub = round(tot - iva, 2)
    elif sub and iva and not tot: tot = round(sub + iva, 2)
    elif calc_sub and sub and not iva and not tot and iva_f == "Sí":
        iva = round(sub * 0.16, 2)
        tot = round(sub + iva, 2)

    if pu is None and sub is not None:
        pu = round(sub / qty, 2) if qty else sub

    return {
        "Fecha": (fecha.isoformat() if fecha else datetime.date.today().isoformat()),
        "Rubro": label, "QT": "Sí", "T. Cambio": "MXN", "(+ IVA)": iva_f,
        "Cantidad": qty, "Precio Unitario": pu, "Subtotal (Sin IVA)": sub,
        "IVA 16%": iva, "Total con IVA": tot, "Diferencia final": None,
        "Monto en Anexo Escrito": None, "Observaciones": obs,
    }

# ─────────────────────────────────────────────────────────────
# EXPORTACIÓN A EXCEL
# ─────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("Conciliación")
    
    MF = '_-"$ "* #,##0.00_-;\-"$ "* #,##0.00_-;_-"$ "* "-"??_-;_-@_-'
    B = {"font_name": "Calibri", "font_size": 11, "border": 1}
    hf = wb.add_format({**B, "bold":True, "bg_color": "#D4C19C", "align": "center", "valign": "vcenter"})
    h2f = wb.add_format({**B, "bold":True, "bg_color": "#EBE2D1", "align": "center", "valign": "vcenter"})
    df_ = wb.add_format({**B, "num_format": "dd/mm/yy"})
    mf = wb.add_format({**B, "num_format": MF})
    tf = wb.add_format({**B})
    nf = wb.add_format({**B, "num_format": "#,##0"})
    tot = wb.add_format({**B, "bold":True, "num_format": MF})
    of = wb.add_format({**B, "text_wrap": True})
    
    for c, h in enumerate(_COLS): ws.write(0, c, h, h2f if c >= 11 else hf)
    for i, w in enumerate(_WIDTHS): ws.set_column(i, i, w)
    ws.set_row(0, 30)
    
    MC = {
        "Precio Unitario": (6, None),
        "Subtotal (Sin IVA)": (7, "=F{r}*G{r}"),
        "IVA 16%": (8, "=H{r}*0.16"),
        "Total con IVA": (9, "=H{r}+I{r}"),
        "Monto en Anexo Escrito": (11, None),
    }
    
    D, n = 1, len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        r0, r1 = D + i, D + i + 1
        fecha_val = row.get("Fecha")
        if isinstance(fecha_val, str):
            try: fecha_val = datetime.date.fromisoformat(fecha_val)
            except: fecha_val = None
        if isinstance(fecha_val, (datetime.date, datetime.datetime)): 
            dt = datetime.datetime.combine(fecha_val, datetime.time()) if isinstance(fecha_val, datetime.date) else fecha_val
            ws.write_datetime(r0, 0, dt, df_)
        else:
            ws.write(r0, 0, str(fecha_val or ""), df_)
        
        ws.write(r0, 1, str(row.get("Rubro", "") or ""), of)
        ws.write(r0, 2, str(row.get("QT", "Sí")), tf)
        ws.write(r0, 3, str(row.get("T. Cambio", "MXN")), tf)
        ws.write(r0, 4, str(row.get("(+ IVA)", "") or ""), tf)
        
        q = _safe_f(row.get("Cantidad"))
        ws.write(r0, 5, int(q) if q is not None else 1, nf)
        
        for cn, (xi, fm) in MC.items():
            v = _safe_f(row.get(cn))
            if v is not None: ws.write_number(r0, xi, v, mf)
            elif fm: ws.write_formula(r0, xi, fm.format(r=r1), mf)
            else: ws.write_blank(r0, xi, None, mf)
        
        ws.write_formula(r0, 10, f"=J{r1}-L{r1}", mf)
        ws.write(r0, 12, str(row.get("Observaciones", "") or ""), of)
    
    s1, e1 = D + 1, D + n
    for ci in (9, 10, 11):
        cl = chr(ord("A") + ci)
        ws.write_formula(D + n, ci, f"=SUM({cl}{s1}:{cl}{e1})", tot)
    
    wb.close()
    buf.seek(0)
    return buf.read()

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
            st.session_state.update(pdf_bytes=raw, pdf_hash=h, extracted=False, df=None, current_page=0)
            with fitz.open(stream=raw, filetype="pdf") as d: st.session_state.total_pages = len(d)
        st.success(f"✅ {st.session_state.total_pages} páginas cargadas")
    
    st.markdown("---")
    st.markdown("### ⚙️ Secciones")
    n = int(st.number_input("Número de secciones", min_value=1, max_value=50, value=st.session_state.num_sec, step=1))
    
    if n != st.session_state.num_sec:
        st.session_state.num_sec = n
        st.session_state.extracted = False
        st.session_state.df = None
    
    cfgs = st.session_state.sec_cfg
    tp = st.session_state.total_pages or 1
    
    while len(cfgs) < n:
        i = len(cfgs) + 1
        cfgs.append({"label": f"Sección {i}", "p0": i, "p1": i, "det_iva": True, "calc_sub": True})
    del cfgs[n:]
    
    for i, c in enumerate(cfgs):
        with st.expander(f"📄 Sección {i+1}", expanded=(n <= 5)):
            c["label"] = st.text_input("Rubro/Concepto", value=c["label"], key=f"lb{i}")
            a, b = st.columns(2)
            c["p0"] = a.number_input("Pág. Inicio", 1, tp, min(c["p0"], tp), key=f"p0{i}")
            c["p1"] = b.number_input("Pág. Fin", c["p0"], tp, max(min(c["p1"], tp), c["p0"]), key=f"p1{i}")
            c["det_iva"] = st.checkbox("Detectar IVA", value=c["det_iva"], key=f"iv{i}")
            c["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=c["calc_sub"], key=f"cs{i}")
    
    st.markdown("---")
    run = st.button("🔍 Extraer Montos", disabled=(st.session_state.pdf_bytes is None), use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
 <h1>📋 Conciliador de Cotizaciones PDF</h1>
 <p>Extrae, edita y exporta montos desde documentos PDF</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# EXTRACCIÓN
# ─────────────────────────────────────────────────────────────
if run and st.session_state.pdf_bytes:
    rows = []
    bar = st.progress(0, text="Procesando…")
    
    for i, c in enumerate(st.session_state.sec_cfg):
        bar.progress((i + .5) / n, text=f"Extrayendo: {c['label']}")
        try:
            rows.append(extract(st.session_state.pdf_bytes, c["label"], c["p0"], c["p1"], c["det_iva"], c["calc_sub"]))
        except Exception as e:
            st.warning(f"⚠️ Sección {i+1}: {str(e)[:80]}")
            rows.append({k: None for k in _COLS} | {
                "Rubro": c["label"], "QT": "Sí", "T. Cambio": "MXN", "Cantidad": 1,
                "Fecha": datetime.date.today().isoformat(), "Observaciones": f"Error: {str(e)[:60]}",
            })
        bar.progress((i + 1) / n)
    
    bar.empty()
    df_new = pd.DataFrame(rows, columns=_COLS)
    df_new["Diferencia final"] = (
        df_new["Total con IVA"].fillna(0) - df_new["Monto en Anexo Escrito"].fillna(0)
    ).where(df_new["Total con IVA"].notna() | df_new["Monto en Anexo Escrito"].notna())
    
    st.session_state.df = df_new
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# LAYOUT VISUAL
# ─────────────────────────────────────────────────────────────
L, R = st.columns(2, gap="medium")

with L:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)
    tp = st.session_state.total_pages
    cp = st.session_state.current_page
    
    c1, c2, c3 = st.columns([1, 4, 1])
    if c1.button("◀", key="pv"): cp = max(0, cp - 1)
    cp = int(c2.number_input("", 1, tp, cp + 1, label_visibility="collapsed", key="pg")) - 1
    if c3.button("▶", key="nx"): cp = min(tp - 1, cp + 1)
    
    st.session_state.current_page = cp
    st.caption(f"Página {cp+1} de {tp}")
    
    for c in st.session_state.sec_cfg:
        if c["p0"] <= cp + 1 <= c["p1"]:
            st.markdown(f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;border-radius:4px;font-size:.8rem">📑 {c["label"]}</span>', unsafe_allow_html=True)
            break
    
    with st.spinner("Cargando…"): st.image(_render(st.session_state.pdf_bytes, cp), use_container_width=True)
    st.download_button("📥 Descargar PDF", data=st.session_state.pdf_bytes, file_name="documento.pdf", mime="application/pdf", use_container_width=True)

with R:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)
    
    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona 🔍 Extraer Montos.")
    else:
        df = st.session_state.df
        ts_ = _safe_f(df["Total con IVA"].sum(skipna=True)) or 0.0
        rs_ = _safe_f(df["Monto en Anexo Escrito"].sum(skipna=True)) or 0.0
        dif = ts_ - rs_
        
        k1, k2, k3 = st.columns(3)
        for col, val, lbl in [(k1, ts_, "Total Extraído"), (k2, rs_, "Monto Referencia"), (k3, dif, "Diferencia")]:
            clr = "#1a2744" if lbl != "Diferencia" else ("#c0392b" if abs(dif) > .01 else "#28a745")
            col.markdown(f'<div class="kpi"><div class="v" style="color:{clr}">${val:,.2f}</div><div class="l">{lbl}</div></div>', unsafe_allow_html=True)
        
       # ==========================================
        # NUEVO BOTÓN: AUDITORÍA CON IA (AUTO-DESCUBRIMIENTO)
        # ==========================================
        if st.button("🤖 Auditar montos con IA", type="secondary", use_container_width=True):
            with st.spinner("La IA está revisando matemáticamente el documento..."):
                try:
                    # Configurar la llave desde los secretos
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    
                    # 1. AUTO-DESCUBRIMIENTO DE MODELOS
                    modelo_correcto = None
                    for m in genai.list_models():
                        if 'generateContent' in m.supported_generation_methods:
                            if 'flash' in m.name: # Preferimos flash por velocidad
                                modelo_correcto = m.name
                                break
                            if modelo_correcto is None:
                                modelo_correcto = m.name
                    
                    if not modelo_correcto:
                        st.error("Error: Tu API Key conectó, pero Google no asignó ningún modelo a tu cuenta.")
                        st.stop()
                        
                    model = genai.GenerativeModel(modelo_correcto)
                    
                    # Convertir el dataframe a diccionario para que la IA lo entienda
                    datos_extraidos = st.session_state.df.to_dict(orient="records")
                    
                    prompt = f"""
                    Eres un auditor financiero experto. Revisa estos datos extraídos de una cotización.
                    Verifica que las matemáticas cuadren perfectamente:
                    1. ¿Cantidad * Precio Unitario = Subtotal?
                    2. ¿El IVA 16% es matemáticamente correcto según el subtotal?
                    3. ¿Subtotal + IVA = Total con IVA?
                    
                    Datos: {datos_extraidos}
                    
                    Responde EXCLUSIVAMENTE con un JSON válido con esta estructura exacta:
                    {{"observacion_general": "Tu análisis corto y directo sobre si los montos cuadran matemáticamente o si hay discrepancias"}}
                    """
                    
                    response = model.generate_content(prompt)
                    # Limpiar el texto devuelto por si la IA le pone formato markdown
                    texto_limpio = response.text.strip().removeprefix("```json").removesuffix("```").strip()
                    resultado_ia = json.loads(texto_limpio)
                    
                    st.success(f"✅ Auditoría completada usando el modelo: {modelo_correcto.replace('models/', '')}")
                    st.info(f"**Veredicto de la IA:** {resultado_ia.get('observacion_general', 'Sin comentarios')}")
                    
                except Exception as e:
                    st.error(f"Error al conectar con la IA: {e}")
        # ==========================================
        # ==========================================
        
        def update_df(): st.session_state.df = st.session_state.editor_datos

        ed = st.data_editor(
            st.session_state.df, key="editor_datos", on_change=update_df, use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "Fecha": st.column_config.TextColumn("Fecha (YYYY-MM-DD)", width="small"),
                "Rubro": st.column_config.TextColumn("Rubro", width="large"),
                "QT": st.column_config.SelectboxColumn("QT", options=["Sí", "No"], width="small"),
                "T. Cambio": st.column_config.SelectboxColumn("T. Cambio", options=["MXN", "USD", "EUR"], width="small"),
                "(+ IVA)": st.column_config.SelectboxColumn("(+ IVA)", options=["Sí", "No", "N/M"], width="small"),
                "Cantidad": st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario": st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)": st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%": st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA": st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Diferencia final": st.column_config.NumberColumn("Diferencia", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
            },
        )
        if ed is not None: st.session_state.df = ed
        
        st.markdown("---")
        try:
            xlsx = to_excel(st.session_state.df)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button("⬇️ Descargar Excel", data=xlsx, file_name=f"Conciliacion_{ts}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e: st.error(f"Error generando Excel: {e}")

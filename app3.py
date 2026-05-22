import hashlib, io, re, datetime
from typing import Optional

import fitz
import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st
import xlsxwriter

# ── Página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliador de Cotizaciones",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"]   { background:#1a2744; }
[data-testid="stSidebar"] * { color:#e8e8e8 !important; }
.hdr {
    background:linear-gradient(90deg,#1a2744,#2d4a8f);
    padding:14px 22px; border-radius:8px; margin-bottom:18px;
}
.hdr h1  { color:#fff; font-size:1.4rem; margin:0; font-weight:700; }
.hdr p   { color:#a8c0ff; font-size:.87rem; margin:4px 0 0; }
.ptitle  {
    background:#2d4a8f; color:#fff !important; font-weight:700;
    font-size:.88rem; padding:7px 14px; border-radius:6px 6px 0 0;
    margin-bottom:4px;
}
.kpi {
    background:#f0f4ff; border:1px solid #c5d3f5; border-radius:8px;
    padding:10px 12px; text-align:center; margin-bottom:6px;
}
.kpi .v { font-size:1.2rem; font-weight:700; color:#1a2744; }
.kpi .l { font-size:.72rem; color:#5566aa; }
</style>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────
_MONEY_RE = re.compile(
    r"\$\s*([\d,]+(?:\.\d{1,2})?)"
    r"|(?<!\d)([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?)"
)
_DATE_RE = re.compile(
    r"(\d{1,2})\s+de\s+(\w+)[,\s]+(?:del?\s+)?(\d{4})"
    r"|(\d{4})[/-](\d{1,2})[/-](\d{1,2})"
    r"|(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"
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

# ── Session state ─────────────────────────────────────────────
for k, v in {
    "pdf_bytes":None,"pdf_hash":None,"total_pages":0,
    "current_page":0,"num_sec":1,"sec_cfg":[],
    "df":None,"extracted":False,
}.items():
    st.session_state.setdefault(k, v)

# ── OCR: cargado UNA sola vez, solo cuando se necesita ────────
# ROOT-CAUSE FIX: ya NO se llama en el sidebar (que corre en cada rerun).
# Se invoca únicamente dentro de extract_section cuando hace falta.
@st.cache_resource(show_spinner=False)
def _ocr_engine():
    """Devuelve instancia RapidOCR o None si no está disponible."""
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR()
    except Exception:
        return None

# ── Helpers ───────────────────────────────────────────────────
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
            try: return float(raw.replace(",",""))
            except ValueError: continue
    return None

def _date(text: str) -> Optional[datetime.date]:
    for m in _DATE_RE.finditer(text.lower()):
        g = m.groups()
        try:
            if g[0]: return datetime.date(int(g[2]), _MESES.get(g[1],1), int(g[0]))
            if g[3]: return datetime.date(int(g[3]), int(g[4]), int(g[5]))
            if g[6]:
                yr = int(g[8])
                return datetime.date(yr+2000 if yr<100 else yr, int(g[7]), int(g[6]))
        except (ValueError,KeyError): continue
    return None

def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

@st.cache_data(max_entries=50, show_spinner=False)
def _render(pdf_bytes: bytes, idx: int) -> bytes:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(1.5,1.5),
                                  colorspace=fitz.csRGB, alpha=False)
    return pix.tobytes("png")

def _ocr_page(pdf_bytes: bytes, idx: int) -> str:
    ocr = _ocr_engine()
    if ocr is None:
        return ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            pix = doc[idx].get_pixmap(matrix=fitz.Matrix(2.0,2.0),
                                      colorspace=fitz.csRGB, alpha=False)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.height,pix.width,3)[:,:,::-1]
        res, _ = ocr(img)
        return "\n".join(r[1] for r in res if r and len(r)>1) if res else ""
    except Exception:
        return ""

# ── Motor de extracción ───────────────────────────────────────
def extract(pdf_bytes: bytes, label: str,
            p0: int, p1: int,
            det_iva: bool, calc_sub: bool) -> dict:
    text, rows = "", []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n = len(pdf.pages)
        pr = range(p0-1, min(p1, n))
        for i in pr:
            text += "\n" + (pdf.pages[i].extract_text() or "")
            for tbl in pdf.pages[i].extract_tables():
                if tbl:
                    rows.extend([str(c or "").strip() for c in r] for r in tbl if r)

    if len(text.strip()) < 30:
        for i in pr:
            text += "\n" + _ocr_page(pdf_bytes, i)

    fecha = _date(text)
    tlow  = text.lower()
    iva_f = "Sí" if det_iva and re.search(r"\biva\b|16%|vat", tlow) else "N/M"

    tot = iva = sub = pu = None
    qty = 1

    for row in rows:
        j, s = "  ".join(row).lower(), "  ".join(row)
        if "total" in j and not re.search(r"sub|parcial", j):
            v = _money(s)
            if v and (tot is None or v > tot): tot = v
        if re.search(r"\biva\b|16%|vat", j):
            v = _money(s)
            if v: iva = v
        if re.search(r"subtotal|sin\s*iva", j):
            v = _money(s)
            if v: sub = v

    if tot is None:
        for ln in text.splitlines():
            if re.search(r"\btotal\b",ln,re.I) and not re.search(r"sub|parcial",ln,re.I):
                v = _money(ln)
                if v and (tot is None or v > tot): tot = v
    if iva is None and iva_f == "Sí":
        for ln in text.splitlines():
            if re.search(r"\biva\b|16%|vat",ln,re.I):
                v = _money(ln)
                if v: iva = v; break

    for row in rows:
        nf = []
        for t in re.findall(r"[\d,]+(?:\.\d+)?", "  ".join(row)):
            try: nf.append(float(t.replace(",","")))
            except ValueError: continue
        if len(nf) >= 2 and 1 <= nf[0] <= 999 and nf[-1] > 0:
            qty = int(nf[0]); pu = nf[-2] if len(nf)>2 else nf[-1]

    obs = ""
    if tot and not sub and not iva and iva_f == "Sí":
        sub = round(tot/1.16,2); iva = round(tot-sub,2); obs = "IVA incluido en precio"
    elif tot and iva and not sub:  sub = round(tot-iva,2)
    elif sub and iva and not tot:  tot = round(sub+iva,2)
    elif calc_sub and sub and not iva and not tot and iva_f == "Sí":
        iva = round(sub*0.16,2); tot = round(sub+iva,2)
    if pu is None and sub: pu = round(sub/qty,2) if qty else sub

    return {
        "Fecha":                  fecha.isoformat() if fecha else datetime.date.today().isoformat(),
        "Rubro":                  label,
        "QT":                     "Sí",
        "T. Cambio":              "MXN",
        "(+ IVA)":                iva_f,
        "Cantidad":               qty,
        "Precio Unitario":        pu,
        "Subtotal (Sin IVA)":     sub,
        "IVA 16%":                iva,
        "Total con IVA":          tot,
        "Diferencia final":       None,
        "Monto en Anexo Escrito": None,
        "Observaciones":          obs,
    }

# ── Excel ─────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory":True,"nan_inf_to_errors":True})
    ws  = wb.add_worksheet("Conciliación")
    MF  = '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'
    B   = {"font_name":"Calibri","font_size":11,"border":1}
    hf  = wb.add_format({**B,"bold":True,"bg_color":"#D4C19C","align":"center","valign":"vcenter"})
    h2f = wb.add_format({**B,"bold":True,"bg_color":"#EBE2D1","align":"center","valign":"vcenter"})
    df_ = wb.add_format({**B,"num_format":"dd/mm/yy"})
    mf  = wb.add_format({**B,"num_format":MF})
    tf  = wb.add_format({**B})
    nf  = wb.add_format({**B,"num_format":"#,##0"})
    tot = wb.add_format({**B,"bold":True,"num_format":MF})
    of  = wb.add_format({**B,"text_wrap":True})

    for c,h in enumerate(_COLS): ws.write(0,c,h, h2f if c>=11 else hf)
    for i,w in enumerate(_WIDTHS): ws.set_column(i,i,w)
    ws.set_row(0,30)

    MC = {
        "Precio Unitario":(6,None),"Subtotal (Sin IVA)":(7,"=F{r}*G{r}"),
        "IVA 16%":(8,"=H{r}*0.16"),"Total con IVA":(9,"=H{r}+I{r}"),
        "Monto en Anexo Escrito":(11,None),
    }
    D,n = 1,len(df)
    for i,(_,row) in enumerate(df.iterrows()):
        r0,r1 = D+i,D+i+1
        fecha = row.get("Fecha")
        if isinstance(fecha,str):
            try: fecha = datetime.date.fromisoformat(fecha)
            except: fecha = None
        if isinstance(fecha,(datetime.date,datetime.datetime)):
            dt = datetime.datetime.combine(fecha,datetime.time()) if isinstance(fecha,datetime.date) else fecha
            ws.write_datetime(r0,0,dt,df_)
        else:
            ws.write(r0,0,str(fecha or ""),df_)
        ws.write(r0,1,str(row.get("Rubro","") or ""),of)
        ws.write(r0,2,str(row.get("QT","Sí")),tf)
        ws.write(r0,3,str(row.get("T. Cambio","MXN")),tf)
        ws.write(r0,4,str(row.get("(+ IVA)","") or ""),tf)
        q = _safe_f(row.get("Cantidad"))
        ws.write(r0,5,int(q) if q is not None else 1,nf)
        for cn,(xi,fm) in MC.items():
            v = _safe_f(row.get(cn))
            if v is not None: ws.write_number(r0,xi,v,mf)
            elif fm: ws.write_formula(r0,xi,fm.format(r=r1),mf)
            else: ws.write_blank(r0,xi,mf)
        ws.write_formula(r0,10,f"=J{r1}-L{r1}",mf)
        ws.write(r0,12,str(row.get("Observaciones","") or ""),of)

    s1,e1 = D+1,D+n
    for ci in (9,10,11):
        cl = chr(ord("A")+ci)
        ws.write_formula(D+n,ci,f"=SUM({cl}{s1}:{cl}{e1})",tot)
    wb.close(); buf.seek(0); return buf.read()

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Documento PDF")
    up = st.file_uploader("Sube el archivo PDF", type=["pdf"])
    if up:
        raw = up.read()
        h   = _md5(raw)
        if h != st.session_state.pdf_hash:
            st.session_state.update(
                pdf_bytes=raw, pdf_hash=h, extracted=False,
                df=None, current_page=0,
            )
            with fitz.open(stream=raw, filetype="pdf") as d:
                st.session_state.total_pages = len(d)
        st.success(f"✅ {st.session_state.total_pages} páginas")

    st.markdown("---")
    st.markdown("### ⚙️ Secciones")
    n = int(st.number_input("Número de secciones", min_value=1, max_value=50,
                             value=st.session_state.num_sec, step=1))
    if n != st.session_state.num_sec:
        st.session_state.num_sec  = n
        st.session_state.extracted = False
        st.session_state.df       = None

    cfgs = st.session_state.sec_cfg
    tp   = st.session_state.total_pages or 1
    while len(cfgs) < n:
        i = len(cfgs)+1
        cfgs.append({"label":f"Sección {i}","p0":i,"p1":i,"det_iva":True,"calc_sub":True})
    del cfgs[n:]

    for i,c in enumerate(cfgs):
        with st.expander(f"📄 Sección {i+1}", expanded=(n<=5)):
            c["label"]    = st.text_input("Rubro/Concepto", value=c["label"], key=f"lb{i}")
            a,b           = st.columns(2)
            c["p0"]       = a.number_input("Pág. Inicio", 1, tp, min(c["p0"],tp), key=f"p0{i}")
            c["p1"]       = b.number_input("Pág. Fin", c["p0"], tp, max(min(c["p1"],tp),c["p0"]), key=f"p1{i}")
            c["det_iva"]  = st.checkbox("Detectar IVA", value=c["det_iva"],  key=f"iv{i}")
            c["calc_sub"] = st.checkbox("Calcular subtotal si falta", value=c["calc_sub"], key=f"cs{i}")

    st.markdown("---")
    run = st.button("🔍 Extraer Montos",
                    disabled=(st.session_state.pdf_bytes is None),
                    use_container_width=True, type="primary")

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
  <h1>📋 Conciliador de Cotizaciones PDF</h1>
  <p>Extrae, edita y exporta montos desde documentos PDF</p>
</div>
""", unsafe_allow_html=True)

# ── Extracción ────────────────────────────────────────────────
if run and st.session_state.pdf_bytes:
    rows = []
    bar  = st.progress(0, text="Procesando…")
    for i,c in enumerate(st.session_state.sec_cfg):
        bar.progress((i+.5)/n, text=f"Extrayendo: {c['label']}")
        try:
            rows.append(extract(st.session_state.pdf_bytes, c["label"],
                                c["p0"],c["p1"],c["det_iva"],c["calc_sub"]))
        except Exception as e:
            st.warning(f"⚠️ Sección {i+1}: {str(e)[:80]}")
            rows.append({k:None for k in _COLS}|{
                "Rubro":c["label"],"QT":"Sí","T. Cambio":"MXN","Cantidad":1,
                "Fecha":datetime.date.today().isoformat(),
                "Observaciones":f"Error: {str(e)[:60]}",
            })
        bar.progress((i+1)/n)
    bar.empty()
    df = pd.DataFrame(rows, columns=_COLS)
    df["Diferencia final"] = (
        df["Total con IVA"].fillna(0)-df["Monto en Anexo Escrito"].fillna(0)
    ).where(df["Total con IVA"].notna()|df["Monto en Anexo Escrito"].notna())
    st.session_state.df        = df
    st.session_state.extracted = True
    st.success(f"✅ {len(rows)} sección(es) procesada(s).")

if st.session_state.pdf_bytes is None:
    st.info("👈 Sube un PDF en la barra lateral para comenzar.")
    st.stop()

# ── Layout ────────────────────────────────────────────────────
L, R = st.columns(2, gap="medium")

with L:
    st.markdown('<p class="ptitle">🔍 Visor de Documento</p>', unsafe_allow_html=True)
    tp = st.session_state.total_pages
    cp = st.session_state.current_page
    c1,c2,c3 = st.columns([1,4,1])
    if c1.button("◀", key="pv"): cp = max(0,cp-1)
    cp = int(c2.number_input("",1,tp,cp+1,label_visibility="collapsed",key="pg"))-1
    if c3.button("▶", key="nx"): cp = min(tp-1,cp+1)
    st.session_state.current_page = cp
    st.caption(f"Página **{cp+1}** de **{tp}**")
    for c in st.session_state.sec_cfg:
        if c["p0"] <= cp+1 <= c["p1"]:
            st.markdown(f'<span style="background:#2d4a8f;color:#fff;padding:3px 10px;'
                        f'border-radius:4px;font-size:.8rem">📑 {c["label"]}</span>',
                        unsafe_allow_html=True); break
    with st.spinner("Cargando…"):
        st.image(_render(st.session_state.pdf_bytes, cp), use_container_width=True)
    st.download_button("📥 Descargar PDF", data=st.session_state.pdf_bytes,
                       file_name="documento.pdf", mime="application/pdf",
                       use_container_width=True)

with R:
    st.markdown('<p class="ptitle">✏️ Editor de Datos</p>', unsafe_allow_html=True)
    if not st.session_state.extracted or st.session_state.df is None:
        st.info("Configura las secciones y presiona **🔍 Extraer Montos**.")
    else:
        df = st.session_state.df
        ts  = _safe_f(df["Total con IVA"].sum(skipna=True)) or 0.0
        rs  = _safe_f(df["Monto en Anexo Escrito"].sum(skipna=True)) or 0.0
        dif = ts - rs
        k1,k2,k3 = st.columns(3)
        for col,val,lbl in [(k1,ts,"Total Extraído"),(k2,rs,"Monto Referencia"),(k3,dif,"Diferencia")]:
            clr = "#1a2744" if lbl!="Diferencia" else ("#c0392b" if abs(dif)>.01 else "#28a745")
            col.markdown(f'<div class="kpi"><div class="v" style="color:{clr}">'
                         f'${val:,.2f}</div><div class="l">{lbl}</div></div>',
                         unsafe_allow_html=True)

        ed = st.data_editor(
            st.session_state.df,
            key="editor_datos",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Fecha":                  st.column_config.TextColumn("Fecha (YYYY-MM-DD)", width="small"),
                "Rubro":                  st.column_config.TextColumn("Rubro", width="large"),
                "QT":                     st.column_config.SelectboxColumn("QT", options=["Sí","No"], width="small"),
                "T. Cambio":              st.column_config.SelectboxColumn("T. Cambio", options=["MXN","USD","EUR"], width="small"),
                "(+ IVA)":               st.column_config.SelectboxColumn("(+ IVA)", options=["Sí","No","N/M"], width="small"),
                "Cantidad":               st.column_config.NumberColumn("Cant.", format="%d", width="small"),
                "Precio Unitario":        st.column_config.NumberColumn("P. Unit.", format="$%.2f", width="medium"),
                "Subtotal (Sin IVA)":     st.column_config.NumberColumn("Subtotal", format="$%.2f", width="medium"),
                "IVA 16%":               st.column_config.NumberColumn("IVA 16%", format="$%.2f", width="medium"),
                "Total con IVA":          st.column_config.NumberColumn("Total", format="$%.2f", width="medium"),
                "Diferencia final":       st.column_config.NumberColumn("Diferencia", format="$%.2f", width="medium"),
                "Monto en Anexo Escrito": st.column_config.NumberColumn("Ref. Escrito", format="$%.2f", width="medium"),
                "Observaciones":          st.column_config.TextColumn("Observaciones", width="large"),
            },
        )
        if ed is not None:
            st.session_state.df = ed

        st.markdown("---")
        try:
            xlsx = to_excel(st.session_state.df)
            ts_  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx,
                file_name=f"Conciliacion_{ts_}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Error generando Excel: {e}")
# ============================================================
# CONCILIADOR DE COTIZACIONES  |  main_qt.py  v6.6
# ============================================================
from __future__ import annotations

import datetime
import os
import sys

import pandas as pd
from PyQt6.QtCore import (
    QByteArray, QSize, Qt, QThread, pyqtSignal
)
from PyQt6.QtGui import (
    QColor, QPixmap, QTransform
)
from PyQt6.QtWidgets import (
    QApplication, QComboBox,
    QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QListWidget, QMainWindow, QMessageBox, QProgressBar,
    QPushButton, QScrollArea, QSpinBox,
    QSplitter, QStatusBar, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
)

import engine

# ─────────────────────────────────────────────────────────────
# PALETA DE COLORES
# ─────────────────────────────────────────────────────────────
ACCENT      = "#FF4B4B"  # Streamlit Primary Red
ACCENT_LITE = "#ff7676"
BANNER_BG   = "#B84A6B"  # Magenta headers from screenshot
BG_DARK     = "#0E1117"  # Streamlit main background
BG_PANEL    = "#0E1117"  # Streamlit main background
BG_SIDEBAR  = "#262730"  # Streamlit secondary background
TEXT_MAIN   = "#FAFAFA"
TEXT_MUTED  = "#A3A8B8"
BORDER      = "#707482"  # Lighter border to stand out
INPUT_BG    = "#262730"
GREEN       = "#09ab3b"
RED_WARN    = "#ff2b2b"

QSS = f"""
QMainWindow, QWidget {{
    background-color: {BG_DARK};
    color: {TEXT_MAIN};
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
}}
QSplitter::handle {{
    background-color: {BORDER};
    width: 2px; height: 2px;
}}
/* ── Dock lateral ── */
#sidebar {{
    background-color: {BG_SIDEBAR};
    border-right: 1px solid {BORDER};
}}
QWidget#sidebar {{
    background-color: {BG_SIDEBAR};
    border-right: 1px solid {BORDER};
}}
QWidget#main_banner {{
    background-color: {BANNER_BG};
    border-radius: 6px;
}}
QWidget#main_banner QLabel {{
    background-color: transparent;
}}
/* ── Paneles interiores ── */
QGroupBox {{
    border: 1px solid {BORDER};
    border-radius: 6px;
    margin-top: 14px;
    padding-top: 10px;
    color: {TEXT_MAIN};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 10px;
    top: 0px;
    color: {TEXT_MUTED};
    font-size: 11px;
    font-weight: bold;
    text-transform: uppercase;
}}
/* ── Inputs ── */
QLineEdit, QSpinBox, QComboBox, QTextEdit {{
    background-color: {INPUT_BG};
    color: {TEXT_MAIN};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 4px 8px;
    min-height: 22px;
}}
QSpinBox::up-button, QSpinBox::down-button {{
    background-color: {TEXT_MUTED};
    width: 16px;
    border-radius: 2px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
    background-color: {TEXT_MAIN};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
    border: 1px solid {ACCENT_LITE};
}}
QComboBox QAbstractItemView {{
    background-color: {BG_PANEL};
    color: {TEXT_MAIN};
    border: 1px solid {BORDER};
    selection-background-color: {ACCENT};
}}
/* ── Botones ── */
QPushButton {{
    background-color: #4A4D59;
    color: {TEXT_MAIN};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 6px 14px;
    font-weight: 500;
    min-height: 24px;
}}
QPushButton:hover  {{ border: 1px solid {ACCENT}; color: {ACCENT}; background-color: #5C5F6D; }}
QPushButton:pressed {{ background-color: {ACCENT}; color: #fff; }}

QPushButton#expander_btn {{
    text-align: left;
    background-color: transparent;
    border: none;
    padding: 8px 10px;
    font-weight: 600;
    font-size: 13px;
    color: {TEXT_MAIN};
    border-radius: 4px;
}}
QPushButton#expander_btn:hover {{
    color: {ACCENT};
    background-color: rgba(255, 75, 75, 0.1);
}}
QWidget#expander_content {{
    border-top: 1px solid {BORDER};
    background-color: transparent;
}}
QWidget#section_widget {{
    border: 1px solid {BORDER};
    border-radius: 6px;
    background-color: {INPUT_BG};
    margin-bottom: 8px;
}}

QPushButton#viewer_btn {{
    background-color: #3C3F4D;
    color: #FAFAFA;
    border: 1px solid #707482;
    border-radius: 4px;
    padding: 2px;
    font-size: 14px;
    font-weight: bold;
    min-width: 30px;
    min-height: 24px;
}}
QPushButton#viewer_btn:hover {{
    background-color: #5C5F6D;
    border: 1px solid #FF4B4B;
    color: #FF4B4B;
}}

QPushButton#btn_primary {{
    background-color: {ACCENT};
    color: #ffffff;
    font-weight: bold;
    border: none;
    padding: 8px 18px;
}}
QPushButton#btn_primary:hover  {{ background-color: {ACCENT_LITE}; border: none; color: #ffffff; }}
QPushButton#btn_primary:pressed {{ background-color: {ACCENT}; }}
QPushButton#btn_primary:disabled {{
    background-color: #5C2828;
    color: #A3A8B8;
}}
/* ── Tabla ── */
QTableWidget {{
    background-color: {BG_PANEL};
    color: {TEXT_MAIN};
    gridline-color: transparent;
    border: 1px solid {BORDER};
    selection-background-color: {ACCENT};
}}
QTableWidget::item {{
    border-bottom: 1px solid {BORDER};
}}
QTableWidget::item:selected {{ background-color: {ACCENT}; color: #fff; }}
QHeaderView::section {{
    background-color: {BG_DARK};
    color: {TEXT_MUTED};
    border: none;
    border-bottom: 1px solid {BORDER};
    padding: 4px 6px;
    font-size: 11px;
    font-weight: bold;
    text-transform: uppercase;
}}
/* ── Barra de progreso ── */
QProgressBar {{
    background-color: {INPUT_BG};
    border: 1px solid {BORDER};
    border-radius: 4px;
    color: {TEXT_MAIN};
    text-align: center;
    height: 20px;
}}
QProgressBar::chunk {{ background-color: {ACCENT}; border-radius: 3px; }}
/* ── Scroll ── */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 0px;
}}
QScrollBar::handle:vertical {{
    background: #555555;
    min-height: 20px;
    border-radius: 5px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: none;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 0px;
}}
QScrollBar::handle:horizontal {{
    background: #555555;
    min-width: 20px;
    border-radius: 5px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0px;
}}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: none;
}}
/* ── Títulos de Sección ── */
QLabel#section_title {{
    background-color: {BANNER_BG};
    color: #ffffff;
    font-weight: bold;
    padding: 8px 12px;
    border-radius: 6px;
}}
QLabel#kpi_value {{
    font-size: 18px;
    font-weight: bold;
    color: {ACCENT_LITE};
}}
QLabel#kpi_label {{
    font-size: 11px;
    color: {TEXT_MUTED};
}}
/* ── Checkbox ── */
QCheckBox {{ color: {TEXT_MAIN}; spacing: 6px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {BORDER};
    border-radius: 3px;
    background-color: {INPUT_BG};
}}
QCheckBox::indicator:checked {{
    background-color: {ACCENT};
    border-color: {ACCENT};
}}
/* ── Menú ── */
QMenuBar {{
    background-color: {BG_SIDEBAR};
    color: {TEXT_MAIN};
    border-bottom: 1px solid {BORDER};
}}
QMenuBar::item:selected {{ background-color: {ACCENT}; }}
QMenu {{
    background-color: {BG_PANEL};
    color: {TEXT_MAIN};
    border: 1px solid {BORDER};
}}
QMenu::item:selected {{ background-color: {ACCENT}; }}
/* ── Status bar ── */
QStatusBar {{ background-color: {BG_SIDEBAR}; color: {TEXT_MUTED}; font-size: 11px; }}
"""


# ─────────────────────────────────────────────────────────────
# WORKERS (hilos de fondo)
# ─────────────────────────────────────────────────────────────

class NoScrollComboBox(QComboBox):
    def wheelEvent(self, e):
        e.ignore()

class NoScrollSpinBox(QSpinBox):
    def wheelEvent(self, e):
        e.ignore()

class RenderWorker(QThread):
    """Renderiza una página del PDF en segundo plano."""
    finished = pyqtSignal(int, QPixmap)   # (page_idx, pixmap)
    failed   = pyqtSignal(int)            # (page_idx)

    def __init__(self, pdf_bytes: bytes, idx: int, scale: float = 1.5):
        super().__init__()
        self._bytes = pdf_bytes
        self._idx   = idx
        self._scale = scale

    def run(self):
        png = engine.render_page(self._bytes, self._idx, self._scale)
        if png:
            pm = QPixmap()
            pm.loadFromData(QByteArray(png))
            self.finished.emit(self._idx, pm)
        else:
            self.failed.emit(self._idx)


class ExtractWorker(QThread):
    """Ejecuta la extracción de montos en segundo plano."""
    progress = pyqtSignal(int, int, str)   # (current, total, label)
    row_done = pyqtSignal(int, dict)       # (section_idx, result_dict)
    warning  = pyqtSignal(str)
    finished = pyqtSignal(list)            # lista completa de dicts

    def __init__(self, pdf_files: list, sec_cfgs: list, bx_token: str):
        super().__init__()
        self._files  = pdf_files
        self._cfgs   = sec_cfgs
        self._token  = bx_token

    def run(self):
        results = []
        n = len(self._cfgs)
        for i, cfg in enumerate(self._cfgs):
            self.progress.emit(i + 1, n, cfg["label"])
            pdf_idx = cfg.get("pdf_idx", 0)
            if pdf_idx < len(self._files):
                pdf_bytes = self._files[pdf_idx]["bytes"]
            else:
                pdf_bytes = self._files[0]["bytes"] if self._files else b""
            try:
                row = engine.extract(
                    pdf_bytes,
                    cfg["label"],
                    cfg["p0"],
                    cfg["p1"],
                    cfg["det_iva"],
                    cfg.get("calc_sub", True),
                    tipo=cfg.get("tipo", "Cotización Proveedor"),
                    moneda=cfg.get("moneda", "AUTO"),
                    bx_token=self._token,
                    sec_num=i + 1,
                    on_warning=lambda m: self.warning.emit(m),
                )
            except Exception as exc:
                row = {
                    **{k: None for k in engine._COLS},
                    "Nº Sec": i + 1,
                    "Tipo":   cfg.get("tipo", "Cotización Proveedor"),
                    "Rubro":  cfg["label"],
                    "QT":     "Sí",
                    "T. Cambio": cfg.get("moneda", "MXN"),
                    "Cantidad":  1,
                    "Fecha":  datetime.date.today().isoformat(),
                    "Observaciones": f"Error: {str(exc)[:120]}",
                }
                self.warning.emit(f"Sección {i+1} «{cfg['label']}»: {exc}")
            results.append(row)
            self.row_done.emit(i, row)
        self.finished.emit(results)


# ─────────────────────────────────────────────────────────────
# WIDGET DE SECCIÓN (barra lateral)
# ─────────────────────────────────────────────────────────────
class SectionWidget(QWidget):
    changed = pyqtSignal()
    delete_requested = pyqtSignal(object)

    def __init__(self, idx: int, pdf_names: list[str], max_pages: list[int], parent=None):
        super().__init__(parent)
        self.setObjectName("section_widget")
        self._idx = idx
        self._pdf_names = pdf_names
        self._max_pages = max_pages
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Botón Colapsable y Eliminar
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(4)
        
        self.btn_toggle = QPushButton(f"  >   📄 Sección {self._idx + 1}")
        self.btn_toggle.setCheckable(True)
        self.btn_toggle.setChecked(False) # Colapsado por defecto
        self.btn_toggle.setObjectName("expander_btn")
        self.btn_toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle.clicked.connect(self._toggle)
        
        self.btn_delete = QPushButton("🗑️")
        self.btn_delete.setFixedWidth(32)
        self.btn_delete.setObjectName("viewer_btn")
        self.btn_delete.clicked.connect(lambda: self.delete_requested.emit(self))
        self.btn_delete.setToolTip("Eliminar sección")
        
        top_row.addWidget(self.btn_toggle, stretch=1)
        top_row.addWidget(self.btn_delete)
        layout.addLayout(top_row)
        
        # Contenedor de Formulario
        self.container = QWidget()
        self.container.setObjectName("expander_content")
        self.container.setVisible(False)
        clayout = QVBoxLayout(self.container)
        clayout.setContentsMargins(14, 14, 14, 14)
        clayout.setSpacing(12)

        lbl_style = "font-size: 11px; font-weight: bold; color: #A3A8B8;"

        # Rubro
        lbl_rubro = QLabel("Rubro / Concepto")
        lbl_rubro.setStyleSheet(lbl_style)
        self.le_label = QLineEdit(f"Sección {self._idx + 1}")
        clayout.addWidget(lbl_rubro)
        clayout.addWidget(self.le_label)

        # Tipo
        lbl_tipo = QLabel("Tipo")
        lbl_tipo.setStyleSheet(lbl_style)
        self.cb_tipo = NoScrollComboBox()
        self.cb_tipo.addItems(["Cotización Proveedor", "Presupuesto Global"])
        clayout.addWidget(lbl_tipo)
        clayout.addWidget(self.cb_tipo)

        # PDF
        self.lbl_pdf = QLabel("PDF")
        self.lbl_pdf.setStyleSheet(lbl_style)
        self.cb_pdf = NoScrollComboBox()
        if self._pdf_names:
            self.cb_pdf.addItems(self._pdf_names)
        else:
            self.cb_pdf.addItem("(sin PDF)")
        clayout.addWidget(self.lbl_pdf)
        clayout.addWidget(self.cb_pdf)
        has_multi = len(self._pdf_names) > 1
        self.lbl_pdf.setVisible(has_multi)
        self.cb_pdf.setVisible(has_multi)

        # Páginas
        h_pages = QHBoxLayout()
        v_p0 = QVBoxLayout(); v_p0.setContentsMargins(0,0,0,0); v_p0.setSpacing(4)
        v_p1 = QVBoxLayout(); v_p1.setContentsMargins(0,0,0,0); v_p1.setSpacing(4)
        lbl_p0 = QLabel("Pág. Inicio"); lbl_p0.setStyleSheet(lbl_style)
        lbl_p1 = QLabel("Pág. Fin");    lbl_p1.setStyleSheet(lbl_style)
        self.sb_p0 = NoScrollSpinBox(); self.sb_p0.setMinimum(1)
        self.sb_p1 = NoScrollSpinBox(); self.sb_p1.setMinimum(1)
        self._update_max_pages()
        v_p0.addWidget(lbl_p0); v_p0.addWidget(self.sb_p0)
        v_p1.addWidget(lbl_p1); v_p1.addWidget(self.sb_p1)
        h_pages.addLayout(v_p0); h_pages.addLayout(v_p1)
        clayout.addLayout(h_pages)

        layout.addWidget(self.container)

        # Propagar cambios
        for w in (self.le_label, self.cb_tipo, self.cb_pdf):
            if hasattr(w, "textChanged"):
                w.textChanged.connect(self.changed)
            elif hasattr(w, "currentIndexChanged"):
                w.currentIndexChanged.connect(self.changed)
        for w in (self.sb_p0, self.sb_p1):
            w.valueChanged.connect(self.changed)
        self.cb_pdf.currentIndexChanged.connect(self._on_pdf_changed)

    def _toggle(self, checked):
        if checked:
            self.btn_toggle.setText(f"  v   📄 Sección {self._idx + 1}")
            self.container.setVisible(True)
        else:
            self.btn_toggle.setText(f"  >   📄 Sección {self._idx + 1}")
            self.container.setVisible(False)

    def _update_max_pages(self):
        pdf_idx = self.cb_pdf.currentIndex() if self._pdf_names else 0
        mp = self._max_pages[pdf_idx] if pdf_idx < len(self._max_pages) else 1
        self.sb_p0.setMaximum(mp)
        self.sb_p1.setMaximum(mp)

    def _on_pdf_changed(self):
        self._update_max_pages()
        self.changed.emit()

    def get_config(self) -> dict:
        return {
            "label":    self.le_label.text() or f"Sección {self._idx + 1}",
            "tipo":     self.cb_tipo.currentText(),
            "pdf_idx":  self.cb_pdf.currentIndex(),
            "p0":       self.sb_p0.value(),
            "p1":       self.sb_p1.value(),
            "det_iva":  True,
            "calc_sub": True,
        }

    def set_config(self, cfg: dict):
        if "label" in cfg:
            self.le_label.setText(cfg["label"])
        if "tipo" in cfg:
            self.cb_tipo.setCurrentText(cfg["tipo"])
        if "pdf_idx" in cfg and cfg["pdf_idx"] < self.cb_pdf.count():
            self.cb_pdf.setCurrentIndex(cfg["pdf_idx"])
        if "p0" in cfg:
            self.sb_p0.setValue(cfg["p0"])
        if "p1" in cfg:
            self.sb_p1.setValue(cfg["p1"])

    def update_pdf_names(self, pdf_names: list[str], max_pages: list[int]):
        self._pdf_names = pdf_names
        self._max_pages = max_pages
        cur = self.cb_pdf.currentIndex()
        self.cb_pdf.blockSignals(True)
        self.cb_pdf.clear()
        self.cb_pdf.addItems(pdf_names if pdf_names else ["(sin PDF)"])
        if cur < self.cb_pdf.count():
            self.cb_pdf.setCurrentIndex(cur)
        self.cb_pdf.blockSignals(False)
        self.cb_pdf.setVisible(len(pdf_names) > 1)
        self._update_max_pages()


# ─────────────────────────────────────────────────────────────
# PANEL LATERAL DE CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
class ConfigPanel(QWidget):
    extract_requested = pyqtSignal(list, str)  # (sec_configs, bx_token)
    pdf_loaded        = pyqtSignal()
    token_update_requested = pyqtSignal(str)
    section_deleted   = pyqtSignal(str)        # Emite el label de la sección eliminada

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(280)
        self._pdf_files: list[dict]  = []
        self._sec_widgets: list[SectionWidget] = []
        self._build()

    def _build(self):
        # ── Panel exterior: solo contiene el scroll area del menú ──
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._panel_scroll = QScrollArea()
        self._panel_scroll.setWidgetResizable(True)
        self._panel_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        # No scroll horizontal: el panel tiene ancho fijo
        self._panel_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self._panel_scroll.setStyleSheet(
            f"QScrollArea {{ background:{BG_SIDEBAR}; border:none; }}"
        )

        _content = QWidget()
        _content.setObjectName("sidebar_content")
        _content.setStyleSheet(f"QWidget#sidebar_content {{ background:{BG_SIDEBAR}; }}")
        main = QVBoxLayout(_content)
        main.setContentsMargins(10, 10, 10, 10)
        main.setSpacing(12)

        # ── PDFs ─────────────────────────────────────────────────────
        grp_pdf = QGroupBox("Documentos PDF")
        vl = QVBoxLayout(grp_pdf)
        self.btn_add_pdf = QPushButton("📂  Agregar PDF(s)")
        self.btn_add_pdf.clicked.connect(self._open_pdfs)
        vl.addWidget(self.btn_add_pdf)

        self.lbl_pdf_info = QLabel("Sin archivos cargados")
        self.lbl_pdf_info.setWordWrap(True)
        self.lbl_pdf_info.setStyleSheet(f"color:{TEXT_MUTED}; font-size:11px;")
        vl.addWidget(self.lbl_pdf_info)

        # Lista compacta de archivos (máx 3 filas visibles, scrolleable)
        self.lst_pdf_files = QListWidget()
        self.lst_pdf_files.setMaximumHeight(72)
        self.lst_pdf_files.setStyleSheet(f"""
            QListWidget {{
                background:{INPUT_BG};
                border:1px solid {BORDER};
                color:{TEXT_MUTED};
                font-size:11px;
                border-radius:4px;
            }}
            QListWidget::item {{ padding:2px 6px; }}
            QListWidget::item:selected {{ background:{ACCENT}; color:#fff; }}
        """)
        self.lst_pdf_files.setVisible(False)
        vl.addWidget(self.lst_pdf_files)

        self.btn_clear_pdf = QPushButton("🗑  Limpiar PDFs")
        self.btn_clear_pdf.clicked.connect(self._clear_pdfs)
        self.btn_clear_pdf.setEnabled(False)
        vl.addWidget(self.btn_clear_pdf)
        main.addWidget(grp_pdf)

        # ── Proyecto ──────────────────────────────────────────
        grp_proy = QGroupBox("Proyecto")
        fl = QFormLayout(grp_proy)
        self.le_proyecto = QLineEdit()
        self.le_proyecto.setPlaceholderText("Nombre del proyecto…")
        fl.addRow("Nombre:", self.le_proyecto)
        main.addWidget(grp_proy)

        # ── Banxico ───────────────────────────────────────────
        grp_bx = QGroupBox("Banxico – Tipo de Cambio")
        fl2 = QFormLayout(grp_bx)
        self.le_token = QLineEdit()
        self.le_token.setPlaceholderText("Pega tu token aquí.")
        self.le_token.setEchoMode(QLineEdit.EchoMode.Password)
        
        token_layout = QHBoxLayout()
        token_layout.addWidget(self.le_token)
        self.btn_update_token = QPushButton("Aplicar")
        self.btn_update_token.setToolTip("Aplicar token a la tabla de datos")
        self.btn_update_token.setFixedHeight(24)
        self.btn_update_token.setStyleSheet("font-size: 11px; font-weight: bold; background: transparent; border: 1px solid white; border-radius: 4px; color: white; padding: 0 6px;")
        self.btn_update_token.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_update_token.clicked.connect(self._on_update_token)
        token_layout.addWidget(self.btn_update_token)
        
        fl2.addRow("Token:", token_layout)
        self.lbl_token_status = QLabel("")
        self.lbl_token_status.setStyleSheet(f"color:{TEXT_MUTED}; font-size:11px;")
        fl2.addRow("", self.lbl_token_status)
        
        self.lbl_token_link = QLabel('<a href="https://www.banxico.org.mx/SieAPIRest/service/v1/token" style="color: #63b3ed; text-decoration: none;">Obtener Token en Banxico</a>')
        self.lbl_token_link.setOpenExternalLinks(True)
        self.lbl_token_link.setStyleSheet("font-size: 11px;")
        fl2.addRow("", self.lbl_token_link)
        
        self.le_token.textChanged.connect(
            lambda t: self.lbl_token_status.setText("🔑 Token activo" if t else "")
        )
        main.addWidget(grp_bx)

        # ── Secciones ─────────────────────────────────────────
        grp_sec = QGroupBox("Secciones (Cotizaciones)")
        vl_sec = QVBoxLayout(grp_sec)
        
        lbl_nsec = QLabel("Número de secciones")
        lbl_nsec.setStyleSheet("font-size: 11px; font-weight: bold; color: #A3A8B8;")
        vl_sec.addWidget(lbl_nsec)

        self.sb_nsec = NoScrollSpinBox()
        self.sb_nsec.setRange(1, 50)
        self.sb_nsec.setValue(1)
        self.sb_nsec.setMinimumWidth(64)
        self.sb_nsec.valueChanged.connect(self._rebuild_sections)
        vl_sec.addWidget(self.sb_nsec)

        # Sin scroll anidado — el panel completo ya scrollea verticalmente
        self._sec_container = QWidget()
        self._sec_layout    = QVBoxLayout(self._sec_container)
        self._sec_layout.setContentsMargins(0, 0, 0, 0)
        self._sec_layout.setSpacing(8)
        self._sec_layout.addStretch()
        vl_sec.addWidget(self._sec_container)
        main.addWidget(grp_sec)

        # ── Botón Extraer ───────────────────────────────────────
        self.btn_extract = QPushButton("🔍  Extraer Montos")
        self.btn_extract.setObjectName("btn_primary")
        self.btn_extract.setEnabled(False)
        self.btn_extract.clicked.connect(self._on_extract)
        main.addWidget(self.btn_extract)
        main.addStretch()

        self._panel_scroll.setWidget(_content)
        outer.addWidget(self._panel_scroll)

        # Construir sección inicial
        self._rebuild_sections(1)

    def _pdf_names(self)  -> list[str]:  return [f["name"] for f in self._pdf_files]
    def _max_pages(self) -> list[int]:   return [f["pages"] for f in self._pdf_files]

    def _rebuild_sections(self, n: int):
        current_n = len(self._sec_widgets)
        if n == current_n:
            return
        
        names = self._pdf_names()
        pages = self._max_pages()
        
        if n > current_n:
            # Solo agregar los nuevos al final
            for i in range(current_n, n):
                sw = SectionWidget(i, names, pages)
                sw.delete_requested.connect(self._remove_section)
                self._sec_layout.insertWidget(i, sw)
                self._sec_widgets.append(sw)
        else:
            # Eliminar solo los que sobran del final
            for i in range(current_n - 1, n - 1, -1):
                w = self._sec_widgets.pop()
                self._sec_layout.removeWidget(w)
                w.deleteLater()

    def _remove_section(self, sw: SectionWidget):
        label = sw.le_label.text().strip()
        if sw in self._sec_widgets:
            self._sec_widgets.remove(sw)
            self._sec_layout.removeWidget(sw)
            sw.deleteLater()
            
            # Ajustar el spinbox sin disparar _rebuild_sections
            self.sb_nsec.blockSignals(True)
            self.sb_nsec.setValue(len(self._sec_widgets))
            self.sb_nsec.blockSignals(False)
            
            self.section_deleted.emit(label)

    def _open_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar PDF(s)", "", "PDF (*.pdf)"
        )
        if not paths:
            return
        import fitz as _fitz
        new_files = []
        for p in paths:
            with open(p, "rb") as fh:
                raw = fh.read()
            h = engine._md5(raw)
            with _fitz.open(stream=raw, filetype="pdf") as doc:
                n = len(doc)
            new_files.append({
                "name":  os.path.basename(p),
                "bytes": raw,
                "hash":  h,
                "pages": n,
            })
        self._pdf_files = new_files
        self._refresh_pdf_ui()
        self.pdf_loaded.emit()

    def _clear_pdfs(self):
        self._pdf_files = []
        self._refresh_pdf_ui()
        self.pdf_loaded.emit()

    def _refresh_pdf_ui(self):
        n = len(self._pdf_files)
        if n == 0:
            self.lbl_pdf_info.setText("Sin archivos cargados")
            self.lst_pdf_files.clear()
            self.lst_pdf_files.setVisible(False)
            self.btn_clear_pdf.setEnabled(False)
            self.btn_extract.setEnabled(False)
        else:
            total = sum(f["pages"] for f in self._pdf_files)
            self.lbl_pdf_info.setText(
                f"✅ {n} archivo{'s' if n > 1 else ''} · {total} págs."
            )
            self.lst_pdf_files.clear()
            for f in self._pdf_files:
                self.lst_pdf_files.addItem(f"• {f['name']}")
            self.lst_pdf_files.setVisible(True)
            self.btn_clear_pdf.setEnabled(True)
            self.btn_extract.setEnabled(True)
        # Actualizar secciones
        names = self._pdf_names()
        pages = self._max_pages()
        
        if n > 0:
            self.sb_nsec.setValue(n)
            for i, sw in enumerate(self._sec_widgets):
                sw.update_pdf_names(names, pages)
                if i < n:
                    # Mapear 1 a 1
                    sw.cb_pdf.setCurrentIndex(i)
                    
                    # Auto-completar Rubro
                    fname = names[i]
                    fname_clean = fname[:-4] if fname.lower().endswith('.pdf') else fname
                    sw.le_label.setText(fname_clean)
                    
                    # Inferir Tipo
                    if "PRESUPUESTO" in fname.upper():
                        sw.cb_tipo.setCurrentText("Presupuesto Global")
                    else:
                        sw.cb_tipo.setCurrentText("Cotización Proveedor")
                        
                    # Auto-completar páginas
                    sw.sb_p0.setValue(1)
                    sw.sb_p1.setValue(pages[i])
        else:
            self.sb_nsec.setValue(1)
            for sw in self._sec_widgets:
                sw.update_pdf_names(names, pages)

    def _on_extract(self):
        cfgs = [sw.get_config() for sw in self._sec_widgets]
        self.extract_requested.emit(cfgs, self.le_token.text().strip())

    def _on_update_token(self):
        self.token_update_requested.emit(self.le_token.text().strip())

    def get_state(self) -> dict:
        import base64
        pdfs = []
        for pdf in self._pdf_files:
            pdfs.append({
                "name": pdf["name"],
                "bytes": base64.b64encode(pdf["bytes"]).decode("utf-8"),
                "pages": pdf["pages"]
            })
        return {
            "proyecto": self.le_proyecto.text(),
            "token": self.le_token.text(),
            "pdf_files": pdfs,
            "secciones": [sw.get_config() for sw in self._sec_widgets]
        }

    def load_state(self, state: dict):
        import base64
        if "proyecto" in state:
            self.le_proyecto.setText(state["proyecto"])
        if "token" in state:
            self.le_token.setText(state["token"])
        
        # Cargar PDFs PRIMERO para que los spinboxes tengan los límites correctos
        if "pdf_files" in state:
            self._pdf_files = []
            for pdf in state["pdf_files"]:
                self._pdf_files.append({
                    "name": pdf["name"],
                    "bytes": base64.b64decode(pdf["bytes"]),
                    "pages": pdf["pages"]
                })
            self._refresh_pdf_ui()
            self.pdf_loaded.emit()

        secciones = state.get("secciones", [])
        self.sb_nsec.setValue(len(secciones))
        for i, cfg in enumerate(secciones):
            if i < len(self._sec_widgets):
                self._sec_widgets[i].set_config(cfg)

    def get_pdf_files(self) -> list[dict]:
        return self._pdf_files

    def get_proyecto(self) -> str:
        return self.le_proyecto.text().strip()


# ─────────────────────────────────────────────────────────────
# VISOR DE PDF
# ─────────────────────────────────────────────────────────────
class PDFViewer(QWidget):
    page_changed = pyqtSignal(int)   # nueva página (0-based)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pdf_bytes:  bytes      = b""
        self._total_pages: int       = 0
        self._current_page: int      = 0
        self._page_cache: dict[int, QPixmap] = {}
        self._render_worker: RenderWorker | None = None
        self._zoom_level = 1.0
        self._rotation_angle = 0
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # Título
        title = QLabel("🔍 Visor de Documento")
        title.setObjectName("section_title")
        layout.addWidget(title)

        # Barra de navegación
        nav = QHBoxLayout()
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setFixedWidth(40)
        self.btn_prev.setObjectName("viewer_btn")
        self.btn_prev.clicked.connect(lambda: self._go_to(self._current_page - 1))

        self.lbl_page = QLabel("Página 0 / 0")
        self.lbl_page.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.sb_jump = NoScrollSpinBox()
        self.sb_jump.setMinimum(1)
        self.sb_jump.setFixedWidth(70)
        self.sb_jump.valueChanged.connect(self._on_jump)

        self.btn_next = QPushButton("▶")
        self.btn_next.setFixedWidth(40)
        self.btn_next.setObjectName("viewer_btn")
        self.btn_next.clicked.connect(lambda: self._go_to(self._current_page + 1))
        
        # Herramientas de vista
        self.btn_zoom_out = QPushButton("➖")
        self.btn_zoom_out.setFixedWidth(30)
        self.btn_zoom_out.setObjectName("viewer_btn")
        self.btn_zoom_out.clicked.connect(self._zoom_out)
        
        self.btn_zoom_reset = QPushButton("1:1")
        self.btn_zoom_reset.setFixedWidth(36)
        self.btn_zoom_reset.setObjectName("viewer_btn")
        self.btn_zoom_reset.clicked.connect(self._zoom_reset)
        
        self.btn_zoom_in = QPushButton("➕")
        self.btn_zoom_in.setFixedWidth(30)
        self.btn_zoom_in.setObjectName("viewer_btn")
        self.btn_zoom_in.clicked.connect(self._zoom_in)
        
        self.btn_rotate_left = QPushButton("⟲")
        self.btn_rotate_left.setFixedWidth(30)
        self.btn_rotate_left.setObjectName("viewer_btn")
        self.btn_rotate_left.clicked.connect(self._rotate_left)
        
        self.btn_rotate_right = QPushButton("⟳")
        self.btn_rotate_right.setFixedWidth(30)
        self.btn_rotate_right.setObjectName("viewer_btn")
        self.btn_rotate_right.clicked.connect(self._rotate_right)

        nav.addWidget(self.btn_prev)
        nav.addWidget(self.lbl_page, stretch=1)
        nav.addWidget(self.sb_jump)
        nav.addWidget(self.btn_next)
        nav.addSpacing(10)
        nav.addWidget(self.btn_zoom_out)
        nav.addWidget(self.btn_zoom_reset)
        nav.addWidget(self.btn_zoom_in)
        nav.addSpacing(4)
        nav.addWidget(self.btn_rotate_left)
        nav.addWidget(self.btn_rotate_right)
        layout.addLayout(nav)

        # Indicador de sección
        self.lbl_section = QLabel("")
        self.lbl_section.setStyleSheet(
            f"background:{ACCENT};color:#fff;padding:3px 10px;"
            f"border-radius:4px;font-size:11px;"
        )
        self.lbl_section.setVisible(False)
        layout.addWidget(self.lbl_section)

        # Área de imagen con QGraphicsView
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.view.setStyleSheet(f"background:{BG_DARK}; border:none;")
        self.pixmap_item = QGraphicsPixmapItem()
        self.scene.addItem(self.pixmap_item)
        layout.addWidget(self.view, stretch=1)

        self._nav_blocked = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, '_zoom_level') and getattr(self, '_zoom_level', 1.0) == 1.0:
            self._refresh_current_view()

    def load_pdf(self, pdf_bytes: bytes, total_pages: int):
        self._pdf_bytes   = pdf_bytes
        self._total_pages = total_pages
        self._current_page = 0
        self._page_cache.clear()
        self.sb_jump.setMaximum(max(total_pages, 1))
        self._update_nav()
        self._request_render(0)

    def clear(self):
        self._pdf_bytes   = b""
        self._total_pages = 0
        self._current_page = 0
        self._page_cache.clear()
        self.pixmap_item.setPixmap(QPixmap())
        self._update_nav()

    def _go_to(self, target: int):
        if self._total_pages == 0:
            return
        target = max(0, min(self._total_pages - 1, target))
        if target == self._current_page:
            return
        self._current_page = target
        self._nav_blocked = True
        self.sb_jump.setValue(target + 1)
        self._nav_blocked = False
        self._update_nav()
        if target in self._page_cache:
            self._show_pixmap(self._page_cache[target])
        else:
            self._request_render(target)
        self.page_changed.emit(target)

    def _on_jump(self, val: int):
        if self._nav_blocked:
            return
        self._go_to(val - 1)

    def _update_nav(self):
        tp = self._total_pages
        cp = self._current_page
        self.btn_prev.setEnabled(tp > 0 and cp > 0)
        self.btn_next.setEnabled(tp > 0 and cp < tp - 1)
        self.lbl_page.setText(
            f"Página {cp + 1} / {tp}" if tp > 0 else "Página 0 / 0"
        )

    def _request_render(self, idx: int):
        """Lanza un hilo de render para el índice dado."""
        if not self._pdf_bytes:
            return
        # Si ya hay un worker corriendo para otra página, se le deja terminar;
        # el resultado se cachea aunque ya no sea la página visible.
        worker = RenderWorker(self._pdf_bytes, idx, scale=1.5)
        worker.finished.connect(self._on_rendered)
        worker.failed.connect(lambda i: print(f"Error al renderizar pág. {i + 1}"))
        # Guardar referencia para evitar GC prematuro
        if self._render_worker is not None:
            # No matamos el worker anterior, sólo actualizamos la referencia
            self._render_worker.finished.disconnect()
            self._render_worker.failed.disconnect()
        self._render_worker = worker
        worker.start()

    def _on_rendered(self, idx: int, pm: QPixmap):
        """Callback cuando el worker terminó. Guarda en caché y muestra si es la página actual."""
        self._page_cache[idx] = pm
        if idx == self._current_page:
            self._show_pixmap(pm)

    def _show_pixmap(self, pm: QPixmap):
        if pm is None or pm.isNull():
            return
            
        # Rotar la imagen primero
        if self._rotation_angle != 0:
            transform = QTransform().rotate(self._rotation_angle)
            pm = pm.transformed(transform, Qt.TransformationMode.SmoothTransformation)
            
        self.pixmap_item.setPixmap(pm)
        self.scene.setSceneRect(self.pixmap_item.boundingRect())
        
        # Reset transform and re-apply
        self.view.resetTransform()
        
        # Apply zoom (1.0 = fit to width initially)
        if self._zoom_level == 1.0:
            self.view.fitInView(self.scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
            # Store the actual scale factor that was applied
            self._base_transform = self.view.transform()
        else:
            if hasattr(self, "_base_transform"):
                self.view.setTransform(self._base_transform)
            self.view.scale(self._zoom_level, self._zoom_level)
        
    def _zoom_in(self):
        self._zoom_level = min(self._zoom_level + 0.25, 3.0)
        self._refresh_current_view()
        
    def _zoom_out(self):
        self._zoom_level = max(self._zoom_level - 0.25, 0.25)
        self._refresh_current_view()
        
    def _zoom_reset(self):
        self._zoom_level = 1.0
        self._rotation_angle = 0
        self._refresh_current_view()
        
    def _rotate_left(self):
        self._rotation_angle = (self._rotation_angle - 90) % 360
        self._refresh_current_view()
        
    def _rotate_right(self):
        self._rotation_angle = (self._rotation_angle + 90) % 360
        self._refresh_current_view()
        
    def _refresh_current_view(self):
        if self._current_page in self._page_cache:
            self._show_pixmap(self._page_cache[self._current_page])

    def set_section_label(self, text: str):
        if text:
            self.lbl_section.setText(text)
            self.lbl_section.setVisible(True)
        else:
            self.lbl_section.setVisible(False)

    def current_page(self) -> int:
        return self._current_page


# ─────────────────────────────────────────────────────────────
# PANEL DE DATOS (tabla + KPIs)
# ─────────────────────────────────────────────────────────────
DROPDOWN_COLS = {
    "Tipo":    ["Cotización Proveedor", "Presupuesto Global"],
    "QT":      ["Sí", "No"],
    "(+ IVA)": ["Sí", "Incluido", "Exento", "N/M"],
    "T. Cambio": ["MXN", "USD", "EUR", "CAD"],
}
READONLY_COLS  = {"Nº Sec", "Sección", "Diferencia final"}
MONEY_COLS     = {"Precio Unitario", "Subtotal (Sin IVA)", "IVA 16%",
                  "Total con IVA", "Diferencia final", "Monto en Anexo Escrito"}
NUMERIC_COLS   = {"Cantidad"} | MONEY_COLS

COL_WIDTHS = {
    "Nº Sec": 60, "Sección": 90, "Tipo": 160, "Fecha": 90, "Rubro": 200, "QT": 50,
    "T. Cambio": 80, "(+ IVA)": 75, "Cantidad": 65,
    "Precio Unitario": 110, "Subtotal (Sin IVA)": 120, "IVA 16%": 90,
    "Total con IVA": 110, "Diferencia final": 110,
    "Monto en Anexo Escrito": 140, "Observaciones": 200,
}


class DataPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._df: pd.DataFrame | None = None
        self._updating = False
        self._manual_edits: set[tuple[int, str]] = set()
        self._build()

    def get_state(self) -> dict:
        import json
        df_dict = json.loads(self._df.to_json(orient="records", date_format="iso")) if self._df is not None else None
        # Convert set of tuples to list of lists for JSON serialization
        manual_edits_list = [list(x) for x in self._manual_edits]
        return {
            "df": df_dict,
            "manual_edits": manual_edits_list
        }

    def load_state(self, state: dict):
        import pandas as pd
        if "df" in state and state["df"] is not None:
            df = pd.DataFrame(state["df"]).reindex(columns=engine._COLS)
        else:
            df = pd.DataFrame(columns=engine._COLS)
            
        edits = state.get("manual_edits", [])
        if isinstance(edits, list):
            self._manual_edits = set(tuple(x) for x in edits)
        else:
            self._manual_edits = set()
            
        self.load_data(df)

    def get_current_state(self):
        return self._df, self._manual_edits

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.data_splitter = QSplitter(Qt.Orientation.Vertical)
        self.data_splitter.setHandleWidth(4)

        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(6)

        # Título
        title = QLabel("✏️ Editor de Datos")
        title.setObjectName("section_title")
        top_layout.addWidget(title)

        # KPIs
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(8)
        self._kpi_prov  = self._make_kpi("Total Proveedores")
        self._kpi_anx   = self._make_kpi("Monto Anexo Manual")
        self._kpi_dif   = self._make_kpi("Diferencia")
        kpi_row.addWidget(self._kpi_prov[0])
        kpi_row.addWidget(self._kpi_anx[0])
        kpi_row.addWidget(self._kpi_dif[0])
        top_layout.addLayout(kpi_row)

        # Advertencias con Scroll
        self.warnings_scroll = QScrollArea()
        self.warnings_scroll.setWidgetResizable(True)
        self.warnings_scroll.setVisible(False)
        self.warnings_scroll.setMaximumHeight(150)
        self.warnings_scroll.setStyleSheet(f"QScrollArea {{ border: none; background: {BG_PANEL}; }}")
        
        self.lbl_warnings = QLabel("")
        self.lbl_warnings.setWordWrap(True)
        self.lbl_warnings.setStyleSheet(
            f"color:{ACCENT_LITE};background:{BG_PANEL};"
            f"border:1px solid {BORDER};border-radius:4px;padding:4px;"
        )
        self.warnings_scroll.setWidget(self.lbl_warnings)
        top_layout.addWidget(self.warnings_scroll)
        self.data_splitter.addWidget(top_container)

        self.bot_container = QWidget()
        self.bot_layout = QVBoxLayout(self.bot_container)
        self.bot_layout.setContentsMargins(0, 0, 0, 0)
        self.bot_layout.setSpacing(6)

        # Tabla
        self.table = QTableWidget(0, len(engine._COLS))
        self.table.setHorizontalHeaderLabels(engine._COLS)
        self.table.horizontalHeader().setSectionsMovable(False)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked |
            QTableWidget.EditTrigger.AnyKeyPressed
        )
        self.table.itemChanged.connect(self._on_item_changed)

        for ci, col in enumerate(engine._COLS):
            w = COL_WIDTHS.get(col, 100)
            self.table.setColumnWidth(ci, w)
        self.table.horizontalHeader().setSectionResizeMode(
            engine._COLS.index("Observaciones"),
            QHeaderView.ResizeMode.Stretch,
        )
        self.bot_layout.addWidget(self.table, stretch=1)

        # Botones de descarga
        btn_row = QHBoxLayout()
        self.btn_export = QPushButton("⬇️  Descargar Excel")
        self.btn_export.setObjectName("btn_primary")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_excel)

        self.btn_template = QPushButton("📄  Plantilla vacía")
        self.btn_template.setEnabled(False)
        self.btn_template.clicked.connect(self._export_template)

        btn_row.addWidget(self.btn_export)
        btn_row.addWidget(self.btn_template)

        self.bot_layout.addLayout(btn_row)
        self.data_splitter.addWidget(self.bot_container)
        self.data_splitter.setSizes([100, 800])
        layout.addWidget(self.data_splitter, stretch=1)

        # Placeholder
        self.lbl_placeholder = QLabel(
            "Configura las secciones y presiona 🔍 Extraer Montos"
        )
        self.lbl_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_placeholder.setStyleSheet(f"color:{TEXT_MUTED};")
        layout.addWidget(self.lbl_placeholder)

    def _make_kpi(self, label: str):
        box = QWidget()
        box.setStyleSheet(
            f"background:{BG_PANEL};border:1px solid {BORDER};"
            f"border-radius:8px;"
        )
        vl = QVBoxLayout(box)
        vl.setContentsMargins(10, 8, 10, 8)
        val_lbl = QLabel("$ 0.00")
        val_lbl.setObjectName("kpi_value")
        val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl = QLabel(label)
        lbl.setObjectName("kpi_label")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        vl.addWidget(val_lbl)
        vl.addWidget(lbl)
        return box, val_lbl, lbl

    def load_data(self, df: pd.DataFrame, proyecto: str = "", token: str = ""):
        self._df = df.copy()
        self._proyecto = proyecto
        self._token = token
        self._updating = True
        self._manual_edits = {(r, c) for (r, c) in self._manual_edits if r < len(self._df)}
        self._populate_table(self._df)
        self._updating = False
        self._refresh_kpis()
        self.btn_export.setEnabled(True)
        self.btn_template.setEnabled(True)
        self.lbl_placeholder.setVisible(False)
        self.table.setVisible(True)
        self._check_warnings()

    def set_token(self, token: str):
        self._token = token

    def apply_token(self, new_token: str):
        if not new_token or self._df is None or self._df.empty:
            QMessageBox.warning(self, "Aviso", "No hay datos para aplicar o falta el token.")
            return
            
        self.set_token(new_token)
        import datetime
        from engine import banxico_tc
        
        updated = False
        self._updating = True
        
        for r, row in self._df.iterrows():
            moneda = row.get("T. Cambio", "MXN")
            if moneda != "MXN":
                obs = str(row.get("Observaciones", ""))
                # Si no tiene la leyenda de conversión, es porque falló previamente y sigue en moneda original
                if f"1 {moneda} =" not in obs:
                    fecha_str = row.get("Fecha")
                    try:
                        d_obj = datetime.date.fromisoformat(str(fecha_str).strip()[:10])
                    except (ValueError, TypeError):
                        d_obj = datetime.date.today()
                        
                    tc = banxico_tc(moneda, d_obj, new_token)
                    if tc:
                        pu = self._df.at[r, "Precio Unitario"]
                        if pd.notna(pu):
                            self._df.at[r, "Precio Unitario"] = round(float(pu) * tc, 2)
                        
                        tot_orig = self._df.at[r, "Total con IVA"]
                        if pd.isna(tot_orig): 
                            tot_orig = self._df.at[r, "Subtotal (Sin IVA)"]
                            
                        new_obs = f"1 {moneda} = ${tc:.4f} MXN | Total orig.: {moneda} {float(tot_orig if pd.notna(tot_orig) else 0):,.2f}"
                        if obs and str(obs).strip():
                            self._df.at[r, "Observaciones"] = str(obs) + " | " + new_obs
                        else:
                            self._df.at[r, "Observaciones"] = new_obs
                            
                        self._recalc_and_refresh_row(r)
                        updated = True
        
        self._updating = False
        if updated:
            self._refresh_kpis()
            self._check_warnings()
            QMessageBox.information(self, "Token Aplicado", "Se han convertido las filas pendientes usando el nuevo token.")
        else:
            QMessageBox.information(self, "Token Aplicado", "El token es válido, pero no había filas pendientes por convertir.")

    def remove_section_row(self, label: str):
        if self._df is None or "Sección" not in self._df.columns:
            return
            
        idx_to_drop = self._df[self._df["Sección"] == label].index
        if not idx_to_drop.empty:
            self._df = self._df.drop(idx_to_drop).reset_index(drop=True)
            self._manual_edits.clear()
            self.load_data(self._df, getattr(self, "_proyecto", ""), getattr(self, "_token", ""))

    def _populate_table(self, df: pd.DataFrame):
        self.table.setRowCount(0)
        self.table.setRowCount(len(df))
        for r, (_, row) in enumerate(df.iterrows()):
            self.table.setRowHeight(r, 26)
            for ci, col in enumerate(engine._COLS):
                val = row.get(col, "")
                if col in DROPDOWN_COLS:
                    cb = NoScrollComboBox()
                    cb.addItems(DROPDOWN_COLS[col])
                    s = str(val) if val is not None and str(val) != "nan" else DROPDOWN_COLS[col][0]
                    idx_cb = cb.findText(s)
                    cb.setCurrentIndex(max(idx_cb, 0))
                    cb.currentIndexChanged.connect(
                        lambda _, r=r, ci=ci, cb=cb: self._on_combo_changed(r, ci, cb)
                    )
                    self.table.setCellWidget(r, ci, cb)
                else:
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        display = ""
                    elif col in MONEY_COLS and val != "":
                        try:
                            display = f"${float(val):,.2f}"
                        except (ValueError, TypeError):
                            display = str(val)
                    else:
                        display = str(val)
                    item = QTableWidgetItem(display)
                    if col in READONLY_COLS:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        item.setForeground(QColor(TEXT_MUTED))
                    self.table.setItem(r, ci, item)

    def _on_combo_changed(self, r: int, ci: int, cb: QComboBox):
        if self._updating or self._df is None:
            return
        col = engine._COLS[ci]
        self._manual_edits.add((r, col))
        old_val = self._df.at[r, col]
        new_val = cb.currentText()
        self._df.at[r, col] = new_val
        if col == "(+ IVA)":
            self._recalc_and_refresh_row(r)
        elif col == "T. Cambio" and str(old_val) != str(new_val):
            self._recalc_currency(r, str(old_val), str(new_val))
        else:
            self._refresh_kpis()

    def _recalc_currency(self, r: int, old_curr: str, new_curr: str):
        if not hasattr(self, "_token") or not self._token:
            self.lbl_warnings.setText(f"⚠️ Se detectó cambio a {new_curr} pero se requiere el Token Banxico para hacer la conversión matemática.")
            self.lbl_warnings.setVisible(True)
            if hasattr(self, 'warnings_scroll'): self.warnings_scroll.setVisible(True)
            return
            
        import datetime
        from engine import banxico_tc
        
        fecha_str = self._df.at[r, "Fecha"]
        try:
            d_obj = datetime.date.fromisoformat(str(fecha_str).strip()[:10])
        except (ValueError, TypeError):
            d_obj = datetime.date.today()
            
        factor = 1.0
        
        # Revertir a MXN base si la vieja moneda tenía tasa
        if old_curr != "MXN":
            old_tc = banxico_tc(old_curr, d_obj, self._token)
            if old_tc:
                factor /= old_tc
                
        # Aplicar nueva tasa si la nueva moneda no es MXN
        if new_curr != "MXN":
            new_tc = banxico_tc(new_curr, d_obj, self._token)
            if new_tc:
                factor *= new_tc
                
        if factor != 1.0:
            pu  = self._df.at[r, "Precio Unitario"]
            if pd.notna(pu):
                self._df.at[r, "Precio Unitario"] = round(float(pu) * factor, 2)
            self._recalc_and_refresh_row(r)
        else:
            self._refresh_kpis()

    def _on_item_changed(self, item: QTableWidgetItem):
        if self._updating or self._df is None:
            return
        r   = item.row()
        ci  = item.column()
        col = engine._COLS[ci]
        if col in READONLY_COLS or col in DROPDOWN_COLS:
            return
            
        self._manual_edits.add((r, col))
        
        txt = item.text().replace("$", "").replace(",", "").strip()
        if col in NUMERIC_COLS:
            try:
                self._df.at[r, col] = float(txt) if txt else None
            except ValueError:
                self._df.at[r, col] = None
        else:
            self._df.at[r, col] = txt
            
        if col == "Total con IVA":
            tot = self._df.at[r, "Total con IVA"]
            if pd.notna(tot):
                tot = float(tot)
                qty = self._df.at[r, "Cantidad"]
                qty = float(qty) if pd.notna(qty) else 1.0
                has_iva = str(self._df.at[r, "(+ IVA)"]).strip().lower() in ["sí", "si", "incluido"]
                
                if has_iva:
                    sub = round(tot / 1.16, 2)
                    iva = round(tot - sub, 2)
                else:
                    sub = tot
                    iva = 0.0
                    
                pu = round(sub / qty, 2)
                
                self._df.at[r, "Subtotal (Sin IVA)"] = sub
                self._df.at[r, "IVA 16%"] = iva
                self._df.at[r, "Precio Unitario"] = pu
            self._recalc_and_refresh_row(r, forward=False)
            
        elif col in ["Cantidad", "Precio Unitario"]:
            self._recalc_and_refresh_row(r, forward=True)
        else:
            self._updating = True
            row = self._df.iloc[r]
            val = row.get(col)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                display = ""
            elif col in MONEY_COLS:
                try:
                    display = f"${float(val):,.2f}"
                except (ValueError, TypeError):
                    display = str(val)
            else:
                display = str(val)
            item.setText(display)
            self._updating = False
            self._refresh_kpis()

    def _recalc_and_refresh_row(self, r: int, forward: bool = True):
        qty = self._df.at[r, "Cantidad"]
        qty = float(qty) if pd.notna(qty) else 1.0
        pu = self._df.at[r, "Precio Unitario"]
        has_iva = str(self._df.at[r, "(+ IVA)"]).strip().lower() in ["sí", "si", "incluido"]
        
        if forward and pd.notna(pu):
            sub = qty * float(pu)
            self._df.at[r, "Subtotal (Sin IVA)"] = sub
            if has_iva:
                self._df.at[r, "IVA 16%"] = sub * 0.16
                self._df.at[r, "Total con IVA"] = sub * 1.16
            else:
                self._df.at[r, "IVA 16%"] = 0.0
                self._df.at[r, "Total con IVA"] = sub

        tot = self._df.at[r, "Total con IVA"]
        anx = self._df.at[r, "Monto en Anexo Escrito"]
        if pd.notna(tot) and pd.notna(anx):
            self._df.at[r, "Diferencia final"] = float(tot) - float(anx)
        else:
            self._df.at[r, "Diferencia final"] = None

        self._updating = True
        row = self._df.iloc[r]
        for ci, col in enumerate(engine._COLS):
            if col in DROPDOWN_COLS:
                continue
            if col in READONLY_COLS or col in MONEY_COLS:
                val = row.get(col)
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    display = ""
                elif col in MONEY_COLS:
                    try:
                        display = f"${float(val):,.2f}"
                    except (ValueError, TypeError):
                        display = str(val)
                else:
                    display = str(val)
                item = self.table.item(r, ci)
                if item:
                    item.setText(display)
        self._updating = False
        self._refresh_kpis()

    def _refresh_kpis(self):
        if self._df is None:
            return
        df_prov   = self._df[self._df["Tipo"] == "Cotización Proveedor"]
        df_presup = self._df[self._df["Tipo"] == "Presupuesto Global"]
        ts = pd.to_numeric(df_prov["Total con IVA"], errors="coerce").sum()
        if not df_presup.empty:
            rs = pd.to_numeric(df_presup["Total con IVA"], errors="coerce").sum()
            self._kpi_anx[2].setText("Ppto. Global Extraído")
        else:
            rs = pd.to_numeric(self._df["Monto en Anexo Escrito"], errors="coerce").sum()
            self._kpi_anx[2].setText("Monto Anexo Manual")
        dif = ts - rs
        self._kpi_prov[1].setText(f"${ts:,.2f}")
        self._kpi_anx[1].setText(f"${rs:,.2f}")
        self._kpi_dif[1].setText(f"${dif:,.2f}")
        color = RED_WARN if abs(dif) > 0.01 else GREEN
        self._kpi_dif[1].setStyleSheet(
            f"font-size:18px; font-weight:bold; color:{color};"
        )

    def _check_warnings(self):
        if self._df is None:
            self.lbl_warnings.setVisible(False)
            if hasattr(self, 'warnings_scroll'): self.warnings_scroll.setVisible(False)
            return
        warn_rows = self._df[
            self._df["Observaciones"].astype(str).str.contains("⚠|OCR|inferido", na=False)
        ]
        if not warn_rows.empty:
            lines = [
                f"• {r['Rubro']}: {r['Observaciones']}"
                for _, r in warn_rows.iterrows()
            ]
            self.lbl_warnings.setText(f"⚠ {len(warn_rows)} aviso(s):\n" + "\n".join(lines))
            self.lbl_warnings.setVisible(True)
            if hasattr(self, 'warnings_scroll'): self.warnings_scroll.setVisible(True)
        else:
            self.lbl_warnings.setVisible(False)
            if hasattr(self, 'warnings_scroll'): self.warnings_scroll.setVisible(False)

    def add_warning(self, msg: str):
        cur = self.lbl_warnings.text()
        self.lbl_warnings.setText((cur + "\n" if cur else "") + f"⚠ {msg}")
        self.lbl_warnings.setVisible(True)
        if hasattr(self, 'warnings_scroll'): self.warnings_scroll.setVisible(True)

    def _export_excel(self):
        if self._df is None:
            return
        proyecto = getattr(self, "_proyecto", "")
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        name = (proyecto or "Cotizaciones").replace(" ", "_")
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel",
            f"Conciliacion_{name}_{ts}.xlsx",
            "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            data = engine.to_excel(self._df, nombre=proyecto)
            with open(path, "wb") as fh:
                fh.write(data)
            QMessageBox.information(self, "Exportado", f"Archivo guardado en:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"No se pudo guardar:\n{exc}")

    def _export_template(self):
        if self._df is None:
            return
        proyecto = getattr(self, "_proyecto", "")
        name = (proyecto or "Cotizaciones").replace(" ", "_")
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Plantilla Vacía",
            f"Plantilla_{name}.xlsx",
            "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            # Enviamos flag blank=True para que to_excel mantenga todo excepto numéricos
            data = engine.to_excel(self._df, nombre=proyecto, blank=True)
            with open(path, "wb") as fh:
                fh.write(data)
            QMessageBox.information(self, "Éxito", f"Plantilla guardada en:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar la plantilla:\n{e}")


# ─────────────────────────────────────────────────────────────
# VENTANA PRINCIPAL
# ─────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conciliador de Cotizaciones v6.6")
        self.resize(1400, 860)
        self._extract_worker: ExtractWorker | None = None
        self._pdf_combined_bytes: bytes = b""
        self._pdf_files: list[dict]     = []
        self._sec_cfgs:  list[dict]     = []
        self._current_project_file: str = ""
        self._save_worker = None  # Keep reference to prevent GC
        self._build()

    def _build(self):
        # ── Barra de menú ─────────────────────────────────────
        menu = self.menuBar()
        file_m = menu.addMenu("Archivo")
        file_m.addAction("Cargar Proyecto...", self._load_project)
        file_m.addAction("Guardar Proyecto...", self._save_project)
        file_m.addSeparator()
        file_m.addAction("Acerca de...", self._show_about)
        file_m.addSeparator()
        file_m.addAction("Salir", self.close)

        # ── Barra de estado ───────────────────────────────────
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Listo · Conciliador v6.6")

        # ── Barra de progreso (oculta por defecto) ────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(18)
        self.status.addPermanentWidget(self.progress_bar, 1)

        # ── Layout principal ─────────────────────────────────
        central = QWidget()
        self.setCentralWidget(central)
        h_main = QHBoxLayout(central)
        h_main.setContentsMargins(0, 0, 0, 0)
        h_main.setSpacing(0)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.setHandleWidth(4)

        # Panel de config (sidebar fijo a la izquierda)
        self.config_panel = ConfigPanel()
        self.config_panel.extract_requested.connect(self._start_extract)
        self.config_panel.pdf_loaded.connect(self._on_pdf_loaded)
        
        # Área derecha (Banner + Splitter)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(15, 15, 15, 15)
        right_layout.setSpacing(15)

        # Banner Superior
        banner = QWidget()
        banner.setObjectName("main_banner")
        banner_layout = QHBoxLayout(banner)
        
        # Botón Hamburguesa
        self.btn_toggle_sidebar = QPushButton("☰")
        self.btn_toggle_sidebar.setFixedWidth(36)
        self.btn_toggle_sidebar.setFixedHeight(36)
        self.btn_toggle_sidebar.setStyleSheet("font-size: 20px; font-weight: bold; background: transparent; color: white; border: none;")
        self.btn_toggle_sidebar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_sidebar.clicked.connect(self._toggle_sidebar)
        banner_layout.addWidget(self.btn_toggle_sidebar)

        # Botón Guardar Rápido
        self.btn_quick_save = QPushButton(" 💾 Guardar")
        self.btn_quick_save.setObjectName("btn_icon")
        self.btn_quick_save.setFixedHeight(36)
        self.btn_quick_save.setToolTip("Guardar Proyecto (Sobreescribir)")
        self.btn_quick_save.setStyleSheet("font-size: 14px; font-weight: bold; background: rgba(255, 255, 255, 0.1); border: 1px solid white; border-radius: 4px; color: white; padding: 0 10px;")
        self.btn_quick_save.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_quick_save.clicked.connect(self._save_project)
        banner_layout.addWidget(self.btn_quick_save)
        banner_layout.addSpacing(15)

        text_layout = QVBoxLayout()
        lbl_title = QLabel("📄 Conciliador de Cotizaciones")
        lbl_title.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        lbl_subtitle = QLabel("Extracción automática PDF · Carga múltiple · Cota de Cordura · v6.6")
        lbl_subtitle.setStyleSheet("font-size: 13px; color: #f0f0f0;")
        text_layout.addWidget(lbl_title)
        text_layout.addWidget(lbl_subtitle)
        banner_layout.addLayout(text_layout)
        banner_layout.addStretch()
        right_layout.addWidget(banner)

        # Splitter central
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(4)

        self.pdf_viewer = PDFViewer()
        self.pdf_viewer.page_changed.connect(self._on_page_changed)
        splitter.addWidget(self.pdf_viewer)

        self.data_panel = DataPanel()
        self.config_panel.section_deleted.connect(self.data_panel.remove_section_row)
        self.config_panel.le_token.textChanged.connect(self.data_panel.set_token)
        self.config_panel.token_update_requested.connect(self.data_panel.apply_token)
        splitter.addWidget(self.data_panel)
        splitter.setSizes([400, 600])

        right_layout.addWidget(splitter, stretch=1)
        
        self.main_splitter.addWidget(self.config_panel)
        self.main_splitter.addWidget(right_widget)
        self.main_splitter.setSizes([280, 800])
        h_main.addWidget(self.main_splitter)

    # ── Carga de PDF ──────────────────────────────────────────
    def _on_pdf_loaded(self):
        files = self.config_panel.get_pdf_files()
        self._pdf_files = files
        if not files:
            self.pdf_viewer.clear()
            return
        combined_bytes, _, total_pages, _ = engine.build_combined_pdf(files)
        self._pdf_combined_bytes = combined_bytes
        self.pdf_viewer.load_pdf(combined_bytes, total_pages)
        self.status.showMessage(
            f"PDF cargado · {len(files)} archivo(s) · {total_pages} páginas"
        )

    def _on_page_changed(self, page_idx: int):
        # Mostrar indicador de sección
        for cfg in self._sec_cfgs:
            pdf_idx = cfg.get("pdf_idx", 0)
            offset = sum(
                f["pages"] for f in self._pdf_files[:pdf_idx]
            )
            g_p0 = offset + cfg["p0"] - 1
            g_p1 = offset + cfg["p1"] - 1
            if g_p0 <= page_idx <= g_p1:
                self.pdf_viewer.set_section_label(
                    f"{'🌐' if cfg.get('tipo') == 'Presupuesto Global' else '📑'} "
                    f"{cfg['label']} · {cfg.get('moneda', 'MXN')}"
                )
                return
        self.pdf_viewer.set_section_label("")

    # ── Extracción ────────────────────────────────────────────
    def _start_extract(self, cfgs: list, bx_token: str):
        if not self._pdf_files:
            QMessageBox.warning(self, "Sin PDF", "Carga al menos un PDF antes de extraer.")
            return
        self._sec_cfgs = cfgs
        self.config_panel.btn_extract.setEnabled(False)
        self.progress_bar.setMaximum(len(cfgs))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status.showMessage("Extrayendo montos…")

        self._extract_worker = ExtractWorker(
            self._pdf_files, cfgs, bx_token
        )
        self._extract_worker.progress.connect(self._on_extract_progress)
        self._extract_worker.warning.connect(self.data_panel.add_warning)
        self._extract_worker.finished.connect(self._on_extract_finished)
        self._extract_worker.start()

    def _on_extract_progress(self, current: int, total: int, label: str):
        self.progress_bar.setValue(current)
        self.status.showMessage(f"Extrayendo {current}/{total}: {label}…")

    def _on_extract_finished(self, results: list):
        self.progress_bar.setVisible(False)
        self.config_panel.btn_extract.setEnabled(True)
        df_new = pd.DataFrame(results).reindex(columns=engine._COLS)
        
        # Merge manual edits before recalcular
        df_old, manual_edits = self.data_panel.get_current_state()
        if df_old is not None and not df_old.empty:
            for r, col in manual_edits:
                if r < len(df_new) and col in df_new.columns:
                    df_new.at[r, col] = df_old.at[r, col]
                    
        df_new = engine.recalc_derived(df_new)
        self.data_panel.load_data(df_new, self.config_panel.get_proyecto(), self.config_panel.le_token.text().strip())
        self.status.showMessage(
            f"Extracción completa · {len(results)} sección(es) procesada(s)"
        )

    def _toggle_sidebar(self):
        is_visible = self.config_panel.isVisible()
        self.config_panel.setVisible(not is_visible)

    def _save_project(self):
        if self._current_project_file:
            file_path = self._current_project_file
        else:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Guardar Proyecto", "", "Conciliador Project (*.cshcp)"
            )
            if not file_path:
                return

        # Build state in main thread (fast, just reading widgets)
        try:
            state = {
                "config": self.config_panel.get_state(),
                "data": self.data_panel.get_state()
            }
        except Exception as e:
            QMessageBox.critical(self, "Error al preparar guardado",
                                 f"No se pudo preparar el estado:\n{str(e)}")
            return

        # Do the heavy JSON + disk write on a background thread so Qt doesn't freeze
        self.status.showMessage("Guardando proyecto...")
        self.btn_quick_save.setEnabled(False)

        class _SaveWorker(QThread):
            done    = pyqtSignal(str)   # success path
            errored = pyqtSignal(str)   # error message

            def __init__(self, path, data):
                super().__init__()
                self._path = path
                self._data = data

            def run(self):
                import json, math

                class _Enc(json.JSONEncoder):
                    def default(self, o):
                        # numpy int / float
                        try:
                            import numpy as np
                            if isinstance(o, (np.integer,)): return int(o)
                            if isinstance(o, (np.floating,)): return None if math.isnan(float(o)) else float(o)
                            if isinstance(o, np.ndarray): return o.tolist()
                        except ImportError:
                            pass
                        # pandas Timestamp / NaT
                        try:
                            import pandas as pd
                            if isinstance(o, pd.Timestamp): return o.isoformat()
                            if o is pd.NaT: return None
                        except ImportError:
                            pass
                        return super().default(o)

                try:
                    text = json.dumps(self._data, cls=_Enc, indent=4, ensure_ascii=False)
                    with open(self._path, "w", encoding="utf-8") as f:
                        f.write(text)
                    self.done.emit(self._path)
                except Exception as exc:
                    self.errored.emit(str(exc))

        def on_done(path):
            self._current_project_file = path
            self.btn_quick_save.setEnabled(True)
            self.status.showMessage(f"Guardado: {path}", 4000)

        def on_error(msg):
            self.btn_quick_save.setEnabled(True)
            self.status.showMessage("Error al guardar.", 4000)
            QMessageBox.critical(self, "Error al guardar",
                                 f"No se pudo guardar el proyecto:\n{msg}")

        worker = _SaveWorker(file_path, state)
        worker.done.connect(on_done)
        worker.errored.connect(on_error)
        self._save_worker = worker   # prevent GC while running
        worker.start()

    def _load_project(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Cargar Proyecto", "", "Conciliador Project (*.cshcp)")
        if file_path:
            import json
            try:
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        state = json.load(f)
                except UnicodeDecodeError:
                    with open(file_path, "r", encoding="latin-1") as f:
                        state = json.load(f)
                
                # Para retrocompatibilidad
                if "config" in state:
                    self.config_panel.load_state(state["config"])
                    if "data" in state:
                        self.data_panel.load_state(state["data"])
                else:
                    self.config_panel.load_state(state)
                
                self._current_project_file = file_path
                self.status.showMessage(f"Proyecto cargado desde {file_path}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo cargar el proyecto:\\n{str(e)}")

    # ── Acerca de ─────────────────────────────────────────────
    def _show_about(self):
        QMessageBox.about(
            self,
            "Acerca del Conciliador",
            "<b>Conciliador de Cotizaciones v6.6</b><br>"
            "Motor de extracción PDF nativo con PyQt6.<br><br>"
            "5 estrategias de extracción + OCR + Banxico API.<br>"
            "Sin dependencias de servidor web."
        )


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────
def main():
    # Habilitar soporte para pantallas HiDPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setApplicationName("Conciliador de Cotizaciones")
    app.setOrganizationName("Conciliador")
    app.setStyle("Fusion")
    app.setStyleSheet(QSS)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

# -*- mode: python ; coding: utf-8 -*-
# ============================================================
# conciliador_qt.spec  –  PyInstaller v6.6 (PyQt6)
# ============================================================
import os
from PyInstaller.utils.hooks import collect_data_files

# RapidOCR modelos ONNX
rapidocr_data = collect_data_files("rapidocr_onnxruntime", include_py_files=False)

# PyQt6 datos (plugins de plataforma, traducciones, etc.)
qt_data = collect_data_files("PyQt6")

extra_datas = [
    *rapidocr_data,
    *qt_data,
]

if os.path.exists("assets/icon.ico"):
    extra_datas.append(("assets/icon.ico", "assets"))

if os.path.isdir("bin/tesseract"):
    extra_datas.append(("bin/tesseract", "bin/tesseract"))
else:
    print("[AVISO] No se encontró bin/tesseract; el ejecutable requerirá Tesseract instalado en el sistema.")

a = Analysis(
    ["src/main_qt.py"],
    pathex=["src"],
    binaries=[],
    datas=extra_datas,
    hiddenimports=[
        "PyQt6",
        "PyQt6.sip",
        "PyQt6.QtCore",
        "PyQt6.QtGui",
        "PyQt6.QtWidgets",
        "rapidocr_onnxruntime",
        "rapidocr_onnxruntime.main",
        "onnxruntime",
        "pdfplumber",
        "pdfplumber.page",
        "pymupdf",
        "fitz",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "pandas",
        "numpy",
        "requests",
        "pytesseract",
        "PIL",
        "engine",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["streamlit", "pywebview", "tkinter", "matplotlib", "scipy"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Conciliador",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,             # Sin ventana de consola
    icon="assets/icon.ico" if os.path.exists("assets/icon.ico") else None,
)

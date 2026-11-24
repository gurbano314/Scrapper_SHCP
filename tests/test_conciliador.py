import datetime
import io
import os
import sys
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

# Asegurar acceso a src/ sin depender de PYTHONPATH externo
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT_DIR, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

# Evitar offscreen QPA en Windows para conservar fuentes nativas
os.environ.pop("QT_QPA_PLATFORM", None)

from PyQt6.QtCore import QDate, Qt
from PyQt6.QtWidgets import QApplication

import engine
import main_qt


def _get_qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
        app.setStyle("Fusion")
        app.setStyleSheet(main_qt.QSS)
    return app


class EngineTests(unittest.TestCase):
    def test_to_excel_alignment_formulas_decimals_and_sheet_title(self):
        df = pd.DataFrame([
            {
                "Nº Sec": 1,
                "Sección": "Ppto Global",
                "Tipo": "Presupuesto Global",
                "Fecha": "2026-03-01",
                "Rubro": "Presupuesto Total Proyecto",
                "QT": "Sí",
                "T. Cambio": "MXN",
                "(+ IVA)": "Sí",
                "Cantidad": 1,
                "Precio Unitario": 100000.0,
                "Subtotal (Sin IVA)": 100000.0,
                "IVA 16%": 16000.0,
                "Total con IVA": 116000.0,
                "Diferencia final": 0.0,
                "Monto en Anexo Escrito": 116000.0,
                "Observaciones": "Presupuesto global",
            },
            {
                "Nº Sec": 2,
                "Sección": "Sec 2",
                "Tipo": "Cotización Proveedor",
                "Fecha": "2026-03-10",
                "Rubro": "Tubos de fibra",
                "QT": "Sí",
                "T. Cambio": "MXN",
                "(+ IVA)": "Sí",
                "Cantidad": 2.5,
                "Precio Unitario": 1000.0,
                "Subtotal (Sin IVA)": 2500.0,
                "IVA 16%": 400.0,
                "Total con IVA": 2900.0,
                "Diferencia final": None,
                "Monto en Anexo Escrito": None,
                "Observaciones": "IVA desglosado",
            },
        ])

        xlsx_bytes = engine.to_excel(df, nombre="PAR / Pista: [2026]*?")
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
        ws = wb.active

        # Título saneado sin caracteres ilegales de Excel
        for bad in ("/", "\\", ":", "*", "?", "[", "]"):
            self.assertNotIn(bad, ws.title)

        # Alineación exacta de encabezados A..P (16 columnas)
        headers = [ws.cell(2, ci).value for ci in range(1, 17)]
        self.assertEqual(headers, engine._COLS)

        # Fila 4 (segunda fila de datos = Cotización Proveedor)
        self.assertEqual(ws.cell(4, 1).value, 2)
        self.assertEqual(ws.cell(4, 2).value, "Sec 2")
        self.assertEqual(ws.cell(4, 3).value, "Cotización Proveedor")
        self.assertEqual(ws.cell(4, 4).value, "2026-03-10")
        self.assertEqual(ws.cell(4, 5).value, "Tubos de fibra")
        self.assertEqual(ws.cell(4, 9).value, 2.5)  # decimal sin truncar
        self.assertEqual(ws.cell(4, 10).value, 1000.0)
        self.assertEqual(ws.cell(4, 11).value, '=IF(ISNUMBER(J4),ROUND(I4*J4,2),"")')
        self.assertEqual(ws.cell(4, 12).value, '=IF(ISNUMBER(K4),ROUND(K4*0.16,2),"")')
        self.assertEqual(ws.cell(4, 13).value, '=IF(ISNUMBER(K4),IF(ISNUMBER(L4),K4+L4,K4),"")')
        self.assertEqual(ws.cell(4, 14).value, '=IF(AND(ISNUMBER(M4),ISNUMBER(O4)),M4-O4,"")')
        self.assertIsNone(ws.cell(4, 15).value)
        self.assertEqual(ws.cell(4, 16).value, "IVA desglosado")

        # Fila de totales con SUMIF excluyendo Presupuesto Global en K, L, M y SUM en O
        self.assertEqual(ws.cell(5, 1).value, "TOTALES")
        self.assertEqual(ws.cell(5, 11).value, '=SUMIF(C3:C4,"<>Presupuesto Global",K3:K4)')
        self.assertEqual(ws.cell(5, 12).value, '=SUMIF(C3:C4,"<>Presupuesto Global",L3:L4)')
        self.assertEqual(ws.cell(5, 13).value, '=SUMIF(C3:C4,"<>Presupuesto Global",M3:M4)')
        self.assertEqual(ws.cell(5, 15).value, "=SUM(O3:O4)")

    def test_validate_normative_tope_admin_with_global_and_all_rules(self):
        # Si Presupuesto Global = 100,000 y Cotizaciones = 100,000, un gasto admin de 7,000 es 7% (> 5%)
        # Antes se sumaban ambos -> 200,000 -> 3.5% y NO alertaba.
        df = pd.DataFrame([
            {
                "Nº Sec": 1, "Tipo": "Presupuesto Global", "Fecha": "2026-06-01",
                "Rubro": "Presupuesto papelería global", "Total con IVA": 100_000.0,
            },
            {
                "Nº Sec": 2, "Tipo": "Cotización Proveedor", "Fecha": "2026-01-15",
                "Rubro": "Honorarios del contador", "Total con IVA": 7_000.0,
            },
            {
                "Nº Sec": 3, "Tipo": "Cotización Proveedor", "Fecha": "2026-06-01",
                "Rubro": "Kayak profesional", "Total con IVA": 93_000.0,
            },
            {
                "Nº Sec": 4, "Tipo": "Cotización Proveedor", "Fecha": "2026-06-01",
                "Rubro": "Kayak profesional", "Total con IVA": 92_000.0,
            },
        ]).reindex(columns=engine._COLS)

        warnings = engine.validate_normative(
            df,
            fecha_solicitud=datetime.date(2026, 6, 15),
            pii_solo_equipo=True,
        )
        joined = "\n".join(warnings)
        self.assertIn("exceden el tope del 5%", joined)
        self.assertIn("7.0%", joined)
        self.assertIn("excede los 3 meses de vigencia", joined)
        self.assertIn("Posible duplicado", joined)
        self.assertIn("No se detectó un rubro de seguro", joined)
        self.assertIn("PII exclusivamente de equipo", joined)
        self.assertIn("CFDI", joined)

    def test_banxico_tc_does_not_cache_failures(self):
        engine._BX_CACHE.clear()
        hoy = datetime.date(2026, 5, 10)

        class _FailResp:
            def raise_for_status(self):
                raise RuntimeError("401 Unauthorized")

        class _OkResp:
            def raise_for_status(self):
                return None

            def json(self):
                return {"bmx": {"series": [{"datos": [{"dato": "20.15"}]}]}}

        with patch.object(engine.requests, "get", side_effect=[_FailResp(), _OkResp()]) as mock_get:
            tc_fail = engine.banxico_tc("USD", hoy, "bad_token")
            self.assertIsNone(tc_fail)
            self.assertNotIn(("USD", hoy.isoformat()), engine._BX_CACHE)

            tc_ok = engine.banxico_tc("USD", hoy, "good_token")
            self.assertEqual(tc_ok, 20.15)
            self.assertEqual(mock_get.call_count, 2)

            # Aciertos válidos sí quedan en caché
            tc_cached = engine.banxico_tc("USD", hoy, "good_token")
            self.assertEqual(tc_cached, 20.15)
            self.assertEqual(mock_get.call_count, 2)

    def test_recalc_derived_diferencia_none_and_iva_incluido(self):
        df = pd.DataFrame([
            {
                "Nº Sec": 1, "Sección": "S1", "Tipo": "Cotización Proveedor",
                "Fecha": "2026-03-01", "Rubro": "Equipo", "QT": "Sí",
                "T. Cambio": "MXN", "(+ IVA)": "Incluido", "Cantidad": 2,
                "Precio Unitario": 500.0, "Subtotal (Sin IVA)": None,
                "IVA 16%": None, "Total con IVA": None,
                "Diferencia final": None, "Monto en Anexo Escrito": None,
                "Observaciones": "",
            }
        ])
        out = engine.recalc_derived(df)
        self.assertEqual(out.at[0, "Subtotal (Sin IVA)"], 1000.0)
        self.assertEqual(out.at[0, "IVA 16%"], 160.0)
        self.assertEqual(out.at[0, "Total con IVA"], 1160.0)
        self.assertTrue(pd.isna(out.at[0, "Diferencia final"]))


class UITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _get_qapp()

    def test_positional_section_deletion_even_with_duplicate_names(self):
        win = main_qt.MainWindow()
        win.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        try:
            win.config_panel._pdf_files = [{"name": "a.pdf", "bytes": b"%PDF", "hash": "h", "pages": 10}]
            win.config_panel._refresh_pdf_ui()
            secs = [
                {"label": "Alimentación", "tipo": "Cotización Proveedor", "pdf_idx": 0, "p0": 1, "p1": 1, "det_iva": True, "calc_sub": True},
                {"label": "Alimentación", "tipo": "Cotización Proveedor", "pdf_idx": 0, "p0": 2, "p1": 2, "det_iva": True, "calc_sub": True},
                {"label": "Alimentación", "tipo": "Cotización Proveedor", "pdf_idx": 0, "p0": 3, "p1": 3, "det_iva": True, "calc_sub": True},
            ]
            win.config_panel.apply_detected_sections(secs)
            rows = [
                {
                    "Nº Sec": i + 1, "Sección": "Alimentación", "Tipo": "Cotización Proveedor",
                    "Fecha": "2026-06-01", "Rubro": f"Comida {i+1}", "QT": "Sí",
                    "T. Cambio": "MXN", "(+ IVA)": "Sí", "Cantidad": 1,
                    "Precio Unitario": 100.0 * (i + 1), "Subtotal (Sin IVA)": 100.0 * (i + 1),
                    "IVA 16%": 16.0 * (i + 1), "Total con IVA": 116.0 * (i + 1),
                    "Diferencia final": None, "Monto en Anexo Escrito": None, "Observaciones": "",
                }
                for i in range(3)
            ]
            win._on_extract_finished(rows)
            self.assertEqual(len(win.data_panel._df), 3)

            # Eliminar la segunda sección homónima (índice 1)
            win.config_panel._remove_section(win.config_panel._sec_widgets[1])
            self.assertEqual(len(win.config_panel._sec_widgets), 2)
            self.assertEqual(win.config_panel.sb_nsec.value(), 2)
            self.assertEqual(list(win.data_panel._df["Rubro"]), ["Comida 1", "Comida 3"])
            self.assertEqual(list(win.data_panel._df["Nº Sec"]), [1, 2])

            # Eliminar todas hasta llegar a la última restante
            win.config_panel._remove_section(win.config_panel._sec_widgets[1])
            win.config_panel._remove_section(win.config_panel._sec_widgets[0])
            self.assertEqual(len(win.config_panel._sec_widgets), 1)
            self.assertEqual(win.config_panel.sb_nsec.value(), 1)
            self.assertTrue(win.data_panel._df.empty)
            self.assertEqual(win.data_panel.table.rowCount(), 0)
        finally:
            win.close()

    def test_zero_quantity_edit_no_zero_division(self):
        win = main_qt.MainWindow()
        win.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        try:
            df = pd.DataFrame([{
                "Nº Sec": 1, "Sección": "S1", "Tipo": "Cotización Proveedor",
                "Fecha": "2026-06-01", "Rubro": "Papelería", "QT": "Sí",
                "T. Cambio": "MXN", "(+ IVA)": "Sí", "Cantidad": 0,
                "Precio Unitario": 100.0, "Subtotal (Sin IVA)": 100.0,
                "IVA 16%": 16.0, "Total con IVA": 116.0,
                "Diferencia final": None, "Monto en Anexo Escrito": None, "Observaciones": "",
            }])
            win.data_panel.load_data(df, "Test", "")
            col_tot = engine._COLS.index("Total con IVA")
            item = win.data_panel.table.item(0, col_tot)
            item.setText("$232.00")
            # No debe lanzar ZeroDivisionError y debe asumir Cantidad = 1
            self.assertEqual(win.data_panel._df.at[0, "Precio Unitario"], 200.0)
        finally:
            win.close()

    def test_save_load_state_and_live_normative_validation(self):
        import pymupdf

        d = pymupdf.open()
        d.new_page()
        d.new_page()
        valid_pdf = d.tobytes()
        d.close()

        win = main_qt.MainWindow()
        win.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        try:
            win.config_panel._pdf_files = [{"name": "doc.pdf", "bytes": valid_pdf, "hash": engine._md5(valid_pdf), "pages": 2}]
            win.config_panel.le_proyecto.setText("PII Kayak")
            win.config_panel.le_token.setText("tok_123")
            win.config_panel.de_fecha_solicitud.setDate(QDate(2026, 6, 15))
            win.config_panel.chk_pii_equipo.setChecked(False)

            rows = [{
                "Nº Sec": 1, "Sección": "S1", "Tipo": "Cotización Proveedor",
                "Fecha": "2026-06-01", "Rubro": "Honorarios del contador", "QT": "Sí",
                "T. Cambio": "MXN", "(+ IVA)": "Sí", "Cantidad": 1,
                "Precio Unitario": 1000.0, "Subtotal (Sin IVA)": 1000.0,
                "IVA 16%": 160.0, "Total con IVA": 1160.0,
                "Diferencia final": None, "Monto en Anexo Escrito": None, "Observaciones": "",
            }]
            win._on_extract_finished(rows)
            state = {"config": win.config_panel.get_state(), "data": win.data_panel.get_state()}

            win2 = main_qt.MainWindow()
            win2.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
            try:
                win2.show()
                win2.config_panel.load_state(state["config"])
                win2.data_panel.load_state(
                    state["data"],
                    proyecto=win2.config_panel.get_proyecto(),
                    token=win2.config_panel.le_token.text().strip(),
                )
                # Saltos rápidos de página en el visor no deben crashear RenderWorker
                win2.pdf_viewer._go_to(1)
                win2.pdf_viewer._go_to(0)
                self.app.processEvents()

                self.assertEqual(win2.data_panel._proyecto, "PII Kayak")
                self.assertEqual(win2.data_panel._token, "tok_123")
                self.assertIn("hash", win2.config_panel._pdf_files[0])

                # Activar checkbox PII en vivo -> debe alertar inmediatamente sin re-extraer
                win2.config_panel.chk_pii_equipo.setChecked(True)
                self.assertTrue(win2.data_panel.warnings_scroll.isVisible())
                self.assertIn("PII exclusivamente de equipo", win2.data_panel.lbl_warnings.text())
            finally:
                win2.close()
        finally:
            win.close()

    def test_template_button_enabled_before_extraction(self):
        win = main_qt.MainWindow()
        win.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        try:
            self.assertTrue(win.data_panel.btn_template.isEnabled())
            self.assertFalse(win.data_panel.btn_export.isEnabled())
            sanitized = win.data_panel._sanitize_filename("PAR: Pista / <2026>*?")
            self.assertEqual(sanitized, "PAR_Pista_2026")
        finally:
            win.close()


if __name__ == "__main__":
    unittest.main()

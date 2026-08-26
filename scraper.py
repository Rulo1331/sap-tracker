"""
Automatización para extraer solicitudes de pedido (SP) desde SAP WebGUI,
usando la función nativa de exportar a Excel en vez de leer la pantalla.

Flujo replicado (grabado con playwright codegen):
1. Login en WebGUI
2. Abrir la transacción "Visualizar solicitud de pedido" (una sola vez)
3. Por cada SP: usar "Otra solicitud de pedido" -> ingresar número -> Enter
4. Exportar -> Hoja de cálculo -> nombre de archivo -> descargar
5. Leer el Excel descargado con pandas

AJUSTAR si tu flujo cambia:
- Los nombres de botones ("Transacciones SAP GUI -> Visualizar solicitud
  de pedido", "Ejecutar", "Otra solicitud de pedido (May") son exactos a
  tu grabación; si tu menú se ve distinto, corrige los textos abajo.
"""

import os
import shutil
import tempfile
from dataclasses import dataclass, asdict
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

try:
    import streamlit as st
    _secrets = st.secrets
except Exception:
    _secrets = {}


def _get_config(key: str, default: str = None) -> str:
    """Busca primero en st.secrets (Streamlit Cloud) y si no, en variables
    de entorno (para correrlo en tu máquina local con un .env)."""
    if key in _secrets:
        return _secrets[key]
    return os.getenv(key, default)


FIORI_URL = _get_config(
    "FIORI_URL",
    "https://tu-servidor.com/sap/bc/gui/sap/its/webgui",
)
SAP_USER = _get_config("SAP_USER")
SAP_PASSWORD = _get_config("SAP_PASSWORD")

# Nombre exacto del botón de menú que abre la transacción de solicitudes.
# AJUSTAR si tu menú/favoritos tiene otro texto.
TRANSACTION_BUTTON_NAME = "Transacciones SAP GUI -> Visualizar solicitud de pedido"


@dataclass
class OrderExportResult:
    order_number: str
    file_path: str = ""
    error: str = ""


def login_and_open_transaction(page):
    """Login + abrir la transacción de 'Visualizar solicitud de pedido'.
    Se hace UNA sola vez por sesión de navegador; luego cada SP se
    consulta reutilizando la misma pantalla con 'Otra solicitud de pedido'.
    """
    page.goto(FIORI_URL)

    page.get_by_role("textbox", name="Usuario Obligatorio").fill(SAP_USER)
    page.locator("#sap-password-r").click()
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").fill(SAP_PASSWORD)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").press("Enter")

    page.get_by_role("button", name=TRANSACTION_BUTTON_NAME).click()
    page.get_by_role("button", name="Ejecutar  Resaltado").click()


def export_order(page, order_number: str, download_dir: Path) -> OrderExportResult:
    """Busca una SP específica y la exporta a Excel, descargando el archivo
    a download_dir. Devuelve la ruta del archivo descargado."""
    try:
        page.get_by_role("button", name="Otra solicitud de pedido (May").click()
        search_box = page.get_by_role("textbox", name="Solicitud de pedido")
        search_box.click()
        search_box.fill(order_number)
        search_box.press("Enter")

        page.get_by_role("button", name="Exportar").click()
        page.get_by_text("Hoja de cálculo").click()

        file_name = f"{order_number}.xlsx"
        page.get_by_role("textbox", name="Fichero").click()
        page.get_by_role("textbox", name="Fichero").fill(file_name)

        with page.expect_download() as download_info:
            with page.expect_popup() as popup_info:
                page.get_by_role("button", name="OK").click()
            popup = popup_info.value
        download = download_info.value
        popup.close()

        save_path = download_dir / file_name
        download.save_as(save_path)

        return OrderExportResult(order_number=order_number, file_path=str(save_path))

    except PlaywrightTimeout:
        return OrderExportResult(
            order_number=order_number,
            error="No se encontró la SP o la exportación tardó demasiado",
        )
    except Exception as e:
        return OrderExportResult(order_number=order_number, error=str(e))


def read_exported_order(file_path: str) -> pd.DataFrame:
    """Lee el Excel exportado por SAP. AJUSTAR una vez que revises la
    estructura real del archivo (SAP a veces agrega filas de encabezado
    antes de la tabla de datos -> usar skiprows si hace falta)."""
    return pd.read_excel(file_path)


def get_multiple_order_exports(order_numbers: list[str]) -> list[OrderExportResult]:
    """Abre una sola sesión, hace login una vez, y exporta varias SP en fila.
    headless=True corre sin mostrar ventana -> puedes seguir trabajando
    en otra cosa mientras se ejecuta."""
    results = []
    with tempfile.TemporaryDirectory() as tmp_dir:
        download_dir = Path(tmp_dir)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            login_and_open_transaction(page)

            for order in order_numbers:
                results.append(export_order(page, order, download_dir))

            browser.close()

        # Copiamos los archivos fuera del directorio temporal antes de que
        # se borre, a una carpeta persistente para revisarlos si hace falta.
        persistent_dir = Path("descargas_sap")
        persistent_dir.mkdir(exist_ok=True)
        for r in results:
            if r.file_path:
                new_path = persistent_dir / Path(r.file_path).name
                shutil.move(r.file_path, new_path)
                r.file_path = str(new_path)

    return results


def results_to_dicts(results: list[OrderExportResult]) -> list[dict]:
    return [asdict(r) for r in results]


# ---------------------------------------------------------------------------
# Transacción 2: "Reporte de seguimiento" - exporta a texto con tabuladores.
# Es una transacción distinta a la de arriba, así que tiene su propio login
# y su propio recorrido de menú.
# ---------------------------------------------------------------------------

# AJUSTAR si el texto de tu menú es distinto:
REPORT_MENU_TITLE = "Desplegar nodo"          # título del ícono que abre el árbol de menú
REPORT_BUTTON_NAME = "Reporte seguimiento de"  # nombre del botón dentro del árbol


def login_and_open_report_transaction(page):
    """Login + abrir la transacción 'Reporte de seguimiento'. Se hace UNA
    sola vez por sesión; luego cada SP se consulta reutilizando la misma
    pantalla (se cambia el número y se vuelve a ejecutar)."""
    page.goto(FIORI_URL)

    page.get_by_role("textbox", name="Usuario Obligatorio").fill(SAP_USER)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").fill(SAP_PASSWORD)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").press("Enter")

    page.get_by_title(REPORT_MENU_TITLE).click()
    page.get_by_role("button", name=REPORT_BUTTON_NAME).click()
    page.get_by_role("button", name="Ejecutar  Resaltado").click()


def export_report(page, order_number: str, download_dir: Path) -> OrderExportResult:
    """Busca una SP en el Reporte de seguimiento y lo exporta como texto
    con tabuladores (.txt)."""
    try:
        search_box = page.get_by_role("textbox", name="Solicitud de Pedido")
        search_box.click()
        search_box.fill(order_number)
        page.get_by_role("button", name="Ejecutar  Resaltado").click()

        page.get_by_role("button", name="Local File... (Control+Mayús+").click()
        page.get_by_role("radio", name="Texto con tabuladores").click()
        page.get_by_role("button", name="Continuar (Entrada)").click()

        file_name = f"{order_number}.txt"
        fichero_box = page.get_by_role("textbox", name="Fichero")
        fichero_box.click()
        fichero_box.fill(file_name)

        with page.expect_download() as download_info:
            fichero_box.press("Enter")
        download = download_info.value

        save_path = download_dir / file_name
        download.save_as(save_path)

        return OrderExportResult(order_number=order_number, file_path=str(save_path))

    except PlaywrightTimeout:
        return OrderExportResult(
            order_number=order_number,
            error="No se encontró la SP o la exportación tardó demasiado",
        )
    except Exception as e:
        return OrderExportResult(order_number=order_number, error=str(e))


def read_report_txt(file_path: str) -> pd.DataFrame:
    """Lee el archivo de texto con tabuladores exportado por SAP.
    AJUSTAR encoding si el archivo real no calza (algunos exports de SAP
    usan latin-1 en vez de utf-8)."""
    try:
        return pd.read_csv(file_path, sep="\t", encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(file_path, sep="\t", encoding="latin-1")


def get_multiple_report_exports(order_numbers: list[str]) -> list[OrderExportResult]:
    """Igual que get_multiple_order_exports, pero para la transacción de
    Reporte de seguimiento (exporta .txt en vez de .xlsx)."""
    results = []
    with tempfile.TemporaryDirectory() as tmp_dir:
        download_dir = Path(tmp_dir)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            login_and_open_report_transaction(page)

            for order in order_numbers:
                results.append(export_report(page, order, download_dir))

            browser.close()

        persistent_dir = Path("descargas_sap_reporte")
        persistent_dir.mkdir(exist_ok=True)
        for r in results:
            if r.file_path:
                new_path = persistent_dir / Path(r.file_path).name
                shutil.move(r.file_path, new_path)
                r.file_path = str(new_path)

    return results

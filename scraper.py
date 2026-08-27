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
import re
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
    screenshot_path: str = ""
    html_path: str = ""


def _capture_diagnostics(page, persistent_dir: Path, label: str) -> tuple[str, str]:
    """Guarda una captura de pantalla y el HTML de la página en el momento
    de un error, para poder diagnosticar sin poder ver el navegador
    headless corriendo en la nube. Cada una se intenta de forma
    independiente: si la imagen falla (pasa con páginas pesadas con
    iframes), igual queda el HTML guardado como respaldo."""
    screenshot_path = ""
    html_path = ""
    try:
        shot = persistent_dir / f"{label}.png"
        # Sin full_page=True: en pantallas con iframes/dynpros de SAP a
        # veces falla o se corrompe al medir el alto real de la página.
        # Una captura del viewport visible es más confiable para diagnóstico.
        page.screenshot(path=str(shot))
        screenshot_path = str(shot)
    except Exception:
        pass
    try:
        html_file = persistent_dir / f"{label}.html"
        html_file.write_text(page.content(), encoding="utf-8")
        html_path = str(html_file)
    except Exception:
        pass
    return screenshot_path, html_path


def _get_visible(locator, timeout: int = 30000):
    """Devuelve el elemento realmente VISIBLE de un locator con varios
    matches (común en SAP WebGUI cuando quedan pantallas o menús
    anteriores 'fantasma' en el HTML). Si ninguno pasa el filtro, devuelve
    el último como mejor esfuerzo. Devuelve None si no hay ningún match."""
    locator.first.wait_for(state="attached", timeout=timeout)
    candidates = locator.all()
    for c in candidates:
        if c.is_visible():
            return c
    return candidates[-1] if candidates else None


def _click_visible(page, role: str, name: str, timeout: int = 30000, **kwargs):
    """Hace clic en el elemento VISIBLE que matchee ese rol/nombre (ver
    _get_visible)."""
    el = _get_visible(page.get_by_role(role, name=name, **kwargs), timeout)
    if el is None:
        raise Exception(f"No se encontró ningún elemento {role} con nombre '{name}'")
    el.click()


def _click_visible_text(page, text: str, timeout: int = 30000):
    """Igual que _click_visible pero buscando por texto en vez de rol
    (para opciones de menú tipo 'Hoja de cálculo' que no siempre exponen
    un rol de botón claro)."""
    el = _get_visible(page.get_by_text(text), timeout)
    if el is None:
        raise Exception(f"No se encontró ningún texto visible '{text}'")
    el.click()


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
        _click_visible(page, "button", "Otra solicitud de pedido (May")
        search_box = page.get_by_role("textbox", name="Solicitud de pedido")
        search_box.click()
        search_box.fill(order_number)
        search_box.press("Enter")

        _click_visible(page, "button", "Exportar")
        _click_visible_text(page, "Hoja de cálculo")

        file_name = f"{order_number}.xlsx"
        fichero_box = _get_visible(page.get_by_role("textbox", name="Fichero"))
        fichero_box.click()
        fichero_box.fill(file_name)

        with page.expect_download() as download_info:
            with page.expect_popup() as popup_info:
                _click_visible(page, "button", "OK")
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
        persistent_dir = Path("descargas_sap")
        persistent_dir.mkdir(exist_ok=True)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            try:
                login_and_open_transaction(page)
            except Exception as e:
                shot_path, html_path = _capture_diagnostics(page, persistent_dir, "error_login")
                browser.close()
                return [
                    OrderExportResult(
                        order_number="(login/apertura de transacción)",
                        error=f"No se pudo abrir la transacción: {e}",
                        screenshot_path=shot_path,
                        html_path=html_path,
                    )
                ]

            for order in order_numbers:
                result = export_order(page, order, download_dir)
                if result.error and not result.screenshot_path:
                    shot_path, html_path = _capture_diagnostics(page, persistent_dir, f"error_{order}")
                    result.screenshot_path = shot_path
                    result.html_path = html_path
                results.append(result)

            browser.close()

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
    pantalla (se cambia el número y se vuelve a ejecutar).

    Cada paso espera explícitamente a que la pantalla termine de recargar
    (networkidle) antes de seguir, porque SAP WebGUI recarga la pantalla
    completa en cada acción (a diferencia de una app moderna que solo
    actualiza un fragmento) y eso a veces tarda más desde la nube que
    desde tu red local.
    """
    page.goto(FIORI_URL)
    page.wait_for_load_state("networkidle")

    page.get_by_role("textbox", name="Usuario Obligatorio").fill(SAP_USER)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").fill(SAP_PASSWORD)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").press("Enter")
    page.wait_for_load_state("networkidle")

    page.get_by_title(REPORT_MENU_TITLE).click()
    page.wait_for_load_state("networkidle")

    page.locator(".lsScrollbar--expandSize.lsScrollbar--absolute").click()

    page.get_by_role("button", name=REPORT_BUTTON_NAME).click(timeout=60000)
    page.wait_for_load_state("networkidle")

    page.get_by_role("button", name="Ejecutar  Resaltado").click()
    page.wait_for_load_state("networkidle")


def export_report(page, order_number: str, download_dir: Path) -> OrderExportResult:
    """Busca una SP en el Reporte de seguimiento y lo exporta como texto
    con tabuladores (.txt)."""
    try:
        search_box = _get_visible(page.get_by_role("textbox", name="Solicitud de Pedido"))
        search_box.click()
        search_box.fill(order_number)
        _click_visible(page, "button", "Ejecutar  Resaltado")
        page.wait_for_load_state("networkidle")

        _click_visible(page, "button", "Local File... (Control+Mayús+")
        _click_visible(page, "radio", "Texto con tabuladores")
        _click_visible(page, "button", "Continuar (Entrada)")

        file_name = f"{order_number}.txt"
        fichero_box = _get_visible(page.get_by_role("textbox", name="Fichero"))
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
        persistent_dir = Path("descargas_sap_reporte")
        persistent_dir.mkdir(exist_ok=True)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            try:
                login_and_open_report_transaction(page)
            except Exception as e:
                shot_path, html_path = _capture_diagnostics(
                    page, persistent_dir, "error_login_reporte"
                )
                browser.close()
                return [
                    OrderExportResult(
                        order_number="(login/apertura de transacción)",
                        error=f"No se pudo abrir la transacción de reporte: {e}",
                        screenshot_path=shot_path,
                        html_path=html_path,
                    )
                ]

            for order in order_numbers:
                result = export_report(page, order, download_dir)
                if result.error and not result.screenshot_path:
                    shot_path, html_path = _capture_diagnostics(page, persistent_dir, f"error_{order}")
                    result.screenshot_path = shot_path
                    result.html_path = html_path
                results.append(result)

            browser.close()

        for r in results:
            if r.file_path:
                new_path = persistent_dir / Path(r.file_path).name
                shutil.move(r.file_path, new_path)
                r.file_path = str(new_path)

    return results


# ---------------------------------------------------------------------------
# Transacción 3: "Reporte seguim. serv. y otros"
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Transacción 3: "Reporte seguim. serv. y otros"
# ---------------------------------------------------------------------------

REPORT3_MENU_TITLE = "Desplegar nodo"
REPORT3_BUTTON_NAME = "Reporte seguim. serv. y otros"

def login_and_open_report3_transaction(page):
    """Login + abrir la transacción 'Reporte seguim. serv. y otros'."""
    page.goto(FIORI_URL)
    page.wait_for_load_state("networkidle")

    page.get_by_role("textbox", name="Usuario Obligatorio").fill(SAP_USER)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").fill(SAP_PASSWORD)
    page.get_by_role("textbox", name="Clave de acceso Obligatorio").press("Enter")
    page.wait_for_load_state("networkidle")

    page.get_by_title(REPORT3_MENU_TITLE).click()
    page.wait_for_load_state("networkidle")

    page.locator(".lsScrollbar--expandSize.lsScrollbar--absolute").click()

    page.get_by_role("button", name=REPORT3_BUTTON_NAME).click(timeout=60000)
    page.wait_for_load_state("networkidle")

    page.get_by_role("button", name="Ejecutar  Resaltado").click()
    page.wait_for_load_state("networkidle")


def export_report3(page, order_number: str, download_dir: Path) -> OrderExportResult:
    """Busca una SP en el Reporte 3 y lo exporta usando locators específicos."""
    try:
        search_box = _get_visible(page.get_by_role("textbox", name="Solicitud de Pedido"))
        search_box.click()
        search_box.fill(order_number)
        _click_visible(page, "button", "Ejecutar  Resaltado")
        page.wait_for_load_state("networkidle")

        # AQUÍ USAMOS TUS LOCATORS ESPECÍFICOS PARA LA DESCARGA
        page.get_by_role("button", name="Fichero local... (Control+May").click()
        page.locator("div").filter(has_text=re.compile(r"^Texto con tabuladores$")).click()
        _click_visible(page, "button", "Continuar (Entrada)")

        file_name = f"{order_number}.txt"
        fichero_box = _get_visible(page.get_by_role("textbox", name="Fichero"))
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


def get_multiple_report3_exports(order_numbers: list[str]) -> list[OrderExportResult]:
    """Exporta múltiples SP usando la tercera transacción."""
    results = []
    with tempfile.TemporaryDirectory() as tmp_dir:
        download_dir = Path(tmp_dir)
        persistent_dir = Path("descargas_sap_reporte3")
        persistent_dir.mkdir(exist_ok=True)

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            try:
                login_and_open_report3_transaction(page)
            except Exception as e:
                shot_path, html_path = _capture_diagnostics(
                    page, persistent_dir, "error_login_reporte3"
                )
                browser.close()
                return [
                    OrderExportResult(
                        order_number="(login/apertura de transacción)",
                        error=f"No se pudo abrir la transacción de reporte 3: {e}",
                        screenshot_path=shot_path,
                        html_path=html_path,
                    )
                ]

            for order in order_numbers:
                # AHORA LLAMAMOS A export_report3 EN LUGAR DE export_report
                result = export_report3(page, order, download_dir)
                if result.error and not result.screenshot_path:
                    shot_path, html_path = _capture_diagnostics(page, persistent_dir, f"error_{order}")
                    result.screenshot_path = shot_path
                    result.html_path = html_path
                results.append(result)

            browser.close()

        for r in results:
            if r.file_path:
                new_path = persistent_dir / Path(r.file_path).name
                shutil.move(r.file_path, new_path)
                r.file_path = str(new_path)

    return results

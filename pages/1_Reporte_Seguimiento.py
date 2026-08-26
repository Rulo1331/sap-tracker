import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Reporte de Seguimiento SAP", page_icon="📋")


@st.cache_resource
def _ensure_playwright_browser():
    """Mismo chequeo que en la página principal - si ya se instaló ahí,
    esto es casi instantáneo (no reinstala de nuevo)."""
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
    )


with st.spinner("Preparando el navegador headless (solo la primera vez)..."):
    _ensure_playwright_browser()

from scraper import get_multiple_report_exports, read_report_txt  # noqa: E402

st.title("📋 Reporte de Seguimiento SAP")

st.markdown(
    "Ingresa números de solicitud de pedido para consultar el "
    "**Reporte de seguimiento** (se exporta como texto con tabuladores)."
)

tab1, tab2 = st.tabs(["✍️ Ingreso manual", "📁 Subir Excel"])

order_numbers: list[str] = []

with tab1:
    text_input = st.text_area(
        "Números de solicitud de pedido (uno por línea)",
        placeholder="4000005467\n4000005535",
        key="report_text_input",
    )
    if text_input:
        order_numbers = [line.strip() for line in text_input.splitlines() if line.strip()]

with tab2:
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel", type=["xlsx", "xls"], key="report_file_uploader"
    )
    if uploaded_file:
        df_upload = pd.read_excel(uploaded_file)
        col = next(
            (c for c in df_upload.columns if c.lower() in ("pedido", "solped", "order")),
            None,
        )
        if col:
            order_numbers = df_upload[col].astype(str).tolist()
            st.success(f"Se detectaron {len(order_numbers)} pedidos en la columna '{col}'")
        else:
            st.error("No se encontró una columna 'pedido', 'solped' u 'order' en el archivo.")

if st.button("🔍 Consultar reporte", disabled=not order_numbers, key="report_button"):
    with st.spinner(f"Exportando {len(order_numbers)} reporte(s) desde SAP..."):
        export_results = get_multiple_report_exports(order_numbers)

    resumen = pd.DataFrame([r.__dict__ for r in export_results])
    st.dataframe(resumen, use_container_width=True)

    st.divider()
    st.subheader("Detalle por solicitud")

    for r in export_results:
        with st.expander(f"SP {r.order_number}"):
            if r.error:
                st.error(r.error)
                if r.screenshot_path and Path(r.screenshot_path).exists():
                    st.image(
                        r.screenshot_path,
                        caption="Así se veía la pantalla de SAP en el momento del error",
                    )
                elif not r.html_path:
                    st.caption("No se pudo generar ninguna captura de diagnóstico.")
                if r.html_path and Path(r.html_path).exists():
                    with open(r.html_path, "rb") as f:
                        st.download_button(
                            "⬇️ Descargar HTML de esa pantalla (ábrelo en tu navegador)",
                            f.read(),
                            Path(r.html_path).name,
                            key=f"html_report_{r.order_number}",
                        )
                continue
            try:
                df_detalle = read_report_txt(r.file_path)
                st.dataframe(df_detalle, use_container_width=True)
                with open(r.file_path, "rb") as f:
                    st.download_button(
                        f"⬇️ Descargar TXT original ({r.order_number})",
                        f.read(),
                        Path(r.file_path).name,
                        key=f"dl_report_{r.order_number}",
                    )
            except Exception as e:
                st.warning(f"El archivo se descargó pero no se pudo leer automáticamente: {e}")

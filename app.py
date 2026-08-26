import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Seguimiento de Pedidos SAP", page_icon="📦")


@st.cache_resource
def _ensure_playwright_browser():
    """Streamlit Cloud no trae Chromium preinstalado. Esto lo instala una
    sola vez por sesión del contenedor (se cachea con @st.cache_resource,
    así no se reinstala en cada recarga de la app)."""
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
    )


with st.spinner("Preparando el navegador headless (solo la primera vez)..."):
    _ensure_playwright_browser()

from scraper import get_multiple_order_exports, read_exported_order  # noqa: E402
st.title("📦 Seguimiento de Pedidos / Solped SAP")

st.markdown(
    "Ingresa números de pedido manualmente o sube un Excel con una columna "
    "llamada **`pedido`** (o `solped` / `order`) para consultar el status en Fiori."
)

tab1, tab2 = st.tabs(["✍️ Ingreso manual", "📁 Subir Excel"])

order_numbers: list[str] = []

with tab1:
    text_input = st.text_area(
        "Números de pedido (uno por línea)",
        placeholder="4500001234\n4500001235",
    )
    if text_input:
        order_numbers = [line.strip() for line in text_input.splitlines() if line.strip()]

with tab2:
    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])
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

if st.button("🔍 Consultar status", disabled=not order_numbers):
    with st.spinner(f"Exportando {len(order_numbers)} solicitud(es) desde SAP..."):
        export_results = get_multiple_order_exports(order_numbers)

    resumen = pd.DataFrame([r.__dict__ for r in export_results])
    st.dataframe(resumen, use_container_width=True)

    st.divider()
    st.subheader("Detalle por solicitud")

    for r in export_results:
        with st.expander(f"SP {r.order_number}"):
            if r.error:
                st.error(r.error)
                if r.screenshot_path:
                    st.image(
                        r.screenshot_path,
                        caption="Así se veía la pantalla de SAP en el momento del error",
                    )
                continue
            try:
                df_detalle = read_exported_order(r.file_path)
                st.dataframe(df_detalle, use_container_width=True)
                with open(r.file_path, "rb") as f:
                    st.download_button(
                        f"⬇️ Descargar Excel original ({r.order_number})",
                        f.read(),
                        Path(r.file_path).name,
                        key=f"dl_{r.order_number}",
                    )
            except Exception as e:
                st.warning(f"El archivo se descargó pero no se pudo leer automáticamente: {e}")

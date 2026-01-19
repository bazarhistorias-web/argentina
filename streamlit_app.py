import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="BH | Cruce Pedido vs Factura", layout="wide")

# =========================
# Helpers
# =========================
EDITORIALES_VALIDAS = ["ivrea", "ovni", "planeta", "panini", "kemuri"]

def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))

def norm_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().replace("\u00a0", " ")
    s = strip_accents(s).lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def to_int(x):
    try:
        return int(float(str(x).replace(",", ".")))
    except:
        return 0

def to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except:
        return np.nan

def excel_download(sheets: dict, filename="reporte_cruce_bh.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)
    return output, filename

# =========================
# Parsers
# =========================
def parse_base(df, c_pais, c_semana, c_nombre, c_editorial, c_cantidad):
    b = df.rename(columns={
        c_pais: "Pais",
        c_semana: "Semana",
        c_nombre: "Nombre",
        c_editorial: "Editorial",
        c_cantidad: "Cantidad_pedida"
    })

    b["Cantidad_pedida"] = b["Cantidad_pedida"].apply(to_int)

    for c in ["Pais", "Editorial", "Nombre"]:
        b[c] = b[c].astype("string").fillna("")
        b[f"{c}_norm"] = b[c].map(norm_text)

    return b.groupby(
        ["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"],
        as_index=False
    ).agg({"Cantidad_pedida": "sum"})

def parse_factura(df, c_pais, c_editorial, c_titulo, c_cantidad, c_precio):
    f = df.rename(columns={
        c_pais: "Pais",
        c_editorial: "Editorial",
        c_titulo: "Nombre",
        c_cantidad: "Cantidad_facturada",
        c_precio: "Precio_unitario"
    })

    f["Cantidad_facturada"] = f["Cantidad_facturada"].apply(to_int)
    f["Precio_unitario"] = f["Precio_unitario"].apply(to_float)
    f["Total_factura_bruto"] = f["Cantidad_facturada"] * f["Precio_unitario"]

    for c in ["Pais", "Editorial", "Nombre"]:
        f[c] = f[c].astype("string").fillna("")
        f[f"{c}_norm"] = f[c].map(norm_text)

    return f.groupby(
        ["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"],
        as_index=False
    ).agg({
        "Cantidad_facturada": "sum",
        "Precio_unitario": "median",
        "Total_factura_bruto": "sum"
    })

# =========================
# UI
# =========================
st.title("BH · Cruce Pedido vs Factura + Descuentos por Editorial")

st.subheader("📂 Carga de archivos")
c1, c2 = st.columns(2)
with c1:
    f_base = st.file_uploader("Excel BASE (pedido)", type=["xlsx", "xls"])
with c2:
    f_fact = st.file_uploader("Excel FACTURA", type=["xlsx", "xls"])

if not f_base or not f_fact:
    st.stop()

df_base = pd.read_excel(f_base)
df_fact = pd.read_excel(f_fact)

st.subheader("🧩 Mapeo de columnas")
cA, cB = st.columns(2)

with cA:
    st.markdown("**BASE**")
    b_pais = st.selectbox("Pais", df_base.columns)
    b_semana = st.selectbox("Semana", df_base.columns)
    b_nombre = st.selectbox("Nombre", df_base.columns)
    b_editorial = st.selectbox("Editorial", df_base.columns)
    b_cantidad = st.selectbox("Cantidad", df_base.columns)

with cB:
    st.markdown("**FACTURA**")
    f_pais = st.selectbox("Pais (factura)", df_fact.columns)
    f_editorial = st.selectbox("Editorial (factura)", df_fact.columns)
    f_titulo = st.selectbox("Título", df_fact.columns)
    f_cantidad = st.selectbox("Cantidad", df_fact.columns)
    f_precio = st.selectbox("Precio unitario", df_fact.columns)

# =========================
# DESCUENTOS
# =========================
st.subheader("💸 Descuentos por editorial (%)")

descuentos = {}
cols = st.columns(len(EDITORIALES_VALIDAS))
for i, ed in enumerate(EDITORIALES_VALIDAS):
    descuentos[ed] = cols[i].number_input(
        ed.capitalize(),
        min_value=0.0,
        max_value=100.0,
        value=0.0,
        step=1.0
    )

# =========================
# Procesamiento
# =========================
base = parse_base(df_base, b_pais, b_semana, b_nombre, b_editorial, b_cantidad)
fact = parse_factura(df_fact, f_pais, f_editorial, f_titulo, f_cantidad, f_precio)

merged = pd.merge(
    base,
    fact,
    on=["Pais_norm", "Nombre_norm"],
    how="outer",
    suffixes=("_base", "_fact"),
    indicator=True
)

rep = pd.DataFrame()
rep["Pais"] = merged["Pais_fact"].fillna(merged["Pais_base"])
rep["Editorial"] = merged["Editorial_fact"].fillna(merged["Editorial_base"])
rep["Nombre"] = merged["Nombre_fact"].fillna(merged["Nombre_base"])

rep["Cantidad_pedida"] = merged["Cantidad_pedida"].fillna(0).astype(int)
rep["Cantidad_facturada"] = merged["Cantidad_facturada"].fillna(0).astype(int)
rep["Diferencia"] = rep["Cantidad_facturada"] - rep["Cantidad_pedida"]

rep["Precio_unitario"] = merged["Precio_unitario"]
rep["Total_bruto"] = merged["Total_factura_bruto"]

# aplicar descuento
rep["Editorial_norm"] = rep["Editorial"].map(norm_text)
rep["Descuento_%"] = rep["Editorial_norm"].map(descuentos).fillna(0.0)
rep["Precio_con_desc"] = rep["Precio_unitario"] * (1 - rep["Descuento_%"] / 100)
rep["Total_con_desc"] = rep["Cantidad_facturada"] * rep["Precio_con_desc"]

rep["Estado"] = np.select(
    [
        merged["_merge"] == "both",
        merged["_merge"] == "left_only",
        merged["_merge"] == "right_only"
    ],
    ["OK / DIF", "NO LLEGÓ", "NO PEDIDO"],
    default="REVISAR"
)

# =========================
# Resultados
# =========================
st.subheader("📊 Resumen por país (con descuento)")
pago_pais = rep.groupby("Pais", as_index=False)[
    ["Total_bruto", "Total_con_desc"]
].sum()

st.dataframe(pago_pais, use_container_width=True)

st.subheader("📋 Detalle completo")
st.dataframe(rep, use_container_width=True, height=520)

# =========================
# Export
# =========================
st.subheader("⬇️ Descargar reporte")
file_bytes, fname = excel_download({
    "Cruce_completo": rep,
    "Pago_por_pais": pago_pais
})

st.download_button(
    "Descargar Excel",
    data=file_bytes,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

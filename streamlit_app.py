import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="BH | Cruce Pedido vs Factura", layout="wide")

# -------------------------
# Helpers
# -------------------------
PAISES_OK = {"chile", "peru", "perú", "colombia", "argentina", "mexico", "méxico", "espana", "españa"}

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
        if pd.isna(x) or str(x).strip() == "":
            return 0
        s = str(x).replace(".", "").replace(",", ".")
        return int(round(float(s)))
    except:
        return 0

def to_float(x):
    try:
        if pd.isna(x) or str(x).strip() == "":
            return np.nan
        s = str(x).replace(".", "").replace(",", ".")
        return float(s)
    except:
        return np.nan

def excel_download(sheets: dict, filename="reporte_cruce.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)
    return output, filename

# -------------------------
# Parsers
# -------------------------
def parse_base(df_base: pd.DataFrame, c_pais, c_semana, c_nombre, c_editorial, c_cantidad):
    b = df_base.copy()
    b = b.rename(columns={
        c_pais: "Pais",
        c_semana: "Semana",
        c_nombre: "Nombre",
        c_editorial: "Editorial",
        c_cantidad: "Cantidad_pedida",
    })

    b["Pais"] = b["Pais"].astype("string").fillna("")
    b["Semana"] = b["Semana"].astype("string").fillna("")
    b["Nombre"] = b["Nombre"].astype("string").fillna("")
    b["Editorial"] = b["Editorial"].astype("string").fillna("")
    b["Cantidad_pedida"] = b["Cantidad_pedida"].apply(to_int)

    b["Pais_norm"] = b["Pais"].map(norm_text)
    b["Editorial_norm"] = b["Editorial"].map(norm_text)
    b["Nombre_norm"] = b["Nombre"].map(norm_text)

    b_g = b.groupby(
        ["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"],
        as_index=False
    ).agg({"Cantidad_pedida": "sum"})

    return b_g

def parse_factura_simple(df, c_pais, c_editorial, c_titulo, c_cantidad, c_precio):
    f = df.copy()
    f = f.rename(columns={
        c_pais: "Pais",
        c_editorial: "Editorial",
        c_titulo: "Nombre",
        c_cantidad: "Cantidad_facturada",
        c_precio: "Precio_unitario",
    })

    f["Pais"] = f["Pais"].astype("string").fillna("")
    f["Editorial"] = f["Editorial"].astype("string").fillna("")
    f["Nombre"] = f["Nombre"].astype("string").fillna("")

    f["Cantidad_facturada"] = f["Cantidad_facturada"].apply(to_int)
    f["Precio_unitario"] = f["Precio_unitario"].apply(to_float)
    f["Total_factura"] = f["Cantidad_facturada"] * f["Precio_unitario"]

    f["Pais_norm"] = f["Pais"].map(norm_text)
    f["Editorial_norm"] = f["Editorial"].map(norm_text)
    f["Nombre_norm"] = f["Nombre"].map(norm_text)

    f_g = f.groupby(
        ["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"],
        as_index=False
    ).agg({
        "Cantidad_facturada": "sum",
        "Precio_unitario": "median",
        "Total_factura": "sum"
    })

    return f_g

# -------------------------
# UI
# -------------------------
st.title("Cruce BH: Pedido (Base) vs Factura (Recibido)")
st.caption("Match por nombre (normalizado) + cálculo de diferencias + total por país según factura (cantidad × precio).")

c1, c2 = st.columns(2)
with c1:
    f_base = st.file_uploader("Sube Excel BASE (pedido)", type=["xlsx", "xls"])
with c2:
    f_fact = st.file_uploader("Sube Excel FACTURA (recibido / facturación)", type=["xlsx", "xls"])

if not f_base or not f_fact:
    st.stop()

xls_base = pd.ExcelFile(f_base)
xls_fact = pd.ExcelFile(f_fact)

cA, cB = st.columns(2)
with cA:
    sheet_base = st.selectbox("Hoja BASE", xls_base.sheet_names, index=0)
with cB:
    sheet_fact = st.selectbox("Hoja FACTURA", xls_fact.sheet_names, index=0)

df_base_raw = pd.read_excel(xls_base, sheet_name=sheet_base)
df_fact_raw = pd.read_excel(xls_fact, sheet_name=sheet_fact)

st.subheader("1) Mapear columnas")

col_map1, col_map2 = st.columns(2)

with col_map1:
    st.markdown("**BASE (pedido)**")
    c_pais = st.selectbox("Pais (base)", df_base_raw.columns, index=0)
    c_semana = st.selectbox("Semana (base)", df_base_raw.columns, index=1 if len(df_base_raw.columns) > 1 else 0)
    c_nombre = st.selectbox("Nombre (base)", df_base_raw.columns, index=2 if len(df_base_raw.columns) > 2 else 0)
    c_editorial = st.selectbox("Editorial (base)", df_base_raw.columns, index=3 if len(df_base_raw.columns) > 3 else 0)
    c_cantidad = st.selectbox("Cantidad (base)", df_base_raw.columns, index=4 if len(df_base_raw.columns) > 4 else 0)

with col_map2:
    st.markdown("**FACTURA (final)**")
    f_pais = st.selectbox("Pais (factura)", df_fact_raw.columns, index=0)
    f_editorial = st.selectbox("Editorial (factura)", df_fact_raw.columns, index=1 if len(df_fact_raw.columns) > 1 else 0)
    f_titulo = st.selectbox("titulo (factura)", df_fact_raw.columns, index=2 if len(df_fact_raw.columns) > 2 else 0)
    f_cantidad = st.selectbox("cantidad (factura)", df_fact_raw.columns, index=3 if len(df_fact_raw.columns) > 3 else 0)
    f_precio = st.selectbox("Precio (factura)", df_fact_raw.columns, index=4 if len(df_fact_raw.columns) > 4 else 0)

st.divider()

# -------------------------
# Parsear y cruzar
# -------------------------
try:
    base = parse_base(df_base_raw, c_pais, c_semana, c_nombre, c_editorial, c_cantidad)
    fact = parse_factura_simple(df_fact_raw, f_pais, f_editorial, f_titulo, f_cantidad, f_precio)
except Exception as e:
    st.error(f"Error leyendo / procesando los archivos: {e}")
    st.stop()

# Validaciones simples de factura
fact_errors = []
if (fact["Pais_norm"] == "").any():
    fact_errors.append("Hay filas en FACTURA con País vacío.")
if (fact["Nombre_norm"] == "").any():
    fact_errors.append("Hay filas en FACTURA con título vacío.")
if (fact["Cantidad_facturada"] <= 0).any():
    fact_errors.append("Hay filas en FACTURA con cantidad 0 o negativa (revisar).")
if fact["Precio_unitario"].isna().any():
    fact_errors.append("Hay filas en FACTURA con Precio vacío o no numérico.")

# Match recomendado: Pais + Nombre (título)
key_cols = ["Pais_norm", "Nombre_norm"]

merged = pd.merge(
    base,
    fact,
    on=key_cols,
    how="outer",
    suffixes=("_base", "_fact"),
    indicator=True
)

def pick_col(df, base_col, fact_col, default=""):
    out = df[base_col].copy() if base_col in df.columns else pd.Series([default] * len(df))
    if fact_col in df.columns:
        out = out.fillna(df[fact_col])
    return out.fillna(default)

rep = pd.DataFrame()
rep["Pais"] = pick_col(merged, "Pais_base", "Pais_fact")
rep["Editorial_base"] = merged.get("Editorial_base", "")
rep["Editorial_factura"] = merged.get("Editorial_fact", "")

rep["Nombre"] = pick_col(merged, "Nombre_base", "Nombre_fact")
rep["Cantidad_pedida"] = merged.get("Cantidad_pedida", 0).fillna(0).astype(int)
rep["Cantidad_facturada"] = merged.get("Cantidad_facturada", 0).fillna(0).astype(int)
rep["Diferencia_cantidad"] = rep["Cantidad_facturada"] - rep["Cantidad_pedida"]

rep["Precio_unitario"] = merged.get("Precio_unitario", np.nan)
rep["Total_factura"] = merged.get("Total_factura", np.nan)

rep["Editorial_distinta"] = (
    merged.get("Editorial_norm_base", "").fillna("") != merged.get("Editorial_norm_fact", "").fillna("")
)
rep["Editorial_distinta"] = rep["Editorial_distinta"].fillna(False)

rep["Estado"] = np.select(
    [
        merged["_merge"].eq("both") & (rep["Diferencia_cantidad"] == 0),
        merged["_merge"].eq("both") & (rep["Diferencia_cantidad"] < 0),
        merged["_merge"].eq("both") & (rep["Diferencia_cantidad"] > 0),
        merged["_merge"].eq("left_only"),
        merged["_merge"].eq("right_only"),
    ],
    [
        "OK (coincide)",
        "FALTANTE (llegó menos)",
        "SOBRANTE (llegó más)",
        "NO LLEGÓ (pedido sin factura)",
        "NO PEDIDO (factura sin pedido)",
    ],
    default="REVISAR"
)

# -------------------------
# Resultados
# -------------------------
st.subheader("2) Resumen")
r1, r2, r3, r4 = st.columns(4)
r1.metric("Ítems OK", int((rep["Estado"] == "OK (coincide)").sum()))
r2.metric("Ítems faltantes", int(rep["Estado"].isin(["FALTANTE (llegó menos)", "NO LLEGÓ (pedido sin factura)"]).sum()))
r3.metric("Ítems sobrantes / no pedidos", int(rep["Estado"].isin(["SOBRANTE (llegó más)", "NO PEDIDO (factura sin pedido)"]).sum()))
r4.metric("Total líneas", len(rep))

st.subheader("3) Errores detectados (calidad de datos)")
if fact_errors:
    for e in fact_errors:
        st.warning(e)
else:
    st.success("No se detectaron problemas obvios en el formato de FACTURA.")

st.subheader("4) Cuánto debe pagar cada país (según factura)")
if rep["Total_factura"].notna().any():
    pago_pais = rep.dropna(subset=["Total_factura"]).groupby("Pais", as_index=False)["Total_factura"].sum()
    pago_pais = pago_pais.sort_values("Total_factura", ascending=False)
    st.dataframe(pago_pais, use_container_width=True)
else:
    st.info("No se pudo calcular Total_factura. Revisa que Precio y cantidad sean numéricos.")

st.subheader("5) Detalle del cruce")
filtro = st.multiselect(
    "Filtrar por estado",
    sorted(rep["Estado"].unique().tolist()),
    default=sorted(rep["Estado"].unique().tolist())
)

solo_editorial_distinta = st.checkbox("Mostrar solo casos con editorial distinta (warning)", value=False)
rep_view = rep[rep["Estado"].isin(filtro)].copy()
if solo_editorial_distinta:
    rep_view = rep_view[rep_view["Editorial_distinta"] == True]

st.dataframe(rep_view, use_container_width=True, height=520)

ok = rep[rep["Estado"] == "OK (coincide)"].copy()
faltantes = rep[rep["Estado"].isin(["FALTANTE (llegó menos)", "NO LLEGÓ (pedido sin factura)"])].copy()
sobrantes = rep[rep["Estado"].isin(["SOBRANTE (llegó más)", "NO PEDIDO (factura sin pedido)"])].copy()
warn_editorial = rep[rep["Editorial_distinta"] == True].copy()

st.subheader("6) Descargar reporte")
sheets = {
    "Cruce_completo": rep,
    "OK": ok,
    "Faltantes": faltantes,
    "Sobrantes_NoPedido": sobrantes,
    "Editorial_distinta": warn_editorial
}

if rep["Total_factura"].notna().any():
    sheets["Pago_por_pais"] = pago_pais

file_bytes, fname = excel_download(sheets, filename="reporte_cruce_bh.xlsx")
st.download_button(
    "Descargar Excel del reporte",
    data=file_bytes,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Match: Pais + titulo (normalizado). Consejo: mantén títulos idénticos en ambos archivos para cruce perfecto.")

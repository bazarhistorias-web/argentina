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
# Parser factura (externo) estilo "Ivrea Chile" + items
# -------------------------
HEADER_RE = re.compile(
    r"^\s*([A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9\.\-]+)\s+(Chile|Peru|Perú|Colombia|Argentina|Mexico|México|Espana|España)\s*$",
    re.IGNORECASE
)

def parse_factura_externa(df_raw: pd.DataFrame, col_titulo: str, col_cantidad: str, col_pvp: str | None, col_total: str | None):
    df = df_raw.copy()

    if col_titulo not in df.columns or col_cantidad not in df.columns:
        raise ValueError("No se encuentran columnas requeridas en la factura.")

    current_editorial = ""
    current_pais = ""

    rows = []
    for _, r in df.iterrows():
        titulo = r[col_titulo]
        if pd.isna(titulo) or str(titulo).strip() == "":
            continue

        t = str(titulo).strip()

        # Si parece encabezado "Editorial País"
        m = HEADER_RE.match(t)
        if m and (pd.isna(r[col_cantidad]) or to_int(r[col_cantidad]) == 0):
            current_editorial = m.group(1).strip()
            current_pais = m.group(2).strip()
            current_pais = current_pais.replace("Peru", "Perú").replace("Mexico", "México").replace("Espana", "España")
            continue

        cant = to_int(r[col_cantidad])

        # Extraer precio/total si existen
        pvp = to_float(r[col_pvp]) if col_pvp and col_pvp in df.columns else np.nan
        tot = to_float(r[col_total]) if col_total and col_total in df.columns else np.nan

        nombre = t
        editorial = current_editorial
        pais = current_pais

        # Si viene "Nombre | Ivrea Argentina" en la misma celda
        if "|" in t:
            left, right = [p.strip() for p in t.split("|", 1)]
            if left:
                nombre = left
            tail = right

            m2 = HEADER_RE.match(tail)
            if m2:
                editorial = m2.group(1).strip()
                pais = m2.group(2).strip()
                pais = pais.replace("Peru", "Perú").replace("Mexico", "México").replace("Espana", "España")
            else:
                tokens = tail.split()
                if len(tokens) >= 2 and norm_text(tokens[-1]) in PAISES_OK:
                    pais = tokens[-1]
                    editorial = " ".join(tokens[:-1])

        rows.append({
            "Pais": pais,
            "Editorial": editorial,
            "Nombre": nombre,
            "Cantidad_facturada": cant,
            "PVP_factura": pvp,
            "Total_factura": tot,
        })

    out = pd.DataFrame(rows)

    # Normalización
    out["Pais"] = out["Pais"].replace({"Peru": "Perú", "Mexico": "México", "Espana": "España"})
    out["Pais_norm"] = out["Pais"].map(norm_text)
    out["Editorial_norm"] = out["Editorial"].map(norm_text)
    out["Nombre_norm"] = out["Nombre"].map(norm_text)

    # Agrupar (por si viene duplicado)
    grp_cols = ["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"]
    out_g = out.groupby(grp_cols, dropna=False, as_index=False).agg({
        "Cantidad_facturada": "sum",
        "Total_factura": "sum",
        "PVP_factura": "median"
    })

    return out_g

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

    # Agrupar por si el base tiene repetidos por semana/editorial
    b_g = b.groupby(["Pais", "Editorial", "Nombre", "Pais_norm", "Editorial_norm", "Nombre_norm"], as_index=False).agg({
        "Cantidad_pedida": "sum"
    })

    return b_g

# -------------------------
# UI
# -------------------------
st.title("Cruce BH: Pedido (Base) vs Factura (Recibido)")
st.caption("Match por nombre (normalizado) y cálculo de diferencias + total por país según factura.")

col1, col2 = st.columns(2)
with col1:
    f_base = st.file_uploader("Sube Excel BASE (pedido)", type=["xlsx", "xls"])
with col2:
    f_fact = st.file_uploader("Sube Excel FACTURA (recibido / facturación)", type=["xlsx", "xls"])

if not f_base or not f_fact:
    st.stop()

# Leer excels
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
    c_pais = st.selectbox("Pais", df_base_raw.columns, index=0)
    c_semana = st.selectbox("Semana", df_base_raw.columns, index=1 if len(df_base_raw.columns) > 1 else 0)
    c_nombre = st.selectbox("Nombre", df_base_raw.columns, index=2 if len(df_base_raw.columns) > 2 else 0)
    c_editorial = st.selectbox("Editorial", df_base_raw.columns, index=3 if len(df_base_raw.columns) > 3 else 0)
    c_cantidad = st.selectbox("Cantidad", df_base_raw.columns, index=4 if len(df_base_raw.columns) > 4 else 0)

with col_map2:
    st.markdown("**FACTURA (recibido)**")
    f_titulo = st.selectbox("Columna título (Nombre o 'Nombre | Editorial País')", df_fact_raw.columns, index=0)
    f_cantidad = st.selectbox("Columna cantidad", df_fact_raw.columns, index=1 if len(df_fact_raw.columns) > 1 else 0)
    f_pvp = st.selectbox("Columna PVP (opcional)", ["(no)"] + list(df_fact_raw.columns), index=0)
    f_total = st.selectbox("Columna Total (recomendado)", ["(no)"] + list(df_fact_raw.columns), index=0)

match_mode = st.radio(
    "Modo de match",
    ["Pais + Nombre (recomendado)", "Pais + Editorial + Nombre (más estricto)", "Solo Nombre (riesgoso si hay homónimos)"],
    index=0,
    horizontal=True
)

st.divider()

# Parsear
try:
    base = parse_base(df_base_raw, c_pais, c_semana, c_nombre, c_editorial, c_cantidad)

    fact = parse_factura_externa(
        df_fact_raw,
        col_titulo=f_titulo,
        col_cantidad=f_cantidad,
        col_pvp=None if f_pvp == "(no)" else f_pvp,
        col_total=None if f_total == "(no)" else f_total
    )
except Exception as e:
    st.error(f"Error leyendo / parseando: {e}")
    st.stop()

# Validaciones factura
fact_errors = []
if (fact["Pais_norm"] == "").any():
    fact_errors.append("Hay filas en FACTURA con País vacío (no se detectó encabezado tipo 'Editorial País').")
if (fact["Editorial_norm"] == "").any():
    fact_errors.append("Hay filas en FACTURA con Editorial vacía (no se detectó 'Editorial País' o 'Nombre | Editorial País').")
if fact["Nombre_norm"].eq("").any():
    fact_errors.append("Hay filas en FACTURA con Nombre vacío.")

# Claves de match
if match_mode == "Pais + Nombre (recomendado)":
    key_cols = ["Pais_norm", "Nombre_norm"]
elif match_mode == "Pais + Editorial + Nombre (más estricto)":
    key_cols = ["Pais_norm", "Editorial_norm", "Nombre_norm"]
else:
    key_cols = ["Nombre_norm"]

base_key = base.copy()
fact_key = fact.copy()

# merge principal
merged = pd.merge(
    base_key,
    fact_key,
    on=key_cols,
    how="outer",
    suffixes=("_base", "_fact"),
    indicator=True
)

# Completar columnas visibles
def pick_col(df, base_col, fact_col, default=""):
    out = df[base_col].copy() if base_col in df.columns else pd.Series([default] * len(df))
    if fact_col in df.columns:
        out = out.fillna(df[fact_col])
    return out.fillna(default)

# Armar reporte
rep = pd.DataFrame()
rep["Pais"] = pick_col(merged, "Pais_base", "Pais_fact")
rep["Editorial"] = pick_col(merged, "Editorial_base", "Editorial_fact")
rep["Nombre"] = pick_col(merged, "Nombre_base", "Nombre_fact")

rep["Cantidad_pedida"] = merged.get("Cantidad_pedida", 0).fillna(0).astype(int)
rep["Cantidad_facturada"] = merged.get("Cantidad_facturada", 0).fillna(0).astype(int)
rep["Diferencia_cantidad"] = rep["Cantidad_facturada"] - rep["Cantidad_pedida"]

rep["PVP_factura"] = merged.get("PVP_factura", np.nan)
rep["Total_factura"] = merged.get("Total_factura", np.nan)

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
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ítems OK", int((rep["Estado"] == "OK (coincide)").sum()))
c2.metric("Ítems faltantes", int(rep["Estado"].isin(["FALTANTE (llegó menos)", "NO LLEGÓ (pedido sin factura)"]).sum()))
c3.metric("Ítems sobrantes / no pedidos", int(rep["Estado"].isin(["SOBRANTE (llegó más)", "NO PEDIDO (factura sin pedido)"]).sum()))
c4.metric("Total líneas", len(rep))

st.subheader("3) Errores detectados (calidad de datos)")
if fact_errors:
    for e in fact_errors:
        st.warning(e)
else:
    st.success("No se detectaron problemas obvios en el formato de FACTURA.")

# Total a pagar por país (según factura)
st.subheader("4) Cuánto debe pagar cada país (según factura)")
if rep["Total_factura"].notna().any():
    pago_pais = rep.dropna(subset=["Total_factura"]).groupby("Pais", as_index=False)["Total_factura"].sum()
    pago_pais = pago_pais.sort_values("Total_factura", ascending=False)
    st.dataframe(pago_pais, use_container_width=True)
else:
    st.info("No hay columna Total_factura disponible (o viene vacía). Selecciona la columna 'Total' en el mapeo de FACTURA para calcular pagos por país.")

# Tablas principales con filtros
st.subheader("5) Detalle del cruce")
filtro = st.multiselect(
    "Filtrar por estado",
    sorted(rep["Estado"].unique().tolist()),
    default=sorted(rep["Estado"].unique().tolist())
)
rep_view = rep[rep["Estado"].isin(filtro)].copy()
st.dataframe(rep_view, use_container_width=True, height=520)

# Hojas separadas
ok = rep[rep["Estado"] == "OK (coincide)"].copy()
faltantes = rep[rep["Estado"].isin(["FALTANTE (llegó menos)", "NO LLEGÓ (pedido sin factura)"])].copy()
sobrantes = rep[rep["Estado"].isin(["SOBRANTE (llegó más)", "NO PEDIDO (factura sin pedido)"])].copy()

# Descargar reporte
st.subheader("6) Descargar reporte")
sheets = {
    "Cruce_completo": rep,
    "OK": ok,
    "Faltantes": faltantes,
    "Sobrantes_NoPedido": sobrantes
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

st.caption("Tip: Si te aparecen muchos 'NO PEDIDO', revisa el modo de match o si el País/Editorial no se detectó bien en la factura.")

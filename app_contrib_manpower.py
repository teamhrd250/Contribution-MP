# app_contrib_manpower.py
# Streamlit Dashboard: Kontribusi Manpower Terhadap Profit

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Kontribusi Manpower → Profit", page_icon="💼", layout="wide")

# ==============================
# ---------- UTILITIES ----------
# ==============================

@st.cache_data(show_spinner=False)
def load_excel(file):
    try:
        xl = pd.ExcelFile(file)
        df = xl.parse(xl.sheet_names[0], header=None)
        return df
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return None

def find_header_row(df_raw):
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if any("nama karyawan" in str(x) for x in row):
            return i
    return 3  # fallback

def clean_dataframe(df_raw):
    header_row = find_header_row(df_raw)
    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)

    # Rename kolom
    rename_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if "nama karyawan" in cl:
            rename_map[c] = "Nama Karyawan"
        elif "gaji" in cl:
            rename_map[c] = "Gaji Tahunan (Rp)"
        elif "bop" in cl:
            rename_map[c] = "BOP MP/tahun (Rp)"
        elif "tax" in cl or "pensiun" in cl:
            rename_map[c] = "TAX + PENSIUN/tahun"
        elif "total cost" in cl:
            rename_map[c] = "Total Cost per-MP (Rp)"
        elif "profit per" in cl:
            rename_map[c] = "Profit per-MP (Rp)"
        elif "rasio efisiensi" in cl:
            rename_map[c] = "Rasio Efisiensi"
        elif "range efisiensi" in cl:
            rename_map[c] = "Range Efisiensi"
        elif "revenue/mp" in cl:
            rename_map[c] = "Revenue/MP (dynamic)"
        elif "kontribusi" in cl:
            rename_map[c] = "Kontribusi to Profit (%)"
    df = df.rename(columns=rename_map)

    # Keep hanya baris valid
    if "Nama Karyawan" in df.columns:
        df = df[df["Nama Karyawan"].notna()]

    # Konversi angka
    num_cols = [
        "Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun",
        "Total Cost per-MP (Rp)", "Profit per-MP (Rp)", "Revenue/MP (dynamic)",
        "Rasio Efisiensi", "Kontribusi to Profit (%)"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^0-9\.-]', '', regex=True), errors="coerce")

    # Hitung tambahan
    if "Total Cost per-MP (Rp)" not in df.columns:
        if all(col in df.columns for col in ["Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun"]):
            df["Total Cost per-MP (Rp)"] = df[["Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun"]].sum(axis=1)

    if "Profit per-MP (Rp)" not in df.columns and all(c in df.columns for c in ["Revenue/MP (dynamic)", "Total Cost per-MP (Rp)"]):
        df["Profit per-MP (Rp)"] = df["Revenue/MP (dynamic)"] - df["Total Cost per-MP (Rp)"]

    if "Rasio Efisiensi" not in df.columns and all(c in df.columns for c in ["Profit per-MP (Rp)", "Total Cost per-MP (Rp)"]):
        df["Rasio Efisiensi"] = df["Profit per-MP (Rp)"] / df["Total Cost per-MP (Rp)"]

    if "Range Efisiensi" not in df.columns and "Rasio Efisiensi" in df.columns:
        df["Range Efisiensi"] = np.where(df["Rasio Efisiensi"] >= 1, "Efisien", "Tidak Efisien")

    if "Kontribusi to Profit (%)" not in df.columns and "Profit per-MP (Rp)" in df.columns:
        total_profit = df["Profit per-MP (Rp)"].sum()
        if total_profit != 0:
            df["Kontribusi to Profit (%)"] = (df["Profit per-MP (Rp)"] / total_profit) * 100

    return df

def kpi_card(label, value):
    with st.container(border=True):
        st.markdown(f"**{label}**")
        st.markdown(f"<div style='font-size:28px;font-weight:700;margin-top:-6px'>{value}</div>", unsafe_allow_html=True)

def currency_idr(x):
    try:
        return "Rp{:,.0f}".format(float(x)).replace(",", ".")
    except:
        return "-"

# ==============================
# ------------ UI --------------
# ==============================

st.title("💼 Kontribusi Manpower → Profit")
st.caption("Analitik biaya & efisiensi manpower terhadap profit perusahaan")

with st.sidebar:
    st.header("⚙️ Pengaturan & Data")
    uploaded = st.file_uploader("Upload data Excel", type=["xlsx", "xls"])
    st.subheader("Filter")
    show_only_efficient = st.checkbox("Hanya tampilkan yang Efisien (Rasio ≥ 1)", value=False)
    min_profit = st.number_input("Minimal Profit per-MP (Rp)", value=0, step=1_000_000)

# Load data (upload only)
if uploaded is not None:
    raw = load_excel(uploaded)
else:
    raw = None

if raw is None:
    st.warning("Data belum tersedia. Upload file Excel terlebih dahulu.")
    st.stop()

df = clean_dataframe(raw)

if df.empty or "Nama Karyawan" not in df.columns:
    st.error("Struktur data tidak sesuai.")
    st.stop()

# Filter
if "Profit per-MP (Rp)" in df.columns:
    df = df[df["Profit per-MP (Rp)"].fillna(0) >= min_profit]
if show_only_efficient and "Range Efisiensi" in df.columns:
    df = df[df["Range Efisiensi"] == "Efisien"]

# ==============================
# ---------- METRICS -----------
# ==============================
col1, col2, col3 = st.columns(3)
with col1:
    kpi_card("Total Manpower Cost", currency_idr(df["Total Cost per-MP (Rp)"].sum()))
with col2:
    kpi_card("Total Profit", currency_idr(df["Profit per-MP (Rp)"].sum()))
with col3:
    kpi_card("Jumlah Karyawan", f"{df['Nama Karyawan'].nunique()} orang")

st.markdown("### 📈 Visualisasi")

if "Profit per-MP (Rp)" in df.columns:
    fig = px.bar(df.sort_values("Profit per-MP (Rp)", ascending=False),
                 x="Nama Karyawan", y="Profit per-MP (Rp)", title="Profit per Karyawan")
    st.plotly_chart(fig, use_container_width=True)

if "Revenue/MP (dynamic)" in df.columns and "Profit per-MP (Rp)" in df.columns:
    fig_scatter = px.scatter(df, x="Revenue/MP (dynamic)", y="Profit per-MP (Rp)",
                             color="Range Efisiensi", hover_name="Nama Karyawan",
                             size="Total Cost per-MP (Rp)", title="Revenue vs Profit")
    st.plotly_chart(fig_scatter, use_container_width=True)

st.markdown("### 📋 Data Tabel")
st.dataframe(df)

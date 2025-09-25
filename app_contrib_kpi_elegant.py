# app_contrib_kpi_elegant.py
# KPI Dashboard: Kontribusi Manpower → Profit (Elegant + Excel-Style Table)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="KPI Manpower → Profit", page_icon="💼", layout="wide")

# ==============================
# ---------- THEME -------------
# ==============================
st.markdown("""
<style>
.stApp { background: radial-gradient(1200px 600px at 10% 10%, #16222A 0%, #3A6073 60%); color: #f6f7fb; }
.block-container { padding-top: 1rem; padding-bottom: 2rem; }
h1,h2,h3,h4,h5,h6 { color: #f6f7fb; }
.kpi-card {
  background: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.18);
  border-radius: 16px;
  padding: 14px 16px;
  box-shadow: 0 10px 30px rgba(0,0,0,0.15);
  backdrop-filter: blur(6px);
}
.kpi-label { font-size: 13px; font-weight: 600; color: #d7dbe6; }
.kpi-value { font-size: 28px; font-weight: 800; color: #fff; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)

# ==============================
# ---------- UTILS -------------
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
    for i in range(min(25, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if any("nama karyawan" in str(x) for x in row):
            return i
    return 3

def clean_dataframe(df_raw):
    header_row = find_header_row(df_raw)
    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)

    rename_map = {}
    for c in df.columns:
        cl_low = str(c).strip().lower()
        if "nama" in cl_low and "karyawan" in cl_low: rename_map[c] = "Nama Karyawan"
        elif "gaji" in cl_low: rename_map[c] = "Gaji Tahunan (Rp)"
        elif "bop" in cl_low: rename_map[c] = "BOP MP/tahun (Rp)"
        elif "tax" in cl_low or "pensiun" in cl_low: rename_map[c] = "TAX + PENSIUN/tahun"
        elif "total cost" in cl_low: rename_map[c] = "Total Cost per-MP (Rp)"
        elif "profit per" in cl_low: rename_map[c] = "Profit per-MP (Rp)"
        elif "rasio efisiensi" in cl_low: rename_map[c] = "Rasio Efisiensi"
        elif "range efisiensi" in cl_low: rename_map[c] = "Range Efisiensi"
        elif "revenue/mp" in cl_low: rename_map[c] = "Revenue/MP (dynamic)"
        elif "kontribusi" in cl_low: rename_map[c] = "Kontribusi to Profit (%)"
    df = df.rename(columns=rename_map)

    num_cols = [
        "Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun",
        "Total Cost per-MP (Rp)", "Profit per-MP (Rp)", "Revenue/MP (dynamic)",
        "Rasio Efisiensi", "Kontribusi to Profit (%)"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^0-9\.-]', '', regex=True), errors="coerce")

    return df

def kpi_card(label, value):
    st.markdown(f"<div class='kpi-card'><div class='kpi-label'>{label}</div><div class='kpi-value'>{value}</div></div>", unsafe_allow_html=True)

def currency_idr(x):
    try: return "Rp{:,.0f}".format(float(x)).replace(",", ".")
    except: return "-"

# ==============================
# ------------ UI --------------
# ==============================
st.title("💼 KPI Dashboard — Kontribusi Manpower → Profit")
st.caption("Elegant Edition • Styled Table • Siap Presentasi")

with st.sidebar:
    st.header("📥 Data & Filter")
    uploaded = st.file_uploader("Upload data Excel", type=["xlsx", "xls"])
    search_name = st.text_input("Cari karyawan (opsional)", "")
    show_only_efficient = st.checkbox("Hanya yang Efisien (Rasio ≥ 1)", value=False)
    min_profit = st.number_input("Minimal Profit per-MP (Rp)", value=0, step=1_000_000)

if uploaded is None:
    st.warning("Upload file Excel dulu.")
    st.stop()

raw = load_excel(uploaded)
if raw is None: st.stop()
df = clean_dataframe(raw)
if df.empty or "Nama Karyawan" not in df.columns:
    st.error("Struktur data tidak sesuai.")
    st.stop()

if search_name:
    df = df[df["Nama Karyawan"].str.contains(search_name, case=False, na=False)]
if "Profit per-MP (Rp)" in df.columns:
    df = df[df["Profit per-MP (Rp)"].fillna(0) >= min_profit]
if show_only_efficient and "Range Efisiensi" in df.columns:
    df = df[df["Range Efisiensi"].str.contains("Efisien", case=False, na=False)]

# ==============================
# ---------- METRICS -----------
# ==============================
col1, col2, col3, col4, col5 = st.columns(5)
with col1: kpi_card("Total Revenue", currency_idr(df["Revenue/MP (dynamic)"].sum()))
with col2: kpi_card("Total Cost", currency_idr(df["Total Cost per-MP (Rp)"].sum()))
with col3: kpi_card("Total Profit", currency_idr(df["Profit per-MP (Rp)"].sum()))
with col4: kpi_card("Avg Efisiensi", f"{df['Rasio Efisiensi'].mean():.2f}")
with col5: kpi_card("Jumlah Karyawan", f"{df['Nama Karyawan'].nunique()}")

# ==============================
# ---------- TABS --------------
# ==============================
tab1, tab2, tab3 = st.tabs(["📈 Summary", "🔎 Detail", "📋 Data Tabel"])

with tab1:
    st.subheader("Profit per Karyawan")
    fig = px.bar(df.sort_values("Profit per-MP (Rp)", ascending=False),
                 x="Nama Karyawan", y="Profit per-MP (Rp)",
                 color="Range Efisiensi",
                 color_discrete_map={"Efisien":"#2ecc71","Tidak Efisien":"#e74c3c","Cukup Efisien":"#f1c40f"})
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Revenue vs Profit")
    fig2 = px.scatter(df, x="Revenue/MP (dynamic)", y="Profit per-MP (Rp)",
                      size="Total Cost per-MP (Rp)", color="Range Efisiensi",
                      hover_name="Nama Karyawan")
    st.plotly_chart(fig2, use_container_width=True)

with tab2:
    selected = st.selectbox("Pilih karyawan:", df["Nama Karyawan"].unique())
    drow = df[df["Nama Karyawan"]==selected].iloc[0]
    c1,c2,c3 = st.columns(3)
    with c1: kpi_card("Total Cost", currency_idr(drow["Total Cost per-MP (Rp)"]))
    with c2: kpi_card("Profit", currency_idr(drow["Profit per-MP (Rp)"]))
    with c3: kpi_card("Efisiensi", f"{drow['Rasio Efisiensi']:.2f}" if "Rasio Efisiensi" in drow else "-")
    pie_cols = [c for c in ["Gaji Tahunan (Rp)","BOP MP/tahun (Rp)","TAX + PENSIUN/tahun"] if c in df.columns]
    if len(pie_cols)>=2:
        pie_df = pd.DataFrame({"Komponen":pie_cols,"Biaya":[drow[c] for c in pie_cols]})
        fig_pie = px.pie(pie_df, names="Komponen", values="Biaya", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

with tab3:
    st.subheader("📋 Data Lengkap")
    df_display = df.fillna("")  # fix error NaN
    styled = df_display.style.format({
        "Gaji Tahunan (Rp)": lambda x: currency_idr(x),
        "BOP MP/tahun (Rp)": lambda x: currency_idr(x),
        "TAX + PENSIUN/tahun": lambda x: currency_idr(x),
        "Total Cost per-MP (Rp)": lambda x: currency_idr(x),
        "Revenue/MP (dynamic)": lambda x: currency_idr(x),
        "Profit per-MP (Rp)": lambda x: currency_idr(x),
        "Rasio Efisiensi": "{:.2f}",
        "Kontribusi to Profit (%)": "{:.2f}%"
    }).background_gradient(
        subset=["Profit per-MP (Rp)"], cmap="Greens"
    ).background_gradient(
        subset=["Rasio Efisiensi"], cmap="RdYlGn"
    )

    with st.expander("Klik untuk lihat tabel detail ➡️"):
        st.dataframe(styled, use_container_width=True)

    # ringkasan total
    st.markdown("### 📊 Ringkasan Total")
    st.write({
        "Total Revenue": currency_idr(df["Revenue/MP (dynamic)"].sum()),
        "Total Cost": currency_idr(df["Total Cost per-MP (Rp)"].sum()),
        "Total Profit": currency_idr(df["Profit per-MP (Rp)"].sum()),
        "Rata-rata Efisiensi": f"{df['Rasio Efisiensi'].mean():.2f}"
    })

    # download
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CSV", data=csv, file_name="kpi_manpower.csv", mime="text/csv")

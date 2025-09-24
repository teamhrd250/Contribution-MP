
# app_contrib_manpower_mantap.py
# Streamlit Dashboard: Kontribusi Manpower Terhadap Profit (Mantap Jiwa Edition)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Manpower Contribution Dashboard", page_icon="💼", layout="wide")

# ==============================
# ---------- STYLE -------------
# ==============================

st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        color: #fff;
    }
    .block-container {padding-top:1rem; padding-bottom:2rem;}
    h1,h2,h3,h4,h5,h6 {color: #f5f5f5;}
    .metric-card {
        background: rgba(255,255,255,0.1);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 14px;
        padding: 16px;
        text-align: center;
        margin: 5px;
    }
    .metric-label {font-size:14px; font-weight:600; color:#ddd;}
    .metric-value {font-size:26px; font-weight:700; color:#fff; margin-top:4px;}
    </style>
""", unsafe_allow_html=True)

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
    return 3

def clean_dataframe(df_raw):
    header_row = find_header_row(df_raw)
    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)

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

    if "Nama Karyawan" in df.columns:
        df = df[df["Nama Karyawan"].notna()]

    num_cols = [
        "Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun",
        "Total Cost per-MP (Rp)", "Profit per-MP (Rp)", "Revenue/MP (dynamic)",
        "Rasio Efisiensi", "Kontribusi to Profit (%)"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^0-9\.-]', '', regex=True), errors="coerce")

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
    st.markdown(
        f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div></div>",
        unsafe_allow_html=True
    )

def currency_idr(x):
    try:
        return "Rp{:,.0f}".format(float(x)).replace(",", ".")
    except:
        return "-"

# ==============================
# ------------ UI --------------
# ==============================

st.title("💼 Dashboard Kontribusi Manpower → Profit")
st.caption("Edisi Mantap Jiwa 🚀 | Analitik biaya & efisiensi manpower terhadap profit perusahaan")

with st.sidebar:
    st.header("⚙️ Data & Filter")
    uploaded = st.file_uploader("Upload data Excel", type=["xlsx", "xls"])
    show_only_efficient = st.checkbox("Hanya Efisien (Rasio ≥ 1)", value=False)
    min_profit = st.number_input("Minimal Profit per-MP (Rp)", value=0, step=1_000_000)

if uploaded is None:
    st.warning("Upload file Excel dulu untuk mulai analisis.")
    st.stop()

raw = load_excel(uploaded)
if raw is None:
    st.stop()

df = clean_dataframe(raw)
if df.empty or "Nama Karyawan" not in df.columns:
    st.error("Struktur data tidak sesuai.")
    st.stop()

if "Profit per-MP (Rp)" in df.columns:
    df = df[df["Profit per-MP (Rp)"].fillna(0) >= min_profit]
if show_only_efficient and "Range Efisiensi" in df.columns:
    df = df[df["Range Efisiensi"] == "Efisien"]

# ==============================
# ---------- LAYOUT ------------
# ==============================

colL, colR = st.columns([2,2])

with colL:
    st.subheader("📊 KPI Summary")
    c1, c2, c3 = st.columns(3)
    with c1: kpi_card("Total Cost", currency_idr(df["Total Cost per-MP (Rp)"].sum()))
    with c2: kpi_card("Total Profit", currency_idr(df["Profit per-MP (Rp)"].sum()))
    with c3: kpi_card("Avg Efisiensi", f"{df['Rasio Efisiensi'].mean():.2f}" if "Rasio Efisiensi" in df.columns else "-")

with colR:
    st.subheader("🏆 Top 5 Contributor")
    if "Kontribusi to Profit (%)" in df.columns:
        top5 = df.sort_values("Kontribusi to Profit (%)", ascending=False).head(5)
        st.dataframe(top5[["Nama Karyawan","Profit per-MP (Rp)","Kontribusi to Profit (%)"]].style.format({
            "Profit per-MP (Rp)": lambda x: currency_idr(x),
            "Kontribusi to Profit (%)": "{:.2f}%"
        }), use_container_width=True)

# ==============================
# ---------- TABS --------------
# ==============================
tab1, tab2, tab3 = st.tabs(["📈 Visualisasi", "🔎 Detail Per Karyawan", "📋 Data Tabel"])

with tab1:
    st.markdown("### Profit per Karyawan")
    if "Profit per-MP (Rp)" in df.columns:
        fig = px.bar(df.sort_values("Profit per-MP (Rp)", ascending=False),
                     x="Nama Karyawan", y="Profit per-MP (Rp)", color="Range Efisiensi",
                     color_discrete_map={"Efisien":"#2ecc71","Tidak Efisien":"#e74c3c"})
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Revenue vs Profit")
    if "Revenue/MP (dynamic)" in df.columns and "Profit per-MP (Rp)" in df.columns:
        fig_scatter = px.scatter(df, x="Revenue/MP (dynamic)", y="Profit per-MP (Rp)",
                                 size="Total Cost per-MP (Rp)", color="Range Efisiensi",
                                 hover_name="Nama Karyawan")
        st.plotly_chart(fig_scatter, use_container_width=True)

with tab2:
    selected = st.selectbox("Pilih karyawan:", df["Nama Karyawan"].unique())
    drow = df[df["Nama Karyawan"]==selected].iloc[0]
    c1, c2, c3 = st.columns(3)
    with c1: kpi_card("Total Cost", currency_idr(drow["Total Cost per-MP (Rp)"]))
    with c2: kpi_card("Profit", currency_idr(drow["Profit per-MP (Rp)"]))
    with c3: kpi_card("Efisiensi", f"{drow['Rasio Efisiensi']:.2f}" if "Rasio Efisiensi" in drow else "-")
    pie_cols = [c for c in ["Gaji Tahunan (Rp)","BOP MP/tahun (Rp)","TAX + PENSIUN/tahun"] if c in df.columns]
    if len(pie_cols)>=2:
        pie_df = pd.DataFrame({"Komponen":pie_cols,"Biaya":[drow[c] for c in pie_cols]})
        fig_pie = px.pie(pie_df, names="Komponen", values="Biaya", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

with tab3:
    st.dataframe(df, use_container_width=True)
    st.download_button("⬇️ Download CSV", data=df.to_csv(index=False).encode("utf-8"),
                       file_name="contrib_manpower.csv", mime="text/csv")

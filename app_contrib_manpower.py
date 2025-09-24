
# app_contrib_manpower.py
# Streamlit Dashboard: Kontribusi Manpower Terhadap Profit
# Author: ChatGPT (for Arie Hermansyah - HR Project)
# How to run: streamlit run app_contrib_manpower.py

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
    # Cari baris yang memuat string "Nama Karyawan" sebagai header
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if any("nama karyawan" in str(x) for x in row):
            return i
    # fallback: baris ke-3 seperti sampel
    return 3

def clean_dataframe(df_raw):
    header_row = find_header_row(df_raw)
    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)

    # Normalisasi nama kolom
    rename_map = {}
    for c in df.columns:
        cl = str(c).strip()
        cl_low = cl.lower()
        if "nama" in cl_low and "karyawan" in cl_low:
            rename_map[c] = "Nama Karyawan"
        elif "gaji" in cl_low and "tahun" in cl_low:
            rename_map[c] = "Gaji Tahunan (Rp)"
        elif "bop" in cl_low:
            rename_map[c] = "BOP MP/tahun (Rp)"
        elif ("tax" in cl_low or "pensiun" in cl_low) and "tahun" in cl_low:
            rename_map[c] = "TAX + PENSIUN/tahun"
        elif "total cost" in cl_low:
            rename_map[c] = "Total Cost per-MP (Rp)"
        elif "profit per" in cl_low:
            rename_map[c] = "Profit per-MP (Rp)"
        elif "rasio efisiensi" in cl_low:
            rename_map[c] = "Rasio Efisiensi"
        elif "range efisiensi" in cl_low:
            rename_map[c] = "Range Efisiensi"
        elif "revenue/mp" in cl_low:
            rename_map[c] = "Revenue/MP (dynamic)"
        elif "kontribusi" in cl_low and "profit" in cl_low:
            rename_map[c] = "Kontribusi to Profit (%)"
        elif "total revenue perusahaan" in cl_low:
            rename_map[c] = "Total Revenue Perusahaan"
        elif "jumlah karyawan" in cl_low:
            rename_map[c] = "Jumlah Karyawan"
        else:
            # biarkan kolom yang tidak dikenal
            pass

    df = df.rename(columns=rename_map)

    # Keep only rows yang punya Nama Karyawan (menghindari baris input/notes di bawah)
    if "Nama Karyawan" in df.columns:
        df = df[df["Nama Karyawan"].notna() & (df["Nama Karyawan"].astype(str).str.strip() != "")]

    # Konversi numeric
    numeric_cols = [
        "Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun",
        "Total Cost per-MP (Rp)", "Profit per-MP (Rp)", "Revenue/MP (dynamic)",
        "Rasio Efisiensi", "Kontribusi to Profit (%)"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^0-9\.-]', '', regex=True), errors="coerce")

    # Hitung kolom turunan jika perlu
    if "Total Cost per-MP (Rp)" not in df.columns:
        need_cols = ["Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun"]
        if all(col in df.columns for col in need_cols):
            df["Total Cost per-MP (Rp)"] = df[need_cols].sum(axis=1)

    if "Profit per-MP (Rp)" not in df.columns and "Revenue/MP (dynamic)" in df.columns and "Total Cost per-MP (Rp)" in df.columns:
        df["Profit per-MP (Rp)"] = df["Revenue/MP (dynamic)"] - df["Total Cost per-MP (Rp)"]

    if "Rasio Efisiensi" not in df.columns and "Profit per-MP (Rp)" in df.columns and "Total Cost per-MP (Rp)" in df.columns:
        df["Rasio Efisiensi"] = df["Profit per-MP (Rp)"] / df["Total Cost per-MP (Rp)"]

    if "Range Efisiensi" not in df.columns and "Rasio Efisiensi" in df.columns:
        df["Range Efisiensi"] = np.where(df["Rasio Efisiensi"] >= 1, "Efisien", "Tidak Efisien")

    # Kontribusi ke profit (persen) = Profit per-MP / Total Profit (sum profit) * 100
    if "Kontribusi to Profit (%)" not in df.columns and "Profit per-MP (Rp)" in df.columns:
        total_profit = df["Profit per-MP (Rp)"].sum(skipna=True)
        if total_profit and total_profit != 0:
            df["Kontribusi to Profit (%)"] = (df["Profit per-MP (Rp)"] / total_profit) * 100

    return df

def kpi_card(label, value, help_text=None):
    with st.container(border=True):
        st.markdown(f"**{label}**")
        st.markdown(f"<div style='font-size:28px;font-weight:700;margin-top:-6px'>{value}</div>", unsafe_allow_html=True)
        if help_text:
            st.caption(help_text)

def currency_idr(x):
    try:
        return "Rp{:,.0f}".format(float(x)).replace(",", ".")
    except:
        return "-"


# ==============================
# ------------ UI --------------
# ==============================

st.markdown(
    """
    <style>
    /* Soft glassmorphism look */
    .stApp { background: radial-gradient(1200px 600px at 10% 10%, #f0f4ff 0%, #ffffff 60%) }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    .metric-card { background: rgba(255,255,255,0.7); border: 1px solid rgba(0,0,0,0.05); border-radius: 18px; padding: 14px 16px; }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("💼 Kontribusi Manpower → Profit")
st.caption("Analitik biaya & efisiensi manpower terhadap profit perusahaan")

with st.sidebar:
    st.header("⚙️ Pengaturan & Data")
    uploaded = st.file_uploader("Upload data Excel", type=["xlsx", "xls"])
    st.markdown("---")
    st.subheader("Filter")
    show_only_efficient = st.checkbox("Hanya tampilkan yang Efisien (Rasio ≥ 1)", value=False)
    min_profit = st.number_input("Minimal Profit per-MP (Rp)", value=0, step=1_000_000)
    st.markdown("---")
    st.caption("Tips: Kosongkan untuk pakai file contoh atau gunakan file asli Anda.")

# Load data (pakai file contoh jika tidak upload)
if uploaded is not None:
    raw = load_excel(uploaded)
else:
    # fallback: load dari lokasi default (sesuai file yang diupload user sebelumnya)
    try:
        raw = load_excel("/mnt/data/CONTRIBUTION MANPOWER TO PROFIT (1).xlsx")
    except:
        raw = None

if raw is None:
    st.error("Data belum tersedia. Upload file Excel terlebih dahulu.")
    st.stop()

df = clean_dataframe(raw)

if df.empty or "Nama Karyawan" not in df.columns:
    st.error("Struktur data tidak sesuai. Pastikan ada kolom 'Nama Karyawan' dan kolom biaya/profit.")
    st.dataframe(df.head(20))
    st.stop()

# Terapkan filter sederhana
if "Profit per-MP (Rp)" in df.columns:
    df = df[df["Profit per-MP (Rp)"].fillna(0) >= min_profit]
if show_only_efficient and "Range Efisiensi" in df.columns:
    df = df[df["Range Efisiensi"] == "Efisien"]

# ==============================
# ---------- METRICS -----------
# ==============================
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    total_cost = df.get("Total Cost per-MP (Rp)", pd.Series(dtype=float)).sum(skipna=True)
    kpi_card("Total Manpower Cost", currency_idr(total_cost))
with col2:
    total_profit = df.get("Profit per-MP (Rp)", pd.Series(dtype=float)).sum(skipna=True)
    kpi_card("Total Profit (sum MP)", currency_idr(total_profit))
with col3:
    avg_profit = df.get("Profit per-MP (Rp)", pd.Series(dtype=float)).mean(skipna=True)
    kpi_card("Avg Profit per-MP", currency_idr(avg_profit))
with col4:
    avg_ratio = df.get("Rasio Efisiensi", pd.Series(dtype=float)).mean(skipna=True)
    kpi_card("Rasio Efisiensi (avg)", f"{avg_ratio:,.2f}" if pd.notna(avg_ratio) else "-")
with col5:
    count_emp = df["Nama Karyawan"].nunique()
    kpi_card("Jumlah Karyawan (records)", f"{count_emp} orang")

st.markdown("### 📈 Visualisasi & Analitik")

# ------------------ Bar: Profit per Karyawan ------------------
if "Profit per-MP (Rp)" in df.columns:
    fig = px.bar(
        df.sort_values("Profit per-MP (Rp)", ascending=False),
        x="Nama Karyawan",
        y="Profit per-MP (Rp)",
        text="Profit per-MP (Rp)",
        title="Profit per Karyawan",
    )
    fig.update_traces(texttemplate='%{text:.2s}', textposition='outside', hovertemplate="%{x}<br>Profit: %{y:,.0f}")
    fig.update_layout(xaxis_title="", yaxis_title="Rupiah", margin=dict(t=60, b=50))
    st.plotly_chart(fig, use_container_width=True)

# ------------------ Stacked Bar: Komposisi Cost ------------------
cost_cols = [c for c in ["Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun"] if c in df.columns]
if len(cost_cols) >= 2:
    cost_df = df[["Nama Karyawan"] + cost_cols].melt(id_vars=["Nama Karyawan"], var_name="Komponen", value_name="Biaya")
    fig_cost = px.bar(
        cost_df, x="Nama Karyawan", y="Biaya", color="Komponen",
        title="Komposisi Biaya per Karyawan (Gaji, BOP, Tax)",
    )
    fig_cost.update_layout(xaxis_title="", yaxis_title="Rupiah", barmode="stack", margin=dict(t=60, b=50))
    st.plotly_chart(fig_cost, use_container_width=True)

# ------------------ Scatter: Revenue vs Profit ------------------
if "Revenue/MP (dynamic)" in df.columns and "Profit per-MP (Rp)" in df.columns:
    fig_scatter = px.scatter(
        df, x="Revenue/MP (dynamic)", y="Profit per-MP (Rp)",
        size=df.get("Total Cost per-MP (Rp)", pd.Series(np.ones(len(df)))),
        color="Range Efisiensi" if "Range Efisiensi" in df.columns else None,
        hover_name="Nama Karyawan",
        title="Revenue vs Profit per Karyawan (bubble ~ Total Cost)"
    )
    fig_scatter.update_layout(xaxis_title="Revenue/MP", yaxis_title="Profit/MP", margin=dict(t=60, b=50))
    st.plotly_chart(fig_scatter, use_container_width=True)

# ------------------ Top Contributors Table ------------------
if "Kontribusi to Profit (%)" in df.columns:
    top5 = df.sort_values("Kontribusi to Profit (%)", ascending=False).head(5).copy()
    st.markdown("### 🏆 Top 5 Kontributor ke Profit")
    st.dataframe(top5[["Nama Karyawan", "Profit per-MP (Rp)", "Kontribusi to Profit (%)"]].style.format({
        "Profit per-MP (Rp)": lambda x: currency_idr(x),
        "Kontribusi to Profit (%)": "{:.2f}%"
    }), use_container_width=True)

# ------------------ Detail per Karyawan ------------------
st.markdown("### 🔎 Detail Breakdown per Karyawan")
selected_name = st.selectbox("Pilih karyawan untuk melihat breakdown:", options=df["Nama Karyawan"].unique())
detail_cols = [c for c in [
    "Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun",
    "Total Cost per-MP (Rp)", "Revenue/MP (dynamic)", "Profit per-MP (Rp)",
    "Rasio Efisiensi", "Range Efisiensi", "Kontribusi to Profit (%)"
] if c in df.columns]
detail_row = df[df["Nama Karyawan"] == selected_name][["Nama Karyawan"] + detail_cols].reset_index(drop=True)

# Mini cards
c1, c2, c3, c4 = st.columns(4)
if "Total Cost per-MP (Rp)" in df.columns:
    with c1: kpi_card("Total Cost", currency_idr(detail_row.at[0, "Total Cost per-MP (Rp)"]))
if "Profit per-MP (Rp)" in df.columns:
    with c2: kpi_card("Profit", currency_idr(detail_row.at[0, "Profit per-MP (Rp)"]))
if "Rasio Efisiensi" in df.columns:
    with c3: kpi_card("Rasio Efisiensi", f"{detail_row.at[0, 'Rasio Efisiensi']:.2f}")
if "Kontribusi to Profit (%)" in df.columns:
    with c4: kpi_card("Kontribusi → Profit", f"{detail_row.at[0, 'Kontribusi to Profit (%)']:.2f}%")

# Pie breakdown
pie_cols = [c for c in ["Gaji Tahunan (Rp)", "BOP MP/tahun (Rp)", "TAX + PENSIUN/tahun"] if c in df.columns]
if len(pie_cols) >= 2:
    pie_df = detail_row.melt(id_vars=["Nama Karyawan"], value_vars=pie_cols, var_name="Komponen", value_name="Biaya")
    fig_pie = px.pie(pie_df, names="Komponen", values="Biaya", title=f"Komposisi Biaya • {selected_name}", hole=0.45)
    st.plotly_chart(fig_pie, use_container_width=True)

# ------------------ Data Table ------------------
st.markdown("### 📋 Data Tabel (Filterable)")
st.dataframe(
    df.style.format({
        "Gaji Tahunan (Rp)": lambda x: currency_idr(x),
        "BOP MP/tahun (Rp)": lambda x: currency_idr(x),
        "TAX + PENSIUN/tahun": lambda x: currency_idr(x),
        "Total Cost per-MP (Rp)": lambda x: currency_idr(x),
        "Revenue/MP (dynamic)": lambda x: currency_idr(x),
        "Profit per-MP (Rp)": lambda x: currency_idr(x),
        "Rasio Efisiensi": "{:.2f}",
        "Kontribusi to Profit (%)": "{:.2f}%"
    }),
    use_container_width=True
)

# ------------------ Export ------------------
st.markdown("### ⬇️ Export")
def to_excel_bytes(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='DashboardData')
    return output.getvalue()

colx1, colx2 = st.columns(2)
with colx1:
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV", data=csv, file_name="contrib_manpower_clean.csv", mime="text/csv")
with colx2:
    xlsb = to_excel_bytes(df)
    st.download_button("Download Excel", data=xlsb, file_name="contrib_manpower_clean.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("© 2025 • HR Analytics • Kontribusi Manpower → Profit")

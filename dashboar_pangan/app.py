import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="FoodFlow Analytics",
    layout="wide",
    initial_sidebar_state="expanded"
)

# KUSTOMISASI CSS: Kontras Super Tinggi untuk Orang Awam & Orang Tua
st.markdown("""
    <style>
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Judul Besar Dashboard */
    .main-title {
        font-size: 2.8rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.5rem;
        font-family: 'Segoe UI', Arial, sans-serif;
    }
    
    /* Judul Grafik Utama */
    .chart-header {
        font-size: 1.6rem;
        font-weight: 800;
        color: #0f172a;
        margin-top: 2rem;
        margin-bottom: 0.5rem;
        border-left: 6px solid #22c55e;
        padding-left: 12px;
    }
    
    .chart-subheader {
        font-size: 1.2rem;
        font-weight: 700;
        color: #334155;
        margin-bottom: 1rem;
    }

    /* Kotak Informasi Insight Berwarna Cerah */
    .insight-box {
        background-color: #ffffff;
        padding: 1.2rem;
        border-radius: 8px;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-top: 0.8rem;
        margin-bottom: 1.5rem;
    }
    
    .insight-title {
        font-weight: 800;
        color: #1e3a8a;
        font-size: 1.05rem;
        margin-bottom: 0.4rem;
    }
    
    .insight-text {
        color: #1e293b !important; /* Memaksa warna teks gelap konstan ber-kontras tinggi */
        font-size: 0.95rem;
        line-height: 1.5;
        margin: 0;
    }

    /* Kartu Metrik Utama */
    .metric-card-custom {
        background-color: #ffffff;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        border: 2px solid #cbd5e1;
        text-align: center;
    }
    
    .metric-label-custom {
        font-size: 1rem;
        color: #334155;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 0.5rem;
    }
    
    .metric-value-custom {
        font-size: 2.4rem;
        font-weight: 800;
        color: #0f172a; 
    }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNGSI MEMBACA DATA EXCEL LOKAL
@st.cache_data
def load_local_data():

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    path_padi = os.path.join(BASE_DIR, "produksi padi (ton).xlsx")
    path_jagung = os.path.join(BASE_DIR, "PRODUKSI JAGUNG (TON).xlsx")
    path_prod_cabe = os.path.join(BASE_DIR, "produksi cabe rawit.xlsx")
    path_prod_kacang = os.path.join(BASE_DIR, "produksi kacang panjang.xlsx")
    path_prod_ketimun = os.path.join(BASE_DIR, "produksi ketimun.xlsx")

    df_padi = pd.DataFrame()
    df_jagung = pd.DataFrame()

    def clean_vegetable_df(path):
        try:
            df_raw = pd.read_excel(path, header=None)

            header_idx = 0
            for idx, row in df_raw.iterrows():
                if row.astype(str).str.contains("PROVINSI", case=False).any():
                    header_idx = idx
                    break

            df = pd.read_excel(path, header=header_idx)

            df.columns.values[0] = "Provinsi"
            df.columns = df.columns.astype(str).str.strip()

            df = df.dropna(subset=["Provinsi"])

            df["Provinsi"] = (
                df["Provinsi"]
                .astype(str)
                .str.upper()
                .str.strip()
            )

            df = df[
                ~df["Provinsi"].str.contains(
                    "INDONESIA|PROVINSI",
                    na=False
                )
            ]

            for col in ["2021", "2022", "2023", "2024"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(
                        df[col],
                        errors="coerce"
                    ).fillna(0)

            return df

        except Exception as e:
            st.error(f"ERROR membaca {path}: {e}")
            return pd.DataFrame()

    try:
        df_padi = pd.read_excel(path_padi)

        df_padi.columns = (
            df_padi.columns.astype(str).str.strip()
        )

        if "No" in df_padi.columns:
            df_padi = df_padi.drop(columns=["No"])

        df_padi["Provinsi"] = (
            df_padi["Provinsi"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        for col in ["2021", "2022", "2023", "2024", "2025"]:
            if col in df_padi.columns:
                df_padi[col] = pd.to_numeric(
                    df_padi[col],
                    errors="coerce"
                ).fillna(0)

    except Exception as e:
        st.error(f"ERROR PADI: {e}")

    try:
        df_jagung = pd.read_excel(path_jagung)

        df_jagung.columns = (
            df_jagung.columns.astype(str).str.strip()
        )

        if "No" in df_jagung.columns:
            df_jagung = df_jagung.drop(columns=["No"])

        df_jagung["Provinsi"] = (
            df_jagung["Provinsi"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        for col in ["2021", "2022", "2023", "2024"]:
            if col in df_jagung.columns:
                df_jagung[col] = pd.to_numeric(
                    df_jagung[col],
                    errors="coerce"
                ).fillna(0)

    except Exception as e:
        st.error(f"ERROR JAGUNG: {e}")

    df_cabe = clean_vegetable_df(path_prod_cabe)
    df_kacang = clean_vegetable_df(path_prod_kacang)
    df_ketimun = clean_vegetable_df(path_prod_ketimun)

    return (
        df_padi,
        df_jagung,
        df_cabe,
        df_kacang,
        df_ketimun
    )
df_padi, df_jagung, df_cabe, df_kacang, df_ketimun = load_local_data()

# 3. SIDEBAR CONTROLLER
st.sidebar.markdown("### Navigasi Data")
komoditas = st.sidebar.radio("Pilih Komoditas Pangan:", ["Padi (Beras)", "Jagung", "Cabai Rawit", "Kacang Panjang", "Mentimun"])

if komoditas == "Padi (Beras)":
    df_aktif = df_padi
    tahun_cols = ['2021', '2022', '2023', '2024', '2025']
    tahun_terbaru = "2025"
    warna_line = "#16a34a"
elif komoditas == "Jagung":
    df_aktif = df_jagung
    tahun_cols = ['2021', '2022', '2023', '2024']
    tahun_terbaru = "2024"
    warna_line = "#ea580c"
elif komoditas == "Cabai Rawit":
    df_aktif = df_cabe
    tahun_cols = ['2021', '2022', '2023', '2024']
    tahun_terbaru = "2024"
    warna_line = "#dc2626"
elif komoditas == "Kacang Panjang":
    df_aktif = df_kacang
    tahun_cols = ['2021', '2022', '2023', '2024']
    tahun_terbaru = "2024"
    warna_line = "#15803d"
else:
    df_aktif = df_ketimun
    tahun_cols = ['2021', '2022', '2023', '2024']
    tahun_terbaru = "2024"
    warna_line = "#0284c7"

if df_aktif is None or df_aktif.empty:
    st.error(f"File excel untuk komoditas {komoditas} tidak ditemukan, salah nama, atau strukturnya kosong.")
    st.stop()

# 4. JUDUL BESAR UTAMA DASHBOARD
st.markdown('<div class="main-title">Food Flow Analytics</div>', unsafe_allow_html=True)

tahun_sebelumnya = str(int(tahun_terbaru) - 1)
total_terbaru = df_aktif[tahun_terbaru].sum()
total_sebelumnya = df_aktif[tahun_sebelumnya].sum()
perubahan = ((total_terbaru - total_sebelumnya) / total_sebelumnya) * 100 if total_sebelumnya > 0 else 0

m1, m2, m3 = st.columns(3)
with m1:
    st.markdown(f"""
        <div class="metric-card-custom" style="border-top: 6px solid #22c55e;">
            <div class="metric-label-custom">Total Produksi ({tahun_terbaru})</div>
            <div class="metric-value-custom">{total_terbaru:,.0f} Ton</div>
        </div>
    """, unsafe_allow_html=True)
with m2:
    warna_yoy = "#22c55e" if perubahan >= 0 else "#ef4444"
    tanda_panah = "▲ Naik" if perubahan >= 0 else "▼ Turun"
    st.markdown(f"""
        <div class="metric-card-custom" style="border-top: 6px solid #3b82f6;">
            <div class="metric-label-custom">Kenaikan / Penurunan Per Tahun</div>
            <div class="metric-value-custom" style="color: {warna_yoy};">{tanda_panah} {abs(perubahan):.2f}%</div>
        </div>
    """, unsafe_allow_html=True)
with m3:
    status_stabilitas = "Pasokan Aman" if perubahan > -1.5 else "Butuh Logistik"
    warna_status = "#22c55e" if status_stabilitas == "Pasokan Aman" else "#ef4444"
    st.markdown(f"""
        <div class="metric-card-custom" style="border-top: 6px solid #64748b;">
            <div class="metric-label-custom">Status Stabilitas Pasokan</div>
            <div class="metric-value-custom" style="color: {warna_status};">{status_stabilitas}</div>
        </div>
    """, unsafe_allow_html=True)

# 5. GRAFIK 1: TREN BERKELANJUTAN
st.markdown('<div class="chart-header">1. Tren Perkembangan Produksi Pangan Nasional</div>', unsafe_allow_html=True)
total_tren = df_aktif[tahun_cols].sum().reset_index()
total_tren.columns = ['Tahun', 'Total Produksi']

fig_tren = go.Figure()
fig_tren.add_trace(go.Scatter(
    x=total_tren['Tahun'], 
    y=total_tren['Total Produksi'],
    mode='lines+markers+text',
    text=total_tren['Total Produksi'],
    textposition='top center',
    texttemplate='%{text:,.0f} Ton',
    textfont=dict(color='#0f172a', size=12, weight='bold'),
    line=dict(color=warna_line, width=4),
    marker=dict(size=12, color='#0f172a'),
))
fig_tren.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    height=360,
    margin=dict(l=60, r=40, t=40, b=20),
    xaxis=dict(type='category', showgrid=False, tickfont=dict(color='#0f172a', size=14, weight="bold")),
    yaxis=dict(showgrid=True, gridcolor='#cbd5e1', tickfont=dict(color='#475569', size=12),
               title=dict(text="Total Produksi (Ton)", font=dict(color='#0f172a', size=12, weight='bold')))
)
st.plotly_chart(fig_tren, use_container_width=True)

# INSIGHT GRAFIK 1 (Teks disesuaikan presisi & dinamis)
kondisi_tren_text = "peningkatan yang baik" if perubahan >= 0 else "Mengalami Penurunan Produksi"
st.markdown(f"""
    <div class="insight-box">
        <div class="insight-title">Informasi Tren Produksi:</div>
        <p class="insight-text">
            Grafik di samping menunjukkan total hasil panen komoditas <b>{komoditas}</b> di seluruh Indonesia.<br><br>Tahun <b>{tahun_terbaru}</b> terkumpul <b>{total_terbaru:,.0f} Ton</b>. Secara umum, tren ketersediaan pangan nasional kita saat ini dinilai <b>{kondisi_tren_text}</b> dibandingkan tahun lalu.
        </p>
    </div>
""", unsafe_allow_html=True)

st.markdown("---")

# 6. GRAFIK 2: TOP 5 SURPLUS & DEFISIT
st.markdown(f'<div class="chart-header">2. Analisis Distribusi Wilayah Produksi Pangan Utama ({tahun_terbaru})</div>', unsafe_allow_html=True)

if komoditas == "Padi (Beras)":
    provinsi_top_padi = ['JAWA TIMUR', 'JAWA TENGAH', 'JAWA BARAT', 'SULAWESI SELATAN', 'SUMATERA SELATAN']
    top_5 = df_aktif[df_aktif['Provinsi'].isin(provinsi_top_padi)].sort_values(by=tahun_terbaru, ascending=False).head(5)
else:
    df_sorted = df_aktif[['Provinsi', tahun_terbaru]].sort_values(by=tahun_terbaru, ascending=False)
    top_5 = df_sorted.head(5)

df_sorted_bottom = df_aktif[['Provinsi', tahun_terbaru]].sort_values(by=tahun_terbaru, ascending=False)
bottom_5 = df_sorted_bottom[df_sorted_bottom[tahun_terbaru] > 0].tail(5).sort_values(by=tahun_terbaru, ascending=True)

col_left, col_right = st.columns(2)

with col_left:
    st.markdown('<div class="chart-subheader">🟢 Top 5 Daerah Surplus Produksi Tertinggi</div>', unsafe_allow_html=True)
    fig_top = px.bar(top_5, x=tahun_terbaru, y='Provinsi', orientation='h', text=tahun_terbaru)
    fig_top.update_traces(
        marker_color='#22c55e', 
        texttemplate='%{text:,.0f} Ton',
        textposition='auto', 
        textfont=dict(color='#0f172a', weight='bold', size=13),
        insidetextfont=dict(color='#ffffff')
    )
    fig_top.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=450, 
        margin=dict(l=240, r=80, t=10, b=10), 
        xaxis=dict(title=None, showgrid=True, gridcolor='#e2e8f0'),
        yaxis=dict(title=None, categoryorder='total ascending', tickfont=dict(color='#0f172a', size=14, weight='bold'))
    )
    st.plotly_chart(fig_top, use_container_width=True)

with col_right:
    st.markdown('<div class="chart-subheader">🔴 5 Daerah Produksi Terendah </div>', unsafe_allow_html=True)
    fig_bottom = px.bar(bottom_5, x=tahun_terbaru, y='Provinsi', orientation='h', text=tahun_terbaru)
    fig_bottom.update_traces(
        marker_color='#ef4444', 
        texttemplate='%{text:,.0f} Ton',
        textposition='auto', 
        textfont=dict(color='#0f172a', weight='bold', size=13),
        insidetextfont=dict(color='#ffffff')
    )
    fig_bottom.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=450, 
        margin=dict(l=240, r=80, t=10, b=10), 
        xaxis=dict(title=None, showgrid=True, gridcolor='#e2e8f0'),
        yaxis=dict(title=None, categoryorder='total descending', tickfont=dict(color='#0f172a', size=14, weight='bold'))
    )
    st.plotly_chart(fig_bottom, use_container_width=True)

if not bottom_5.empty and not top_5.empty:
    daerah_penerima = bottom_5['Provinsi'].iloc[0]
    daerah_pemasok = top_5['Provinsi'].iloc[0]

    st.markdown(
        f"""
        <div style="background-color: #ffffff; padding: 1.5rem; border-left: 6px solid #22c55e; border-radius: 12px; margin-top: 1rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #cbd5e1;">
            <h4 style="margin-top:0; color:#0f172a; font-size:1.2rem; font-weight:800;">Framework Aliran Distribusi Logistik Antar-Daerah:</h4>
            <p style="margin-bottom:0; color:#334155; font-size:1.05rem; line-height: 1.6;">
                • <b>Daerah yang Paling Membutuhkan Kiriman Tambahan (Penerima):</b> Wilayah koridor <b>{daerah_penerima}</b> mencatat angka hasil panen paling kecil mandiri. Wilayah ini diprioritaskan utama untuk dikirimi pasokan stok.<br>
                • <b>Daerah Lumbung yang Menjadi Pemasok Utama (Pengirim):</b> Wilayah <b>{daerah_pemasok}</b> sebagai produsen hasil panen terbesar nasional direkomendasikan menyuplai langsung daerah penerima untuk menstabilkan harga pasaran.
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.markdown("---")

# ======================================================
# 3. REKOMENDASI DISTRIBUSI PANGAN OTOMATIS & EARLY WARNING
# ======================================================
st.markdown('<div class="chart-header">3. Smart Distribution & Early Warning System</div>', unsafe_allow_html=True)

top_sender = top_5.iloc[0]
top_receiver = bottom_5.iloc[0]
jumlah_rekomendasi = min(top_sender[tahun_terbaru] * 0.10, top_receiver[tahun_terbaru] * 5)

rekom_df = pd.DataFrame({
    "Komoditas": [komoditas],
    "Daerah Pengirim": [top_sender["Provinsi"]],
    "Daerah Penerima": [top_receiver["Provinsi"]],
    "Volume Rekomendasi (Ton)": [round(jumlah_rekomendasi)]
})

# Menampilkan dataframe tanpa index
# KODE PERBAIKAN UNTUK TABEL HTML CUSTOM
html_table = f"""
<div style="overflow-x:auto; margin-bottom: 1rem; border: 2px solid #cbd5e1; border-radius: 8px; background-color: #ffffff;">
    <table style="width:100%; border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; text-align: left;">
        <thead>
            <tr style="background-color: #f1f5f9; border-bottom: 2px solid #cbd5e1;">
                <th style="padding: 12px 16px; color: #0f172a; font-weight: 800; font-size: 0.95rem;">Komoditas</th>
                <th style="padding: 12px 16px; color: #0f172a; font-weight: 800; font-size: 0.95rem;">Daerah Pengirim</th>
                <th style="padding: 12px 16px; color: #0f172a; font-weight: 800; font-size: 0.95rem;">Daerah Penerima</th>
                <th style="padding: 12px 16px; color: #0f172a; font-weight: 800; font-size: 0.95rem; text-align: right;">Volume Rekomendasi (Ton)</th>
            </tr>
        </thead>
        <tbody>
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 14px 16px; color: #334155; font-weight: 700; font-size: 0.95rem;">{komoditas}</td>
                <td style="padding: 14px 16px; color: #16a34a; font-weight: 800; font-size: 0.95rem;"> {top_sender["Provinsi"]}</td>
                <td style="padding: 14px 16px; color: #dc2626; font-weight: 800; font-size: 0.95rem;"> {top_receiver["Provinsi"]}</td>
                <td style="padding: 14px 16px; color: #0f172a; font-weight: 800; font-size: 1.05rem; text-align: right;">{int(round(jumlah_rekomendasi)):,} Ton</td>
            </tr>
        </tbody>
    </table>
</div>
"""

# Render tabel HTML ke Streamlit
st.markdown(html_table, unsafe_allow_html=True)


# Penyatuan Kontainer EWS & Logistik agar tidak berjarak jauh
if perubahan > 5:
    status, warna, bg_warna = "🟢 PASOKAN AMAN", "#22c55e", "#f0fdf4"
elif perubahan > 0:
    status, warna, bg_warna = "🟡 WASPADA SURPLUS", "#eab308", "#fefce8"
else:
    status, warna, bg_warna = "🔴 RISIKO PENURUNAN PASOKAN", "#ef4444", "#fef2f2"

col_box1, col_box2 = st.columns(2)

with col_box1:
    st.markdown(f"""
    <div class="insight-box" style="margin-top: 0px; height: 100%;">
        <div class="insight-title"> Rekomendasi Distribusi Prioritas</div>
        <p class="insight-text">
        Berdasarkan analisis produksi tahun {tahun_terbaru}, wilayah <b>{top_sender['Provinsi']}</b> merupakan produsen terbesar komoditas <b>{komoditas}</b>.<br><br>
        Sementara itu, wilayah <b>{top_receiver['Provinsi']}</b> memiliki produksi terendah sehingga diprioritaskan sebagai penerima distribusi.
        Sistem merekomendasikan pengiriman awal sebesar <b>{round(jumlah_rekomendasi):,.0f} ton</b> untuk membantu mengurangi potensi food loss akibat penumpukan surplus di wilayah produsen.
        </p>
    </div>
    """, unsafe_allow_html=True)

with col_box2:
    st.markdown(f"""
    <div style="background:{bg_warna}; padding:1.2rem; border-left:6px solid {warna}; border-radius:8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); height: 100%;">
        <h3 style="color:{warna}; margin-top:0; font-weight:800; font-size:1.15rem;">⚠️ Status: {status}</h3>
        <p class="insight-text" style="color:#1e293b; margin-top:0.5rem;">
        Monitoring produksi menunjukkan kondisi komoditas <b>{komoditas}</b> tahun {tahun_terbaru} berada pada kategori <b>{status}</b>.<br><br>
        Sistem menyarankan optimalisasi manajemen logistik antar koridor guna menyeimbangkan neraca pangan wilayah secara presisi.
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 7. GRAFIK 3: RISK INDEX
st.markdown('<div class="chart-header">4. Food Loss Risk Index (Indeks Kerentanan Risiko Panen Wilayah)</div>', unsafe_allow_html=True)
df_aktif['Volatilitas'] = df_aktif[tahun_cols].std(axis=1) / (df_aktif[tahun_cols].mean(axis=1) + 1)
df_risk = df_aktif[['Provinsi', 'Volatilitas']].sort_values(by='Volatilitas', ascending=False).head(10)

fig_risk = px.bar(df_risk, x='Volatilitas', y='Provinsi', orientation='h', color='Volatilitas',
                  color_continuous_scale=[[0, '#cbd5e1'], [1, '#ef4444']])
fig_risk.update_layout(
    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=400,
    margin=dict(l=240, r=40, t=10, b=10),
    xaxis=dict(title="Tingkat Ketidakstabilan Naik-Turun Data", showgrid=True, gridcolor='#cbd5e1'),
    yaxis=dict(title=None, categoryorder='total ascending', tickfont=dict(color='#0f172a', size=14, weight='bold')),
    coloraxis_showscale=False
)
st.plotly_chart(fig_risk, use_container_width=True)

# INSIGHT GRAFIK 3
st.markdown(f"""
    <div class="insight-box" style="border-left-color: #ef4444;">
        <div class="insight-title" style="color: #991b1b;">⚠️ Informasi Risiko Kerentanan Panen:</div>
        <p class="insight-text">
            Grafik bar di atas menunjukkan tingkat kerentanan atau ketidakstabilan panen di suatu wilayah. 
            Semakin <b>panjang batang merahnya</b>, berarti daerah tersebut hasil produksinya sering naik-turun secara drastis setiap tahunnya (tidak stabil). 
            Daerah di peringkat teratas membutuhkan perhatian ekstra dari pemerintah karena pasokan pangannya paling berisiko mengalami gagal panen mendadak.
        </p>
    </div>
""", unsafe_allow_html=True)

# FOOTER
st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748b; font-size: 0.85rem; font-weight:600;'>Food Flow Analytics System | Laporan Publikasi Distribusi Komoditas Agraria Nasional</p>", unsafe_allow_html=True)

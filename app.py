import streamlit as st
import pandas as pd
import numpy as np
import joblib
import plotly.express as px
from datetime import datetime
import base64
import io

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Executive Dashboard | UNMUH BABEL",
    page_icon="🎓",
    layout="wide"
)

# ==========================================
# 2. FUNCTION UTILITY (DOWNLOAD & CSS)
# ==========================================
def get_base64_file(file):
    try:
        with open(file, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        return None

def create_template():
    kolom_template = [
        "Nama", "Jenis Kelamin", "Program Studi", "Tahun Masuk",
        "IPK", "IPS 1", "IPS 2", "IPS 3", "IPS 4", "IPS 5", "IPS 6", "IPS 7", "IPS 8", "IPS 9", "IPS 10", "IPS 11",
        "Jumlah SKS", "Jumlah Mata Kuliah yang Diulang", "Motivasi Belajar", 
        "Dukungan Keluarga", "Tingkat Stres", "Sosial-Ekonomi", 
        "Pekerjaan Paruh Waktu", "Keaktifan dalam Berorganisasi"
    ]
    df_template = pd.DataFrame(columns=kolom_template)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 'border': 1, 'align': 'center'
        })
        
        for col_num, value in enumerate(df_template.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 22)

        # Dropdowns
        worksheet.data_validation('B2:B500', {'validate': 'list', 'source': ['Laki-laki', 'Perempuan']})
        worksheet.data_validation('C2:C500', {'validate': 'list', 'source': ['PGSD', 'PJKR', 'PMTK', 'PBI', 'Ilmu Komputer', 'KSDA', 'Teknik Sipil', 'Kriminologi', 'Pariwisata', 'Kewirausahaan']})
        
        opsi_5 = ['sangat rendah', 'rendah', 'sedang', 'tinggi', 'sangat tinggi']
        worksheet.data_validation('S2:S500', {'validate': 'list', 'source': opsi_5})
        worksheet.data_validation('T2:T500', {'validate': 'list', 'source': opsi_5})
        worksheet.data_validation('U2:U500', {'validate': 'list', 'source': ['rendah', 'sedang', 'tinggi']})
        worksheet.data_validation('V2:V500', {'validate': 'list', 'source': ['rendah', 'menengah', 'tinggi']})
        worksheet.data_validation('W2:W500', {'validate': 'list', 'source': ['bekerja', 'tidak bekerja']})
        worksheet.data_validation('X2:X500', {'validate': 'list', 'source': ['aktif', 'tidak aktif']})

    return buffer.getvalue()

banner_64 = get_base64_file("campus.jpg")
logo_64 = get_base64_file("logo.png")

banner_css = f"""
    background-image: linear-gradient(rgba(10,25,47,0.5), rgba(10,25,47,0.6)), url('data:image/jpg;base64,{banner_64}');
    background-size: cover; background-position: center center; min-height: 400px;
""" if banner_64 else "background-color: #0a192f; min-height: 300px;"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    .stApp {{ background-color: #f8fafc; font-family: 'Inter', sans-serif; }}
    .main-header {{
        {banner_css}
        padding: 60px 20px; border-radius: 0 0 40px 40px; color: white;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.1); margin-bottom: 40px;
    }}
    .header-content h1 {{ font-size: 2.8rem; font-weight: 800; margin-top: 15px; text-shadow: 2px 2px 8px rgba(0,0,0,0.7); }}
    .step-box {{
        background: white; border-radius: 15px; padding: 20px; text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; height: 100%;
    }}
    .step-icon {{ font-size: 40px; margin-bottom: 10px; }}
    .step-title {{ color: #1e3a8a; font-weight: 800; margin-bottom: 8px; font-size: 1.1rem; }}
    .step-desc {{ color: #64748b; font-size: 0.85rem; line-height: 1.4; }}
    .guide-card {{
        background: white; padding: 25px; border-radius: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-left: 6px solid #1e3a8a;
        height: 100%;
    }}
    .step-number {{
        background: #1e3a8a; color: white; width: 28px; height: 28px;
        border-radius: 50%; display: inline-flex; align-items: center;
        justify-content: center; font-weight: bold; margin-right: 10px; font-size: 0.9rem;
    }}
    .stat-card {{
        background: white; padding: 25px; border-radius: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); text-align: center; border-bottom: 5px solid #1e3a8a;
    }}
    .stButton>button {{
        width: 100%; border-radius: 12px; background-color: #1e3a8a; color: white;
        font-weight: 600; height: 3.5em; transition: 0.3s; border: none;
    }}
    .stButton>button:hover {{ background-color: #3b82f6; transform: translateY(-2px); color: white !important; }}
    div.stDownloadButton > button {{
        background-color: white !important; color: #334155 !important;
        border: 1px solid #e2e8f0 !important; box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. HEADER UI
# ==========================================
logo_html = f'<img src="data:image/png;base64,{logo_64}" style="height: 110px;">' if logo_64 else "🎓"
st.markdown(f"""
<div class="main-header">
    {logo_html}
    <div class="header-content">
        <h1>SISTEM PREDIKSI KELULUSAN MAHASISWA</h1>
        <p style="font-size: 1.1rem; opacity: 0.9; letter-spacing: 3px; font-weight: 600;">UNIVERSITAS MUHAMMADIYAH BANGKA BELITUNG</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 4. PETUNJUK VISUAL
# ==========================================
st.markdown("<h3 style='text-align: center; color: #1e3a8a; margin-bottom: 30px;'>📖 Petunjuk Alur Penggunaan Sistem</h3>", unsafe_allow_html=True)

s1, s2, s3, s4, s5 = st.columns(5)
with s1:
    st.markdown('<div class="step-box"><div class="step-icon">📥</div><div class="step-title">1. Unduh Template</div><div class="step-desc">Klik tombol <b>"UNDUH TEMPLATE"</b> untuk mendapatkan format file Excel kosong.</div></div>', unsafe_allow_html=True)
with s2:
    st.markdown('<div class="step-box"><div class="step-icon">📝</div><div class="step-title">2. Isi Data</div><div class="step-desc">Gunakan <b>dropdown</b> pada Excel untuk memilih kategori agar tidak typo.</div></div>', unsafe_allow_html=True)
with s3:
    st.markdown('<div class="step-box"><div class="step-icon">📤</div><div class="step-title">3. Unggah File</div><div class="step-desc">Klik kotak <b>"Upload"</b> lalu pilih file Excel yang sudah diisi datanya secara lengkap.</div></div>', unsafe_allow_html=True)
with s4:
    st.markdown('<div class="step-box"><div class="step-icon">⚙️</div><div class="step-title">4. Proses</div><div class="step-desc">Klik tombol <b>"MULAI PROSES PREDIKSI"</b> untuk menganalisis data secara otomatis.</div></div>', unsafe_allow_html=True)
with s5:
    st.markdown('<div class="step-box"><div class="step-icon">📄</div><div class="step-title">5. Unduh Hasil</div><div class="step-desc">Setelah hasil muncul, klik tombol <b>"Download File Excel"</b> untuk menyimpan laporan prediksi.</div></div>', unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# ==========================================
# 5. PANDUAN TEKNIS
# ==========================================
st.markdown("### 📋 Panduan Pengisian Data")
g1, g2 = st.columns([1, 1.3])

with g1:
    st.markdown(f"""
    <div class="guide-card">
        <h4 style="color: #1e3a8a; margin-bottom: 20px;">🚀 Validasi Data</h4>
        <p><span class="step-number">✓</span> <b>Pilihan Otomatis:</b> Sudah tersedia menu pilihan (dropdown) di Excel.</p>
        <p><span class="step-number">✓</span> <b>Bebas Kesalahan:</b> Tidak perlu mengetik manual untuk data kategori.</p>
        <p><span class="step-number">✓</span> <b>Format Akurat:</b> Data otomatis sesuai dengan kebutuhan sistem.</p>
        <div style="background-color: #fff4f4; padding: 12px; border-radius: 10px; border: 1px solid #fecaca; margin-top: 15px;">
            <p style="color: #b91c1c; font-size: 0.9rem; margin-bottom: 0;">
                <b>Note:</b> Gunakan Template Excel Kosong agar fitur menu pilihan tetap muncul saat pengisian. 
                Sistem ini hanya bisa memprediksi mahasiswa semester 5-8, untuk mahasiswa semester 5 kebawah dan mahasiswa semester 8 keatas tidak bisa diprediksi menggunakan sistem ini 🙏.
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(label="📥 UNDUH TEMPLATE EXCEL KOSONG", data=create_template(), file_name="Template_Prediksi_Unmuh_Babel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with g2:
    st.markdown('<div class="guide-card">', unsafe_allow_html=True)
    st.markdown('<h4 style="color: #1e3a8a; margin-bottom: 15px;">🔍 Opsi Yang Tersedia di Excel</h4>', unsafe_allow_html=True)
    petunjuk_data = {
        "Kategori": ["Jenis Kelamin", "Program Studi", "Motivasi & Dukungan", "Tingkat Stres", "Sosial-Ekonomi", "Pekerjaan", "Organisasi"],
        "Opsi Pilihan (Tinggal Pilih)": ["Laki-laki, Perempuan", "10 Program Studi di Unmuh Babel", "sangat rendah s/d sangat tinggi", "rendah, sedang, tinggi", "rendah, menengah, tinggi", "bekerja, tidak bekerja", "aktif, tidak aktif"]
    }
    st.table(pd.DataFrame(petunjuk_data))
    st.caption("⚠️ Anda tidak perlu mengetik manual pada variabel yang bersifat kategori, cukup klik sel di Excel dan pilih opsi yang muncul.")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# ==========================================
# 6. MODEL & PROCESSING
# ==========================================
@st.cache_resource
def load_model():
    try:
        model = joblib.load("model_random_forest.pkl")
        fitur = joblib.load("fitur_model.pkl")
        return model, fitur
    except Exception:
        return None, None

model, fitur_sistem = load_model()

st.markdown("### 📂 Unggah File yang Sudah Diisi")
file_up = st.file_uploader("Pilih file Excel", type=["xlsx"], label_visibility="collapsed")

if file_up:
    try:
        df_ori = pd.read_excel(file_up)
        df_ori.columns = df_ori.columns.str.strip()
        st.success(f"✔️ Berhasil mengimpor {len(df_ori)} data mahasiswa.")
        
        with st.expander("🔍 Klik untuk menampilkan data mentah"):
            st.dataframe(df_ori, use_container_width=True)

        if st.button("🚀 MULAI PROSES PREDIKSI"):
            if model is None:
                st.error("❌ File model (.pkl) tidak ditemukan. Pastikan file model tersedia di server.")
            else:
                with st.spinner("Sedang menganalisis data..."):
                    df_proc = df_ori.copy()
                    
                    # --- PREPROCESSING ---
                    mapping_5 = {'sangat rendah':1, 'rendah':2, 'sedang':3, 'tinggi':4, 'sangat tinggi':5}
                    mapping_stres = {'rendah':1, 'sedang':2, 'tinggi':3}
                    mapping_ekonomi = {'rendah':1, 'menengah':2, 'tinggi':3}
                    mapping_bin = {'tidak bekerja':0, 'bekerja':1, 'tidak aktif':0, 'aktif':1}

                    cols_to_map = {
                        'Motivasi Belajar': mapping_5,
                        'Dukungan Keluarga': mapping_5,
                        'Tingkat Stres': mapping_stres,
                        'Sosial-Ekonomi': mapping_ekonomi,
                        'Pekerjaan Paruh Waktu': mapping_bin,
                        'Keaktifan dalam Berorganisasi': mapping_bin
                    }

                    for col, mapping in cols_to_map.items():
                        if col in df_proc.columns:
                            temp_col = df_proc[col].astype(str).str.lower().str.strip()
                            df_proc[f'{col}_Num'] = temp_col.map(mapping).fillna(0)
                            df_proc[col] = df_proc[f'{col}_Num']

                    # IPS Handling
                    all_ips = [f'IPS {i}' for i in range(1, 12)]
                    for col in all_ips:
                        if col not in df_proc.columns:
                            df_proc[col] = np.nan
                    
                    def fill_ips_dynamic(row):
                        existing = [row[c] for c in all_ips if pd.notna(row[c]) and row[c] > 0]
                        avg = sum(existing)/len(existing) if existing else row.get('IPK', 0.0)
                        for c in all_ips:
                            if pd.isna(row[c]) or row[c] <= 0:
                                row[c] = avg
                        return row
                    
                    df_proc = df_proc.apply(fill_ips_dynamic, axis=1)

                    # --- PREDIKSI DENGAN LOGIKA ADAPTIF ---
                    X = df_proc.reindex(columns=fitur_sistem, fill_value=0)
                    y_pred = model.predict(X)
                    
                    final_status, final_years = [], []
                    for i in range(len(y_pred)):
                        # Data Dasar
                        years = float(y_pred[i])
                        ipk_val = df_proc.iloc[i]['IPK']
                        tahun_masuk = df_proc.iloc[i]['Tahun Masuk']
                        
                        # Hitung Semester
                        existing_ips = [df_ori.iloc[i][c] for c in all_ips if c in df_ori.columns and pd.notna(df_ori.iloc[i][c]) and df_ori.iloc[i][c] > 0]
                        jml_semester = len(existing_ips)

                        # 1. Validasi Angkatan Lama/Semester Awal (Filter Utama)
                        if jml_semester < 4 or tahun_masuk <= 2021:
                            years = 4.5
                        else:
                            # 2. Ambil Nilai Faktor Pendukung
                            stres = df_proc.iloc[i].get('Tingkat Stres_Num', 2)
                            dukungan = df_proc.iloc[i].get('Dukungan Keluarga_Num', 3)
                            motivasi = df_proc.iloc[i].get('Motivasi Belajar_Num', 3)
                            pekerjaan = df_proc.iloc[i].get('Pekerjaan Paruh Waktu', 0)
                            keaktifan = df_proc.iloc[i].get('Keaktifan dalam Berorganisasi',0)
                            
                            # 3. Tentukan Baseline (Standar Awal)
                            if ipk_val >= 3.25:
                                baseline = 3.8 # Potensi tepat waktu kuat
                            elif ipk_val >= 3.0:
                                baseline = 3.95 # Di ambang batas
                            else:
                                baseline = 4.2 # Potensi terlambat
                            
                            # 4. Logika Adaptif "Sedang" & Penyeimbang
                            penyesuaian = 0
                            
                            # Jika Kondisi Sedang Semua (Logika yang Anda minta)
                            if stres == 2 and dukungan == 3 and motivasi == 3 and pekerjaan == 0 and keaktifan == 0:
                                # Status murni ditentukan baseline IPK
                                penyesuaian = 0 
                            else:
                                # Bonus (Mengurangi estimasi tahun)
                                if motivasi >= 4: penyesuaian -= 0.2
                                if stres == 1: penyesuaian -= 0.1
                        
                                
                                # Penalti (Menambah estimasi tahun)
                                if dukungan <= 2: penyesuaian += 0.2
                                if stres == 3: penyesuaian += 0.3
                                if keaktifan == 1: penyesuaian += 0.1
                                if pekerjaan == 1: penyesuaian += 0.1
                            
                            years = baseline + penyesuaian

                        status = "TEPAT WAKTU" if years <= 4.0 else "TERLAMBAT"
                        final_status.append(status)
                        final_years.append(years)

                    df_ori['Masa Studi'] = [f"{y:.1f} Tahun" for y in final_years]
                    df_ori['Status'] = final_status

                    # --- DASHBOARD HASIL ---
                    st.markdown("---")
                    t1, t2 = st.columns(2)
                    tepat_count = final_status.count("TEPAT WAKTU")
                    terlambat_count = final_status.count("TERLAMBAT")
                    
                    with t1: st.markdown(f'<div class="stat-card"><p style="color:#64748b">TEPAT WAKTU</p><h2 style="color:#1e3a8a">{tepat_count} Mahasiswa</h2></div>', unsafe_allow_html=True)
                    with t2: st.markdown(f'<div class="stat-card" style="border-bottom-color:#ef4444"><p style="color:#64748b">TERLAMBAT</p><h2 style="color:#ef4444">{terlambat_count} Mahasiswa</h2></div>', unsafe_allow_html=True)

                    st.subheader("📋 Laporan Hasil Prediksi")
                    df_final = df_ori[['Nama', 'Jenis Kelamin', 'Program Studi', 'Masa Studi', 'Status']]
                    st.dataframe(df_final.style.map(lambda v: f'color: {"#1e3a8a" if v == "TEPAT WAKTU" else "#ef4444"}; font-weight: bold', subset=['Status']), use_container_width=True)

                    # Visualisasi
                    st.markdown("---")
                    v1, v2, v3 = st.columns(3)
                    with v1:
                        fig = px.pie(df_ori, names='Status', hole=0.4, title='Status Kelulusan', color='Status', color_discrete_map={'TEPAT WAKTU': '#1e3a8a', 'TERLAMBAT': '#ef4444'})
                        st.plotly_chart(fig, use_container_width=True)
                    with v2:
                        df_jk = df_ori['Jenis Kelamin'].value_counts().reset_index()
                        fig = px.bar(df_jk, x='Jenis Kelamin', y='count', title='Sebaran Jenis Kelamin', color_discrete_sequence=['#1e3a8a'])
                        st.plotly_chart(fig, use_container_width=True)
                    with v3:
                        df_prodi = df_ori['Program Studi'].value_counts().reset_index()
                        fig = px.bar(df_prodi, y='Program Studi', x='count', orientation='h', title='Sebaran Program Studi', color_discrete_sequence=['#10b981'])
                        st.plotly_chart(fig, use_container_width=True)

                    # --- Unduh Hasil ---
                    st.markdown("---")
                    st.markdown("<h4 style='text-align: center;'>📥 Unduh Hasil Prediksi</h4>", unsafe_allow_html=True)
                    
                    excel_data = io.BytesIO()
                    with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Hasil Prediksi')
                        workbook  = writer.book
                        worksheet = writer.sheets['Hasil Prediksi']
                        header_format = workbook.add_format({'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 'border': 1, 'align': 'center'})
                        for col_num, value in enumerate(df_final.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            worksheet.set_column(col_num, col_num, 20)

                    _, db_col, _ = st.columns([1, 1, 1])
                    with db_col:
                        st.download_button(
                            label="Download File Excel (.xlsx)",
                            data=excel_data.getvalue(),
                            file_name=f"hasil_prediksi_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

    except Exception as e:
        st.error(f"❌ Terjadi kesalahan pada file: {e}")

st.markdown(f"<div style='text-align:center; margin-top:50px; color:#94a3b8;'>© {datetime.now().year} UNIVERSITAS MUHAMMADIYAH BANGKA BELITUNG</div>", unsafe_allow_html=True)
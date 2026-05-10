import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import io
import os
import requests
import re
from werkzeug.security import generate_password_hash, check_password_hash

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

try:
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    WORDCLOUD_AVAILABLE = True
except ImportError:
    WORDCLOUD_AVAILABLE = False

st.set_page_config(page_title="Aplikasi Pelaporan Blokade", page_icon="🚧", layout="wide", initial_sidebar_state="expanded")

# Menggunakan nama file baru agar tidak bentrok dengan password plaintext dari DB sebelumnya
conn = sqlite3.connect('blokade_pro_v3.db', check_same_thread=False)

def init_db():
    c = conn.cursor()
    # Tabel Users dengan password hash
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (username TEXT PRIMARY KEY, password TEXT, role TEXT, nama_lengkap TEXT)''')
    
    # Tabel Audit Logs
    c.execute('''CREATE TABLE IF NOT EXISTS audit_logs
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, user TEXT, action TEXT, details TEXT)''')
    
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        hashed_pw = generate_password_hash("123") 
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", ('admin', hashed_pw, 'Administrator', 'Super Admin'))
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", ('spv', hashed_pw, 'Supervisor', 'Jhon Doe (SPV)'))
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?)", ('mgr', hashed_pw, 'Manager', 'Bapak Manager'))
        conn.commit()

    try:
        df_cek = pd.read_sql("SELECT * FROM laporan LIMIT 1", conn)
    except:
        DUMMY_DATA = [
            {"Tanggal": datetime.date(2024, 2, 10), "Pelaku": "Andi", "No HP": "0812345", "Kabupaten": "Bolaang Mongondow", "Kecamatan": "Lolayan", "Desa": "Bakan", "Lokasi": "Simpang 3 Masjid Bakan", "Waktu Mulai": "08:00:00", "Waktu Selesai": "13:00:00", "Durasi (Jam)": 5.0, "Kategori Durasi": "Cepat", "Isu": "Tuntutan Pekerjaan", "Target": "PT JRBM & PT SMA", "Deskripsi": "Masyarakat menuntut pekerjaan untuk warga lokal. Tenda didirikan di simpang 3.", "File_Bukti": ""},
            {"Tanggal": datetime.date(2024, 5, 20), "Pelaku": "Budi", "No HP": "0812345", "Kabupaten": "Bolaang Mongondow", "Kecamatan": "Lolayan", "Desa": "Matali Baru", "Lokasi": "Lokasi Umum Matali Baru", "Waktu Mulai": "07:00:00", "Waktu Selesai": "18:00:00", "Durasi (Jam)": 11.0, "Kategori Durasi": "Lambat", "Isu": "Kualitas Lingkungan", "Target": "PT JRBM", "Deskripsi": "Kompensasi debu jalan belum dibayarkan. Warga memblokade jalan pakai kayu.", "File_Bukti": ""},
            {"Tanggal": datetime.date(2024, 8, 15), "Pelaku": "Citra", "No HP": "0812345", "Kabupaten": "Bolaang Mongondow Selatan", "Kecamatan": "Pinolosian Tengah", "Desa": "Tobayagan", "Lokasi": "Port Motandoi", "Waktu Mulai": "09:00:00", "Waktu Selesai": "13:30:00", "Durasi (Jam)": 4.5, "Kategori Durasi": "Cepat", "Isu": "Kompensasi Lahan", "Target": "PT JRBM", "Deskripsi": "Ganti rugi lahan area pelabuhan. Mediasi berhasil dilakukan jam 1 siang.", "File_Bukti": ""},
            {"Tanggal": datetime.date(2025, 1, 15), "Pelaku": "Eko", "No HP": "0812345", "Kabupaten": "Bolaang Mongondow", "Kecamatan": "Lolayan", "Desa": "Bakan", "Lokasi": "Parkiran Blok C", "Waktu Mulai": "06:00:00", "Waktu Selesai": "10:00:00", "Durasi (Jam)": 4.0, "Kategori Durasi": "Cepat", "Isu": "Tuntutan Pekerjaan", "Target": "PT JRBM & PT SMA", "Deskripsi": "Protes penerimaan karyawan baru. Karyawan lokal merasa diabaikan.", "File_Bukti": ""},
            {"Tanggal": datetime.date(2025, 6, 20), "Pelaku": "Gita", "No HP": "0812345", "Kabupaten": "Bolaang Mongondow Selatan", "Kecamatan": "Pinolosian Tengah", "Desa": "Tobayagan", "Lokasi": "Simpang 2 Akses Masuk Motandoi", "Waktu Mulai": "08:00:00", "Waktu Selesai": "20:00:00", "Durasi (Jam)": 12.0, "Kategori Durasi": "Lambat", "Isu": "Ganti Rugi Tanaman/Banjir", "Target": "PT JRBM", "Deskripsi": "Tuntutan ganti rugi tanaman cengkeh yang terkena dampak. Negosiasi alot.", "File_Bukti": ""},
        ]
        df_dummy = pd.DataFrame(DUMMY_DATA)
        df_dummy.to_sql('laporan', conn, index=False)
        add_audit_log("System", "INIT", "Sistem menginisialisasi database laporan pertama kali.")

def load_data():
    return pd.read_sql("SELECT * FROM laporan", conn)

def save_new_data(new_dict):
    new_df = pd.DataFrame([new_dict])
    new_df.to_sql('laporan', conn, if_exists='append', index=False)

def add_audit_log(user, action, details):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c = conn.cursor()
    c.execute("INSERT INTO audit_logs (timestamp, user, action, details) VALUES (?, ?, ?, ?)", (now, user, action, details))
    conn.commit()

def send_telegram_notif(data_dict, pelapor):
    try:
        # Mengambil dari st.secrets (Praktik Keamanan Terbaik)
        BOT_TOKEN = st.secrets["BOT_TOKEN"]
        CHAT_ID = st.secrets["CHAT_ID"]
    except Exception:
        # Fallback jika belum di-setup di lokal/server
        return False 
    
    pesan = f"🚨 *LAPORAN BLOKADE BARU* 🚨\n\n"
    pesan += f"📍 *Lokasi:* {data_dict['Desa']}, {data_dict['Kabupaten']}\n"
    pesan += f"🏢 *Target:* {data_dict['Target']}\n"
    pesan += f"📢 *Isu/Tuntutan:* {data_dict['Isu']}\n"
    pesan += f"⏱️ *Durasi:* {data_dict['Durasi (Jam)']} Jam ({data_dict['Kategori Durasi']})\n"
    pesan += f"📝 *Deskripsi:* {data_dict['Deskripsi']}\n\n"
    pesan += f"👤 *Dilaporkan oleh:* {pelapor}"

    try:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
        payload = {"chat_id": CHAT_ID, "text": pesan, "parse_mode": "Markdown"}
        response = requests.post(url, json=payload)
        return response.status_code == 200
    except:
        return False

def authenticate(username, password):
    c = conn.cursor()
    c.execute("SELECT password, role, nama_lengkap FROM users WHERE username=?", (username,))
    user = c.fetchone()
    
    if user:
        stored_hash = user[0]
        if check_password_hash(stored_hash, password):
            add_audit_log(user[2], "LOGIN", f"User {username} berhasil login.")
            return (user[1], user[2]) 
    return None

init_db()

if not os.path.exists("uploads"):
    os.makedirs("uploads")

LOKASI_DATA = {
    "Bolaang Mongondow": {
        "Lolayan": {
            "Bakan": ["Simpang 3 Masjid Bakan", "Parkiran Blok C", "Simpang 3 Tapagale"],
            "Lolayan": ["Lokasi Umum Lolayan"],
            "Matali Baru": ["Lokasi Umum Matali Baru"],
            "Mopusi": ["Lokasi Umum Mopusi"]
        }
    },
    "Bolaang Mongondow Selatan": {
        "Pinolosian Tengah": {
            "Tobayagan": ["Port Motandoi", "Simpang 2 Akses Masuk Motandoi"],
            "Tobayagan Selatan": ["Lokasi Umum Tobayagan Selatan"]
        },
        "Pinolosian Timur": {
            "Motandoi": ["Lokasi Umum Motandoi"],
            "Motandoi Selatan": ["Lokasi Umum Motandoi Selatan"],
            "Dumagin A": ["Lokasi Umum Dumagin A"],
            "Dumagin B": ["Lokasi Umum Dumagin B"],
            "Onggunoi Induk": ["Lokasi Umum Onggunoi Induk"],
            "Onggunoi Selatan": ["Lokasi Umum Onggunoi Selatan"],
            "Pidung": ["Lokasi Umum Pidung"],
            "Dayow": ["Lokasi Umum Dayow"]
        }
    }
}

TUNTUTAN_DATA = {
    "Ganti Rugi Tanaman/Banjir": "PT JRBM",
    "Kompensasi Lahan": "PT JRBM",
    "Tuntutan Pekerjaan": "PT JRBM & PT SMA",
    "Kualitas Lingkungan": "PT JRBM",
    "Sanksi Indisipliner / Hubungan Industrial": "PT SMA",
    "Isu Sosial": "PT JRBM",
    "Lainnya": "Lainnya"
}

def get_coordinates(desa, kab):
    coords = {
        "Bakan": (0.655, 124.050), "Lolayan": (0.670, 124.010),
        "Matali Baru": (0.630, 123.980), "Mopusi": (0.690, 124.030),
        "Tobayagan": (0.380, 123.950), "Tobayagan Selatan": (0.370, 123.960),
        "Dumagin A": (0.350, 123.900), "Motandoi": (0.400, 124.000)
    }
    if desa in coords: return coords[desa]
    if kab == "Bolaang Mongondow": return (0.650, 124.000)
    return (0.350, 123.950)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'role' not in st.session_state: st.session_state['role'] = None
if 'nama_lengkap' not in st.session_state: st.session_state['nama_lengkap'] = None

def inject_custom_css():
    st.markdown("""
        <style>
        .stApp, .main .block-container { background-color: #f8f9fa !important; }
        section[data-testid="stSidebar"] { background-color: #ffffff !important; border-right: 1px solid #e2e8f0 !important; }
        p, h1, h2, h3, h4, h5, h6, span, label, div.stMarkdown, .stSelectbox label, .stTextInput label, .stDateInput label, .stTimeInput label { color: #000000 !important; }

        .metric-card {
            padding: 20px; border-radius: 12px; margin-bottom: 20px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); font-family: 'Segoe UI', Tahoma;
            position: relative; overflow: hidden; border: 1px solid rgba(255,255,255,0.2);
        }
        .metric-card p, .metric-card div, .metric-card span { color: #ffffff !important; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); }
        
        .card-bm    { background: linear-gradient(135deg, #3b82f6, #2563eb); } 
        .card-bms   { background: linear-gradient(135deg, #ec4899, #db2777); } 
        .card-sma   { background: linear-gradient(135deg, #f59e0b, #d97706); } 
        .card-jrbm  { background: linear-gradient(135deg, #10b981, #059669); } 
        .card-lost  { background: linear-gradient(135deg, #ef4444, #dc2828); } 
        .card-total { background: linear-gradient(135deg, #6366f1, #4f46e5); } 
        .card-cepat { background: linear-gradient(135deg, #06b6d4, #0891b2); } 
        .card-lambat{ background: linear-gradient(135deg, #f97316, #ea580c); } 
        .card-avg-bm { background: linear-gradient(135deg, #8b5cf6, #7c3aed); } 
        .card-avg-bms { background: linear-gradient(135deg, #64748b, #475569); } 
        
        .metric-title { font-size: 13px; opacity: 0.95; margin-bottom: 8px; font-weight: 600; text-transform: uppercase; }
        .metric-value { font-size: 34px; font-weight: 800; line-height: 1.2;}
        .metric-subtitle { font-size: 12px; opacity: 0.85; margin-top: 6px; }

        button span, button p, .stAlert p, .stAlert span { color: inherit !important; }

        .login-box {
            background-color: #ffffff; padding: 40px; border-radius: 16px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; margin-top: 50px;
        }
        .sidebar-profile {
            background: linear-gradient(135deg, #0f172a, #1e293b); padding: 20px;
            border-radius: 12px; text-align: center; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .sidebar-profile p, .sidebar-profile h4, .sidebar-profile div { color: #ffffff !important; }
        
        .summary-box {
            background-color: #e0f2fe; border-left: 5px solid #0284c7; padding: 15px 20px;
            border-radius: 8px; margin-bottom: 20px; color: #0c4a6e !important;
        }

        @media print {
            section[data-testid="stSidebar"] { display: none !important; } 
            header[data-testid="stHeader"] { display: none !important; }   
            .stButton, button, .stDownloadButton { display: none !important; } 
            .main .block-container { max-width: 100% !important; padding: 0 !important; margin: 0 !important; }
            * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
            .js-plotly-plot, .metric-card { page-break-inside: avoid; margin-bottom: 15px;}
        }
        </style>
    """, unsafe_allow_html=True)

def force_black_text_on_plot(fig):
    fig.update_layout(
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
        font=dict(color="#000000", family="Segoe UI"),
        title_font=dict(color="#000000", size=16, family="Segoe UI", weight="bold"),
        legend_font=dict(color="#000000"),
        margin=dict(l=65, r=20, t=50, b=50)
    )
    fig.update_xaxes(showgrid=False, title_font=dict(color="#000000"), tickfont=dict(color="#000000"))
    fig.update_yaxes(showgrid=True, gridcolor="#e2e8f0", title_font=dict(color="#000000"), tickfont=dict(color="#000000"), title_standoff=15)
    return fig

def create_kpi_card(title, value, subtitle, css_class):
    return f"""
    <div class="metric-card {css_class}">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-subtitle">{subtitle}</div>
    </div>
    """

def generate_pdf(df):
    pdf = FPDF(orientation='L', unit='mm', format='A3')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Laporan Detail Kejadian Blokade", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [22, 25, 25, 35, 35, 30, 45, 20, 20, 20, 25, 45]
    columns = list(df.columns)
    
    for i, col_name in enumerate(columns):
        w = col_widths[i] if i < len(col_widths) else 30
        pdf.cell(w, 8, str(col_name), border=1, align='C')
    pdf.ln()

    pdf.set_font("Arial", '', 8)
    for _, row in df.iterrows():
        for i, value in enumerate(row):
            w = col_widths[i] if i < len(col_widths) else 30
            text_val = str(value)[:50] + ".." if len(str(value)) > 50 else str(value)
            pdf.cell(w, 8, text_val, border=1, align='C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laporan Blokade')
    return output.getvalue()

def generate_smart_summary(df):
    if df.empty: return "Tidak ada data pada periode ini untuk dianalisis."
    
    total_kasus = len(df)
    total_jam = df['Durasi (Jam)'].sum()
    top_isu = df['Isu'].mode()[0] if not df['Isu'].empty else "-"
    top_kabupaten = df['Kabupaten'].mode()[0] if not df['Kabupaten'].empty else "-"
    top_desa = df['Desa'].mode()[0] if not df['Desa'].empty else "-"
    lambat_pct = (len(df[df['Kategori Durasi'] == 'Lambat']) / total_kasus) * 100 if total_kasus > 0 else 0

    narasi = f"Pada periode yang difilter, tercatat sebanyak **{total_kasus} insiden blokade** yang mengakibatkan terbuangnya waktu operasional (*Lost Time*) sebesar **{total_jam:.1f} jam**. "
    narasi += f"Secara demografis, area paling rawan saat ini berpusat di wilayah **{top_kabupaten} (khususnya Desa {top_desa})**. "
    narasi += f"Akar masalah utama yang paling sering memicu insiden adalah terkait **'{top_isu}'**. "
    
    if lambat_pct > 50:
        narasi += f"⚠️ **Perhatian:** Sekitar **{lambat_pct:.0f}%** dari total insiden masuk dalam kategori penanganan lambat (> 9 Jam), mengindikasikan perlunya evaluasi strategi respons tim di lapangan."
    else:
        narasi += f"✅ Mayoritas insiden berhasil ditangani dengan respons cepat (≤ 9 Jam), menunjukkan performa tim mediasi lapangan yang cukup efektif."
        
    return narasi

def login_page():
    inject_custom_css()
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div class="login-box">
            <h2 style='text-align: center; font-weight: 800; color: #1e293b !important;'>🚧 PORTAL BLOKADE</h2>
            <p style='text-align: center; color: #64748b !important; margin-bottom: 25px;'>Sistem Pelaporan & Analitik Interaktif</p>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("👤 Username")
            password = st.text_input("🔑 Password", type="password")
            submit = st.form_submit_button("🚀 Masuk Ke Sistem", use_container_width=True)
            
            if submit:
                user_data = authenticate(username, password)
                if user_data:
                    st.session_state['logged_in'] = True
                    st.session_state['role'] = user_data[0]
                    st.session_state['nama_lengkap'] = user_data[1]
                    st.rerun()
                else:
                    st.error("❌ Username atau Password salah!")
        
        st.info("💡 **Kredensial Default:**\n- `admin` / `123` (Full Akses)\n- `spv` / `123` (Input Data)\n- `mgr` / `123` (Dashboard)")
        st.markdown("</div>", unsafe_allow_html=True)

def input_form_page():
    st.header("📝 Form Input Pelaporan Blokade")
    st.markdown("Silakan isi form di bawah ini dengan lengkap untuk mencatat insiden baru.")
    
    with st.container():
        st.subheader("🧑‍🤝‍🧑 Data Pelaku & Lokasi")
        col1, col2 = st.columns(2)
        with col1:
            nama_pelaku = st.text_input("👤 Nama Pelaku")
            no_hp = st.text_input("📱 Nomor HP Pelaku", placeholder="Contoh: 08123456789")
        
        with col2:
            kabupaten = st.selectbox("📍 Kabupaten", list(LOKASI_DATA.keys()))
            kecamatan = st.selectbox("🏘️ Kecamatan", list(LOKASI_DATA[kabupaten].keys()))
            desa = st.selectbox("🏡 Desa", list(LOKASI_DATA[kabupaten][kecamatan].keys()))
            lokasi = st.selectbox("📌 Lokasi Kejadian", LOKASI_DATA[kabupaten][kecamatan][desa])

    st.markdown("---")
    
    with st.container():
        st.subheader("⏱️ Waktu Kejadian")
        col3, col4, col5 = st.columns(3)
        with col3: tanggal = st.date_input("📅 Tanggal Kejadian")
        with col4: waktu_mulai = st.time_input("⏳ Waktu Mulai")
        with col5: waktu_selesai = st.time_input("⌛ Waktu Selesai")
            
        dt_mulai = datetime.datetime.combine(tanggal, waktu_mulai)
        dt_selesai = datetime.datetime.combine(tanggal, waktu_selesai)
        if dt_selesai < dt_mulai: dt_selesai += datetime.timedelta(days=1)
            
        durasi_detik = (dt_selesai - dt_mulai).total_seconds()
        durasi_jam = round(durasi_detik / 3600, 2)
        kategori_durasi = "Cepat" if durasi_jam <= 9 else "Lambat"
        st.info(f"⚙️ **Kalkulasi Durasi Otomatis:** {durasi_jam} Jam (Kategori: **{kategori_durasi}**)")

    st.markdown("---")
    
    with st.container():
        st.subheader("🎯 Tuntutan & Follow Up")
        col6, col7 = st.columns(2)
        with col6: isu = st.selectbox("📢 Jenis Isu / Tuntutan", list(TUNTUTAN_DATA.keys()))
        with col7: target_perusahaan = st.text_input("🏢 Target Perusahaan (Otomatis)", value=TUNTUTAN_DATA[isu], disabled=True)
        
        deskripsi = st.text_area("✍️ Deskripsi / Follow Up Kejadian", height=100)
        bukti = st.file_uploader("📎 Upload Bukti Foto (Opsional)", type=['jpg', 'png', 'jpeg'])
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Simpan Data Blokade ke Database", type="primary", use_container_width=True):
            
            # Validasi Ketat
            if not nama_pelaku.strip():
                st.error("❌ Nama Pelaku tidak boleh kosong!")
            elif no_hp and not re.match(r'^\d+$', no_hp):
                st.error("❌ Nomor HP hanya boleh berisi angka!")
            elif len(deskripsi.strip()) < 10:
                st.error("❌ Deskripsi minimal 10 karakter untuk kelengkapan laporan.")
            else:
                file_path = ""
                if bukti:
                    file_path = os.path.join("uploads", bukti.name)
                    with open(file_path, "wb") as f: f.write(bukti.getbuffer())

                new_data = {
                    "Tanggal": str(tanggal), "Pelaku": nama_pelaku, "No HP": no_hp,
                    "Kabupaten": kabupaten, "Kecamatan": kecamatan, "Desa": desa, "Lokasi": lokasi,
                    "Waktu Mulai": str(waktu_mulai), "Waktu Selesai": str(waktu_selesai), "Durasi (Jam)": durasi_jam,
                    "Kategori Durasi": kategori_durasi, "Isu": isu, "Target": target_perusahaan, "Deskripsi": deskripsi,
                    "File_Bukti": file_path
                }
                save_new_data(new_data)
                add_audit_log(st.session_state['nama_lengkap'], "CREATE", f"Menambahkan laporan blokade baru di {desa}.")
                st.success("✅ Data Blokade Berhasil Disimpan Permanen ke Sistem!")
                
                st.toast("Mempersiapkan Notifikasi...", icon="⏳")
                is_sent = send_telegram_notif(new_data, st.session_state['nama_lengkap'])
                if is_sent: 
                    st.info("📲 Notifikasi Telegram berhasil dikirim ke Grup Manajemen!")
                else: 
                    st.warning("🔔 Simulasi Notifikasi: Pesan peringatan terkirim (Harap setup `st.secrets` untuk live Telegram).")
                st.balloons()

def kelola_data_page():
    st.header("⚙️ Kelola Data & Pengaturan")
    tab1, tab2, tab3, tab4 = st.tabs(["📝 Edit & Hapus Laporan", "📥 Import Data Massal", "👥 Manajemen User", "🕵️‍♂️ Log Aktivitas"])
    
    with tab1:
        st.markdown("### Edit atau Hapus Data Laporan (Full CRUD)")
        df = load_data()
        if not df.empty:
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Tabel Laporan", type="primary"):
                edited_df.to_sql('laporan', conn, if_exists='replace', index=False)
                add_audit_log(st.session_state['nama_lengkap'], "UPDATE/DELETE", "Mengubah atau menghapus baris data laporan.")
                st.success("Tabel Laporan berhasil diperbarui!")
                st.rerun()
        else:
            st.info("Belum ada data laporan.")

    with tab2:
        st.markdown("### 📥 Import Data Historis")
        uploaded_file = st.file_uploader("Pilih file data (.csv atau .xlsx)", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                df_import = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
                st.dataframe(df_import.head())
                if st.button("🚀 Eksekusi Import Data", type="primary"):
                    df_import.to_sql('laporan', conn, if_exists='append', index=False)
                    add_audit_log(st.session_state['nama_lengkap'], "IMPORT", f"Mengimpor {len(df_import)} baris dari {uploaded_file.name}.")
                    st.success(f"✅ Berhasil mengimpor {len(df_import)} baris data!")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan saat membaca file: {e}")

    with tab3:
        if st.session_state['role'] == 'Administrator':
            st.markdown("### Manajemen Akun Pengguna")
            df_users = pd.read_sql("SELECT * FROM users", conn)
            # Menyembunyikan password di tampilan agar lebih aman
            df_users_display = df_users.copy()
            df_users_display['password'] = '********'
            st.dataframe(df_users_display, use_container_width=True)
            st.info("💡 Edit user secara langsung melalui database SQLite di server untuk keamanan hash.")
        else:
            st.error("🚫 Akses Ditolak: Harus login sebagai Administrator.")
            
    with tab4:
        if st.session_state['role'] in ['Administrator', 'Manager']:
            st.markdown("### 🕵️‍♂️ Log Riwayat Aktivitas (Audit Trail)")
            df_logs = pd.read_sql("SELECT * FROM audit_logs ORDER BY id DESC", conn)
            st.dataframe(df_logs, use_container_width=True, hide_index=True)
        else:
            st.error("🚫 Akses Ditolak: Hanya Manager dan Administrator yang dapat melihat log audit.")

def dashboard_page():
    inject_custom_css()
    
    col_head1, col_head2 = st.columns([3, 1])
    with col_head1:
        st.header("📊 Dashboard Analitik Blokade")
        st.caption("Ringkasan Performa dan Analisa Kendala di Lapangan secara Real-time")
    with col_head2:
        st.markdown("""<div style="margin-top: 15px;"><button onclick="window.print()" style="background-color: #0f172a; color: white; padding: 12px 20px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; width: 100%; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">🖨️ Cetak Full Dashboard</button></div>""", unsafe_allow_html=True)
    
    df = load_data()
    if df.empty:
        st.warning("⚠️ Belum ada data laporan di Database.")
        return

    df['Datetime_Kejadian'] = pd.to_datetime(df['Tanggal'])
    df['Tanggal_Date'] = df['Datetime_Kejadian'].dt.date
    df['Tahun'] = df['Datetime_Kejadian'].dt.year
    df['Bulan_Tahun'] = df['Datetime_Kejadian'].dt.strftime('%b %Y')
    df['YearMonth_Sort'] = df['Datetime_Kejadian'].dt.strftime('%Y-%m') # Untuk sorting tren
    df['Quarter'] = 'Q' + df['Datetime_Kejadian'].dt.quarter.astype(str)
    
    # Ekstraksi Hari untuk Heatmap
    hari_map = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
    df['Hari'] = df['Datetime_Kejadian'].dt.dayofweek.map(hari_map)
    
    # Ekstraksi Jam untuk Heatmap
    df['Jam'] = df['Waktu Mulai'].apply(lambda x: str(x).split(':')[0] if pd.notna(x) else '00')
    df['Jam'] = df['Jam'].str.zfill(2) + ":00"

    st.markdown("#### 🔍 Filter Analitik Canggih")
    col_f1, col_f2, col_f3 = st.columns([1, 1, 1.5])
    with col_f1: selected_year = st.selectbox("📆 Pilih Tahun", ["Semua Tahun"] + list(sorted(df['Tahun'].unique(), reverse=True)))
    with col_f2: selected_q = st.selectbox("🗂️ Periode Triwulan", ["Semua", "Q1", "Q2", "Q3", "Q4"])
    with col_f3: date_range = st.date_input("🗓️ Filter Rentang Waktu (Awal - Akhir)", [])
    
    df_raw = df.copy()
    if selected_year != "Semua Tahun": df = df[df['Tahun'] == selected_year]
    if selected_q != "Semua": df = df[df['Quarter'] == selected_q]
    if len(date_range) == 2:
        start_date, end_date = date_range
        df = df[(df['Tanggal_Date'] >= start_date) & (df['Tanggal_Date'] <= end_date)]

    st.markdown("<hr/>", unsafe_allow_html=True)
    st.markdown("#### 🤖 Ringkasan Eksekutif Otomatis")
    st.markdown(f"<div class='summary-box'>{generate_smart_summary(df)}</div>", unsafe_allow_html=True)

    if selected_year != "Semua Tahun":
        df_prev = df_raw[df_raw['Tahun'] == (selected_year - 1)]
        m1, m2, m3 = st.columns(3)
        m1.metric("🛡️ Total Blokade", f"{len(df)} Kasus", f"{len(df) - len(df_prev)} Kasus vs Tahun Lalu", "inverse")
        m2.metric("⏳ Total Lost Time", f"{df['Durasi (Jam)'].sum():.1f} Jam", f"{df['Durasi (Jam)'].sum() - df_prev['Durasi (Jam)'].sum():.1f} Jam vs Tahun Lalu", "inverse")
        m3.metric("⚡ Rata-rata Durasi", f"{df['Durasi (Jam)'].mean():.1f} Jam" if len(df)>0 else "0 Jam", "Evaluasi Kecepatan", "off")
        st.markdown("<br>", unsafe_allow_html=True)

    if df.empty: return

    st.markdown("#### 🎯 Key Performance Indicators")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(create_kpi_card("Total Blokade BM", len(df[df['Kabupaten'] == "Bolaang Mongondow"]), "Kejadian Area BM", "card-bm"), unsafe_allow_html=True)
    with c2: st.markdown(create_kpi_card("Total Blokade BMS", len(df[df['Kabupaten'] == "Bolaang Mongondow Selatan"]), "Kejadian Area BMS", "card-bms"), unsafe_allow_html=True)
    with c3: st.markdown(create_kpi_card("Blokade PT SMA", len(df[df['Target'].str.contains('PT SMA', na=False)]), "Melibatkan PT SMA", "card-sma"), unsafe_allow_html=True)
    with c4: st.markdown(create_kpi_card("Blokade PT JRBM", len(df[df['Target'].str.contains('PT JRBM', na=False)]), "Melibatkan PT JRBM", "card-jrbm"), unsafe_allow_html=True)

    st.markdown("---")
    
    st.markdown("#### 🗺️ Pemetaan Geospasial Titik Blokade")
    df_map = df.groupby(['Kabupaten', 'Desa']).size().reset_index(name='Jumlah Kasus')
    df_map['lat'] = df_map.apply(lambda r: get_coordinates(r['Desa'], r['Kabupaten'])[0], axis=1)
    df_map['lon'] = df_map.apply(lambda r: get_coordinates(r['Desa'], r['Kabupaten'])[1], axis=1)
    fig_map = px.scatter_mapbox(df_map, lat="lat", lon="lon", size="Jumlah Kasus", color="Kabupaten", hover_name="Desa", zoom=8, mapbox_style="carto-positron", color_discrete_map={"Bolaang Mongondow": "#3b82f6", "Bolaang Mongondow Selatan": "#ef4444"})
    fig_map.update_layout(margin={"r":0,"t":10,"l":0,"b":0}, paper_bgcolor="#ffffff")
    st.plotly_chart(fig_map, use_container_width=True)
    st.markdown("---")

    st.markdown("#### 🔥 Analisis Waktu Rawan & Tren Isu (Advanced)")
    col_adv1, col_adv2 = st.columns(2)
    
    with col_adv1:
        # Fitur Baru 1: Heatmap Waktu
        heatmap_data = df.groupby(['Hari', 'Jam']).size().reset_index(name='Jumlah')
        fig_heat = px.density_heatmap(
            heatmap_data, x="Jam", y="Hari", z="Jumlah",
            category_orders={
                "Hari": ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"],
                "Jam": [f"{str(i).zfill(2)}:00" for i in range(24)]
            },
            title="Heatmap Jam & Hari Kejadian Rawan",
            color_continuous_scale="Reds"
        )
        st.plotly_chart(force_black_text_on_plot(fig_heat), use_container_width=True, theme=None)
        
    with col_adv2:
        # Fitur Baru 2: Time-Series Tren Isu
        df_trend_isu = df.groupby(['YearMonth_Sort', 'Bulan_Tahun', 'Isu']).size().reset_index(name='Jumlah').sort_values('YearMonth_Sort')
        fig_trend = px.bar(df_trend_isu, x="Bulan_Tahun", y="Jumlah", color="Isu", title="Tren Pergerakan Isu per Bulan", barmode="stack", color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(force_black_text_on_plot(fig_trend), use_container_width=True, theme=None)

    st.markdown("#### 📊 Distribusi Kasus")
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        isu_count = df['Isu'].value_counts().reset_index()
        isu_count.columns = ['Jenis Isu', 'Jumlah']
        fig_isu = px.bar(isu_count.head(3), x='Jenis Isu', y='Jumlah', title="Top 3 Jenis Isu Utama", text_auto=True, color='Jenis Isu', color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(force_black_text_on_plot(fig_isu), use_container_width=True, theme=None)

    with col_chart2:
        target_count = df['Target'].value_counts().reset_index()
        target_count.columns = ['Target Perusahaan', 'Jumlah']
        fig_target = px.pie(target_count, values='Jumlah', names='Target Perusahaan', title="Analisis Target Perusahaan", hole=0.45)
        st.plotly_chart(force_black_text_on_plot(fig_target), use_container_width=True, theme=None)

    if WORDCLOUD_AVAILABLE:
        st.markdown("#### ☁️ Analisis Teks (Word Cloud)")
        text_data = " ".join(df['Deskripsi'].dropna().astype(str))
        if text_data.strip():
            wordcloud = WordCloud(width=800, height=300, background_color='white', colormap='viridis', max_words=100).generate(text_data)
            fig_wc, ax_wc = plt.subplots(figsize=(10, 4))
            ax_wc.imshow(wordcloud, interpolation='bilinear')
            ax_wc.axis("off")
            st.pyplot(fig_wc)

    st.markdown("---")
    st.markdown("### 🗃️ Data Detail Blokade & Galeri")
    display_df = df.drop(columns=['Tahun', 'Bulan_Tahun', 'Quarter', 'Datetime_Kejadian', 'Tanggal_Date', 'YearMonth_Sort', 'Hari', 'Jam'], errors='ignore')
    
    col_d1, col_d2, col_d3 = st.columns([1, 1, 4])
    with col_d1:
        if FPDF_AVAILABLE: st.download_button("📄 Export PDF", data=generate_pdf(df), file_name="Laporan_Blokade.pdf", mime="application/pdf", type="primary")
    with col_d2:
        if EXCEL_AVAILABLE: st.download_button("📥 Export Excel", data=generate_excel(display_df), file_name="Laporan_Blokade.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    
    st.dataframe(display_df, use_container_width=True, hide_index=True)

def main():
    inject_custom_css()
    
    if not st.session_state['logged_in']:
        login_page()
    else:
        st.sidebar.markdown(f"""
        <div class="sidebar-profile">
            <div style="font-size: 50px; margin-bottom: 5px;">🧑‍💼</div>
            <h4 style="margin: 0;">{st.session_state['nama_lengkap']}</h4>
            <p style="margin: 5px 0 0 0; font-size: 12px; color: #94a3b8 !important; border-top: 1px solid #334155; padding-top: 8px;">🟢 {st.session_state['role']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.sidebar.markdown("### 🧭 Navigasi Utama")
        
        if st.session_state['role'] == "Administrator":
            menu = ["📝 Input Data Blokade", "⚙️ Kelola Data (Admin)", "📊 Dashboard Analitik"]
        elif st.session_state['role'] == "Supervisor": 
            menu = ["📝 Input Data Blokade", "⚙️ Kelola Data (Admin)"]
        else: 
            menu = ["📊 Dashboard Analitik", "⚙️ Kelola Data (Admin)"]

        choice = st.sidebar.radio("Pilih Menu:", menu, label_visibility="collapsed")
        
        st.sidebar.markdown("---")
        if st.sidebar.button("🚪 Logout / Keluar", use_container_width=True):
            add_audit_log(st.session_state['nama_lengkap'], "LOGOUT", "User telah keluar dari sistem.")
            for key in ['logged_in', 'role', 'nama_lengkap']: st.session_state[key] = None
            st.rerun()

        if choice == "📝 Input Data Blokade": input_form_page()
        elif choice == "📊 Dashboard Analitik": dashboard_page()
        elif choice == "⚙️ Kelola Data (Admin)": kelola_data_page()

if __name__ == "__main__":
    main()
from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import numpy as np
import random
import math
import pickle
import os
import openpyxl
from datetime import datetime


# ============================================================
# IMPORT FUNGSI DARI MODULE LAIN
# ============================================================
from implementasi_cmab import (
    hitung_bmr,
    hitung_tdee,
    inisialisasi_q_values,
    epsilon_greedy_selection,
    update_q_value,
    decay_epsilon_exponential,
    get_top_arms,
    MU, EPSILON_MAX, EPSILON_MIN
)
from hitung_gizi_dan_penambahan_nasi import (
    proses_penambahan_nasi_ke_arms
)

# ============================================================
# KONFIGURASI
# ============================================================

INPUT_ARMS_FILE    = "arms_kombinasi_menu.xlsx"
INPUT_NUTRISI_FILE = "Data_Nutrisi_Selected_Clean.xlsx"
N_EPISODES         = 30      
ALPHA              = 0.5     # Mixed: 50% nutrisi + 50% rasa
SESSION_FILE       = "cmab_session.pkl"   # Simpan state CMAB antar request
ARMS_FILE          = "cmab_arms.pkl"
HASIL_FILE         = "hasil_cmab_realuser.xlsx"

app = Flask(__name__)
app.secret_key = "dash-cmab-secret-2025"

# ============================================================
# LOAD DATA ARMS SAAT STARTUP (1x saja)
# ============================================================

print("\n" + "="*60)
print("EVALUASI CMAB — REAL USER")
print("="*60)
df_arms_base = pd.read_excel(INPUT_ARMS_FILE, sheet_name="arms")
df_nasi      = pd.read_excel(INPUT_NUTRISI_FILE, sheet_name="nasi")
print(f"Arms (tanpa nasi): {len(df_arms_base):,}")

# LOAD & CACHE DATA NUTRISI 
print("\nLoading data nutrisi...")
DATA_NUTRISI_CACHE = pd.read_excel(INPUT_NUTRISI_FILE, sheet_name=None)

# GABUNGKAN JUS/SMOOTHIE KE BUAH (sama seperti di pembentukan_arm.py)
if 'jus_smoothie' in DATA_NUTRISI_CACHE:
    jus = DATA_NUTRISI_CACHE['jus_smoothie'].copy()
    jus = jus.rename(columns={'nama_menu': 'nama_buah'})
    
    # Tambah kolom kosong ke buah asli
    if 'bahan-bahan' not in DATA_NUTRISI_CACHE['buah'].columns:
        DATA_NUTRISI_CACHE['buah']['bahan-bahan'] = '-'
    if 'cara_pembuatan' not in DATA_NUTRISI_CACHE['buah'].columns:
        DATA_NUTRISI_CACHE['buah']['cara_pembuatan'] = '-'
    if 'waktu_memasak' not in DATA_NUTRISI_CACHE['buah'].columns:
        DATA_NUTRISI_CACHE['buah']['waktu_memasak'] = '-'
    
    # Concat
    DATA_NUTRISI_CACHE['buah'] = pd.concat([DATA_NUTRISI_CACHE['buah'], jus], ignore_index=True)

# ============================================================
# HELPER: SIMPAN & LOAD SESSION STATE
# ============================================================

def save_state(state: dict):
    """Simpan state CMAB ke file pickle"""
    with open(SESSION_FILE, "wb") as f:
        pickle.dump(state, f)

def load_state() -> dict:
    """Load state CMAB dari file pickle"""
    with open(SESSION_FILE, "rb") as f:
        return pickle.load(f)

def state_exists() -> bool:
    return os.path.exists(SESSION_FILE)

def delete_state():
    for f in [SESSION_FILE, ARMS_FILE]:
        if os.path.exists(f):
            os.remove(f)

def save_arms(df):
    df.to_pickle(ARMS_FILE)

def load_arms():
    return pd.read_pickle(ARMS_FILE)

# ============================================================
# FUNGSI FILTER ALERGI 
# ============================================================

def check_alergi(menu_names, alergi_list):
    """Cek apakah ada alergi dalam list menu"""
    for menu in menu_names:
        if pd.isna(menu):
            continue
        menu_lower = str(menu).lower()
        for alergi in alergi_list:
            if alergi in menu_lower:
                return True
    return False

def filter_arms_by_alergi(df_arms, alergi_list):
    """Filter arms yang mengandung bahan alergi/pantangan user"""
    if not alergi_list:
        return df_arms
    
    print(f"\n🔍 Filtering arms dengan alergi: {alergi_list}")
    print(f"   Total arms sebelum filter: {len(df_arms)}")
    
    def has_alergi(row):
        menu_names = [
            row.get("sarapan_lauk", ""),
            row.get("sarapan_sayur", ""),
            row.get("sarapan_buah", ""),
            row.get("siang_lauk", ""),
            row.get("siang_sayur", ""),
            row.get("siang_buah", ""),
            row.get("malam_lauk", ""),
            row.get("malam_sayur", ""),
            row.get("malam_buah", "")
        ]
        return check_alergi(menu_names, alergi_list)
    
    df_filtered = df_arms[~df_arms.apply(has_alergi, axis=1)].copy()
    print(f"   Total arms setelah filter: {len(df_filtered)}")
    return df_filtered

# ============================================================
# HELPER FUNCTIONS - UNTUK DETAIL MENU
# ============================================================

def safe_float(value):
    """Konversi nilai ke float"""
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except:
        return 0.0

def safe_str(value):
    """Handle NaN/None untuk string"""
    if pd.isna(value) or value is None:
        return '-'
    return str(value)

def get_menu_details(menu_name, sheet_name):
    """Dapatkan detail lengkap menu dari dataset"""
    try:
        df_menu = DATA_NUTRISI_CACHE[sheet_name] 
        if sheet_name == 'buah':
            menu = df_menu[df_menu['nama_buah'] == menu_name]
            nama_col = 'nama_buah'
        else:
            menu = df_menu[df_menu['nama_menu'] == menu_name]
            nama_col = 'nama_menu'   

        if len(menu) > 0:
            menu_data = menu.iloc[0]
            return {
                'nama': str(menu_data.get(nama_col, menu_name)),
                'bahan': str(menu_data.get('bahan-bahan', '-')),
                'cara_pembuatan': str(menu_data.get('cara_pembuatan', '-')),
                'waktu_memasak': str(menu_data.get('waktu_memasak', '-')),
                'porsi': str(menu_data.get('porsi', '-')),
                'nutrisi': {
                    'kalori': safe_float(menu_data.get('kalori', 0)),
                    'protein': safe_float(menu_data.get('protein', 0)),
                    'karbohidrat': safe_float(menu_data.get('karbohidrat', 0)),
                    'lemak': safe_float(menu_data.get('lemak', 0)),
                    'lemak_jenuh': safe_float(menu_data.get('lemak_jenuh', 0)),
                    'kolesterol': safe_float(menu_data.get('kolesterol', 0)),
                    'natrium': safe_float(menu_data.get('natrium', 0)),
                    'kalium': safe_float(menu_data.get('kalium', 0)),
                    'serat': safe_float(menu_data.get('serat', 0))
                }
            }
    except Exception as e:
        print(f"Error loading menu {menu_name}: {e}")
    
    return {
        'nama': safe_str(menu_name),
        'bahan': '-',
        'cara_pembuatan': '-',
        'waktu_memasak': '-',
        'porsi': '-',
        'nutrisi': {k: 0.0 for k in ['kalori', 'protein', 'karbohidrat', 'lemak', 
                                      'lemak_jenuh', 'kolesterol', 'natrium', 'kalium', 'serat']}
    }

def get_susu_details(susu_name):
    """Dapatkan detail susu dari dataset"""
    if pd.isna(susu_name) or susu_name == '-' or susu_name == 'nan':
        return None
    try:
        df_susu = DATA_NUTRISI_CACHE['susu']  
        susu = df_susu[df_susu['nama'] == susu_name]
        if len(susu) > 0:
            susu_data = susu.iloc[0]
            return {
                'nama': str(susu_data.get('nama', susu_name)),
                'porsi': '1 gelas (250ml)',
                'nutrisi': {
                    'kalori': safe_float(susu_data.get('kalori', 0)),
                    'protein': safe_float(susu_data.get('protein', 0)),
                    'karbohidrat': safe_float(susu_data.get('karbohidrat', 0)),
                    'lemak': safe_float(susu_data.get('lemak', 0)),
                    'lemak_jenuh': safe_float(susu_data.get('lemak_jenuh', 0)),
                    'kolesterol': safe_float(susu_data.get('kolesterol', 0)),
                    'natrium': safe_float(susu_data.get('natrium', 0)),
                    'kalium': safe_float(susu_data.get('kalium', 0)),
                    'serat': safe_float(susu_data.get('serat', 0))
                }
            }
    except Exception as e:
        print(f"Error loading susu {susu_name}: {e}")
    return None

def format_nasi(nasi_info):
    """Format informasi nasi jadi user-friendly"""
    if pd.isna(nasi_info) or nasi_info == "Tidak perlu" or "Tidak perlu" in str(nasi_info):
        return "Tidak perlu"
    try:
        import re
        match = re.search(r'\(([0-9.]+)\s*porsi\)', str(nasi_info))
        if match:
            porsi = float(match.group(1))
            gram = porsi * 100  # 1 porsi = 100g
            centong = gram / 50  # 1 centong = 50g
            nama_nasi = nasi_info.split('(')[0].strip()
            return f"{nama_nasi}: {centong:.1f} centong (~{gram:.0f}g)"
        else:
            return str(nasi_info)
    except:
        return str(nasi_info)

def cek_kelebihan_natrium(arm_row):
    """Cek kelebihan natrium dan berikan warning"""
    TOLERANSI = 0.10
    NATRIUM_PER_SDT_GARAM = 2325
    total_natrium = arm_row["total_natrium"]
    batas_natrium = 2300  # mg
    ambang_natrium = batas_natrium * (1 + TOLERANSI)  # 2530 mg
    
    if total_natrium <= ambang_natrium:
        return {
            'excess': False,
            'excess_mg': 0,
            'saran_sdt': 0,
            'warna': '#28a745',
            'kategori': 'Baik ✓',
            'pesan': 'Natrium dalam batas aman'
        }
    else:
        excess_mg = total_natrium - ambang_natrium
        saran_sdt = excess_mg / NATRIUM_PER_SDT_GARAM
        return {
            'excess': True,
            'excess_mg': excess_mg,
            'saran_sdt': saran_sdt,
            'warna': '#ffc107',
            'kategori': 'Perlu Dikurangi ⚠️',
            'pesan': f'Natrium berlebih (+{excess_mg:.0f}mg dari batas toleransi). Kurangi garam/penyedap rasa sebanyak {saran_sdt:.1f} sdt.'
        }

# ============================================================
# ROUTES
# ============================================================

@app.route("/", methods=["GET"])
def index():
    """Halaman awal: input profil pasien"""
    resume = None
    # Jika ada session yang belum selesai, redirect ke episode
    if state_exists():
        try:
            state = load_state()
            current_ep = state.get("current_episode", 1)
            if current_ep <= N_EPISODES:
                resume = {"nama": state["nama"], "episode": current_ep}
        except Exception:
            pass
    return render_template("input.html", resume = resume, error=None)

@app.route("/lanjut")
def lanjut():
    """Langsung lanjut ke episode terakhir tanpa input profil ulang."""
    if not state_exists():
        return redirect(url_for("index"))
    return redirect(url_for("episode"))

@app.route("/mulai", methods=["POST"])
def mulai():
    """Proses input profil -> fase offline -> simpan state -> lalu mulai episode 1"""

    # 1. Ambil data form
    nama           = request.form.get("nama", "").strip()
    jenis_kelamin  = request.form.get("jenis_kelamin", "")
    usia           = int(request.form.get("usia", 0))
    berat          = float(request.form.get("berat", 0))
    tinggi         = float(request.form.get("tinggi", 0))
    aktivitas      = request.form.get("aktivitas", "sedentary")
    alergi_raw     = request.form.get("alergi", "")
    alergi_list    = [a.strip().lower() for a in alergi_raw.split(",") if a.strip()]

    if not all([nama, jenis_kelamin, usia, berat, tinggi]):
        return render_template("input.html", resume = None, error="Semua data wajib diisi!")

    # 2. Hitung TDEE & kebutuhan nutrisi 
    from hitung_gizi_dan_penambahan_nasi import hitung_kebutuhan_makronutrien
    bmr             = hitung_bmr(berat, tinggi, usia, jenis_kelamin)
    tdee = hitung_tdee(bmr, aktivitas, berat, tinggi) 
    kebutuhan       = hitung_kebutuhan_makronutrien(tdee) # kebutuhan seluruh nutrisi

    print(f"\n{'='*60}")
    print(f"PASIEN BARU: {nama}")
    print(f"   TDEE: {tdee:.2f} kkal")
    print(f"   Alergi: {alergi_list or 'tidak ada'}")
    print(f"{'='*60}")

    # 3. Tambah nasi ke arms berdasarkan TDEE 
    print("\nMenambahkan nasi ke arms...")
    df_arms = proses_penambahan_nasi_ke_arms(df_arms_base.copy(), tdee, df_nasi)
    print(f"Arms siap: {len(df_arms):,}")

    # 4. Filter alergi 
    if alergi_list:
        df_arms = filter_arms_by_alergi(df_arms, alergi_list)

    if len(df_arms) == 0:
        return render_template("input.html", resume = None, 
            error="Tidak ada menu yang tersedia setelah filter alergi.")

    # 5. FASE OFFLINE: Inisialisasi Q(a) 
    print("\nFASE OFFLINE: Menginisialisasi Q(a)...")
    q_values = inisialisasi_q_values(df_arms, kebutuhan, debug_first_n=0)
    print(f"Q(a) diinisialisasi untuk {len(q_values):,} arms")

    # 6. Simpan state ke file 
    state = {
        "nama"             : nama,
        "tdee"             : round(tdee, 2),
        "kebutuhan"        : kebutuhan,
        "alergi"           : alergi_list,
        "q_values"         : q_values,
        "df_arms_index"    : df_arms["arm_id"].tolist(),
        "current_episode"  : 1,
        "history"          : [],
        "n_explore"        : 0,
        "n_exploit"        : 0,
    }
    save_state(state)
    save_arms(df_arms)

    print("\n✓ Fase offline selesai. Mulai episode 1...")
    return redirect(url_for("episode"))


@app.route("/episode", methods=["GET"])
def episode():
    """Tampilkan rekomendasi untuk episode saat ini"""

    if not state_exists():
        return redirect(url_for("index"))

    state      = load_state()
    current_ep = state["current_episode"]

    # Sudah selesai?
    if current_ep > N_EPISODES:
        return redirect(url_for("hasil"))

    q_values   = state["q_values"]
    kebutuhan  = state["kebutuhan"]  
    df_arms  = load_arms()

    # ε-greedy pilih arm 
    epsilon = decay_epsilon_exponential(current_ep - 1)  # episode dimulai dari 0 di fungsi decay
    arm_id, is_explore = epsilon_greedy_selection(q_values, epsilon)

    # # Untuk tracking apakah eksplorasi atau eksploitasi, kita cek ulang
    # # (ini hanya untuk statistik, bukan untuk memilih arm)
    # is_explore = random.random() < epsilon
    arm_row = df_arms[df_arms["arm_id"] == arm_id].iloc[0]

    # Simpan arm_id yang dipilih ke state (supaya saat submit tahu arm mana)
    state["pending_arm_id"] = arm_id
    state["pending_epsilon"] = epsilon
    state["pending_is_explore"] = is_explore
    save_state(state)

    print(f"Episode {current_ep:2d}: Arm #{arm_id:5d} | ε={epsilon:.4f} | "
        f"{'EKSPLORASI' if is_explore else 'EKSPLOITASI'}")
    
    # Natrium warning
    natrium_warning = cek_kelebihan_natrium(arm_row)

    rekomendasi = {
        'arm_id': int(arm_id),
        'sarapan': {
            'lauk': get_menu_details(arm_row['sarapan_lauk'], 'sarapan'),
            'sayur': get_menu_details(arm_row['sarapan_sayur'], 'sayur'),
            'buah': get_menu_details(arm_row['sarapan_buah'], 'buah'),
            'susu': get_susu_details(arm_row.get('sarapan_susu')),
            'nasi': format_nasi(arm_row.get('nasi_sarapan', 'Tidak perlu'))
        },
        'siang': {
            'lauk': get_menu_details(arm_row['siang_lauk'], 'makan_siang'),
            'sayur': get_menu_details(arm_row['siang_sayur'], 'sayur'),
            'buah': get_menu_details(arm_row['siang_buah'], 'buah'),
            'susu': get_susu_details(arm_row.get('siang_susu')),
            'nasi': format_nasi(arm_row.get('nasi_siang', 'Tidak perlu'))
        },
        'malam': {
            'lauk': get_menu_details(arm_row['malam_lauk'], 'makan_malam'),
            'sayur': get_menu_details(arm_row['malam_sayur'], 'sayur'),
            'buah': get_menu_details(arm_row['malam_buah'], 'buah'),
            'susu': get_susu_details(arm_row.get('malam_susu')),
            'nasi': format_nasi(arm_row.get('nasi_malam', 'Tidak perlu'))
        }
    }

    return render_template(
        "episode.html",
        episode    = current_ep,
        total      = N_EPISODES,
        nama       = state["nama"],
        tdee       = round(state["tdee"]),
        epsilon    = round(epsilon, 4), 
        is_explore = is_explore,
        arm_id     = arm_id,
        rekomendasi = rekomendasi,       
        kebutuhan = kebutuhan,            
        natrium_warning = natrium_warning 
    )


@app.route("/submit_rating", methods=["POST"])
def submit():
    """Terima rating dari user -> update Q(a) -> simpan history -> lanjut ke episode berikutnya"""

    if not state_exists():
        return redirect(url_for("index"))

    state      = load_state()
    arm_id     = int(request.form.get("arm_id"))
    episode_no = int(request.form.get("episode"))
    rating     = int(request.form.get("rating", 3))

    q_values  = state["q_values"]
    epsilon   = state["pending_epsilon"]
    is_explore= state["pending_is_explore"]

    # Hitung Reward 
    reward_rasa    = rating / 5.0
    reward_nutrisi = q_values[arm_id]["cosine"] - q_values[arm_id]["penalti"]
    reward_gabungan = ALPHA * reward_nutrisi + (1 - ALPHA) * reward_rasa

    # Update Q(a) - pakai fungsi dari implementasi_cmab
    update_q_value(q_values, arm_id, reward_gabungan)

    # Simpan history episode ini
    state["history"].append({
        "episode"         : episode_no,
        "arm_id"          : arm_id,
        "rating"          : rating,
        "reward_nutrisi"  : reward_nutrisi,
        "reward_rasa"     : reward_rasa,
        "reward_gabungan" : reward_gabungan,
        "q_value"         : q_values[arm_id]["q_value"],
        "epsilon"         : epsilon,
        "is_explore"      : is_explore,
    })

    # Update counter eksplorasi/eksploitasi 
    if is_explore:
        state["n_explore"] += 1
    else:
        state["n_exploit"] += 1

    # Update Q dan lanjut ke episode berikutnya 
    state["q_values"]        = q_values
    state["current_episode"] = episode_no + 1
    save_state(state)

    print(f"Episode {episode_no}: Arm #{arm_id}, Rating {rating}/5, "
          f"R_gabungan={reward_gabungan:.4f}, Q(a)={q_values[arm_id]['q_value']:.4f}")

    # Jika sudah 30 episode → ke halaman hasil
    if episode_no >= N_EPISODES:
        return redirect(url_for("hasil"))

    return redirect(url_for("episode"))


@app.route("/hasil")
def hasil():
    """Halaman hasil akhir setelah 30 episode"""

    if not state_exists():
        return redirect(url_for("index"))

    state    = load_state()
    nama     = state["nama"]
    history  = state["history"]
    q_values = state["q_values"]
    df_arms  = load_arms()


    if not history:
        return redirect(url_for("episode"))


    # Top 5 Arms - pakai fungsi dari implementasi_cmab
    kebutuhan = state["kebutuhan"]  
    top_arms_df = get_top_arms(q_values, df_arms, kebutuhan, n=5)
    top_arms = []
    for _, row in top_arms_df.iterrows():
        top_arms.append({
            "arm_id"     : int(row["arm_id"]),
            "q_value"    : round(row["q_value"], 4),
            "reward_nutrisi": round(row["reward_nutrisi"], 4),  
            "mape"       : round(row["mape"], 1),  
            "n_selected" : int(row["n_selected"]),
            "sarapan_lauk": str(row.get("sarapan_lauk", "-")),
            "siang_lauk"  : str(row.get("siang_lauk", "-")),
            "malam_lauk"  : str(row.get("malam_lauk", "-")),
        })

    # Statistik 
    selected_arm_ids = [h["arm_id"] for h in history]
    q_vals_final = [q_values[aid]["q_value"] for aid in selected_arm_ids]
    ratings      = [h["rating"] for h in history]

    # Save ke Excel 
    _save_hasil_excel(history, top_arms_df, nama)

    return render_template(
        "hasil.html",
        nama           = state["nama"],
        total_ep       = len(history),
        avg_rating     = round(sum(ratings) / len(ratings), 2),
        max_r          = max(ratings),
        min_r          = min(ratings),
        avg_q          = round(sum(q_vals_final) / len(q_vals_final), 4),
        n_explore      = state.get("n_explore", 0),
        n_exploit      = state.get("n_exploit", 0),
        top5           = top_arms,
        history        = history,
    )

@app.route("/detail_mape/<int:rank>")
def detail_mape(rank):
    """Halaman detail MAPE untuk Top 1-5 arms"""
    if not state_exists():
        return redirect(url_for("index"))
    
    state = load_state()
    df_arms = load_arms()
    kebutuhan = state["kebutuhan"]
    q_values = state["q_values"]
    
    # Get Top 5 arms
    top_arms_df = get_top_arms(q_values, df_arms, kebutuhan, n=5)
    
    # Validate rank
    if rank < 1 or rank > 5:
        return "Invalid rank", 404
    
    # Get specific arm data
    arm_row = top_arms_df.iloc[rank-1]
    
    # Prepare Top 5 list for sidebar
    top_arms_list = []
    for idx, row in top_arms_df.iterrows():
        top_arms_list.append({
            'arm_id': int(row['arm_id']),
            'q_value': round(row['q_value'], 4),
            'mape': round(row['mape'], 2)
        })
    
    return render_template(
        "detail_mape.html",
        current_rank=rank,
        arm=arm_row.to_dict(),
        kebutuhan=kebutuhan,
        top_arms=top_arms_list
    )


@app.route("/download")
def download():
    """Download hasil Excel"""
    from flask import send_file
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, HASIL_FILE)
    if os.path.exists(file_path):
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        return send_file(file_path, as_attachment=True,
                         download_name=f"hasil_cmab_{ts}.xlsx")
    return "File tidak ditemukan.", 404


@app.route("/reset")
def reset():
    """Reset session (mulai dari awal)"""
    delete_state()
    return redirect(url_for("index"))


# ============================================================
# HELPER: SIMPAN HASIL KE EXCEL
# ============================================================

def _save_hasil_excel(history: list, top_arms_df, nama: str):
    """
    Simpan history + top 5 arms ke Excel (MULTI-USER SUPPORT)
    Setiap user akan punya 3 sheets sendiri: UserN_History, UserN_Top5, UserN_Stats
    """
    try:
        # Gunakan path absolute (sama dengan directory script)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, HASIL_FILE)
        
        # === MULTI-USER SUPPORT ===
        # Cek apakah file sudah ada
        if os.path.exists(file_path):
            # Baca file existing
            existing_wb = openpyxl.load_workbook(file_path)
            
            # Hitung jumlah user yang sudah ada
            existing_sheets = existing_wb.sheetnames
            user_numbers = []
            for sheet in existing_sheets:
                if sheet.startswith("User") and "_" in sheet:
                    try:
                        num = int(sheet.split("_")[0].replace("User", ""))
                        user_numbers.append(num)
                    except:
                        pass
            
            # User baru adalah nomor terakhir + 1
            new_user_num = max(user_numbers) + 1 if user_numbers else 1
            existing_wb.close()  # Close workbook
        else:
            # File belum ada, ini user pertama
            new_user_num = 1
        
        # Nama sheet untuk user baru
        prefix = f"User{new_user_num}"
        
        # Buat DataFrame Stats
        stats_data = {
            "Nama": [nama],
            "Total Episode": [len(history)],
            "Avg Rating": [round(sum(h["rating"] for h in history) / len(history), 2)],
            "Avg Q(a)": [round(sum(h["q_value"] for h in history) / len(history), 4)],
            "Max Rating": [max(h["rating"] for h in history)],
            "Min Rating": [min(h["rating"] for h in history)],
            "N Explore": [sum(1 for h in history if h["is_explore"])],
            "N Exploit": [sum(1 for h in history if not h["is_explore"])],
        }
        stats_df = pd.DataFrame(stats_data)
        
        # ============================================================
        # PILIH KOLOM PENTING UNTUK TOP 5 ARMS
        # ============================================================
        kolom_top5 = [
            # Identifikasi
            "arm_id",
            "q_value",
            "reward_nutrisi",
            "mape",
            "n_selected",
            
            # Menu Sarapan
            "sarapan_lauk",
            "sarapan_sayur",
            "sarapan_buah",
            "sarapan_susu",
            "nasi_sarapan",  

            # Nutrisi Sarapan (sudah include nasi)
            "sarapan_kalori",
            "sarapan_karbohidrat",
            "sarapan_protein",
            "sarapan_lemak",
            "sarapan_lemak_jenuh",
            "sarapan_kolesterol",
            "sarapan_natrium",
            "sarapan_kalium",
            "sarapan_serat",
            
            # Menu Siang
            "siang_lauk",
            "siang_sayur",
            "siang_buah",
            "siang_susu",
            "nasi_siang",  

            # Nutrisi Siang (sudah include nasi)
            "siang_kalori",
            "siang_karbohidrat",
            "siang_protein",
            "siang_lemak",
            "siang_lemak_jenuh",
            "siang_kolesterol",
            "siang_natrium",
            "siang_kalium",
            "siang_serat",
            
            # Menu Malam
            "malam_lauk",
            "malam_sayur",
            "malam_buah",
            "malam_susu",
            "nasi_malam",  

            # Nutrisi Malam (sudah include nasi)
            "malam_kalori",
            "malam_karbohidrat",
            "malam_protein",
            "malam_lemak",
            "malam_lemak_jenuh",
            "malam_kolesterol",
            "malam_natrium",
            "malam_kalium",
            "malam_serat",
            
            # Total Nutrisi (SETELAH ditambah nasi)
            "total_kalori",
            "total_karbohidrat",
            "total_protein",
            "total_lemak",
            "total_lemak_jenuh",
            "total_kolesterol",
            "total_natrium",
            "total_kalium",
            "total_serat"
        ]
        
        # Filter hanya kolom yang ada di top_arms_df
        kolom_ada = [col for col in kolom_top5 if col in top_arms_df.columns]
        top5_clean = top_arms_df[kolom_ada].copy()
        
        # ============================================================
        # Mode append jika file ada, write jika file baru
        mode = "a" if os.path.exists(file_path) else "w"
        
        # Tulis ke Excel dengan mode append
        with pd.ExcelWriter(file_path, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
            pd.DataFrame(history).to_excel(writer, sheet_name=f"{prefix}_History", index=False)
            top5_clean.to_excel(writer, sheet_name=f"{prefix}_Top5", index=False)  # ← Pakai top5_clean
            stats_df.to_excel(writer, sheet_name=f"{prefix}_Stats", index=False)
        
        print(f"✓ Data {nama} disimpan sebagai {prefix} ke {file_path}")
        print(f"  → Total user dalam file: {new_user_num}")
        print(f"  → Kolom Top5: {len(kolom_ada)} kolom")
        return file_path
    except Exception as e:
        print(f"⚠️ Gagal simpan Excel: {e}")
        import traceback
        traceback.print_exc()
        return None


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print("Jalankan: http://localhost:5001")
    app.run(debug=True, port=5001)
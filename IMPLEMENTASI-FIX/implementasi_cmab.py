import pandas as pd
import numpy as np
import random
import math
from tqdm import tqdm

# ============================================================
# KONFIGURASI
# ============================================================

INPUT_ARMS_BASE_FILE = "arms_kombinasi_menu.xlsx"  # Arms TANPA nasi
INPUT_NUTRISI_FILE = "Data_Nutrisi_Selected_Clean.xlsx"

TOLERANSI = 0.10  # 10%

# Konfigurasi ε-greedy
EPSILON_MAX = 0.1
EPSILON_MIN = 0.02
N_EPISODES = 30

# Hitung faktor decay (μ) untuk exponential decay
MU = -math.log(EPSILON_MIN / EPSILON_MAX) / N_EPISODES
print(f"Faktor decay (μ) dihitung: {MU:.6f}")

random.seed(42)
np.random.seed(42)

# ============================================================
# FUNGSI HELPER MATEMATIKA (CORE)
# ============================================================

def hitung_bmr(berat_kg, tinggi_cm, usia, jenis_kelamin):

    if jenis_kelamin.lower() == "laki-laki":
        bmr = (9.99 * berat_kg) + (6.25 * tinggi_cm) - (4.92 * usia) + 5
    else:  # perempuan
        bmr = (9.99 * berat_kg) + (6.25 * tinggi_cm) - (4.92 * usia) - 161
    
    return bmr

def hitung_bmi(berat_kg, tinggi_cm):

    # Konversi tinggi dari cm ke meter
    tinggi_m = tinggi_cm / 100
    
    # Hitung BMI
    bmi = berat_kg / (tinggi_m ** 2)
    
    # Klasifikasi BMI
    if bmi < 17.0:
        kategori = "Kurus (tingkat berat)"
        adjustment = +500 
    elif 17.0 <= bmi < 18.5:
        kategori = "Kurus (tingkat ringan)"
        adjustment = +500  
    elif 18.5 <= bmi <= 25.0:
        kategori = "Normal"
        adjustment = 0  
    elif 25.1 <= bmi <= 27.0:
        kategori = "Gemuk (tingkat ringan)"
        adjustment = -500  
    else:  # bmi > 27.0
        kategori = "Gemuk (tingkat berat)"
        adjustment = -500  
    
    return bmi, kategori, adjustment

def hitung_tdee(bmr, aktivitas_level, berat_kg, tinggi_cm):

    PAL_MAP = {
        "sedentary": 1.4,
        "moderately_active": 1.7,
        "very_active": 2.0,
        "extremely_active": 2.4
    }
    
    pal = PAL_MAP.get(aktivitas_level.lower(), 1.4)
    tdee = bmr * pal
    
    # Adjustment berdasarkan BMI
    bmi, kategori, adjustment = hitung_bmi(berat_kg, tinggi_cm)
    
    print(f"\n INFO BMI:")
    print(f"   BMI: {bmi:.1f}")
    print(f"   Kategori: {kategori}")
    
    if adjustment != 0:
        print(f"   TDEE base: {tdee:.0f} kkal")
        print(f"   Adjustment: {adjustment:+d} kkal")
        tdee = tdee + adjustment
        print(f"   TDEE final: {tdee:.0f} kkal")
    else:
        print(f"   TDEE: {tdee:.0f} kkal (tidak ada adjustment)")
    
    return tdee

def cosine_similarity(v1, v2):

    v1 = np.array(v1)
    v2 = np.array(v2)
    
    if np.linalg.norm(v1) == 0 or np.linalg.norm(v2) == 0:
        return 0
    
    return np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2))

def hitung_penalti(total_nutrisi, kebutuhan, debug=False):
    """
    Penalti = rata-rata persentase kelebihan nutrisi
    Hanya untuk: lemak jenuh, kolesterol, natrium
    """
    BATAS_NUTRISI = {
        "lemak_jenuh": kebutuhan["lemak_jenuh"],
        "kolesterol": 150,
        "natrium": 2300
    }
    
    excess_list = []
    
    if debug:
        print("\n      🔍 DEBUG PENALTI:")
    
    for nutrisi in ["lemak_jenuh", "kolesterol", "natrium"]:
        nilai = total_nutrisi.get(nutrisi, 0)
        batas = BATAS_NUTRISI[nutrisi]
        ambang = batas * (1 + TOLERANSI)
        
        if nilai > ambang:
            excess = (nilai - ambang) / ambang
        else:
            excess = 0
        
        excess_list.append(excess)
        
        if debug:
            print(f"         {nutrisi:15} : nilai={nilai:7.2f}, batas={batas:7.2f}, ambang={ambang:7.2f}, excess={excess:.4f}")
    
    penalti = sum(excess_list) / len(excess_list)
    
    if debug:
        print(f"         PENALTI TOTAL   : sum={sum(excess_list):.4f} / {len(excess_list)} = {penalti:.4f}")
    
    return penalti

# ============================================================
# FUNGSI MAPE DENGAN TOLERANSI +/- 10%
# ============================================================

def hitung_mape(arm_row, kebutuhan_nutrisi, toleransi=0.10):
    """
    Hitung MAPE dengan toleransi range ±10%
    
    NUTRISI BIASA (harus dipenuhi):
    - Range: 90-110% dari target
    - Error jika < 90% ATAU > 110%
    
    NUTRISI DIBATASI (lemak jenuh, kolesterol, natrium):
    - Range: 0-110% dari target (semakin rendah semakin bagus)
    - Error HANYA jika > 110%
    """
    
    # Daftar nutrisi yang harus dibatasi (semakin rendah semakin bagus)
    NUTRISI_DIBATASI = {
        "lemak_jenuh": 3,    # index di array
        "kolesterol": 4,
        "natrium": 5
    }
    
    target = [
        kebutuhan_nutrisi["karbohidrat"],     # 0
        kebutuhan_nutrisi["protein"],          # 1
        kebutuhan_nutrisi["lemak"],            # 2
        kebutuhan_nutrisi["lemak_jenuh"],      # 3 ← DIBATASI
        kebutuhan_nutrisi["kolesterol"],       # 4 ← DIBATASI
        kebutuhan_nutrisi["natrium"],          # 5 ← DIBATASI
        kebutuhan_nutrisi["kalium"],           # 6
        kebutuhan_nutrisi["serat"]             # 7
    ]
    
    aktual = [
        arm_row["total_karbohidrat"],          # 0
        arm_row["total_protein"],              # 1
        arm_row["total_lemak"],                # 2
        arm_row["total_lemak_jenuh"],          # 3 ← DIBATASI
        arm_row["total_kolesterol"],           # 4 ← DIBATASI
        arm_row["total_natrium"],              # 5 ← DIBATASI
        arm_row["total_kalium"],               # 6
        arm_row["total_serat"]                 # 7
    ]
    
    percentage_errors = []
    
    for idx, (a, t) in enumerate(zip(aktual, target)):
        if t <= 0:
            continue
        
        # Cek apakah nutrisi ini termasuk yang harus dibatasi
        is_dibatasi = idx in NUTRISI_DIBATASI.values()
        
        if is_dibatasi:
            # ============================================
            # NUTRISI DIBATASI (lemak jenuh, kolesterol, natrium)
            # Semakin rendah semakin bagus
            # ============================================
            batas_atas = t * (1 + toleransi)  # 110% dari target
            
            if a <= batas_atas:
                # ≤ 110% → BAGUS! Error = 0%
                error = 0
            else:
                # > 110% → KEBANYAKAN! Hitung error
                error = abs(a - batas_atas) / batas_atas * 100
        
        else:
            # ============================================
            # NUTRISI BIASA (harus dipenuhi, tidak boleh terlalu rendah/tinggi)
            # ============================================
            batas_bawah = t * (1 - toleransi)  # 90%
            batas_atas = t * (1 + toleransi)   # 110%
            
            if batas_bawah <= a <= batas_atas:
                # DALAM RANGE → error = 0%
                error = 0
            elif a < batas_bawah:
                # DI BAWAH RANGE → kurang asupan
                error = abs(a - batas_bawah) / batas_bawah * 100
            else:  # a > batas_atas
                # DI ATAS RANGE → kelebihan asupan
                error = abs(a - batas_atas) / batas_atas * 100
        
        percentage_errors.append(error)
    
    mape = sum(percentage_errors) / len(percentage_errors) if percentage_errors else 0
    return mape


# ============================================================
# FUNGSI DECAY EPSILON
# ============================================================

def decay_epsilon_exponential(episode, epsilon_max=EPSILON_MAX, mu=MU):
    """Decay epsilon menggunakan fungsi eksponensial"""
    epsilon = epsilon_max * math.exp(-mu * episode)
    return max(EPSILON_MIN, epsilon)

# ============================================================
# FASE OFFLINE: INISIALISASI NILAI Q(a)
# ============================================================

def inisialisasi_q_values(df_arms, kebutuhan_nutrisi, debug_first_n=3):
    """
    FASE OFFLINE: Hitung nilai Q(a) awal untuk setiap arm
    """
    print("\n=== FASE OFFLINE: Inisialisasi Q(a) ===")
    
    # Vektor target DASH (8 komponen)
    v_target = [
        kebutuhan_nutrisi["karbohidrat"],          # g
        kebutuhan_nutrisi["protein"],              # g
        kebutuhan_nutrisi["lemak"],                # g
        kebutuhan_nutrisi["lemak_jenuh"],          # g
        kebutuhan_nutrisi["kolesterol"] / 1000,    # mg → g (150mg = 0.15g)
        kebutuhan_nutrisi["natrium"] / 1000,       # mg → g (2300mg = 2.3g)
        kebutuhan_nutrisi["kalium"] / 1000,        # mg → g (4700mg = 4.7g)
        kebutuhan_nutrisi["serat"]                 # g
    ]
    
    print(f"\nVEKTOR TARGET (kebutuhan nutrisi):")
    print(f"   Karbohidrat  : {v_target[0]:.2f} g")
    print(f"   Protein      : {v_target[1]:.2f} g")
    print(f"   Lemak        : {v_target[2]:.2f} g")
    print(f"   Lemak Jenuh  : {v_target[3]:.2f} g")
    print(f"   Kolesterol   : {v_target[4]:.2f} g")
    print(f"   Natrium      : {v_target[5]:.2f} g")
    print(f"   Kalium       : {v_target[6]:.2f} g")
    print(f"   Serat        : {v_target[7]:.2f} g")
    
    q_values = {}
    debug_count = 0
    
    for idx, row in tqdm(df_arms.iterrows(), total=len(df_arms), desc="Inisialisasi Q(a)"):
        arm_id = row["arm_id"]
        
        # Vektor nutrisi aktual dari arm
        v_aktual = [
            row["total_karbohidrat"],
            row["total_protein"],
            row["total_lemak"],
            row["total_lemak_jenuh"],
            row["total_kolesterol"] / 1000,
            row["total_natrium"] / 1000,
            row["total_kalium"] / 1000,
            row["total_serat"]
        ]
        
        # DEBUG: Print detail untuk 3 arm pertama
        show_debug = (debug_count < debug_first_n)
        
        if show_debug:
            print(f"\n{'='*70}")
            print(f"DEBUG ARM #{arm_id}")
            print(f"{'='*70}")
            print(f"   Sarapan: {row['sarapan_lauk']}")
            print(f"   Siang  : {row['siang_lauk']}")
            print(f"   Malam  : {row['malam_lauk']}")
            print(f"\n   VEKTOR AKTUAL (nutrisi arm):")
            print(f"      Karbohidrat  : {v_aktual[0]:.2f} g  (target: {v_target[0]:.2f})")
            print(f"      Protein      : {v_aktual[1]:.2f} g  (target: {v_target[1]:.2f})")
            print(f"      Lemak        : {v_aktual[2]:.2f} g  (target: {v_target[2]:.2f})")
            print(f"      Lemak Jenuh  : {v_aktual[3]:.2f} g  (target: {v_target[3]:.2f})")
            print(f"      Kolesterol   : {v_aktual[4]:.2f} g (target: {v_target[4]:.2f})")
            print(f"      Natrium      : {v_aktual[5]:.2f} g (target: {v_target[5]:.2f})")
            print(f"      Kalium       : {v_aktual[6]:.2f} g (target: {v_target[6]:.2f})")
            print(f"      Serat        : {v_aktual[7]:.2f} g  (target: {v_target[7]:.2f})")
        
        # Hitung cosine similarity
        cosine = cosine_similarity(v_target, v_aktual)
        
        if show_debug:
            print(f"\n   COSINE SIMILARITY:")
            print(f"      dot(v_target, v_aktual) = {np.dot(v_target, v_aktual):.2f}")
            print(f"      ||v_target||            = {np.linalg.norm(v_target):.2f}")
            print(f"      ||v_aktual||            = {np.linalg.norm(v_aktual):.2f}")
            print(f"      cosine                  = {cosine:.6f}")
        
        # Hitung penalti
        total_nutrisi = {
            "lemak_jenuh": row["total_lemak_jenuh"],
            "kolesterol": row["total_kolesterol"],
            "natrium": row["total_natrium"]
        }
        
        penalti = hitung_penalti(total_nutrisi, kebutuhan_nutrisi, debug=show_debug)
        
        # Reward nutrisi = cosine - penalti
        reward_nutrisi = cosine - penalti
        
        if show_debug:
            print(f"\n   REWARD NUTRISI:")
            print(f"      cosine  : {cosine:.6f}")
            print(f"      penalti : {penalti:.6f}")
            print(f"      reward  : {cosine:.6f} - {penalti:.6f} = {reward_nutrisi:.6f}")
            debug_count += 1
        
        # Simpan ke q_values
        q_values[arm_id] = {
            "q_value": reward_nutrisi,  # Q(a) awal = reward_nutrisi
            "cosine": cosine,
            "penalti": penalti,
            "n_selected": 0,
            "total_reward": 0
        }
    
    print(f"\n✓ Inisialisasi {len(q_values):,} arms selesai")
    
    # Statistik Q(a) awal
    q_vals = [v["q_value"] for v in q_values.values()]
    print(f"   Q(a) min: {min(q_vals):.4f}")
    print(f"   Q(a) max: {max(q_vals):.4f}")
    print(f"   Q(a) mean: {np.mean(q_vals):.4f}")
    
    return q_values

# ============================================================
# EPSILON-GREEDY SELECTION
# ============================================================

def epsilon_greedy_selection(q_values, epsilon):
    rand_val = random.random()  
    if rand_val < epsilon:
        arm_id = random.choice(list(q_values.keys()))
        is_explore = True
    else:
        arm_id = max(q_values.items(), key=lambda x: x[1]["q_value"])[0]
        is_explore = False
    
    return arm_id, is_explore  

# ============================================================
# UPDATE Q(a)
# ============================================================

def update_q_value(q_values, arm_id, reward):
    """Update Q(a) dengan rata-rata reward"""
    q_values[arm_id]["n_selected"] += 1
    q_values[arm_id]["total_reward"] += reward
    
    # Q(a) = rata-rata reward
    q_values[arm_id]["q_value"] = (
        q_values[arm_id]["total_reward"] / q_values[arm_id]["n_selected"]
    )

# ============================================================
# OUTPUT DAN EVALUASI
# ============================================================

def get_top_arms(q_values, df_arms, kebutuhan_nutrisi, n=5):
    """Ambil n arm terbaik berdasarkan Q(a) tertinggi yang pernah dipilih"""
    selected_arms = {aid: info for aid, info in q_values.items() if info["n_selected"] > 0}

    sorted_arms = sorted(selected_arms.items(), 
                        key=lambda x: x[1]["q_value"], 
                        reverse=True)
        
    top_arm_ids = [arm_id for arm_id, _ in sorted_arms[:n]]
    top_arms_df = df_arms[df_arms["arm_id"].isin(top_arm_ids)].copy()
    
    # Tambahkan info Q(a)
    top_arms_df["q_value"] = top_arms_df["arm_id"].map(
        lambda x: q_values[x]["q_value"]
    )
    top_arms_df["n_selected"] = top_arms_df["arm_id"].map(
        lambda x: q_values[x]["n_selected"]
    )
    top_arms_df["reward_nutrisi"] = top_arms_df["arm_id"].map(  
        lambda x: q_values[x]["cosine"] - q_values[x]["penalti"]
    )
    top_arms_df["mape"] = top_arms_df.apply(  
        lambda row: hitung_mape(row, kebutuhan_nutrisi), axis=1
    )
    
    return top_arms_df.sort_values("q_value", ascending=False)


import pandas as pd
import numpy as np
from tqdm import tqdm

# ============================================================
# FUNGSI HITUNG KEBUTUHAN NUTRISI
# ============================================================

def hitung_kebutuhan_makronutrien(tdee):
    """
    Menghitung kebutuhan makronutrien berdasarkan TDEE.
    (dari proposal Bab 2.3 dan Tabel 2.1)
    """
    kebutuhan = {
        "kalori": tdee,
        "karbohidrat": (tdee * 0.55) / 4,    # 55% kalori, 1g = 4 kkal
        "protein": (tdee * 0.18) / 4,         # 18% kalori, 1g = 4 kkal
        "lemak": (tdee * 0.27) / 9,           # 27% kalori, 1g = 9 kkal
        "lemak_jenuh": (tdee * 0.06) / 9,     # 6% kalori, 1g = 9 kkal
        "kolesterol": 150,                    # mg (tetap)
        "natrium": 2300,                      # mg (tetap)
        "kalium": 4700,                       # mg (tetap)
        "serat": 30                           # g (tetap)
    }
    return kebutuhan

# ============================================================
# FUNGSI CEK MENU TINGGI KARBO
# ============================================================

def cek_menu_tinggi_karbo(nama_menu):
    """
    Cek apakah menu adalah makanan tinggi karbohidrat yang tidak butuh nasi tambahan.
    (dari proposal Bab 3.7)
    """
    
    nama_lower = str(nama_menu).lower()
    
    # Daftar menu tinggi karbo yang tidak perlu nasi
    keywords_tinggi_karbo = ["nasi goreng", "mi goreng", "mie goreng", "sushi", "sandwich", "bakmi goreng"]
    
    return any(keyword in nama_lower for keyword in keywords_tinggi_karbo)


# ============================================================
# FUNGSI TAMBAH NASI KE ARM (PER WAKTU MAKAN)
# ============================================================

def tambah_nasi_ke_arm(arm_row, target_karbo, df_nasi, toleransi_karbo=5):
    
    # Distribusi karbo per waktu makan (30-40-30) 
    distribusi = {
        "sarapan": 0.30,
        "siang": 0.40,
        "malam": 0.30
    }
    
    # Inisialisasi hasil
    hasil = {
        "nasi_sarapan": None,
        "nasi_siang": None,
        "nasi_malam": None,
        "total_karbohidrat": arm_row["total_karbohidrat"],
        "total_kalori": arm_row["total_kalori"],
        "total_protein": arm_row["total_protein"],
        "total_lemak": arm_row["total_lemak"],
        "total_lemak_jenuh": arm_row["total_lemak_jenuh"],
        "total_natrium": arm_row["total_natrium"],
        "total_kalium": arm_row["total_kalium"],
        "total_serat": arm_row["total_serat"],
        "total_kolesterol": arm_row.get("total_kolesterol", 0)
    }
    
    # ========================================
    # LOOP PER WAKTU MAKAN
    # ========================================
    for waktu, proporsi in distribusi.items():
        
        # 1. CEK: Apakah menu waktu INI tinggi karbo?
        nama_lauk = arm_row.get(f"{waktu}_lauk", "")
        
        if cek_menu_tinggi_karbo(nama_lauk):
            # HANYA waktu INI yang tidak perlu nasi
            hasil[f"nasi_{waktu}"] = "Tidak perlu (menu sudah tinggi karbo)"
            continue  # Skip ke waktu berikutnya
        
        # 2. HITUNG target karbo untuk waktu INI
        target_karbo_waktu = target_karbo * proporsi
        
        # 3. AMBIL karbo menu waktu INI (LANGSUNG DARI KOLOM!)
        kolom_karbo = f"{waktu}_karbohidrat"
        karbo_menu_waktu = arm_row[kolom_karbo]
        
        # 4. HITUNG kekurangan karbo waktu INI
        sisa_karbo_waktu = target_karbo_waktu - karbo_menu_waktu
        
        # 5. CEK: Apakah perlu tambah nasi?
        if sisa_karbo_waktu <= toleransi_karbo:
            hasil[f"nasi_{waktu}"] = "Tidak perlu"
            continue
        
        # 6. HITUNG porsi nasi
        nasi = df_nasi.sample(1).iloc[0]
        porsi = sisa_karbo_waktu / nasi["karbohidrat"]
        
        if porsi <= 0:
            hasil[f"nasi_{waktu}"] = "Tidak perlu"
            continue
        
        # 7. SIMPAN info nasi
        hasil[f"nasi_{waktu}"] = f"{nasi['nama']} ({porsi} porsi)"
        
        # 8. UPDATE total nutrisi
        hasil["total_karbohidrat"] += nasi["karbohidrat"] * porsi
        hasil["total_kalori"] += nasi["kalori"] * porsi
        hasil["total_protein"] += nasi["protein"] * porsi
        hasil["total_lemak"] += nasi["lemak"] * porsi
        hasil["total_lemak_jenuh"] += nasi["lemak_jenuh"] * porsi
        hasil["total_natrium"] += nasi["natrium"] * porsi
        hasil["total_kalium"] += nasi["kalium"] * porsi
        hasil["total_serat"] += nasi["serat"] * porsi

        # 9. UPDATE nutrisi PER WAKTU MAKAN 
        # Kalo ditambah nasi, nutrisi waktu makan juga harus nambah
        hasil[f"{waktu}_karbohidrat"] = arm_row[f"{waktu}_karbohidrat"] + (nasi["karbohidrat"] * porsi)
        hasil[f"{waktu}_kalori"] = arm_row[f"{waktu}_kalori"] + (nasi["kalori"] * porsi)
        hasil[f"{waktu}_protein"] = arm_row[f"{waktu}_protein"] + (nasi["protein"] * porsi)
        hasil[f"{waktu}_lemak"] = arm_row[f"{waktu}_lemak"] + (nasi["lemak"] * porsi)
        hasil[f"{waktu}_lemak_jenuh"] = arm_row[f"{waktu}_lemak_jenuh"] + (nasi["lemak_jenuh"] * porsi)
        hasil[f"{waktu}_natrium"] = arm_row[f"{waktu}_natrium"] + (nasi["natrium"] * porsi)
        hasil[f"{waktu}_kalium"] = arm_row[f"{waktu}_kalium"] + (nasi["kalium"] * porsi)
        hasil[f"{waktu}_serat"] = arm_row[f"{waktu}_serat"] + (nasi["serat"] * porsi)
        hasil[f"{waktu}_kolesterol"] = arm_row.get(f"{waktu}_kolesterol", 0) + (nasi.get("kolesterol", 0) * porsi)
    
    return hasil

# ============================================================
# FUNGSI UTAMA: TAMBAH NASI KE SEMUA ARMS
# ============================================================

def proses_penambahan_nasi_ke_arms(df_arms, tdee, df_nasi):
    """
    Fungsi utama untuk menambahkan nasi ke semua arms.
    Fungsi ini dipanggil dari evaluasi_cmab_realuser.py setelah user input profil.
    """
    print(f"\nMenambahkan nasi ke {len(df_arms):,} arms berdasarkan TDEE {tdee:.0f} kkal...")
    
    # Hitung kebutuhan karbohidrat
    kebutuhan = hitung_kebutuhan_makronutrien(tdee)
    target_karbo = kebutuhan["karbohidrat"]
    
    print(f"   Target karbohidrat: {target_karbo:.1f} g/hari")
    print(f"   Distribusi: Sarapan {target_karbo*0.3:.1f}g | Siang {target_karbo*0.4:.1f}g | Malam {target_karbo*0.3:.1f}g")
    
    # Proses setiap arm
    arms_dengan_nasi = []
    
    for idx, row in tqdm(df_arms.iterrows(), total=len(df_arms), desc="   Progress", unit=" arm", leave=False):
        # Tambahkan nasi PER WAKTU MAKAN
        hasil = tambah_nasi_ke_arm(row, target_karbo, df_nasi)
        
        # Gabungkan data
        arm_updated = row.to_dict()
        arm_updated.update(hasil)
        arms_dengan_nasi.append(arm_updated)
    
    df_hasil = pd.DataFrame(arms_dengan_nasi)
    
    # ========================================
    # STATISTIK PER WAKTU MAKAN
    # ========================================
    arms_sarapan_nasi = 0
    arms_siang_nasi = 0
    arms_malam_nasi = 0
    
    for idx, row in df_hasil.iterrows():
        # Cek sarapan
        if row['nasi_sarapan'] and 'Tidak perlu' not in str(row['nasi_sarapan']):
            arms_sarapan_nasi += 1
        
        # Cek siang
        if row['nasi_siang'] and 'Tidak perlu' not in str(row['nasi_siang']):
            arms_siang_nasi += 1
        
        # Cek malam
        if row['nasi_malam'] and 'Tidak perlu' not in str(row['nasi_malam']):
            arms_malam_nasi += 1
    
    print(f"Selesai!")
    print(f"Sarapan dengan nasi: {arms_sarapan_nasi:,} / {len(df_hasil):,} arms")
    print(f"Siang dengan nasi  : {arms_siang_nasi:,} / {len(df_hasil):,} arms")
    print(f"Malam dengan nasi  : {arms_malam_nasi:,} / {len(df_hasil):,} arms")
    
    return df_hasil

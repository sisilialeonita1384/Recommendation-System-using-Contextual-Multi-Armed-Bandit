import pandas as pd
import random
from tqdm import tqdm

# ============================================================
# KONFIGURASI
# ============================================================

INPUT_FILE = "Data_Nutrisi_Selected_Clean.xlsx"
OUTPUT_FILE = "arms_kombinasi_menu.xlsx"

NUTRISI_COLS = [
    "kalori", "protein", "lemak", "lemak_jenuh",
    "kolesterol", "karbohidrat", "natrium",
    "kalium", "serat"
]

random.seed(42)  # Untuk reproducibility

# ============================================================
# LOAD DATA
# ============================================================

def load_data():
    data = pd.read_excel(INPUT_FILE, sheet_name=None)
    
    # GABUNGKAN JUS/SMOOTHIE KE BUAH 
    if 'jus_smoothie' in data:
        jus = data['jus_smoothie'].copy()
        
        # Rename kolom jus agar match dengan buah
        jus = jus.rename(columns={'nama_menu': 'nama_buah'})
        
        # Buah asli tidak punya bahan/cara_pembuatan, jadi tambahkan kolom kosong
        if 'bahan-bahan' not in data['buah'].columns:
            data['buah']['bahan-bahan'] = '-'
        if 'cara_pembuatan' not in data['buah'].columns:
            data['buah']['cara_pembuatan'] = '-'
        if 'waktu_memasak' not in data['buah'].columns:
            data['buah']['waktu_memasak'] = '-'
        
        # Concat buah + jus (sekarang kolom sudah match)
        data['buah'] = pd.concat([data['buah'], jus], ignore_index=True)
        
        print(f"Menggabungkan buah ({len(data['buah']) - len(jus)}) + jus ({len(jus)}) = {len(data['buah'])} items")

    # FILTER BUAH KERING (karbo tinggi)
    buah_sebelum = len(data['buah'])
    data['buah'] = data['buah'][~data['buah']['nama_buah'].str.contains('kering', case=False, na=False)]
    buah_filtered = buah_sebelum - len(data['buah'])
    if buah_filtered > 0:
        print(f"Filter buah kering: {buah_filtered} items dihapus (karbo terlalu tinggi)")

    # FILTER BUAH TINGGI KARBO (kurma, kismis, melon)
    buah_sebelum = len(data['buah'])
    buah_tinggi_karbo = ['kurma', 'kismis', 'melon']
    pattern = '|'.join(buah_tinggi_karbo)
    data['buah'] = data['buah'][~data['buah']['nama_buah'].str.contains(pattern, case=False, na=False)]
    buah_filtered = buah_sebelum - len(data['buah'])
    if buah_filtered > 0:
        print(f"Filter buah tinggi karbo (kurma, kismis, melon): {buah_filtered} items dihapus")
    
    # Tampilkan jumlah data per kategori
    for kategori, df in data.items():
        if kategori not in ['jus_smoothie', 'nasi']:
            print(f"{kategori}: {len(df)} items")
    
    return data

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def pilih_slot_susu():

    waktu = ["sarapan", "siang", "malam"]
    jumlah = random.choice([2, 3])
    
    if jumlah == 3:
        return {w: True for w in waktu}
    
    terpilih = random.sample(waktu, 2)
    return {w: (w in terpilih) for w in waktu}

def ambil_susu(df_susu):
    """Pilih satu jenis susu secara acak"""
    return df_susu.sample(1).iloc[0]

def hitung_total_nutrisi(items):
    """
    Hitung total nutrisi dari seluruh komponen menu
    """
    total = {col: 0 for col in NUTRISI_COLS}
    
    for item in items:
        for col in NUTRISI_COLS:
            total[col] += item[col]
    
    return total

# ============================================================
# FUNGSI UTAMA: PEMBENTUKAN ARM
# ============================================================

def create_arms_stok_harian(data, n_arms=10000):
    """
    Membuat N arms melalui random sampling dengan konsep stok harian
    
    Konsep Stok Harian:
    -------------------
    Setiap arm merepresentasikan menu 1 hari dengan:
    - 3 sayur berbeda (dipilih sekali, lalu dibagi ke 3 waktu makan)
    - 3 buah berbeda (dipilih sekali, lalu dibagi ke 3 waktu makan)
    - Lauk berbeda untuk setiap waktu makan (dipilih independent)
    """
    
    # Extract datasets
    sarapan = data["sarapan"]
    siang = data["makan_siang"]
    malam = data["makan_malam"]
    sayur = data["sayur"]
    buah = data["buah"]
    susu = data["susu"]

    # Header
    print("\n" + "="*60)
    print("PEMBENTUKAN ARM - RANDOM SAMPLING")
    print("="*60)
    
    # Informasi konsep
    print("\nKONSEP 'STOK HARIAN':")
    print("   Setiap arm = 1 hari menu dengan:")
    print("   • 3 sayur BERBEDA → dibagi ke sarapan, siang, malam")
    print("   • 3 buah BERBEDA → dibagi ke sarapan, siang, malam")
    print("   • Lauk BERBEDA untuk setiap waktu makan")
    print("   • Susu 2-3 porsi per hari (random)")
    
    # Informasi teknis
    print(f"\nPARAMETER:")
    print(f"   Target jumlah arms: {n_arms:,}")
    print(f"   Metode: Random sampling")
    print(f"   - Lauk: dengan pengembalian (dapat berulang antar arm)")
    print(f"   - Sayur & Buah: tanpa pengembalian dalam 1 arm")

    # Inisialisasi
    arms = []
    
    print(f"\nMembuat {n_arms:,} arms...")
    
    # Loop pembuatan arms dengan progress bar
    for arm_id in tqdm(range(1, n_arms + 1), desc="Progress", unit="arm"):
        
        # 1. SAMPLING LAUK 
        lauk_pagi = sarapan.sample(1).iloc[0]
        lauk_siang = siang.sample(1).iloc[0]
        lauk_malam = malam.sample(1).iloc[0]
        
        # 2. SAMPLING STOK HARIAN (tanpa pengembalian)
        sayur_harian = sayur.sample(3, replace=False).reset_index(drop=True)
        buah_harian = buah.sample(3, replace=False).reset_index(drop=True)
        
        # 3. DISTRIBUSI SAYUR KE WAKTU MAKAN
        sayur_pagi = sayur_harian.iloc[0]
        sayur_siang = sayur_harian.iloc[1]
        sayur_malam = sayur_harian.iloc[2]
        
        # 4. DISTRIBUSI BUAH KE WAKTU MAKAN
        buah_pagi = buah_harian.iloc[0]
        buah_siang = buah_harian.iloc[1]
        buah_malam = buah_harian.iloc[2]
        
        # 5. PENENTUAN SLOT SUSU
        slot_susu = pilih_slot_susu()
        susu_pagi = ambil_susu(susu) if slot_susu["sarapan"] else None
        susu_siang = ambil_susu(susu) if slot_susu["siang"] else None
        susu_malam = ambil_susu(susu) if slot_susu["malam"] else None
        
        # 7. PERHITUNGAN NUTRISI PER WAKTU MAKAN (TANPA SUSU DULU)
        nutrisi_sarapan = hitung_total_nutrisi([lauk_pagi, sayur_pagi, buah_pagi])
        nutrisi_siang = hitung_total_nutrisi([lauk_siang, sayur_siang, buah_siang])
        nutrisi_malam = hitung_total_nutrisi([lauk_malam, sayur_malam, buah_malam])
        
        # Tambahkan susu ke nutrisi waktu makan
        if susu_pagi is not None:
            for col in NUTRISI_COLS:
                nutrisi_sarapan[col] += susu_pagi[col]
        if susu_siang is not None:
            for col in NUTRISI_COLS:
                nutrisi_siang[col] += susu_siang[col]
        if susu_malam is not None:
            for col in NUTRISI_COLS:
                nutrisi_malam[col] += susu_malam[col]
        
        # 8. AGREGASI KOMPONEN UNTUK TOTAL HARIAN
        komponen = [
            lauk_pagi, sayur_pagi, buah_pagi,
            lauk_siang, sayur_siang, buah_siang,
            lauk_malam, sayur_malam, buah_malam
        ]
        
        if susu_pagi is not None:
            komponen.append(susu_pagi)
        if susu_siang is not None:
            komponen.append(susu_siang)
        if susu_malam is not None:
            komponen.append(susu_malam)
        
        total_nutrisi = hitung_total_nutrisi(komponen)
        
        # 9. PEMBENTUKAN ARM
        arm = {
            "arm_id": arm_id,
            
            # Menu
            "sarapan_lauk": lauk_pagi["nama_menu"],
            "sarapan_sayur": sayur_pagi["nama_menu"],
            "sarapan_buah": buah_pagi["nama_buah"],
            "sarapan_susu": susu_pagi["nama"] if susu_pagi is not None else None,
            
            "siang_lauk": lauk_siang["nama_menu"],
            "siang_sayur": sayur_siang["nama_menu"],
            "siang_buah": buah_siang["nama_buah"],
            "siang_susu": susu_siang["nama"] if susu_siang is not None else None,
            
            "malam_lauk": lauk_malam["nama_menu"],
            "malam_sayur": sayur_malam["nama_menu"],
            "malam_buah": buah_malam["nama_buah"],
            "malam_susu": susu_malam["nama"] if susu_malam is not None else None,
        }
        
        # TAMBAHKAN NUTRISI PER WAKTU MAKAN 
        # Otomatis semua nutrisi di NUTRISI_COLS masuk!
        for nutrisi_key in NUTRISI_COLS:
            arm[f"sarapan_{nutrisi_key}"] = nutrisi_sarapan[nutrisi_key]
            arm[f"siang_{nutrisi_key}"] = nutrisi_siang[nutrisi_key]
            arm[f"malam_{nutrisi_key}"] = nutrisi_malam[nutrisi_key]
        
        # Tambahkan total nutrisi harian
        arm.update({f"total_{k}": v for k, v in total_nutrisi.items()})
        
        arms.append(arm)
    
    df_arms = pd.DataFrame(arms)
    return df_arms

# ============================================================
# VALIDASI KONSEP STOK HARIAN
# ============================================================

def validasi_stok_harian(df_arms):
    """
    Cek:
    1. Tidak boleh ada sayur yang sama 3x dalam 1 arm
    2. Tidak boleh ada buah yang sama 3x dalam 1 arm
    """
    print("\n" + "="*60)
    print("VALIDASI KONSEP STOK HARIAN")
    print("="*60)
    
    # Cek sayur sama 3x
    same_sayur = df_arms[
        (df_arms['sarapan_sayur'] == df_arms['siang_sayur']) & 
        (df_arms['siang_sayur'] == df_arms['malam_sayur'])
    ]
    
    # Cek buah sama 3x
    same_buah = df_arms[
        (df_arms['sarapan_buah'] == df_arms['siang_buah']) & 
        (df_arms['siang_buah'] == df_arms['malam_buah'])
    ]
    
    print(f"\nHasil Validasi:")
    print(f"   Total arms: {len(df_arms):,}")
    
    if len(same_sayur) == 0:
        print(f"   Sayur sama 3x dalam 1 arm: {len(same_sayur)} (PASS)")
    else:
        print(f"   Sayur sama 3x dalam 1 arm: {len(same_sayur)} (FAIL)")
    
    if len(same_buah) == 0:
        print(f"   Buah sama 3x dalam 1 arm: {len(same_buah)} (PASS)")
    else:
        print(f"   Buah sama 3x dalam 1 arm: {len(same_buah)} (FAIL)")
    
    return {
        "total_arms": len(df_arms),
        "invalid_sayur": len(same_sayur),
        "invalid_buah": len(same_buah),
        "is_valid": (len(same_sayur) == 0 and len(same_buah) == 0)
    }

# ============================================================
# STATISTIK ARM
# ============================================================

def tampilkan_statistik(df_arms):
 
    print("\n" + "="*60)
    print("STATISTIK ARM")
    print("="*60)
    
    print(f"\n Jumlah & Variasi:")
    print(f"   Total arms: {len(df_arms):,}")
    print(f"   Lauk sarapan unik: {df_arms['sarapan_lauk'].nunique()}")
    print(f"   Lauk siang unik: {df_arms['siang_lauk'].nunique()}")
    print(f"   Lauk malam unik: {df_arms['malam_lauk'].nunique()}")
    print(f"   Sayur unik: {df_arms['sarapan_sayur'].nunique()}")
    print(f"   Buah unik: {df_arms['sarapan_buah'].nunique()}")


# ============================================================
# SAVE ARMS
# ============================================================
def save_arms(df_arms, output_file=OUTPUT_FILE):

    print("\n" + "="*60)
    print("MENYIMPAN HASIL")
    print("="*60)

    print(f"\n Menyimpan {len(df_arms):,} arms ke file...")
    print(f"   Filename: {output_file}")
    print(f"   Estimasi ukuran: ~{len(df_arms) * 0.2 / 1000:.0f} MB")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_arms.to_excel(writer, sheet_name="arms", index=False)

    print(f"\nFile berhasil disimpan!")

# ============================================================
# MAIN PIPELINE
# ============================================================
def main():
    """
    Pipeline utama untuk pembentukan arm
    """
    print("\n" + "="*60)
    print("SISTEM PEMBENTUKAN ARM")
    print("Rekomendasi Diet DASH dengan CMAB")
    print("="*60)
    # 1. Load data
    data = load_data()

    # 2. Buat arms
    df_arms = create_arms_stok_harian(data, n_arms=10000)

    if df_arms is None:
        print("\nProses dibatalkan.")
        return

    # 3. Validasi
    hasil_validasi = validasi_stok_harian(df_arms)

    if not hasil_validasi["is_valid"]:
        print("\n⚠️ WARNING: Validasi gagal! Ada arm yang tidak sesuai konsep stok harian.")
        print("   Silakan periksa implementasi fungsi create_arms_stok_harian()")
        return

    # 4. Statistik
    tampilkan_statistik(df_arms)

    # 5. Simpan
    save_arms(df_arms)

    # 6. Kesimpulan
    print("\n" + "="*60)
    print("RINGKASAN")
    print("="*60)
    print(f"\n✓ Total arms berhasil dibangkitkan: {len(df_arms):,}")
    print(f"✓ Konsep stok harian: VALID")
    print(f"✓ File output: {OUTPUT_FILE}")
    print(f"\nCATATAN:")
    print(f"   • Nutrisi sudah termasuk SUSU (2-3 porsi/hari)")
    print(f"   • Nutrisi BELUM termasuk NASI")
    print(f"   • NASI akan ditambahkan pada tahap berikutnya")
    print(f"     (Bab 3.7: Pembagian Kebutuhan Nutrisi Harian)")
    print("\n" + "="*60)
    
if __name__ == "__main__":
    main()
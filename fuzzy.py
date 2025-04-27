import openpyxl

# --- Membaca Data dari restoran.xlsx ---
def baca_data(filename):
    """
    Membaca data restoran dari file Excel dan menyimpannya dalam list.
    
    Parameter:
    filename (str): Nama file Excel yang berisi data restoran.
    
    Returns:
    list: Daftar data restoran dengan atribut ID, Kualitas Servis, dan Harga.
    """
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header
        id_restoran, kualitas_servis, harga = row
        data.append({
            'id': id_restoran,
            'kualitas_servis': kualitas_servis,
            'harga': harga
        })
    return data

# --- Fuzzification ---
def fuzzifikasi_servis(kualitas):
    """
    Mengubah kualitas servis restoran (1-100) menjadi nilai fuzzy.
    
    Parameter:
    kualitas (int): Nilai kualitas servis restoran.
    
    Returns:
    dict: Nilai keanggotaan fuzzy untuk kategori Buruk, Sedang, dan Bagus.
    """
    if kualitas <= 40:
        return {'Buruk': 1, 'Sedang': 0, 'Bagus': 0}
    elif 40 < kualitas <= 70:
        return {
            'Buruk': (70 - kualitas) / 30,
            'Sedang': (kualitas - 40) / 30,
            'Bagus': 0
        }
    elif 70 < kualitas <= 90:
        return {
            'Buruk': 0,
            'Sedang': (90 - kualitas) / 20,
            'Bagus': (kualitas - 70) / 20
        }
    else:  # kualitas > 90
        return {'Buruk': 0, 'Sedang': 0, 'Bagus': 1}

def fuzzifikasi_harga(harga):
    """
    Mengubah harga restoran (25.000-55.000) menjadi nilai fuzzy untuk kategori Murah, Sedang, dan Mahal.
    
    Parameter:
    harga (float): Nilai harga restoran.
    
    Returns:
    dict: Nilai keanggotaan fuzzy untuk kategori Murah, Sedang, dan Mahal.
    """
    if harga <= 30000:
        return {'Murah': 1.0, 'Sedang': 0.0, 'Mahal': 0.0}
    elif 30000 < harga <= 40000:
        return {
            'Murah': (40000 - harga) / 10000,
            'Sedang': (harga - 30000) / 10000,
            'Mahal': 0.0
        }
    elif 40000 < harga <= 50000:
        return {
            'Murah': 0.0,
            'Sedang': (50000 - harga) / 10000,
            'Mahal': (harga - 40000) / 10000
        }
    else:  # harga > 50000
        return {'Murah': 0.0, 'Sedang': 0.0, 'Mahal': 1.0}

# --- Inferensi ---
def inferensi(servis_fuzzy, harga_fuzzy):
    """
    Menentukan nilai kelayakan berdasarkan aturan inferensi yang melibatkan kualitas servis dan harga restoran.
    
    Parameter:
    servis_fuzzy (dict): Hasil fuzzifikasi kualitas servis restoran.
    harga_fuzzy (dict): Hasil fuzzifikasi harga restoran.
    
    Returns:
    list: Nilai kelayakan untuk setiap aturan yang diterapkan.
    """
    aturan = []
    aturan.append(min(servis_fuzzy['Bagus'], harga_fuzzy['Murah']))  # Servis Bagus, Harga Murah => Layak Tinggi
    aturan.append(min(servis_fuzzy['Sedang'], harga_fuzzy['Murah']))  # Servis Sedang, Harga Murah => Layak Sedang
    aturan.append(min(servis_fuzzy['Bagus'], harga_fuzzy['Mahal']))   # Servis Bagus, Harga Mahal => Layak Sedang
    aturan.append(min(servis_fuzzy['Buruk'], harga_fuzzy['Murah']))   # Servis Buruk, Harga Murah => Layak Rendah
    aturan.append(min(servis_fuzzy['Buruk'], harga_fuzzy['Mahal']))   # Servis Buruk, Harga Mahal => Layak Sangat Rendah
    return aturan

# --- Defuzzification ---
def defuzzifikasi(aturan):
    """
    Mengubah hasil inferensi fuzzy menjadi skor numerik menggunakan metode Weighted Average (rata-rata terbobot).
    
    Parameter:
    aturan (list): Nilai kekuatan untuk setiap aturan yang diterapkan.
    
    Returns:
    float: Skor kelayakan restoran (0-100).
    """
    bobot_kelayakan = [90, 80, 60, 80, 60, 40, 60, 60, 40, 20]  # Bobot aturan kelayakan

    # Hitung rata-rata terbobot
    pembilang = sum(a * b for a, b in zip(aturan, bobot_kelayakan))
    penyebut = sum(aturan)

    # Hindari pembagian dengan nol
    if penyebut == 0:
        return 0.0
    return pembilang / penyebut

# --- Menyimpan Output ke Peringkat.xlsx ---
def simpan_output(filename, hasil):
    """
    Menyimpan hasil peringkat restoran ke dalam file Excel.
    
    Parameter:
    filename (str): Nama file tempat hasil disimpan.
    hasil (list): Daftar hasil peringkat restoran dengan skor kelayakan.
    """
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(['ID', 'Kualitas Servis', 'Harga', 'Skor Kelayakan'])

    for item in hasil:
        sheet.append([item['id'], item['kualitas_servis'], item['harga'], item['skor']])

    wb.save(filename)

# --- Program Utama ---
if __name__ == "__main__":
    # Membaca data restoran dari file restoran.xlsx
    data = baca_data('restoran.xlsx')

    hasil = []
    for d in data:
        # Melakukan fuzzifikasi untuk kualitas servis dan harga
        servis_fuzzy = fuzzifikasi_servis(d['kualitas_servis'])
        harga_fuzzy = fuzzifikasi_harga(d['harga'])

        # Inferensi berdasarkan aturan
        aturan = inferensi(servis_fuzzy, harga_fuzzy)

        # Defuzzifikasi untuk menentukan skor kelayakan
        skor = defuzzifikasi(aturan)

        hasil.append({
            'id': d['id'],
            'kualitas_servis': d['kualitas_servis'],
            'harga': d['harga'],
            'skor': skor,
            'aturan': aturan
        })

    # Sortir hasil berdasarkan skor kelayakan tertinggi
    hasil.sort(key=lambda x: x['skor'], reverse=True)

    # Ambil 3 restoran terbaik (top 3) untuk tracing aturan inferensi
    top_5 = hasil[:5]

    # Simpan hasil ke dalam file peringkat.xlsx
    simpan_output('peringkat.xlsx', hasil[:10])  # Simpan 10 restoran terbaik

    # Tampilkan hasil dalam format tabel dengan tracing untuk top 3
    print(f"{'No.':<4} {'ID':<8} {'Kualitas Servis':<18} {'Harga':<15} {'Skor Kelayakan':<20}")
    print("="*70)
    
    for i, item in enumerate(top_5, 1):
        print(f"{i:<4} {item['id']:<8} {item['kualitas_servis']:<18} {item['harga']:<15} {item['skor']:<20.2f}")
        
        # Menampilkan tracing proses inferensi untuk setiap restoran
        print(f"\nDetail Proses Fuzzy Logic untuk Restoran ID {item['id']}:")
        print("="*70)
        print(f"ID Restoran: {item['id']}")
        print(f"Kualitas Servis: {item['kualitas_servis']} -> {servis_fuzzy}")
        print(f"Harga: Rp {item['harga']} -> {harga_fuzzy}")
        print("\nHasil Inferensi (5 aturan):")
        for idx, val in enumerate(item['aturan'], 1):
            print(f"Aturan {idx}: {val:.2f}")
        
        # Skor akhir dari defuzzifikasi
        print(f"\nSkor Akhir (Defuzzifikasi): {item['skor']:.2f}")
        print("="*70)

    # Tampilkan tabel 10 restoran terbaik
    print("\nDaftar 10 Restoran Terbaik berdasarkan Fuzzy Logic:")
    print(f"{'No.':<4} {'ID':<8} {'Kualitas Servis':<18} {'Harga':<15} {'Skor Kelayakan':<20}")
    print("="*70)
    
    for i, item in enumerate(hasil[:10], 1):
        print(f"{i:<4} {item['id']:<8} {item['kualitas_servis']:<18} {item['harga']:<15} {item['skor']:<20.2f}")

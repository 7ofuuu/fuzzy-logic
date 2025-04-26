import openpyxl

# --- Membaca Data dari restoran.xlsx ---
def baca_data(filename):
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
    if harga <= 30000:
        return {'Murah': 1, 'Sedang': 0, 'Mahal': 0}
    elif 30000 < harga <= 40000:
        return {
            'Murah': (40000 - harga) / 10000,
            'Sedang': (harga - 30000) / 10000,
            'Mahal': 0
        }
    elif 40000 < harga <= 50000:
        return {
            'Murah': 0,
            'Sedang': (50000 - harga) / 10000,
            'Mahal': (harga - 40000) / 10000
        }
    else:  # harga > 50000
        return {'Murah': 0, 'Sedang': 0, 'Mahal': 1}

# --- Inferensi ---
def inferensi(servis_fuzzy, harga_fuzzy):
    aturan = []

    # Definisi aturan (contoh):
    # Jika Servis Bagus dan Harga Murah, maka Layak Tinggi
    aturan.append(min(servis_fuzzy['Bagus'], harga_fuzzy['Murah']))

    # Jika Servis Sedang dan Harga Murah, maka Layak Sedang
    aturan.append(min(servis_fuzzy['Sedang'], harga_fuzzy['Murah']))

    # Jika Servis Bagus dan Harga Mahal, maka Layak Sedang
    aturan.append(min(servis_fuzzy['Bagus'], harga_fuzzy['Mahal']))

    # Jika Servis Buruk dan Harga Murah, maka Layak Rendah
    aturan.append(min(servis_fuzzy['Buruk'], harga_fuzzy['Murah']))

    # Jika Servis Buruk dan Harga Mahal, maka Layak Sangat Rendah
    aturan.append(min(servis_fuzzy['Buruk'], harga_fuzzy['Mahal']))

    return aturan

# --- Defuzzification ---
def defuzzifikasi(aturan):
    # Misal:
    # Aturan 0: Tinggi (80)
    # Aturan 1: Sedang (60)
    # Aturan 2: Sedang (60)
    # Aturan 3: Rendah (40)
    # Aturan 4: Sangat Rendah (20)
    bobot = [80, 60, 60, 40, 20]

    atas = sum(a * b for a, b in zip(aturan, bobot))
    bawah = sum(aturan)
    if bawah == 0:
        return 0
    return atas / bawah

# --- Menyimpan output ke peringkat.xlsx ---
def simpan_output(filename, hasil):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(['ID', 'Kualitas Servis', 'Harga', 'Skor Kelayakan'])

    for item in hasil:
        sheet.append([item['id'], item['kualitas_servis'], item['harga'], item['skor']])

    wb.save(filename)

# --- Program Utama ---
if __name__ == "__main__":
    data = baca_data('restoran.xlsx')

    hasil = []
    for d in data:
        servis_fuzzy = fuzzifikasi_servis(d['kualitas_servis'])
        harga_fuzzy = fuzzifikasi_harga(d['harga'])
        aturan = inferensi(servis_fuzzy, harga_fuzzy)
        skor = defuzzifikasi(aturan)

        hasil.append({
            'id': d['id'],
            'kualitas_servis': d['kualitas_servis'],
            'harga': d['harga'],
            'skor': skor
        })

    # Sortir berdasarkan skor tertinggi
    hasil.sort(key=lambda x: x['skor'], reverse=True)

    # Ambil 5 restoran terbaik
    top_5 = hasil[:5]

    simpan_output('peringkat.xlsx', top_5)

    # Tampilkan hasil di terminal
    for i, item in enumerate(top_5, 1):
        print(f"{i}. ID: {item['id']}, Servis: {item['kualitas_servis']}, Harga: {item['harga']}, Skor: {item['skor']:.2f}")

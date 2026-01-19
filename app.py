from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def load_konversi(df_konv):
    """
    Membaca tabel konversi:
    kolom 0 = batas bawah
    kolom 1 = batas atas
    kolom 2 = nilai konversi
    """
    rules = []
    for _, row in df_konv.iterrows():
        try:
            min_h = float(row.iloc[0])
            max_h = float(row.iloc[1])
            konv  = row.iloc[2]
            rules.append((min_h, max_h, konv))
        except:
            continue
    return rules

def cari_konversi(harga, rules):
    for min_h, max_h, konv in rules:
        if min_h <= harga <= max_h:
            return konv
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        kecamatan = request.form['kecamatan']
        petugas = request.form['petugas']

        file = request.files['file']
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Baca Excel
        xls = pd.ExcelFile(file_path)
        df_rh = xls.parse('RH', header=None)
        df_konv = xls.parse('Konversi', header=None)

        # Ambil aturan konversi
        rules = load_konversi(df_konv)

        hasil = []

        # Asumsi harga ada di kolom tertentu (bisa disesuaikan)
        # Di sini kita ambil semua angka di sheet RH
        for row_idx in range(len(df_rh)):
            for col_idx in range(len(df_rh.columns)):
                val = df_rh.iloc[row_idx, col_idx]
                if isinstance(val, (int, float)):
                    konv = cari_konversi(val, rules)
                    if konv is not None:
                        hasil.append({
                            "Kecamatan": kecamatan,
                            "Petugas": petugas,
                            "Harga": val,
                            "Konversi": konv,
                            "Baris": row_idx + 1,
                            "Kolom": col_idx + 1
                        })

        output_df = pd.DataFrame(hasil)
        output_df["Tanggal"] = datetime.now().strftime("%Y-%m-%d")

        output_name = f"RH_{kecamatan}_{petugas}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_name)
        output_df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

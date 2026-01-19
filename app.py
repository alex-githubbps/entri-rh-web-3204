
from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        kecamatan = request.form['kecamatan']
        petugas = request.form['petugas']

        file = request.files['file']
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        rh = pd.read_excel(filepath, sheet_name='RH')
        konversi = pd.read_excel(filepath, sheet_name='Konversi')

        def cari_konversi(harga):
            for _, row in konversi.iterrows():
                if row.iloc[0] <= harga <= row.iloc[1]:
                    return row.iloc[2]
            return None

        rh['Nilai Konversi'] = rh.iloc[:, 0].apply(cari_konversi)
        rh['Kecamatan'] = kecamatan
        rh['Petugas'] = petugas
        rh['Tanggal'] = datetime.now().strftime('%Y-%m-%d')

        output_name = f"RH_{kecamatan}_{petugas}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_name)
        rh.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

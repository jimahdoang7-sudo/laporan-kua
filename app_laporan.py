import sys
import os
import pandas as pd
from flask import Flask, render_template, request, send_file
import io

# Atur Path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

sys.path.append(BASE_DIR)
sys.path.append(os.path.join(BASE_DIR, 'modules'))

import petugas, kategori, pnbp, wna, wali_hakim

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

MONTH_MAP = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
    7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

def cari_kolom(df, keywords):
    for k in keywords:
        for c in df.columns:
            if k.upper() == str(c).upper().strip(): return c
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    active_tab = request.args.get('tab', 'OVERVIEW')
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data_terakhir.xlsx')
    
    if request.method == 'POST':
        file = request.files.get('file')
        if file:
            file.save(file_path)
    
    if os.path.exists(file_path):
        df_raw = pd.read_excel(file_path, dtype=str)
        df = df_raw.apply(lambda x: x.str.strip().str.upper() if x.dtype == "object" else x).fillna('')
        
        # --- LOGIKA DETEKSI PERIODE ---
        col_tgl = cari_kolom(df, ["TANGGAL AKAD", "TGL AKAD"])
        nama_bulan = "JANUARI"; year_val = "2026"
        try:
            dt_temp = pd.to_datetime(df[col_tgl], dayfirst=True, errors='coerce').dropna()
            nama_bulan = MONTH_MAP.get(dt_temp.dt.month.mode()[0], "JANUARI")
            year_val = int(dt_temp.dt.year.mode()[0])
        except: pass

        col_lokasi = cari_kolom(df, ["NIKAH DI", "LOKASI"])
        col_daftar = cari_kolom(df, ["NO PENDAFTARAN", "PENDAFTARAN"])
        
        
        # --- HITUNG DATA WNA UNTUK KARTU KE-5 (BEKAL SEMUA TAB) ---
        df_wna_tmp, _ = wna.render(df, nama_bulan, year_val)
        total_wna = len(df_wna_tmp) if df_wna_tmp is not None else 0

        # --- HITUNG DATA WALI HAKIM (Untuk Kartu ke-6) ---
        df_wh_tmp, _ = wali_hakim.render(df, nama_bulan, year_val)
        total_wh = len(df_wh_tmp) if df_wh_tmp is not None else 0
        
        # Rekap Dasar (Sekarang ada 6 entri)
        rekap = {
            'total': len(df),
            'kantor': len(df[(df[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df[col_lokasi].str.contains("LUAR", na=False))]) if col_lokasi else 0,
            'luar': len(df[df[col_lokasi].str.contains("LUAR", na=False)]) if col_lokasi else 0,
            'isbat': len(df[df[col_daftar].str.contains("^IB", na=False, case=False)]) if col_daftar else 0,
            'wna': total_wna,
            'wh': total_wh, # Tambahkan ini
            'bulan': nama_bulan, 'tahun': year_val
        }

        chart_data = None
        list_petugas = []
        selected_petugas = 'SEMUA'
        selected_lokasi = 'SEMUA' # Inisialisasi parameter baru

        if active_tab == 'PETUGAS':
            selected_petugas = request.args.get('nama', 'SEMUA')
            selected_lokasi = request.args.get('lokasi', 'SEMUA')
            
            # 1. Panggil render untuk mendapatkan data yang sudah difilter petugas & lokasi
            summary, chart_data, final_df = petugas.render(df, nama_bulan, year_val, selected_petugas, selected_lokasi)
            list_petugas = chart_data['labels']
            
            # 2. LOGIKA PERBAIKAN: Hitung rekap berdasarkan final_df (bukan df global)
            if selected_petugas != 'SEMUA' or selected_lokasi != 'SEMUA':
                s = next((item for item in summary if item["Petugas"] == selected_petugas), None)
                
                # Hitung ulang berdasarkan hasil filter tabel saat ini
                c_kan = len(final_df[(final_df['Nikah Di'].str.contains("KANTOR|KUA", na=False)) & (~final_df['Nikah Di'].str.contains("LUAR", na=False))]) if 'Nikah Di' in final_df.columns else 0
                c_luar = len(final_df[final_df['Nikah Di'].str.contains("LUAR", na=False)]) if 'Nikah Di' in final_df.columns else 0
                
                rekap = {
                    'total': len(final_df),
                    'kantor': c_kan,
                    'luar': c_luar,
                    'isbat': s['Isbat'] if (s and selected_lokasi == 'SEMUA') else len(final_df[final_df['No Pendaftaran'].str.contains("^IB", na=False, case=False)]) if 'No Pendaftaran' in final_df.columns else 0,
                    'wna': total_wna,
                    'bulan': nama_bulan, 'tahun': year_val
                }
                
                if selected_petugas != 'SEMUA':
                    chart_data = {'labels': [selected_petugas], 'kantor': [c_kan], 'luar': [c_luar]}
            
            table_html = final_df.to_html(classes='table table-sm table-hover align-middle', index=False)

        # SEBELUMNYA MUNGKIN ADA KURUNG } ATAU SPASI BERLEBIH DI SINI YANG BIKIN ERROR
        elif active_tab == 'PNBP':
            final_df, err = pnbp.render(df, nama_bulan, year_val)
            rekap['wna'] = total_wna
            table_html = final_df.to_html(classes='table table-sm table-hover', index=False)
            
        elif active_tab == 'ISBAT':
            final_df, err = kategori.render(df, nama_bulan, year_val, pilihan="ISBAT")
            rekap['wna'] = total_wna
            table_html = final_df.to_html(classes='table table-sm table-hover', index=False)
            
        elif active_tab == 'WNA':
            final_df, err = wna.render(df, nama_bulan, year_val)
            rekap['wna'] = total_wna
            rekap['judul_kustom'] = "NIKAH CAMPURAN (WNA)"
            if final_df is not None and not final_df.empty:
                if 'No' not in final_df.columns:
                    final_df.insert(0, 'No', range(1, len(final_df) + 1))
                table_html = final_df.to_html(classes='table table-sm table-hover align-middle', index=False)
            else:
                table_html = f"<div class='alert alert-warning'>{err if err else 'Data WNA tidak ditemukan'}</div>"
        elif active_tab == 'WALI_HAKIM':
            final_df, err = wali_hakim.render(df, nama_bulan, year_val)
            rekap['wna'] = total_wna
            rekap['judul_kustom'] = "DAFTAR PERISTIWA WALI HAKIM"
            if final_df is not None and not final_df.empty:
                if 'No' not in final_df.columns:
                    final_df.insert(0, 'No', range(1, len(final_df) + 1))
                table_html = final_df.to_html(classes='table table-sm table-hover align-middle', index=False)
            else:
                table_html = f"<div class='alert alert-warning'>{err if err else 'Data Wali Hakim tidak ditemukan'}</div>"
        
        else: # OVERVIEW (Pastikan bagian ini ada isinya agar tidak error)
            df_fix = df.copy()
            df_fix['No_Urut'] = range(1, len(df_fix) + 1)
            c_seri = cari_kolom(df, ["NO SERI HURUF", "SERI HURUF"])
            c_perfo = cari_kolom(df, ["NO PERFORASI", "PERFORASI"])
            df_fix['P_Perforasi'] = (df_fix[c_seri].astype(str) + ". " + df_fix[c_perfo].astype(str)) if (c_seri and c_perfo) else (df_fix[c_perfo] if c_perfo else "")
            df_fix['Blank'] = ""
            mapping_all = {
                'No': 'No_Urut', 'No Perforasi': 'P_Perforasi', ' ': 'Blank',
                'No Pemeriksaan': ["NO PEMERIKSAAN", "PEMERIKSAAN"],
                'No Aktanikah': ["NO AKTANIKAH", "AKTA NIKAH"],
                'Nama Suami': ["NAMA SUAMI"], 'Nama Istri': ["NAMA ISTRI"],
                'Tanggal Akad': ["TANGGAL AKAD"], 'Nikah Di': ["NIKAH DI", "LOKASI"]
            }
            rekap_df = pd.DataFrame()
            for label, keys in mapping_all.items():
                if isinstance(keys, list):
                    col_found = cari_kolom(df, keys)
                    rekap_df[label] = df_fix[col_found] if col_found else ""
                else: rekap_df[label] = df_fix[keys]
            table_html = rekap_df.to_html(classes='table table-sm table-hover align-middle', index=False)

        # Baris ini harus sejajar dengan blok 'if os.path.exists(file_path):'
        return render_template('index.html', rekap=rekap, table=table_html, 
                               active_tab=active_tab, chart_data=chart_data,
                               list_petugas=list_petugas, selected_petugas=selected_petugas,
                               selected_lokasi=selected_lokasi)

    return render_template('index.html', rekap=None)

# --- FUNGSI DOWNLOAD ---
@app.route('/download_laporan')
def download_laporan():
    tab = request.args.get('tab', 'OVERVIEW')
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data_terakhir.xlsx')
    if not os.path.exists(file_path): return "File tidak ditemukan", 404
    
    df_raw = pd.read_excel(file_path, dtype=str)
    df = df_raw.apply(lambda x: x.str.strip().str.upper() if x.dtype == "object" else x).fillna('')
    
    # Logika penentuan bulan & tahun tetap sama
    col_tgl = cari_kolom(df, ["TANGGAL AKAD", "TGL AKAD"])
    try:
        dt_temp = pd.to_datetime(df[col_tgl], dayfirst=True, errors='coerce').dropna()
        nama_bulan = MONTH_MAP.get(dt_temp.dt.month.mode()[0], "JANUARI")
        year_val = int(dt_temp.dt.year.mode()[0])
    except: nama_bulan = "JANUARI"; year_val = "2026"

    if tab == 'WNA':
        final_df, _ = wna.render(df, nama_bulan, year_val)
        title_excel = "LAPORAN PERISTIWA NIKAH CAMPURAN (WNA)"
    elif tab == 'ISBAT':
        final_df, _ = kategori.render(df, nama_bulan, year_val, pilihan="ISBAT")
        title_excel = "LAPORAN PERISTIWA NIKAH ISBAT"
    elif tab == 'WALI_HAKIM':
        final_df, _ = wali_hakim.render(df, nama_bulan, year_val)
        title_excel = "LAPORAN PERISTIWA WALI HAKIM"
    elif tab == 'PNBP':
        final_df, _ = pnbp.render(df, nama_bulan, year_val)
        title_excel = "LAPORAN PNBP NR"
    else: return "Tab tidak didukung", 400

    if final_df is not None and not final_df.empty:
        if 'No' not in final_df.columns: final_df.insert(0, 'No', range(1, len(final_df) + 1))
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
            from modules.petugas import format_excel
            format_excel(writer, final_df, title_excel, nama_bulan, year_val)
        output.seek(0)
        return send_file(output, download_name=f"LAPORAN_{tab}.xlsx", as_attachment=True)
    return "Data kosong", 404

if __name__ == '__main__':
    app.run(debug=True)

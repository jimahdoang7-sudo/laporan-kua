import pandas as pd

def get_col(df, keywords):
    """Mencari nama kolom berdasarkan kata kunci (Fungsi Dasar Lo)"""
    for key in keywords:
        for c in df.columns:
            if key.upper() in str(c).upper():
                return c
    return None

def render(df, bulan, tahun, pilih_petugas="SEMUA"):
    """Fungsi Render dengan jaminan kolom 'No' ada di urutan pertama"""
    col_petugas = get_col(df, ["NAMA PENGHULU HADIR", "NAMA PENGHULU", "PENGHULU HADIR"])
    col_lokasi = get_col(df, ["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
    col_daftar = get_col(df, ["NO PENDAFTARAN", "NOMOR PENDAFTARAN", "PENDAFTARAN"])

    if not col_petugas:
        return [], {"labels": [], "kantor": [], "luar": []}, pd.DataFrame()

    # 1. FILTER DATA PETUGAS
    df_f = df.copy()
    if pilih_petugas != "SEMUA":
        df_f = df_f[df_f[col_petugas] == pilih_petugas].copy()

    # 2. DATA UNTUK GRAFIK & KARTU
    all_p = sorted(df[col_petugas].unique())
    summary_list = []
    labels, data_k, data_lk = [], [], []

    for p_name in all_p:
        df_p_all = df[df[col_petugas] == p_name]
        c_kan = len(df_p_all[(df_p_all[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df_p_all[col_lokasi].str.contains("LUAR", na=False))])
        c_luar = len(df_p_all[df_p_all[col_lokasi].str.contains("LUAR", na=False)])
        c_isbat = len(df_p_all[df_p_all[col_daftar].str.contains("^IB", na=False, case=False)]) if col_daftar else 0
        summary_list.append({"Petugas": p_name, "Kantor": c_kan, "Luar": c_luar, "Isbat": c_isbat, "Total": len(df_p_all)})
        labels.append(p_name)
        data_k.append(c_kan)
        data_lk.append(c_luar)
    
    chart_data = {"labels": labels, "kantor": data_k, "luar": data_lk}

    # 3. PROSES MAPPING 13 KOLOM (WAJIB UTUH)
    c_seri = get_col(df, ["NO SERI HURUF", "SERI HURUF"])
    c_perfo = get_col(df, ["NO PERFORASI", "PERFORASI"])
    df_f['P_Perfo'] = (df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)) if (c_seri and c_perfo) else (df_f[c_perfo] if c_perfo else "")

    # Mapping Struktur Utama
    mapping = {
        'No Perforasi': 'P_Perfo',
        'No Pemeriksaan': get_col(df, ["PEMERIKSAAN"]),
        'No Aktanikah': get_col(df, ["AKTANIKAH", "AKTA NIKAH"]),
        'No Pendaftaran': col_daftar,
        'Nama Suami': get_col(df, ["NAMA SUAMI"]),
        'Nama Istri': get_col(df, ["NAMA ISTRI"]),
        'Tanggal': get_col(df, ["TANGGAL AKAD", "TGL AKAD"]),
        'Jam': get_col(df, ["JAM AKAD", "WAKTU AKAD"]),
        'Kelurahan': get_col(df, ["KELURAHAN", "DESA"]),
        'Penghulu': col_petugas,
        'Status Wali': get_col(df, ["STATUS WALI"]),
        'Nikah Di': col_lokasi
    }

    # Buat dataframe berdasarkan mapping (Tanpa No dulu)
    available_cols = [v for v in mapping.values() if v and v in df_f.columns]
    final_df = df_f[available_cols].copy()
    rename_map = {v: k for k, v in mapping.items() if v in final_df.columns}
    final_df.rename(columns=rename_map, inplace=True)

    # --- KUNCI: MASUKKAN KOLOM 'No' SECARA PAKSA DI POSISI PERTAMA ---
    final_df.insert(0, 'No', range(1, len(final_df) + 1))

    return summary_list, chart_data, final_df

def format_excel(writer, final_df, title, month, year):
    """Format Excel: Bukti Judul Tengah & Auto-Width Isi"""
    workbook = writer.book
    worksheet = writer.sheets['Laporan']
    
    fmt_judul = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#9C27B0', 'font_color': 'white'})
    fmt_data = workbook.add_format({'border': 1, 'valign': 'middle', 'align': 'left', 'font_size': 10})
    fmt_center = workbook.add_format({'border': 1, 'valign': 'middle', 'align': 'center', 'font_size': 10})

    # Judul rata tengah berdasarkan jumlah kolom (A sampai M)
    last_col_idx = len(final_df.columns) - 1
    worksheet.merge_range(0, 0, 0, last_col_idx, f'LAPORAN {title.upper()}', fmt_judul)
    worksheet.merge_range(1, 0, 1, last_col_idx, f'BULAN {month} TAHUN {year}', fmt_judul)

    for col_num, col_name in enumerate(final_df.columns):
        # Lebar otomatis menghitung isi data terpanjang
        max_len = max(final_df[col_name].astype(str).map(len).max(), len(str(col_name))) + 3
        actual_width = 6 if col_name == 'No' else min(max(max_len, 12), 50)
        
        worksheet.set_column(col_num, col_num, actual_width)
        worksheet.write(4, col_num, col_name, fmt_header)

    # Penulisan data satu per satu
    for row_num in range(len(final_df)):
        for col_num, col_name in enumerate(final_df.columns):
            val = str(final_df.iloc[row_num, col_num])
            style = fmt_center if col_name in ['No', 'Tanggal', 'Jam'] else fmt_data
            worksheet.write(row_num + 5, col_num, val if val != 'NAN' else '', style)
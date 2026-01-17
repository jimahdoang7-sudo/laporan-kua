import pandas as pd

def get_col(df, keywords):
    for key in keywords:
        for c in df.columns:
            if key.upper() in str(c).upper():
                return c
    return None

def render(df, bulan, tahun, pilihan="SEMUA PERISTIWA"):
    # Deteksi Kolom
    col_lokasi = get_col(df, ["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
    col_daftar = get_col(df, ["NO PENDAFTARAN", "PENDAFTARAN"])
    
    if not col_lokasi:
        return None, "Kolom Lokasi tidak ditemukan!"

    # --- LOGIKA FILTER (Sesuai Permintaan Lo) ---
    if pilihan == "SEMUA PERISTIWA":
        df_f = df.copy()
    elif pilihan == "KUA / KANTOR":
        mask = (df[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df[col_lokasi].str.contains("LUAR", na=False))
        df_f = df[mask].copy()
    elif pilihan == "LUAR KUA / BEDOL":
        df_f = df[df[col_lokasi].str.contains("LUAR", na=False)].copy()
    elif pilihan == "ISBAT":
        # Patokan Isbat: Kode IB di No Pendaftaran
        if col_daftar:
            df_f = df[df[col_daftar].str.contains("^IB", na=False, case=False)].copy()
        else:
            df_f = pd.DataFrame()

    if not df_f.empty:
        df_f['No_Urut'] = range(1, len(df_f) + 1)
        
        # Gabung Seri & Perforasi
        c_seri = get_col(df, ["NO SERI HURUF", "SERI HURUF"])
        c_perfo = get_col(df, ["NO PERFORASI", "PERFORASI"])
        if c_seri and c_perfo:
            df_f['P_Perforasi'] = df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)
        else:
            df_f['P_Perforasi'] = df_f[c_perfo] if c_perfo else ""

        # --- MAPPING TETAP UTUH ---
        mapping = {
            'No': 'No_Urut',
            'No Perforasi': 'P_Perforasi',
            'No Pemeriksaan': get_col(df, ["NO PEMERIKSAAN", "PEMERIKSAAN"]),
            'No Aktanikah': get_col(df, ["AKTANIKAH", "AKTA NIKAH"]),
            'Nama Suami': get_col(df, ["NAMA SUAMI"]),
            'Nama Istri': get_col(df, ["NAMA ISTRI"]),
            'Tanggal': get_col(df, ["TANGGAL AKAD", "TGL AKAD"]),
            'Kelurahan': get_col(df, ["NAMA KELURAHAN", "KELURAHAN", "DESA"]),
            'Penghulu': get_col(df, ["NAMA PENGHULU HADIR", "NAMA PENGHULU"]),
            'Nikah Di': col_lokasi
        }
        
        available_source_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_source_cols].copy()
        rename_map = {v: k for k, v in mapping.items() if v in available_source_cols}
        final_df.rename(columns=rename_map, inplace=True)

        return final_df, None
    else:
        return pd.DataFrame(), f"Data tidak ditemukan untuk kategori {pilihan}"
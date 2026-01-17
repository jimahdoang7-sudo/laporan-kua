import pandas as pd

def get_col(df, keywords):
    for key in keywords:
        for c in df.columns:
            if key.upper() in str(c).upper().strip(): return c
    return None

def render(df, bulan, tahun):
    col_lokasi = get_col(df, ["NIKAH DI", "LOKASI"])
    # Filter otomatis: Hanya yang Luar KUA
    if col_lokasi:
        df_f = df[df[col_lokasi].str.contains("LUAR", na=False)].copy()
    else:
        df_f = df.copy()

    if not df_f.empty:
        df_f['No_Fix'] = range(1, len(df_f) + 1)
        c_pendaftaran = get_col(df, ["NO PENDAFTARAN", "NOMOR PENDAFTARAN"])
        c_ntpn = get_col(df, ["NO NTPN", "NTPN", "NOMOR NTPN"])
        c_jumlah = get_col(df, ["JUMLAH", "SETORAN", "BIAYA"])

        mapping = {
            'No': 'No_Fix',
            'No Pendaftaran': c_pendaftaran,
            'Nama Suami': get_col(df, ["NAMA SUAMI"]),
            'Nama Istri': get_col(df, ["NAMA ISTRI"]),
            'Tanggal Akad': get_col(df, ["TANGGAL AKAD"]),
            'No NTPN': c_ntpn if c_ntpn else None,
            'Penghulu': get_col(df, ["NAMA PENGHULU HADIR"])
        }

        available_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_cols].copy()
        
        # Tambahkan kolom jumlah manual jika tidak ada di excel
        if not c_jumlah:
            final_df['Jumlah Setor'] = "Rp. 600.000"
            
        return final_df, None
    return pd.DataFrame(), "Data PNBP tidak ditemukan."
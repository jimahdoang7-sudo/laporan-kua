import pandas as pd

def get_col(df, keywords):
    for key in keywords:
        for c in df.columns:
            if key.upper() == str(c).upper().strip(): return c
        for c in df.columns:
            if key.upper() in str(c).upper(): return c
    return None

def render(df, bulan, tahun):
    # 1. Cari Kolom Kewarganegaraan
    col_ws = get_col(df, ["WARGANEGARA SUAMI", "WN SUAMI", "WNA S"])
    col_wi = get_col(df, ["WARGANEGARA ISTRI", "WN ISTRI", "WNA I"])
    
    if not col_ws or not col_wi:
        return None, "Kolom Kewarganegaraan tidak ditemukan"

    # 2. FILTER WNA: Salah satu pengantin BUKAN 'WNI'
    df_f = df[(df[col_ws].astype(str).str.strip().str.upper() != 'WNI') | 
              (df[col_wi].astype(str).str.strip().str.upper() != 'WNI')].copy()
    
    if df_f.empty:
        return pd.DataFrame(), "Tidak ada data WNA pada bulan ini."

    # 3. MAPPING KOLOM STANDAR
    c_seri = get_col(df, ["NO SERI HURUF", "SERI HURUF"])
    c_perfo = get_col(df, ["NO PERFORASI", "PERFORASI"])
    df_f['P_Perfo'] = (df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)) if (c_seri and c_perfo) else ""

    mapping = {
        'No Perforasi': 'P_Perfo',
        'No Pendaftaran': get_col(df, ["NO PENDAFTARAN"]),
        'Nama Suami': get_col(df, ["NAMA SUAMI"]),
        'WN S': col_ws,
        'Nama Istri': get_col(df, ["NAMA ISTRI"]),
        'WN I': col_wi,
        'Tanggal': get_col(df, ["TANGGAL AKAD", "TGL AKAD"]),
        'Penghulu': get_col(df, ["NAMA PENGHULU HADIR", "PENGHULU"]),
        'Nikah Di': get_col(df, ["NIKAH DI", "LOKASI"])
    }

    cols = [v for v in mapping.values() if v and v in df_f.columns]
    final_df = df_f[cols].copy()
    final_df.rename(columns={v: k for k, v in mapping.items() if v in final_df.columns}, inplace=True)

    return final_df, None
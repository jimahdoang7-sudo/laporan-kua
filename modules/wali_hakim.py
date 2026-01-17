import pandas as pd

def get_col(df, keywords):
    for key in keywords:
        for c in df.columns:
            if key.upper() in str(c).upper(): return c
    return None

def render(df, bulan, tahun):
    col_status_wali = get_col(df, ["STATUS WALI", "STATUS_WALI"])
    if not col_status_wali:
        return pd.DataFrame(), "Kolom Status Wali tidak ditemukan."

    # Filter kata "HAKIM"
    df_f = df[df[col_status_wali].str.contains("HAKIM", na=False, case=False)].copy()
    
    if not df_f.empty:
        df_f['No_Wali'] = range(1, len(df_f) + 1)
        mapping = {
            'No': 'No_Wali',
            'Nama Suami': get_col(df, ["NAMA SUAMI"]),
            'Nama Istri': get_col(df, ["NAMA ISTRI"]),
            'Status Wali': col_status_wali,
            'Penghulu': get_col(df, ["NAMA PENGHULU HADIR"])
        }
        final_df = df_f[[v for v in mapping.values() if v in df_f.columns]].copy()
        return final_df, None
    return pd.DataFrame(), "Tidak ada data Wali Hakim."
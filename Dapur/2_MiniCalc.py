import pandas as pd
import openpyxl
import configparser
import os
import re

def extract_brand_tag(contact_name, keywords):
    if pd.isna(contact_name):
        return 'DEFAULT'
    c_str = str(contact_name).upper()
    for kw in keywords:
        kw_clean = kw.strip().upper()
        if kw_clean and re.search(r'\b' + re.escape(kw_clean) + r'\b', c_str):
            return kw_clean
    return 'DEFAULT'

def process_minifs(config_file='config.conf', 
                   input_file='Faktur_Penjualan_temp.xlsx', 
                   fallback_file='FallbackCash_temp.xlsx',
                   output_file='Minifs_temp.xlsx'):
    if not os.path.exists(config_file):
        print(f"--> Error: File '{config_file}' tidak ditemukan.")
        return
        
    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')
    
    if not config.has_section('MINIFS'):
        print("--> Error: Seksi [MINIFS] tidak ditemukan.")
        return
        
    mini_stat = config.get('MINIFS', 'mini_stat', fallback='No').strip()
    if mini_stat.lower() != 'ya':
        print(f"--> Informasi: mini_stat = '{mini_stat}'. Proses dilewati.")
        return

    mini_del = [k.strip() for k in config.get('MINIFS', 'mini_del', fallback='').split(',') if k.strip()]
    mini_filter = [k.strip() for k in config.get('MINIFS', 'mini_filter', fallback='').split(',') if k.strip()]
    mini_str_clean = [k.strip() for k in config.get('MINIFS', 'mini_str_clean', fallback='').split(',') if k.strip()]

    if not os.path.exists(input_file):
        print(f"--> Error: File input '{input_file}' tidak ditemukan.")
        return

    print(f"--> Membaca file '{input_file}'...")
    df_faktur = pd.read_excel(input_file)

    if os.path.exists(fallback_file):
        print(f"--> Menggabungkan file '{fallback_file}'...")
        df_fallback = pd.read_excel(fallback_file)
        
        fallback_col_map = {}
        for col in df_fallback.columns:
            col_str = str(col).strip().lower()
            if 'no.' in col_str and 'pelanggan' in col_str:
                fallback_col_map[col] = 'No. Pelanggan'
            elif col_str == 'nama pelanggan':
                fallback_col_map[col] = 'Nama Pelanggan'
            elif 'nama kontak' in col_str:
                fallback_col_map[col] = 'Nama kontak Pelanggan'
                
        df_fallback = df_fallback.rename(columns=fallback_col_map)
        df = pd.concat([df_faktur, df_fallback], ignore_index=True)
    else:
        df = df_faktur

    def normalize_no_pelanggan(val):
        if pd.isna(val):
            return ""
        if isinstance(val, (int, float)):
            return str(int(val))
        val_str = str(val).strip()
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        clean_dots = val_str.replace('.', '')
        if clean_dots.isdigit() and len(val_str) > 3 and '.' in val_str:
            return clean_dots
        return val_str

    df['No. Pelanggan'] = df['No. Pelanggan'].apply(normalize_no_pelanggan)

    if mini_del:
        del_pattern = '|'.join([re.escape(k) for k in mini_del])
        df = df[~df['Nama kontak Pelanggan'].astype(str).str.contains(del_pattern, case=False, na=False)]

    if mini_filter:
        filter_pattern = '|'.join([re.escape(k) for k in mini_filter])
        df = df[df['Nama kontak Pelanggan'].astype(str).str.contains(filter_pattern, case=False, na=False)]

    if df.empty:
        print("--> Tidak ada data tersisa setelah penyaringan.")
        return

    print("--> Menjalankan pembersihan nama dan identifikasi merek...")
    
    default_suffixes = ['IRC', 'ZN', 'TT', 'TL', 'BD', 'GT', 'IX', 'FASTI', 'JIMCO']
    all_clean_keys = list(set(mini_str_clean + default_suffixes))
    sorted_keys = sorted(all_clean_keys, key=len, reverse=True)
    
    pattern = r'(?i)\b(?:' + '|'.join([re.escape(k) for k in sorted_keys]) + r').*$'
    
    def initial_clean(val):
        if not isinstance(val, str) or not val.strip():
            return ""
        cleaned = re.sub(pattern, '', val).strip()
        return re.sub(r'[\s\-_,.]+$', '', cleaned).strip()

    df['Nama_Base_Temp'] = df['Nama kontak Pelanggan'].apply(initial_clean)

    unique_bases = sorted(list(set(df['Nama_Base_Temp'].dropna().tolist())), key=len, reverse=True)
    parent_map = {b: b for b in unique_bases}
    for b1 in unique_bases:
        for b2 in unique_bases:
            if b1 != b2:
                regex_b1 = r'^\s*' + re.escape(b1) + r'(\s+|$)'
                if re.match(regex_b1, b2, re.IGNORECASE):
                    parent_map[b1] = parent_map[b2]

    df['Nama Pelanggan Pengganti'] = df['Nama_Base_Temp'].map(parent_map)

    primary_brand_keywords = [k for k in list(set(mini_filter + mini_str_clean)) if k]
    if not primary_brand_keywords:
        primary_brand_keywords = ['IRC', 'ZN']

    df['Brand_Tag'] = df['Nama kontak Pelanggan'].apply(lambda x: extract_brand_tag(x, primary_brand_keywords))

    extracted_num = df['No. Pelanggan'].astype(str).str.extract(r'(\d+)', expand=False)
    df['No_Pelanggan_Num'] = pd.to_numeric(extracted_num, errors='coerce').fillna(0)

    print("--> Menghitung Min No. Pelanggan dan Penjual Terbaru...")
    
    min_no_map = {}
    for name_pengganti, group in df.groupby('Nama Pelanggan Pengganti'):
        min_row = group.sort_values(by=['No_Pelanggan_Num', 'No. Pelanggan'], ascending=[True, True]).iloc[0]
        min_no_map[name_pengganti] = min_row['No. Pelanggan']

    seller_map = {}
    for (name_pengganti, brand), group in df.groupby(['Nama Pelanggan Pengganti', 'Brand_Tag']):
        valid_sellers = group[group['Nama Penjual'].notna() & (group['Nama Penjual'].astype(str).str.strip() != '')]
        if not valid_sellers.empty:
            latest_row = valid_sellers.sort_values(by=['No_Pelanggan_Num', 'No. Pelanggan'], ascending=[False, False]).iloc[0]
            seller_map[(name_pengganti, brand)] = latest_row['Nama Penjual']

    df['Min No. Pelanggan'] = df['Nama Pelanggan Pengganti'].map(min_no_map)
    df['Nama Penjual Penggganti'] = df.apply(
        lambda r: seller_map.get((r['Nama Pelanggan Pengganti'], r['Brand_Tag']), r['Nama Penjual']), 
        axis=1
    )

    df_result = df.drop_duplicates(subset=['Nama kontak Pelanggan']).copy()

    df_result = df_result.sort_values(by=['Nama Pelanggan Pengganti', 'No_Pelanggan_Num'], ascending=[True, False]).reset_index(drop=True)
    df_result.insert(0, 'No', df_result.index + 1)

    target_cols = [
        'No',
        'Min No. Pelanggan',
        'No. Pelanggan',
        'Nama Pelanggan',
        'Nama kontak Pelanggan',
        'Kota Pelanggan',
        'Nama Penjual',
        'Nama Penjual Penggganti',
        'Nama Pelanggan Pengganti'
    ]

    for col in target_cols:
        if col not in df_result.columns:
            df_result[col] = None

    df_final = df_result[target_cols]

    print(f"--> Menyimpan {len(df_final)} baris data presisi ke '{output_file}'...")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Minifs')
        
        ws = writer.sheets['Minifs']
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    print("--> Selesai! Penjual terbaru berdasarkan merek utama berhasil diperbarui.")

if __name__ == '__main__':
    process_minifs()
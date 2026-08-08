import pandas as pd
import openpyxl

def rapikan_faktur_penjualan(input_path, output_path):
    print("--> Membaca file Excel mentah...")
    df_raw = pd.read_excel(input_path, header=None)
    
    target_columns = [
        'No. Faktur', 
        'Tgl Faktur', 
        'Jatuh Tempo', 
        'No. Pelanggan',
        'Nama Pelanggan', 
        'Alamat 1 Pelanggan', 
        'Nilai Faktur', 
        'Terutang',
        'Jumlah Pembayaran', 
        'Nama Gudang', 
        'Nama Penjual', 
        'Kota Pelanggan',
        'Negara Pelanggan',
        'Nama kontak Pelanggan',
        'Tgl terima Last Payment'
    ]
    
    data_start_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row if pd.notna(v)]).lower()
        if ('01/' in row_str or '02/' in row_str or '03/' in row_str or '260.' in row_str or '250.' in row_str):
            if 'dari' not in row_str and 'daftar' not in row_str:
                data_start_idx = idx
                break
                
    if data_start_idx is None:
        data_start_idx = 7
        
    print(f"--> Baris awal data transaksi terdeteksi pada indeks ke-{data_start_idx + 1}.")
    
    header_rows = df_raw.iloc[:data_start_idx]
    
    col_combined_headers = {}
    for col_idx in range(df_raw.shape[1]):
        col_values = header_rows.iloc[:, col_idx].dropna().astype(str).tolist()
        combined_text = " ".join([v.strip() for v in col_values if v.strip() != ''])
        col_combined_headers[col_idx] = combined_text
        
    col_map = {}
    for target in target_columns:
        target_clean = target.lower().replace(' ', '').replace('.', '')
        best_col = None
        
        for col_idx, text in col_combined_headers.items():
            text_clean = text.lower().replace(' ', '').replace('.', '')
            
            if target.lower() in text.lower() or target_clean in text_clean:
                best_col = col_idx
                break
            elif target == 'No. Pelanggan' and ('nopelanggan' in text_clean or 'nopelang' in text_clean or ('pelanggan' in text_clean and 'nama' not in text_clean and 'alamat' not in text_clean and 'kota' not in text_clean and 'negara' not in text_clean)):
                best_col = col_idx
            elif target == 'Alamat 1 Pelanggan' and 'alamat' in text_clean:
                best_col = col_idx
            elif target == 'Negara Pelanggan' and 'negara' in text_clean:
                best_col = col_idx
            elif target == 'Nama kontak Pelanggan' and 'kontak' in text_clean:
                best_col = col_idx
            elif target == 'Tgl terima Last Payment' and ('tglterima' in text_clean or 'lastpayment' in text_clean or 'payment' in text_clean):
                best_col = col_idx
                
        if best_col is not None:
            col_map[target] = best_col

    print(f"--> Berhasil memetakan {len(col_map)} dari {len(target_columns)} kolom target.")

    df_data = df_raw.iloc[data_start_idx:].copy()
    
    clean_dict = {}
    for col in target_columns:
        if col in col_map:
            clean_dict[col] = df_data.iloc[:, col_map[col]].values
        else:
            clean_dict[col] = None
            
    df_clean = pd.DataFrame(clean_dict)
    
    print("--> Membersihkan baris kosong dan karakter pengganggu...")
    
    df_clean = df_clean.dropna(how='all')
    df_clean = df_clean[df_clean['No. Faktur'].notna()]
    df_clean = df_clean[df_clean['No. Faktur'].astype(str).str.strip() != '']
    df_clean = df_clean[~df_clean['No. Faktur'].astype(str).str.lower().str.contains('no. faktur|total|halaman|page', na=False)]
    
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].apply(lambda x: str(x).strip() if pd.notna(x) else None)
        df_clean[col] = df_clean[col].replace({'nan': None, 'None': None, '': None})
        
    df_clean.reset_index(drop=True, inplace=True)
    
    print(f"--> Menyimpan {len(df_clean)} baris data lengkap ke '{output_path}'...")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Data Faktur')
        worksheet = writer.sheets['Data Faktur']
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
    print("--> Selesai! Seluruh kolom data berhasil terambil dan dirapikan.")
    return df_clean

if __name__ == '__main__':
    FILE_INPUT = 'Daftar Faktur Penjualan.xls'
    FILE_OUTPUT = 'Faktur_Penjualan_temp.xlsx'
    
    rapikan_faktur_penjualan(FILE_INPUT, FILE_OUTPUT)
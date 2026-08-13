import pandas as pd

faktur_path = 'Faktur_Penjualan_temp.xlsx'
minifs_path = 'Minifs_temp.xlsx'

print("--> Membaca data dari file Excel...")
df_faktur = pd.read_excel(faktur_path, sheet_name='Data Faktur')
df_minifs = pd.read_excel(minifs_path, sheet_name='Minifs')

df_faktur['No. Pelanggan'] = df_faktur['No. Pelanggan'].astype(str).str.strip()
df_minifs['No. Pelanggan'] = df_minifs['No. Pelanggan'].astype(str).str.strip()
df_minifs['Min No. Pelanggan'] = df_minifs['Min No. Pelanggan'].astype(str).str.strip()

mapping_min = df_minifs.drop_duplicates(subset=['No. Pelanggan']).set_index('No. Pelanggan')['Min No. Pelanggan']
df_faktur['Min No. Pelanggan'] = df_faktur['No. Pelanggan'].map(mapping_min).fillna(df_faktur['No. Pelanggan'])

df_faktur['Tgl_Faktur_dt'] = pd.to_datetime(df_faktur['Tgl Faktur'], format='%d/%m/%y', errors='coerce')
df_faktur['Tgl_Last_Payment_dt'] = pd.to_datetime(df_faktur['Tgl terima Last Payment'], format='%d/%m/%y', errors='coerce')

df_faktur['Lama Pembayaran'] = (df_faktur['Tgl_Last_Payment_dt'] - df_faktur['Tgl_Faktur_dt']).dt.days
df_faktur['Lama Pembayaran Fix'] = df_faktur['Lama Pembayaran'].apply(lambda x: 1 if x == 0 else x)

df_faktur['Bulan Last Payment'] = df_faktur['Tgl_Last_Payment_dt'].dt.strftime('%m')

pivot_df = pd.pivot_table(
    df_faktur,
    index='Min No. Pelanggan',
    columns='Bulan Last Payment',
    values=['Lama Pembayaran', 'Lama Pembayaran Fix'],
    aggfunc='mean'
).round(2)

output_file = 'Hasil_Lama_Pembayaran_temp.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    output_columns = [
        'No. Faktur', 'No. Pelanggan', 'Min No. Pelanggan',
        'Tgl Faktur', 'Tgl terima Last Payment',
        'Lama Pembayaran', 'Lama Pembayaran Fix'
    ]
    df_faktur[output_columns].to_excel(writer, sheet_name='Sheet1', index=False)
    pivot_df.to_excel(writer, sheet_name='Sheet2')

print(f"--> Selesai! File '{output_file}' berhasil dibuat dengan Sheet1 dan Sheet2.")
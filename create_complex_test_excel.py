#!/usr/bin/env python3
"""
Script untuk membuat file Excel test kompleks dengan struktur data yang mirip contoh user
"""

import pandas as pd
import numpy as np
from datetime import datetime

def create_complex_test_excel():
    """Buat file Excel test dengan struktur data kompleks"""
    
    print("📝 Membuat file Excel test kompleks...")
    
    # Data header yang akan diletakkan di baris pertama
    header_data = {
        'A1': 'Nomor Tagihan, Nomor Registrasi, Tanggal Registrasi',
        'B1': ': IP-00030178, : 2508250425, : 25 Agust 2025',
        'C1': 'Nama Pasien, Terima Dari, Kelas / Kamar',
        'D1': ':,:,: KELAS 1/',
        'E1': 'Penjamin Bayar, Tanggal Keluar, Kelas Dijamin',
        'F1': ': ALLIANZ, : 26 Agust 2025, : KELAS 1'
    }
    
    # Data detail biaya
    detail_data = {
        'A2': 'Jenis Biaya',
        'B2': 'Waktu',
        'C2': 'Tanggal',
        'D2': 'Keterangan/Deskripsi',
        'E2': 'Jumlah/Qty',
        'F2': 'Nilai',
        'G2': 'Sub Total',
        
        'A3': 'Biaya Kamar',
        'B3': '08:00',
        'C3': '25/08/2025',
        'D3': 'Kamar Kelas 1',
        'E3': '1',
        'F3': 'Rp 500,000',
        'G3': 'Rp 500,000',
        
        'A4': 'Biaya Visite',
        'B4': '09:00',
        'C4': '25/08/2025',
        'D4': 'Kunjungan Dokter',
        'E4': '1',
        'F4': 'Rp 150,000',
        'G4': 'Rp 150,000',
        
        'A5': 'Biaya Laboratorium',
        'B5': '10:00',
        'C5': '25/08/2025',
        'D5': 'Tes Darah Lengkap',
        'E5': '1',
        'F5': 'Rp 200,000',
        'G5': 'Rp 200,000',
        
        'A6': 'Biaya Radiologi',
        'B6': '11:00',
        'C6': '25/08/2025',
        'D6': 'X-Ray Thorax',
        'E6': '1',
        'F6': 'Rp 300,000',
        'G6': 'Rp 300,000',
        
        'A7': 'Biaya Peralatan',
        'B7': '12:00',
        'C7': '25/08/2025',
        'D7': 'Pemakaian Oksigen',
        'E7': '2',
        'F7': 'Rp 75,000',
        'G7': 'Rp 150,000',
        
        'A8': 'Biaya Obat',
        'B8': '13:00',
        'C8': '25/08/2025',
        'D8': 'Antibiotik',
        'E8': '3',
        'F8': 'Rp 50,000',
        'G8': 'Rp 150,000',
        
        'A9': 'Administrasi',
        'B9': '14:00',
        'C9': '25/08/2025',
        'D9': 'Biaya Administrasi',
        'E9': '1',
        'F9': 'Rp 100,000',
        'G9': 'Rp 100,000'
    }
    
    # Buat DataFrame dengan struktur yang kompleks
    data = []
    
    # Baris header
    header_row = [''] * 7
    for cell, value in header_data.items():
        col_idx = ord(cell[0]) - ord('A')
        if col_idx < len(header_row):
            header_row[col_idx] = value
    data.append(header_row)
    
    # Baris kosong
    data.append([''] * 7)
    
    # Baris detail biaya
    for row_num in range(2, 10):
        detail_row = [''] * 7
        for cell, value in detail_data.items():
            if cell.endswith(str(row_num)):
                col_idx = ord(cell[0]) - ord('A')
                if col_idx < len(detail_row):
                    detail_row[col_idx] = value
        data.append(detail_row)
    
    # Buat DataFrame
    df = pd.DataFrame(data)
    
    # Buat nama file dengan timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"complex_test_data_{timestamp}.xlsx"
    
    # Simpan ke file Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Billing Data', index=False, header=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Billing Data']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"✅ File Excel test kompleks berhasil dibuat: {filename}")
    print(f"📊 Data: {len(df)} baris x {len(df.columns)} kolom")
    print(f"🔍 Struktur: Header + Detail Biaya + Total")
    
    return filename

if __name__ == "__main__":
    test_file = create_complex_test_excel()
    print(f"\n🎯 File test kompleks siap: {test_file}")
    print("Sekarang Anda bisa test sistem Smart Learning dengan file ini!")

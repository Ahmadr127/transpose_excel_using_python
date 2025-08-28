#!/usr/bin/env python3
"""
Script untuk membuat file Excel test sederhana
"""

import pandas as pd
import numpy as np
from datetime import datetime

def create_test_excel():
    """Buat file Excel test dengan data yang mirip contoh"""
    
    print("📝 Membuat file Excel test...")
    
    # Data test yang mirip dengan contoh
    data = {
        'PROVID': ['', '', '', 'Alkes', 'Alkes', 'Obat', '', ''],
        'PROVIDER_NAME': ['', '', '', '', '', '', '', ''],
        'SERVICECODE': ['010102', '06010120', '06010121', 'Alkes', 'Alkes', 'Obat', '140601', '22101'],
        'SERVICECODE DESCRIPTION': [
            'ADMINISTRATION',
            'GENERAL DOCTOR FEE (TO IP)',
            'SPECIALIST DOCTOR FEE',
            'ALKOHOL SWAB',
            'DISCOFIX 3-WAY TUBING B BRAUN (409810/2)',
            'HEXILON 125 MG / 2 ML INJ',
            'ECG / ELECTRO CARDIOGRAPHY (DR.UMUM)',
            'X-RAY CHEST PA'
        ],
        'KELAS': ['ER', 'ER', 'ER', 'ER', 'ER', 'ER', 'ER', 'ER'],
        'RUANG BEDAH (SURGERY)/NON RUANG BEDAH (NON SURGERY)': [
            'NON OK', 'NON OK', 'NON OK', 'NON OK', 'NON OK', 'NON OK', 'NON OK', 'NON OK'
        ],
        'HELPER': ['', '', '', '', '', '', '', ''],
        'TARIFF': ['75,000', '130,000', '886', '70,630', '886', '70,630', '886', '70,630'],
        'TARIFF DESCRIPTION': ['', '', '', '', '', '', '', ''],
        'QUANTITY': [1, 2, 6, 1, 1, 1, 1, 1],
        'TOTAL BILLED': ['75,000', '260,000', '5,316', '70,630', '886', '70,630', '886', '70,630']
    }
    
    # Buat DataFrame
    df = pd.DataFrame(data)
    
    # Buat nama file dengan timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"test_data_{timestamp}.xlsx"
    
    # Simpan ke file Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Test Data', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Test Data']
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
    
    print(f"✅ File Excel test berhasil dibuat: {filename}")
    print(f"📊 Data: {len(df)} baris x {len(df.columns)} kolom")
    
    return filename

if __name__ == "__main__":
    test_file = create_test_excel()
    print(f"\n🎯 File test siap: {test_file}")
    print("Sekarang Anda bisa test sistem dengan file ini!")

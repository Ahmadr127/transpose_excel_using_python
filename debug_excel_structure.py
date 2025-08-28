#!/usr/bin/env python3
"""
Debug script untuk memeriksa struktur file Excel
"""

import pandas as pd
import os

def debug_excel_structure():
    """Debug struktur file Excel"""
    
    print("🔍 Debugging Excel File Structure...")
    
    if not os.path.exists('sampledata.xlsx'):
        print("❌ File sampledata.xlsx tidak ditemukan")
        return
    
    try:
        # Baca semua sheet
        excel_file = pd.ExcelFile('sampledata.xlsx')
        print(f"📊 File memiliki {len(excel_file.sheet_names)} sheet: {excel_file.sheet_names}")
        
        for sheet_name in excel_file.sheet_names:
            print(f"\n🔍 Sheet: {sheet_name}")
            
            # Baca sheet tanpa header
            df = pd.read_excel('sampledata.xlsx', sheet_name=sheet_name, header=None)
            print(f"  - Shape: {df.shape}")
            
            # Cek 20 baris pertama
            print(f"  - 20 baris pertama:")
            for i in range(min(20, len(df))):
                row_data = df.iloc[i]
                # Filter hanya nilai yang tidak kosong
                non_empty = [f"{j}:{val}" for j, val in enumerate(row_data) if pd.notna(val) and str(val).strip()]
                if non_empty:
                    print(f"    Row {i}: {non_empty[:5]}...")  # Tampilkan max 5 kolom
            
            # Cari baris yang mungkin berisi nama pasien
            print(f"  - Mencari baris dengan nama pasien:")
            for i in range(min(50, len(df))):
                row_data = df.iloc[i]
                row_str = ' '.join([str(val) for val in row_data if pd.notna(val)]).lower()
                
                # Cek apakah ada kata kunci nama pasien
                if any(keyword in row_str for keyword in ['nama', 'pasien', 'client', 'patient']):
                    print(f"    Row {i}: {row_str[:100]}...")
                
                # Cek apakah ada nama yang terlihat seperti nama orang
                for j, val in enumerate(row_data):
                    if pd.notna(val):
                        val_str = str(val).strip()
                        # Cek apakah ini nama (huruf kapital di awal, tidak ada angka, panjang > 3)
                        if (len(val_str) > 3 and 
                            val_str[0].isupper() and 
                            not any(c.isdigit() for c in val_str) and
                            ' ' in val_str):
                            print(f"    Row {i}, Col {j}: Potential name: '{val_str}'")
            
            # Cek kolom yang berisi data penting
            print(f"  - Analisis kolom:")
            for col_idx in range(min(10, len(df.columns))):
                col_data = df.iloc[:, col_idx].dropna()
                if len(col_data) > 0:
                    # Cek apakah ada pola yang menarik
                    sample_values = col_data.head(3).tolist()
                    print(f"    Col {col_idx}: {sample_values}")
                    
                    # Cek apakah ada pola nama
                    name_pattern = any(
                        len(str(val)) > 3 and 
                        str(val)[0].isupper() and 
                        not any(c.isdigit() for c in str(val))
                        for val in sample_values
                    )
                    if name_pattern:
                        print(f"      -> Potential name column!")
    
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_excel_structure()

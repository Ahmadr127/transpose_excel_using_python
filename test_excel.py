#!/usr/bin/env python3
"""
Test script untuk ExcelProcessor
"""

import os
import sys
from excel_processor import ExcelProcessor

def test_excel_processor():
    """Test ExcelProcessor dengan file Excel yang ada"""
    
    print("🧪 Testing ExcelProcessor...")
    
    # Cari file Excel yang ada
    excel_files = []
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
            excel_files.append(file)
    
    if not excel_files:
        print("❌ Tidak ada file Excel yang ditemukan untuk testing")
        return False
    
    print(f"📁 File Excel yang ditemukan: {excel_files}")
    
    # Test dengan file pertama
    test_file = excel_files[0]
    print(f"\n🔍 Testing dengan file: {test_file}")
    
    try:
        processor = ExcelProcessor()
        
        # Test preview
        print("📊 Testing preview_excel...")
        preview = processor.preview_excel(test_file)
        
        print(f"✅ Preview berhasil:")
        if 'summary' in preview:
            print(f"   - Total sheet: {preview['total_sheets']}")
            print(f"   - Total baris: {preview['summary']['total_rows']}")
            print(f"   - Field terdeteksi: {preview['summary']['detected_field_count']}")
        else:
            print(f"   - Total kolom: {preview['total_columns']}")
            print(f"   - Total baris: {preview['total_rows']}")
        
        # Tampilkan field yang terdeteksi
        if 'global_detected_fields' in preview:
            print(f"   - Field: {list(preview['global_detected_fields'].keys())[:5]}...")
        elif 'columns' in preview:
            print(f"   - Kolom: {preview['columns'][:5]}...")
        
        # Test process (tanpa membuat file output)
        print("\n⚙️  Testing process_excel...")
        output_file = processor.process_excel(test_file)
        
        if os.path.exists(output_file):
            print(f"✅ Process berhasil, output file: {output_file}")
            # Hapus file output test
            os.remove(output_file)
            print("🗑️  File output test dihapus")
        else:
            print("❌ Process gagal, file output tidak dibuat")
            return False
        
        print("\n🎉 Semua test berhasil!")
        return True
        
    except Exception as e:
        print(f"❌ Test gagal: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_excel_processor()
    sys.exit(0 if success else 1)

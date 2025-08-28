#!/usr/bin/env python3
"""
Test script untuk verifikasi klasifikasi service code
"""

import os
import sys
from excel_processor import ExcelProcessor

def test_service_code_classification():
    """Test klasifikasi service code berdasarkan jenis_biaya"""
    
    print("🧪 Testing Service Code Classification...")
    
    # Test cases untuk klasifikasi
    test_cases = [
        {
            'jenis_biaya': 'Biaya Peralatan',
            'expected': 'Alkes',
            'description': 'Peralatan medis'
        },
        {
            'jenis_biaya': 'Biaya Alkes',
            'expected': 'Alkes', 
            'description': 'Alat kesehatan'
        },
        {
            'jenis_biaya': 'Biaya Obat',
            'expected': 'Obat',
            'description': 'Obat-obatan'
        },
        {
            'jenis_biaya': 'Biaya Kamar',
            'expected': '',
            'description': 'Biaya kamar (bukan obat/peralatan)'
        },
        {
            'jenis_biaya': 'Biaya Visite',
            'expected': '',
            'description': 'Biaya visite (bukan obat/peralatan)'
        },
        {
            'jenis_biaya': 'Medical Equipment',
            'expected': 'Alkes',
            'description': 'Equipment dalam bahasa Inggris'
        },
        {
            'jenis_biaya': 'Medicine Cost',
            'expected': 'Obat',
            'description': 'Medicine dalam bahasa Inggris'
        }
    ]
    
    # Test ExcelProcessor
    processor = ExcelProcessor()
    
    print("\n📊 Testing Service Code Classification Logic:")
    print("=" * 60)
    
    success_count = 0
    total_count = len(test_cases)
    
    for idx, test_case in enumerate(test_cases, 1):
        print(f"\n🔍 Test Case {idx}: {test_case['description']}")
        print(f"   Input: '{test_case['jenis_biaya']}'")
        print(f"   Expected: '{test_case['expected']}'")
        
        # Buat data row dummy
        data_row = {
            'jenis_biaya': test_case['jenis_biaya'],
            'keterangan': 'Sample keterangan'
        }
        
        # Test classification directly
        try:
            result = processor._classify_service_code(data_row)
            print(f"   Result: '{result}'")
            
            if result == test_case['expected']:
                print(f"   ✅ PASS")
                success_count += 1
            else:
                print(f"   ❌ FAIL - Expected '{test_case['expected']}', got '{result}'")
                
        except Exception as e:
            print(f"   ❌ ERROR: {e}")
    
    print("\n" + "=" * 60)
    print(f"📊 Test Results: {success_count}/{total_count} tests passed")
    
    if success_count == total_count:
        print("🎉 All tests passed! Service code classification is working correctly.")
        return True
    else:
        print("⚠️ Some tests failed. Please check the logic.")
        return False

def test_with_sampledata():
    """Test dengan file sampledata.xlsx yang sebenarnya"""
    
    print("\n🔍 Testing with actual sampledata.xlsx...")
    
    if not os.path.exists('sampledata.xlsx'):
        print("❌ File sampledata.xlsx tidak ditemukan")
        return False
    
    try:
        processor = ExcelProcessor()
        
        # Process file
        output_filepath = processor.process_excel('sampledata.xlsx')
        
        if os.path.exists(output_filepath):
            print(f"✅ Output file created: {output_filepath}")
            
            # Check output content
            import pandas as pd
            output_df = pd.read_excel(output_filepath)
            
            print(f"📊 Output file content:")
            print(f"  - Shape: {output_df.shape}")
            print(f"  - Columns: {list(output_df.columns)[:5]}...")
            
            # Check SERVICECODE DESCRIPTION column
            if 'SERVICECODE DESCRIPTION' in output_df.columns:
                service_codes = output_df['SERVICECODE DESCRIPTION'].dropna().unique()
                print(f"  - Unique Service Codes: {list(service_codes)}")
                
                # Count each type
                alkes_count = (output_df['SERVICECODE DESCRIPTION'] == 'Alkes').sum()
                obat_count = (output_df['SERVICECODE DESCRIPTION'] == 'Obat').sum()
                empty_count = (output_df['SERVICECODE DESCRIPTION'] == '').sum()
                
                print(f"  - Alkes count: {alkes_count}")
                print(f"  - Obat count: {obat_count}")
                print(f"  - Empty count: {empty_count}")
                
                # Show sample rows
                print(f"\n📋 Sample rows with service codes:")
                sample_df = output_df[output_df['SERVICECODE DESCRIPTION'].isin(['Alkes', 'Obat'])].head(5)
                for idx, row in sample_df.iterrows():
                    jenis_biaya = row.get('SERVICECODE DESCRIPTION', 'N/A')
                    keterangan = row.get('SERVICECODE DESCRIPTION', 'N/A')
                    print(f"    Row {idx}: Service Code='{jenis_biaya}', Keterangan='{keterangan}'")
            
            # Clean up
            try:
                os.remove(output_filepath)
                print(f"🗑️ Output file cleaned up")
            except Exception as cleanup_error:
                print(f"⚠️ Warning: Could not clean up output file: {cleanup_error}")
                
        else:
            print(f"❌ Output file not created")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error testing with sampledata.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🧪 Service Code Classification Test Suite")
    print("=" * 60)
    
    # Test 1: Logic classification
    test1_success = test_service_code_classification()
    
    # Test 2: With actual file
    test2_success = test_with_sampledata()
    
    print("\n" + "=" * 60)
    print("🏁 Final Results:")
    print(f"  - Logic Test: {'✅ PASS' if test1_success else '❌ FAIL'}")
    print(f"  - File Test: {'✅ PASS' if test2_success else '❌ FAIL'}")
    
    if test1_success and test2_success:
        print("🎉 All tests completed successfully!")
        sys.exit(0)
    else:
        print("⚠️ Some tests failed. Please review the results.")
        sys.exit(1)

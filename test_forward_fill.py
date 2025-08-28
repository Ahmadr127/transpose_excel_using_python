#!/usr/bin/env python3
"""
Test script untuk verifikasi forward fill functionality
"""

import os
import sys
import pandas as pd
from excel_processor import ExcelProcessor

def test_forward_fill():
    """Test forward fill functionality"""
    
    print("🔄 Testing Forward Fill Functionality...")
    
    # Test 1: Test dengan data dummy
    print("\n📊 Test 1: Forward fill dengan data dummy")
    
    # Buat DataFrame dummy dengan beberapa baris kosong
    dummy_data = {
        'CLIENT NAME': ['Ujang Sunarja', '', '', '', ''],
        'CLIENTS INVOICE NUMBER': ['IP-00030178', '', '', '', ''],
        'CLIENTSREGISTER NUMBER': ['2508250425', '', '', '', ''],
        'TARIFF': ['700.000,-', '130.000,-', '384.000,-', '100.800,-', '512.000,-'],
        'QUANTITY': ['1', '1', '6', '1', '8']
    }
    
    df_dummy = pd.DataFrame(dummy_data)
    print(f"📋 Original DataFrame:")
    print(df_dummy[['CLIENT NAME', 'CLIENTS INVOICE NUMBER', 'CLIENTSREGISTER NUMBER']].to_string())
    
    # Test forward fill method
    processor = ExcelProcessor()
    df_filled = processor._apply_forward_fill(df_dummy)
    
    print(f"\n📋 After Forward Fill:")
    print(df_filled[['CLIENT NAME', 'CLIENTS INVOICE NUMBER', 'CLIENTSREGISTER NUMBER']].to_string())
    
    # Verifikasi hasil
    success = True
    for col in ['CLIENT NAME', 'CLIENTS INVOICE NUMBER', 'CLIENTSREGISTER NUMBER']:
        if df_filled[col].isna().any() or (df_filled[col] == '').any():
            print(f"❌ FAIL: Column {col} still has empty values")
            success = False
        else:
            print(f"✅ PASS: Column {col} is fully filled")
    
    # Test 2: Test dengan file Excel asli
    print(f"\n🔍 Test 2: Forward fill dengan file Excel asli")
    
    if os.path.exists('sampledata.xlsx'):
        try:
            output_filepath = processor.process_excel('sampledata.xlsx')
            
            if os.path.exists(output_filepath):
                print(f"✅ Output file created: {output_filepath}")
                
                # Check output file content
                output_df = pd.read_excel(output_filepath)
                
                print(f"📊 Output file content:")
                print(f"  - Shape: {output_df.shape}")
                
                # Check forward fill columns
                forward_fill_columns = ['CLIENT NAME', 'CLIENTS INVOICE NUMBER', 'CLIENTSREGISTER NUMBER']
                
                for col in forward_fill_columns:
                    if col in output_df.columns:
                        # Cek apakah ada nilai kosong
                        empty_count = (output_df[col].isna() | (output_df[col] == '')).sum()
                        total_count = len(output_df)
                        
                        print(f"  - {col}: {total_count - empty_count}/{total_count} filled ({empty_count} empty)")
                        
                        if empty_count == 0:
                            print(f"    ✅ Column {col} is fully filled")
                        else:
                            print(f"    ⚠️ Column {col} has {empty_count} empty cells")
                            success = False
                    else:
                        print(f"  - ❌ Column {col} not found in output")
                        success = False
                
                # Show sample data
                print(f"\n📋 Sample data from output file:")
                for idx, row in output_df.head(10).iterrows():
                    client_name = row.get('CLIENT NAME', 'N/A')
                    invoice_number = row.get('CLIENTS INVOICE NUMBER', 'N/A')
                    register_number = row.get('CLIENTSREGISTER NUMBER', 'N/A')
                    print(f"    Row {idx}: CLIENT NAME='{client_name}', INVOICE='{invoice_number}', REGISTER='{register_number}'")
                
                # Clean up
                try:
                    os.remove(output_filepath)
                    print(f"\n🗑️ Output file cleaned up")
                except Exception as cleanup_error:
                    print(f"\n⚠️ Warning: Could not clean up output file: {cleanup_error}")
            else:
                print(f"❌ Output file not created")
                success = False
                
        except Exception as e:
            print(f"❌ Error processing Excel file: {e}")
            import traceback
            traceback.print_exc()
            success = False
    else:
        print(f"❌ sampledata.xlsx not found")
        success = False
    
    if success:
        print(f"\n✅ Forward fill test completed successfully!")
    else:
        print(f"\n❌ Forward fill test failed!")
    
    return success

if __name__ == "__main__":
    success = test_forward_fill()
    sys.exit(0 if success else 1)

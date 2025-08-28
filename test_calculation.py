#!/usr/bin/env python3
"""
Test script untuk verifikasi perhitungan total billed
"""

import os
import sys
from excel_processor import ExcelProcessor

def test_calculation():
    """Test perhitungan total billed"""
    
    print("🧮 Testing Total Billed Calculation...")
    
    # Test cases untuk perhitungan
    test_cases = [
        {
            'nilai': '700.000,-',
            'jumlah': '1',
            'expected': '700,000'
        },
        {
            'nilai': '130.000,-',
            'jumlah': '1',
            'expected': '130,000'
        },
        {
            'nilai': '384.000,-',
            'jumlah': '6',
            'expected': '2,304,000'
        },
        {
            'nilai': '100.800,-',
            'jumlah': '1',
            'expected': '100,800'
        },
        {
            'nilai': '512.000,-',
            'jumlah': '8',
            'expected': '4,096,000'
        }
    ]
    
    processor = ExcelProcessor()
    
    for i, test_case in enumerate(test_cases):
        print(f"\n📊 Test Case {i+1}:")
        print(f"  Tarif: {test_case['nilai']}")
        print(f"  Quantity: {test_case['jumlah']}")
        print(f"  Expected: {test_case['expected']}")
        
        # Buat data row untuk testing
        data_row = {
            'nilai': test_case['nilai'],
            'jumlah': test_case['jumlah']
        }
        
        # Hitung total billed
        calculated_total = processor._calculate_total_billed(data_row)
        print(f"  Calculated: {calculated_total}")
        
        if calculated_total == test_case['expected']:
            print(f"  ✅ PASS")
        else:
            print(f"  ❌ FAIL - Expected {test_case['expected']}, got {calculated_total}")
    
    print("\n🔍 Testing with actual Excel file...")
    
    if os.path.exists('sampledata.xlsx'):
        try:
            output_filepath = processor.process_excel('sampledata.xlsx')
            
            if os.path.exists(output_filepath):
                print(f"✅ Output file created: {output_filepath}")
                
                # Check output file content
                import pandas as pd
                output_df = pd.read_excel(output_filepath)
                
                print(f"📊 Output file content:")
                print(f"  - Shape: {output_df.shape}")
                
                # Check TOTAL BILLED column
                if 'TOTAL BILLED' in output_df.columns:
                    print(f"  - TOTAL BILLED column found")
                    
                    # Show some examples
                    for idx, row in output_df.head(5).iterrows():
                        tarif = row.get('TARIFF', 'N/A')
                        quantity = row.get('QUANTITY', 'N/A')
                        total_billed = row.get('TOTAL BILLED', 'N/A')
                        
                        print(f"    Row {idx}: Tarif={tarif}, Qty={quantity}, Total={total_billed}")
                else:
                    print(f"  - ❌ TOTAL BILLED column not found")
                
                # Clean up
                try:
                    os.remove(output_filepath)
                    print(f"🗑️ Output file cleaned up")
                except Exception as cleanup_error:
                    print(f"⚠️ Warning: Could not clean up output file: {cleanup_error}")
            else:
                print(f"❌ Output file not created")
                return False
                
        except Exception as e:
            print(f"❌ Error processing Excel file: {e}")
            import traceback
            traceback.print_exc()
            return False
    else:
        print(f"❌ sampledata.xlsx not found")
        return False
    
    print("\n✅ Total Billed calculation test completed!")
    return True

if __name__ == "__main__":
    success = test_calculation()
    sys.exit(0 if success else 1)

#!/usr/bin/env python3
"""
Test script untuk verifikasi TARIFF DESCRIPTION column kosong
"""

from excel_processor import ExcelProcessor
import pandas as pd
import os

def test_tariff_description_empty():
    """Test bahwa TARIFF DESCRIPTION column kosong"""
    
    print("🧪 Testing TARIFF DESCRIPTION Column is Empty")
    print("=" * 60)
    
    if not os.path.exists('sampledata.xlsx'):
        print("❌ File sampledata.xlsx tidak ditemukan")
        return False
    
    try:
        processor = ExcelProcessor()
        
        # Process file
        output_filepath = processor.process_excel('sampledata.xlsx')
        
        if os.path.exists(output_filepath):
            print(f"✅ Output file created: {output_filepath}")
            
            # Read output file
            output_df = pd.read_excel(output_filepath)
            
            print(f"📊 Output file analysis:")
            print(f"  - Shape: {output_df.shape}")
            
            # Check TARIFF DESCRIPTION column
            if 'TARIFF DESCRIPTION' in output_df.columns:
                print(f"\n🔍 TARIFF DESCRIPTION Analysis:")
                
                # Check if all values are empty
                tariff_desc_values = output_df['TARIFF DESCRIPTION'].dropna().unique()
                print(f"  - Unique TARIFF DESCRIPTION values: {list(tariff_desc_values)}")
                
                # Count empty vs non-empty
                empty_count = (output_df['TARIFF DESCRIPTION'] == '').sum()
                non_empty_count = len(output_df) - empty_count
                
                print(f"  - Empty values: {empty_count}")
                print(f"  - Non-empty values: {non_empty_count}")
                
                # Show sample rows
                print(f"\n📋 Sample rows:")
                sample_df = output_df[['SERVICECODE', 'SERVICECODE DESCRIPTION', 'TARIFF DESCRIPTION']].head(10)
                for idx, row in sample_df.iterrows():
                    service_code = row.get('SERVICECODE', 'N/A')
                    service_desc = row.get('SERVICECODE DESCRIPTION', 'N/A')
                    tariff_desc = row.get('TARIFF DESCRIPTION', 'N/A')
                    print(f"    Row {idx}: Service Code='{service_code}', Service Desc='{service_desc}', Tariff Desc='{tariff_desc}'")
                
                # Check if all values are empty
                if non_empty_count == 0:
                    print(f"\n✅ SUCCESS: TARIFF DESCRIPTION column is completely empty!")
                    return True
                else:
                    print(f"\n⚠️ WARNING: TARIFF DESCRIPTION column has {non_empty_count} non-empty values")
                    return False
                
            else:
                print(f"❌ TARIFF DESCRIPTION column not found")
                print(f"   Available columns: {list(output_df.columns)}")
                return False
            
            # Clean up
            try:
                os.remove(output_filepath)
                print(f"\n🗑️ Output file cleaned up")
            except Exception as cleanup_error:
                print(f"\n⚠️ Warning: Could not clean up output file: {cleanup_error}")
                
        else:
            print(f"❌ Output file not created")
            return False
        
    except Exception as e:
        print(f"❌ Error testing TARIFF DESCRIPTION: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_tariff_description_empty()

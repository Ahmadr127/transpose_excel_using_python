#!/usr/bin/env python3
"""
Test script untuk ExcelProcessor dengan sampledata.xlsx
"""

import os
import sys
from excel_processor import ExcelProcessor

def test_sampledata():
    """Test ExcelProcessor dengan file sampledata.xlsx"""
    
    print("🧪 Testing ExcelProcessor with sampledata.xlsx...")
    
    # Check if file exists
    if not os.path.exists('sampledata.xlsx'):
        print("❌ Error: sampledata.xlsx not found")
        return False
    
    try:
        # Test ExcelProcessor
        processor = ExcelProcessor()
        
        print(f"🔍 Testing preview_excel with sampledata.xlsx...")
        preview_data = processor.preview_excel('sampledata.xlsx')
        
        print(f"📊 Preview data structure:")
        print(f"  - Keys: {list(preview_data.keys())}")
        print(f"  - Total sheets: {preview_data.get('total_sheets', 'N/A')}")
        
        if 'sheets' in preview_data:
            for sheet_name, sheet_data in preview_data['sheets'].items():
                print(f"  - Sheet '{sheet_name}':")
                print(f"    * Rows: {sheet_data.get('total_rows', 'N/A')}")
                print(f"    * Columns: {sheet_data.get('total_columns', 'N/A')}")
                print(f"    * Detected fields: {len(sheet_data.get('detected_fields', {}))}")
        
        print("\n🔄 Testing process_excel...")
        output_filepath = processor.process_excel('sampledata.xlsx')
        
        if os.path.exists(output_filepath):
            print(f"✅ Output file created: {output_filepath}")
            
            # Check if output file has content
            import pandas as pd
            output_df = pd.read_excel(output_filepath)
            print(f"📊 Output file content:")
            print(f"  - Shape: {output_df.shape}")
            print(f"  - Columns: {list(output_df.columns)[:5]}...")
            print(f"  - First few rows:")
            print(output_df.head(3).to_string())
            
            # Clean up output file
            try:
                os.remove(output_filepath)
                print(f"🗑️ Output file cleaned up")
            except Exception as cleanup_error:
                print(f"⚠️ Warning: Could not clean up output file: {cleanup_error}")
        else:
            print(f"❌ Output file not created")
            return False
        
        print("✅ ExcelProcessor test with sampledata.xlsx completed successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Error testing ExcelProcessor: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_sampledata()
    sys.exit(0 if success else 1)

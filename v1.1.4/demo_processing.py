
"""
Demo script to test the mining data processing system
"""

import sys
import os
from datetime import datetime

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from mining_processor import MiningDataProcessor

def run_demo():
    """Run a demo processing of the existing Excel file"""
    
    # File path
    excel_file = '/home/ubuntu/Uploads/July_2025_DAILY_REPORT.xlsx'
    
    if not os.path.exists(excel_file):
        print("[ERROR] Demo Excel file not found. Please ensure July_2025_DAILY_REPORT.xlsx exists in /home/ubuntu/Uploads/")
        return False
    
    print("Mining Data Processing Demo")
    print("=" * 50)
    print(f"Input file: {excel_file}")
    print(f"Demo time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Initialize processor
    print("[INFO] Initializing processor...")
    processor = MiningDataProcessor("demo_outputs")
    
    # Validate file
    print("[INFO] Validating Excel file...")
    is_valid, missing_sheets = processor.validate_excel_file(excel_file)
    
    if is_valid:
        print("[SUCCESS] File validation successful!")
    else:
        print(f"[WARNING] Some sheets missing: {missing_sheets}")
    
    print()
    
    # Process all sheets
    print("[INFO] Starting processing...")
    results = processor.process_all_sheets(excel_file, ['STOPING', 'BENCHES'])  # Test with 2 sheets first
    
    if 'error' in results:
        print(f"[ERROR] Processing failed: {results['error']}")
        return False
    
    print("\n[SUCCESS] Processing completed!")
    
    # Display results summary
    overall = results.get('overall_summary', {})
    print(f"[INFO] Sheets processed: {overall.get('successful_sheets', 0)}/{overall.get('total_sheets_requested', 0)}")
    print(f"[INFO] Total records: {overall.get('total_records', 0)}")
    print(f"[INFO] Output directory: {overall.get('output_directory', 'N/A')}")
    
    print("\n[INFO] Individual Sheet Results:")
    for sheet_type, sheet_result in results.get('sheets_processed', {}).items():
        status = "[SUCCESS]" if sheet_result['success'] else "[ERROR]"
        print(f"  {status} {sheet_type}: {len(sheet_result['data'])} records" if sheet_result['success'] else f"  {status} {sheet_type}: {sheet_result.get('error', 'Failed')}")
    
    print(f"\n[INFO] Output files created in: demo_outputs/")
    
    # List output files
    if os.path.exists('demo_outputs'):
        files = os.listdir('demo_outputs')
        csv_files = [f for f in files if f.endswith('.csv')]
        if csv_files:
            print("📄 CSV files created:")
            for file in csv_files:
                file_path = os.path.join('demo_outputs', file)
                size = os.path.getsize(file_path) / 1024  # Size in KB
                print(f"  - {file} ({size:.1f} KB)")
    
    return True

if __name__ == "__main__":
    success = run_demo()
    print(f"\n{'🎉 Demo completed successfully!' if success else '❌ Demo failed.'}")

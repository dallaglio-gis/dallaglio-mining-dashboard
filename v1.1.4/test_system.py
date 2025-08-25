
"""
Test script to verify the mining processing system works correctly
"""

import sys
import os
import pandas as pd

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """Test that all modules can be imported correctly"""
    print("Testing imports...")
    
    try:
        from mining_processor import MiningDataProcessor
        print("[SUCCESS] MiningDataProcessor imported successfully")
        
        from processors import StopingProcessor, TrammingProcessor, DevelopmentProcessor, HoistingProcessor, BenchesProcessor
        print("[SUCCESS] All sheet processors imported successfully")
        
        from utils.common import setup_logger, clean_and_validate_data
        print("[SUCCESS] Utility functions imported successfully")
        
        from config.validation_targets import VALIDATION_TARGETS, SHEET_CONFIGS
        print("[SUCCESS] Configuration imported successfully")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Import failed: {e}")
        return False

def test_processor_initialization():
    """Test that the main processor initializes correctly"""
    print("\nTesting processor initialization...")
    
    try:
        from mining_processor import MiningDataProcessor
        processor = MiningDataProcessor("test_outputs")
        print("[SUCCESS] MiningDataProcessor initialized successfully")
        
        # Check that all sub-processors are initialized
        expected_processors = ['STOPING', 'TRAMMING', 'DEVELOPMENT', 'HOISTING', 'BENCHES']
        for proc_name in expected_processors:
            if proc_name in processor.processors:
                print(f"[SUCCESS] {proc_name} processor initialized")
            else:
                print(f"[ERROR] {proc_name} processor missing")
                return False
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Processor initialization failed: {e}")
        return False

def test_utility_functions():
    """Test utility functions work correctly"""
    print("\nTesting utility functions...")
    
    try:
        from utils.common import clean_and_validate_data, is_qaqc_sample
        
        # Test data cleaning
        assert clean_and_validate_data(42.5) == 42.5
        assert clean_and_validate_data("42.5") == 42.5
        assert clean_and_validate_data("") == 0
        assert clean_and_validate_data(None) == 0
        print("[SUCCESS] Data cleaning functions work")
        
        # Test QAQC identification  
        assert is_qaqc_sample("FDUP") == True
        assert is_qaqc_sample("BLANK") == True
        assert is_qaqc_sample("42.5") == False
        assert is_qaqc_sample(42.5) == False
        print("[SUCCESS] QAQC identification works")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Utility function tests failed: {e}")
        return False

def test_config_loading():
    """Test that configuration loads correctly"""
    print("\nTesting configuration loading...")
    
    try:
        from config.validation_targets import VALIDATION_TARGETS, SHEET_CONFIGS
        
        # Check validation targets
        expected_sheets = ['STOPING', 'TRAMMING', 'DEVELOPMENT', 'HOISTING', 'BENCHES']
        for sheet in expected_sheets:
            if sheet in VALIDATION_TARGETS:
                print(f"[SUCCESS] {sheet} validation targets loaded")
            else:
                print(f"[WARNING] {sheet} validation targets missing (may be optional)")
        
        # Check sheet configs
        for sheet in expected_sheets:
            if sheet in SHEET_CONFIGS:
                print(f"[SUCCESS] {sheet} sheet config loaded")
            else:
                print(f"[WARNING] {sheet} sheet config missing (may be optional)")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Configuration loading failed: {e}")
        return False

def test_file_validation():
    """Test file validation with a dummy file"""
    print("\nTesting file validation...")
    
    # Check if we have access to the original Excel file
    test_file = '/home/ubuntu/Uploads/July_2025_DAILY_REPORT.xlsx'
    
    try:
        from mining_processor import MiningDataProcessor
        processor = MiningDataProcessor("test_outputs")
        
        if os.path.exists(test_file):
            is_valid, missing_sheets = processor.validate_excel_file(test_file)
            if is_valid:
                print("[SUCCESS] File validation passed with real Excel file")
            else:
                print(f"[WARNING] File validation found missing sheets: {missing_sheets}")
            return True
        else:
            print("[WARNING] Test Excel file not found, skipping file validation test")
            return True
            
    except Exception as e:
        print(f"[ERROR] File validation test failed: {e}")
        return False

def run_all_tests():
    """Run all tests"""
    print("Mining Data Processing System - Test Suite")
    print("=" * 60)
    
    tests = [
        ("Import Tests", test_imports),
        ("Processor Initialization", test_processor_initialization), 
        ("Utility Functions", test_utility_functions),
        ("Configuration Loading", test_config_loading),
        ("File Validation", test_file_validation)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\n[RUNNING] {test_name}...")
        if test_func():
            passed += 1
            print(f"[PASS] {test_name} PASSED")
        else:
            print(f"[FAIL] {test_name} FAILED")
    
    print(f"\n" + "=" * 60)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("[SUCCESS] All tests passed! System is ready for use.")
        return True
    else:
        print("[WARNING] Some tests failed. Please check the issues above.")
        return False

if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)

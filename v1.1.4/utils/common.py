
"""
Common utility functions for mining data processing
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
from typing import Any, Union, List, Dict, Optional
import logging

def setup_logger(name: str, log_file: str, level=logging.INFO):
    """Setup logger for processing operations"""
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    handler = logging.FileHandler(log_file)
    handler.setFormatter(formatter)
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    logger.addHandler(console_handler)
    
    return logger

def clean_and_validate_data(value: Any, data_type: str = 'numeric') -> Union[float, str]:
    """Clean and validate data values with comprehensive error handling"""
    if pd.isna(value) or value == "" or value is None:
        return 0 if data_type == 'numeric' else ""
    
    if data_type == 'numeric':
        str_val = str(value).strip()
        
        # Handle common errors
        if str_val.lower() in ['nan', 'null', 'none', '', '-', 'n/a', '#div/0!', '#value!']:
            return 0
        
        # Clean numeric string
        str_val = re.sub(r'[^\d.-]', '', str_val)
        
        try:
            return float(str_val)
        except (ValueError, TypeError):
            return 0
    
    elif data_type == 'text':
        return str(value).strip()
    
    return value

def detect_date_columns(df: pd.DataFrame, date_row: int = 15, start_col: int = 7) -> tuple:
    """Detect date columns in Excel sheet"""
    dates = []
    date_columns = []
    
    for col_idx in range(start_col, len(df.columns)):
        if col_idx < len(df.index) and date_row < len(df.index):
            date_val = df.iloc[date_row, col_idx] if date_row < len(df) else None
            
            if pd.notna(date_val):
                if isinstance(date_val, datetime):
                    dates.append(date_val.strftime('%Y-%m-%d'))
                    date_columns.append(col_idx)
                elif isinstance(date_val, (int, float)):
                    # Handle Excel serial dates
                    try:
                        excel_date = pd.to_datetime('1899-12-30') + timedelta(days=date_val)
                        dates.append(excel_date.strftime('%Y-%m-%d'))
                        date_columns.append(col_idx)
                    except:
                        continue
    
    return dates, date_columns

def generate_date_range(base_date: datetime = None) -> List[datetime]:
    """Generate complete date range for the month"""
    if not base_date:
        base_date = datetime(2025, 7, 31)  # Default fallback
    
    year = base_date.year
    month = base_date.month
    
    dates = []
    for day in range(1, 32):
        try:
            date = datetime(year, month, day)
            dates.append(date)
        except ValueError:
            break
    
    return dates

def is_qaqc_sample(depth_value: Any) -> bool:
    """Identify QAQC samples based on text content"""
    if pd.isna(depth_value):
        return False
    
    depth_str = str(depth_value).strip()
    
    # Check for QAQC indicators
    qaqc_indicators = ['FDUP', 'BLANK', 'STD', 'QC', 'DUP']
    if any(indicator in depth_str.upper() for indicator in qaqc_indicators):
        return True
    
    # Try to convert to float - if it fails, it contains text (potential QAQC)
    try:
        float(depth_str)
        return False  # Pure numeric value
    except ValueError:
        return True   # Contains text - QAQC sample

def calculate_monthly_statistics(df: pd.DataFrame, value_columns: List[str]) -> Dict:
    """Calculate monthly statistics for validation"""
    stats = {}
    
    for col in value_columns:
        if col in df.columns:
            col_data = pd.to_numeric(df[col], errors='coerce')
            stats[col] = {
                'sum': col_data.sum(),
                'mean': col_data.mean(),
                'count': col_data.count(),
                'std': col_data.std(),
                'min': col_data.min(),
                'max': col_data.max()
            }
    
    return stats

def validate_against_targets(actual_stats: Dict, targets: Dict) -> Dict:
    """Validate extracted data against known targets"""
    validation_results = {
        'passed': 0,
        'failed': 0,
        'details': {}
    }
    
    tolerance = targets.get('tolerance', 0.02)
    
    for metric, target_value in targets.items():
        if metric == 'tolerance':
            continue
            
        if metric in actual_stats:
            actual_value = actual_stats[metric]['sum'] if isinstance(actual_stats[metric], dict) else actual_stats[metric]
            
            # Calculate percentage difference
            if target_value != 0:
                diff_pct = abs((actual_value - target_value) / target_value)
                passed = diff_pct <= tolerance
            else:
                passed = actual_value == target_value
            
            validation_results['details'][metric] = {
                'target': target_value,
                'actual': actual_value,
                'difference': actual_value - target_value,
                'diff_percentage': diff_pct if target_value != 0 else 0,
                'passed': passed
            }
            
            if passed:
                validation_results['passed'] += 1
            else:
                validation_results['failed'] += 1
    
    return validation_results

def create_summary_report(sheet_name: str, df: pd.DataFrame, validation_results: Dict = None) -> str:
    """Create a comprehensive summary report"""
    report = f"""
{sheet_name.upper()} DATA EXTRACTION SUMMARY
{'=' * 60}

Extraction Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Sheet: {sheet_name}

EXTRACTION RESULTS:
==================
Total Records: {len(df)}
Date Range: {df['Date'].min() if 'Date' in df.columns else 'N/A'} to {df['Date'].max() if 'Date' in df.columns else 'N/A'}

DATA QUALITY:
=============
"""
    
    # Add column information
    report += f"Columns: {', '.join(df.columns)}\n"
    
    # Add data types info
    report += f"Data Types:\n"
    for col in df.columns:
        non_null_count = df[col].notna().sum()
        report += f"  {col}: {non_null_count}/{len(df)} non-null values\n"
    
    # Add validation results if available
    if validation_results:
        report += f"\nVALIDATION RESULTS:\n"
        report += f"==================\n"
        report += f"Tests Passed: {validation_results['passed']}\n"
        report += f"Tests Failed: {validation_results['failed']}\n"
        
        for metric, results in validation_results['details'].items():
            status = "[PASS]" if results['passed'] else "[FAIL]"
            report += f"{metric}: {status} (Target: {results['target']}, Actual: {results['actual']})\n"
    
    return report

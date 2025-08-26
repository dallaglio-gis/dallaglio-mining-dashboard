
"""
Master Mining Data Processing Engine
Orchestrates all sheet processors and provides unified interface
"""

import pandas as pd
import numpy as np
import os
import sys
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import logging

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from processors import (
    StopingProcessor, TrammingProcessor, DevelopmentProcessor, 
    HoistingProcessor, BenchesProcessor, PlantProcessor
)
from utils.common import (
    setup_logger, calculate_monthly_statistics, 
    validate_against_targets, create_summary_report
)
from config.validation_targets import VALIDATION_TARGETS, SHEET_CONFIGS

class MiningDataProcessor:
    """
    Master processor that orchestrates all mining data extraction
    """
    
    def __init__(self, output_dir: str = "outputs"):
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs("logs", exist_ok=True)
        
        self.logger = setup_logger('MiningDataProcessor', 'logs/master.log')
        
        # Initialize processors
        self.processors = {
            'STOPING': StopingProcessor(self.logger),
            'TRAMMING': TrammingProcessor(self.logger),
            'DEVELOPMENT': DevelopmentProcessor(self.logger),
            'HOISTING': HoistingProcessor(self.logger),
            'BENCHES': BenchesProcessor(self.logger),
            'PLANT': PlantProcessor(self.logger)
        }
        
        self.results = {}
        self.validation_results = {}
        
    def validate_excel_file(self, file_path: str) -> Tuple[bool, List[str]]:
        """Validate Excel file and check for required sheets"""
        try:
            xl_file = pd.ExcelFile(file_path)
            available_sheets = xl_file.sheet_names
            
            required_sheets = ['Stoping', 'Tramming', 'Hoisting', 'Development', 'BENCHES.']
            missing_sheets = [sheet for sheet in required_sheets if sheet not in available_sheets]
            
            self.logger.info(f"Available sheets: {available_sheets}")
            if missing_sheets:
                self.logger.warning(f"Missing sheets: {missing_sheets}")
            
            return len(missing_sheets) == 0, missing_sheets
            
        except Exception as e:
            self.logger.error(f"Error validating Excel file: {e}")
            return False, [str(e)]
    
    def process_single_sheet(self, file_path: str, sheet_type: str) -> Dict:
        """Process a single sheet and return results"""
        result = {
            'success': False,
            'data': pd.DataFrame(),
            'error': None,
            'validation': None,
            'summary': ""
        }
        
        try:
            self.logger.info(f"Processing {sheet_type} sheet...")
            
            if sheet_type == 'STOPING':
                df = self.processors['STOPING'].extract_stoping_data(file_path)
                result['data'] = df
                
            elif sheet_type == 'TRAMMING':
                df = self.processors['TRAMMING'].extract_tramming_data(file_path)
                result['data'] = df
                
            elif sheet_type == 'DEVELOPMENT':
                df = self.processors['DEVELOPMENT'].extract_development_data(file_path)
                result['data'] = df
                
            elif sheet_type == 'HOISTING':
                df = self.processors['HOISTING'].extract_hoisting_data(file_path)
                result['data'] = df
                
            elif sheet_type == 'PLANT':
                df = self.processors['PLANT'].extract_plant_data(file_path)
                result['data'] = df
                
            elif sheet_type == 'BENCHES':
                df_raw, df_avg = self.processors['BENCHES'].extract_benches_data(file_path)
                result['data'] = df_raw
                result['data_avg'] = df_avg
                
                # Save both files for benches
                if not df_raw.empty:
                    raw_file = os.path.join(self.output_dir, 'benches_raw_data.csv')
                    df_raw.to_csv(raw_file, index=False)
                
                if not df_avg.empty:
                    avg_file = os.path.join(self.output_dir, 'benches_average_grades.csv')
                    df_avg.to_csv(avg_file, index=False)
            
            # Save main data file
            if not result['data'].empty:
                output_file = os.path.join(self.output_dir, f'{sheet_type.lower()}_data.csv')
                result['data'].to_csv(output_file, index=False)
                result['output_file'] = output_file
                result['success'] = True
                
                # Perform validation if targets exist
                if sheet_type in VALIDATION_TARGETS:
                    result['validation'] = self._validate_sheet_data(result['data'], sheet_type)
                
                # Create summary
                result['summary'] = create_summary_report(
                    sheet_type, result['data'], result['validation']
                )
                
                self.logger.info(f"{sheet_type} processing completed successfully")
            else:
                result['error'] = f"No data extracted from {sheet_type} sheet"
                self.logger.warning(result['error'])
                
        except Exception as e:
            result['error'] = str(e)
            self.logger.error(f"Error processing {sheet_type}: {e}")
            
        return result
    
    def process_all_sheets(self, file_path: str, selected_sheets: List[str] = None) -> Dict:
        """Process all selected sheets"""
        # Validate file first
        is_valid, missing_sheets = self.validate_excel_file(file_path)
        
        if not is_valid and missing_sheets:
            self.logger.error(f"File validation failed: {missing_sheets}")
            return {'error': f"Missing required sheets: {missing_sheets}"}
        
        # Process selected sheets (or all if none specified)
        if selected_sheets is None:
            selected_sheets = ['STOPING', 'TRAMMING', 'DEVELOPMENT', 'HOISTING', 'PLANT', 'BENCHES']
        
        results = {
            'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'file_path': file_path,
            'sheets_processed': {},
            'overall_summary': {}
        }
        
        total_records = 0
        successful_sheets = 0
        
        for sheet_type in selected_sheets:
            self.logger.info(f"Starting {sheet_type} processing...")
            sheet_result = self.process_single_sheet(file_path, sheet_type)
            results['sheets_processed'][sheet_type] = sheet_result
            
            if sheet_result['success']:
                successful_sheets += 1
                total_records += len(sheet_result['data'])
        
        # Generate overall summary
        results['overall_summary'] = {
            'total_sheets_requested': len(selected_sheets),
            'successful_sheets': successful_sheets,
            'total_records': total_records,
            'output_directory': self.output_dir
        }
        
        # Save master summary report
        self._save_master_summary(results)
        
        self.logger.info(f"Processing complete. {successful_sheets}/{len(selected_sheets)} sheets processed successfully")
        
        return results
    
    def _validate_sheet_data(self, df: pd.DataFrame, sheet_type: str) -> Dict:
        """Validate extracted data against known targets"""
        if sheet_type not in VALIDATION_TARGETS:
            return {'message': 'No validation targets defined for this sheet'}
        
        targets = VALIDATION_TARGETS[sheet_type]
        
        # Calculate statistics based on sheet type
        if sheet_type in ['STOPING', 'TRAMMING']:
            stats = {
                f'tonnes_actual_target': df[df.columns[2]].astype(float).sum(),  # Actual tonnes
                f'tonnes_budget_target': df[df.columns[5]].astype(float).sum(),  # Budget tonnes
                f'gold_actual_target': df[df.columns[4]].astype(float).sum(),    # Actual gold
                f'gold_budget_target': df[df.columns[7]].astype(float).sum()     # Budget gold
            }
        elif sheet_type == 'DEVELOPMENT':
            stats = {
                'budget_metres_target': df['Budget_Metres'].astype(float).sum(),
                'actual_metres_target': df['Actual_Metres'].astype(float).sum()
            }
        elif sheet_type == 'BENCHES':
            stats = {
                'total_samples': len(df),
                'qaqc_samples': df['is_qaqc'].sum() if 'is_qaqc' in df.columns else 0
            }
        else:
            stats = {'total_records': len(df)}
        
        return validate_against_targets(stats, targets)
    
    def _save_master_summary(self, results: Dict):
        """Save comprehensive master summary report"""
        summary_file = os.path.join(self.output_dir, 'mining_extraction_master_report.txt')
        
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("MINING DAILY REPORT - MASTER EXTRACTION SUMMARY\n")
            f.write("=" * 60 + "\n\n")
            
            f.write(f"Processing Time: {results['processing_time']}\n")
            f.write(f"Input File: {results['file_path']}\n")
            f.write(f"Output Directory: {self.output_dir}\n\n")
            
            f.write("PROCESSING RESULTS:\n")
            f.write("-" * 30 + "\n")
            
            for sheet_type, sheet_result in results['sheets_processed'].items():
                status = "[SUCCESS]" if sheet_result['success'] else "[FAILED]"
                f.write(f"{sheet_type}: {status}\n")
                
                if sheet_result['success']:
                    f.write(f"  Records: {len(sheet_result['data'])}\n")
                    f.write(f"  Output: {sheet_result.get('output_file', 'N/A')}\n")
                    
                    if sheet_result.get('validation'):
                        val_result = sheet_result['validation']
                        f.write(f"  Validation: {val_result['passed']} passed, {val_result['failed']} failed\n")
                else:
                    f.write(f"  Error: {sheet_result.get('error', 'Unknown error')}\n")
                
                f.write("\n")
            
            f.write("OVERALL SUMMARY:\n")
            f.write("-" * 20 + "\n")
            overall = results['overall_summary']
            f.write(f"Sheets Requested: {overall['total_sheets_requested']}\n")
            f.write(f"Sheets Successful: {overall['successful_sheets']}\n")
            f.write(f"Total Records: {overall['total_records']}\n")
            
            f.write("\nOUTPUT FILES:\n")
            f.write("-" * 15 + "\n")
            for sheet_type, sheet_result in results['sheets_processed'].items():
                if sheet_result['success']:
                    f.write(f"- {sheet_type.lower()}_data.csv\n")
                    if sheet_type == 'BENCHES':
                        f.write(f"- benches_raw_data.csv\n")
                        f.write(f"- benches_average_grades.csv\n")
        
        self.logger.info(f"Master summary saved to: {summary_file}")

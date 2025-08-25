
"""
BENCHES Sheet Data Processor
Based on the corrected benches processing logic with forward fill and QAQC handling
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Tuple, List, Dict
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.common import clean_and_validate_data, setup_logger, is_qaqc_sample

class BenchesProcessor:
    """Processor for BENCHES sheet data extraction with forward fill and QAQC"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('BenchesProcessor', 'logs/benches.log')
    
    def extract_benches_data(self, file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Extract benches data with proper forward fill and QAQC handling
        Returns: (raw_data_df, average_grades_df)
        """
        self.logger.info("Loading BENCHES sheet...")
        
        try:
            df = pd.read_excel(file_path, sheet_name='BENCHES.')
        except Exception as e:
            self.logger.error(f"Failed to load BENCHES sheet: {e}")
            return pd.DataFrame(), pd.DataFrame()
        
        self.logger.info(f"Original data shape: {df.shape}")
        
        # Step 1: Process the data with forward fill
        df_processed = self._apply_forward_fill(df.copy())
        
        # Step 2: QAQC Sample Identification
        df_processed['is_qaqc'] = df_processed['DEPTH (m)'].apply(is_qaqc_sample)
        
        qaqc_count = df_processed['is_qaqc'].sum()
        total_samples = len(df_processed)
        
        self.logger.info(f"Total samples: {total_samples}")
        self.logger.info(f"QAQC samples identified: {qaqc_count}")
        self.logger.info(f"Regular samples: {total_samples - qaqc_count}")
        
        # Step 3: Create ID column for grouping (CORRECTED with 5 components)
        df_processed['ID'] = self._create_comprehensive_id(df_processed)
        
        # Step 4: Save Raw Data
        df_raw = df_processed.copy()
        
        # Step 5: Calculate Average Grades (excluding QAQC)
        avg_grades_df = self._calculate_average_grades(df_processed)
        
        self.logger.info(f"Raw data records: {len(df_raw)}")
        self.logger.info(f"Average grades records: {len(avg_grades_df)}")
        
        return df_raw, avg_grades_df
    
    def _apply_forward_fill(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply forward fill logic as per corrected version"""
        self.logger.info("Applying forward fill...")
        
        # Step 1: Forward fill DATE first
        df['DATE'] = df['DATE'].ffill()
        
        # Step 2: Forward fill LEVEL
        df['LEVEL'] = df['LEVEL'].ffill()
        
        # Step 3: Forward fill SECTION
        df['SECTION'] = df['SECTION'].ffill()
        
        # Step 4: Forward fill DIRECTION (CORRECTED - was missing)
        df['DIRECTION'] = df['DIRECTION'].ffill()
        
        # Step 5: Conditional forward fill for other columns
        # Create a grouping key based on DATE, LEVEL, SECTION, DIRECTION
        df['group_key'] = (
            df['DATE'].astype(str) + '_' + 
            df['LEVEL'].astype(str) + '_' + 
            df['SECTION'].astype(str) + '_' +
            df['DIRECTION'].astype(str)
        )
        
        # Identify columns for conditional forward fill
        exclude_cols = ['DATE', 'LEVEL', 'SECTION', 'DIRECTION', 'group_key', 
                       'SAMPLE ID', 'AU', 'DEPTH (m)', 'FROM', 'TO']
        fill_cols = [col for col in df.columns 
                    if col not in exclude_cols and not col.startswith('Unnamed:')]
        
        # Apply conditional forward fill within each group
        for col in fill_cols:
            if col in df.columns:
                df[col] = df.groupby('group_key')[col].ffill()
        
        # Also fill Unnamed: 3 (Column D) as it's part of the ID structure
        if 'Unnamed: 3' in df.columns:
            df['Unnamed: 3'] = df.groupby('group_key')['Unnamed: 3'].ffill()
        
        # Remove temporary group key
        df = df.drop('group_key', axis=1)
        
        return df
    
    def _create_comprehensive_id(self, df: pd.DataFrame) -> pd.Series:
        """Create CORRECTED ID using all 5 components"""
        self.logger.info("Creating comprehensive ID with 5 components...")
        
        # CORRECTION: Using LEVEL + SECTION + Unnamed: 3 + LOCATION + DIRECTION
        id_series = (
            df['LEVEL'].astype(str) + '_' + 
            df['SECTION'].astype(str) + '_' +
            df['Unnamed: 3'].fillna('').astype(str) + '_' +
            df['LOCATION'].fillna('').astype(str) + '_' +
            df['DIRECTION'].fillna('').astype(str)
        )
        
        # Clean up the ID (remove extra underscores)
        id_series = id_series.str.replace('_+', '_', regex=True).str.strip('_')
        
        unique_ids = id_series.nunique()
        self.logger.info(f"Created {unique_ids} unique IDs")
        
        return id_series
    
    def _calculate_average_grades(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate average grades excluding QAQC samples"""
        self.logger.info("Calculating average grades...")
        
        # Filter out QAQC samples for grade calculations
        df_regular = df[~df['is_qaqc']].copy()
        
        self.logger.info(f"Regular samples for averaging: {len(df_regular)}")
        
        # Convert AU to numeric
        df_regular['AU_numeric'] = pd.to_numeric(df_regular['AU'], errors='coerce')
        
        # Group by DATE and ID, calculate averages
        avg_grades = df_regular.groupby(['DATE', 'ID']).agg({
            'LEVEL': 'first',
            'SECTION': 'first',
            'Unnamed: 3': 'first',
            'LOCATION': 'first',
            'DIRECTION': 'first',  # CORRECTED: Include DIRECTION
            'SAMPLE TYPE': 'first',
            'PEG ID': 'first',
            'AU_numeric': ['mean', 'count'],
            'SAMPLE ID': 'count'
        }).reset_index()
        
        # Flatten column names
        avg_grades.columns = ['DATE', 'ID', 'LEVEL', 'SECTION', 'Column_D', 'LOCATION', 
                             'DIRECTION', 'SAMPLE_TYPE', 'PEG_ID', 'AU_avg', 'AU_count', 'Sample_count']
        
        # Round averages
        avg_grades['AU_avg'] = avg_grades['AU_avg'].round(4)
        
        return avg_grades

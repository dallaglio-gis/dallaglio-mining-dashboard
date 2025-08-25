
"""
HOISTING Sheet Data Processor
Handles the unique Source/METRIC1/METRIC2/Value format
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Tuple, List, Dict
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.common import clean_and_validate_data, setup_logger, detect_date_columns

class HoistingProcessor:
    """Processor for HOISTING sheet data extraction"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('HoistingProcessor', 'logs/hoisting.log')
    
    def extract_hoisting_data(self, file_path: str) -> pd.DataFrame:
        """Extract hoisting data in Source/METRIC1/METRIC2/Value format"""
        self.logger.info("Loading HOISTING sheet...")
        
        try:
            df = pd.read_excel(file_path, sheet_name='Hoisting')
        except Exception as e:
            self.logger.error(f"Failed to load HOISTING sheet: {e}")
            return pd.DataFrame()
        
        # Find date columns
        date_row = 15
        dates, date_columns = detect_date_columns(df, date_row, 7)
        
        if not dates:
            self.logger.error("No dates found in HOISTING sheet")
            return pd.DataFrame()
        
        self.logger.info(f"Found {len(dates)} dates from {dates[0]} to {dates[-1]}")
        
        all_data = []
        
        # Find data rows by looking for valid sources
        for i in range(len(df)):
            if len(df.columns) > 4:
                source = clean_and_validate_data(df.iloc[i, 2], 'text') if i < len(df) else ""

                # Only consider valid source names. Exclude numeric values, dates,
                # or empty placeholders. A valid source should contain at least
                # one alphabetic character.
                if (source and
                    source.lower() not in ['source', 'nan', '', '0'] and
                    any(ch.isalpha() for ch in source)):
                    metric1 = clean_and_validate_data(df.iloc[i, 3], 'text') if i < len(df) and len(df.columns) > 3 else ""
                    metric2 = clean_and_validate_data(df.iloc[i, 4], 'text') if i < len(df) and len(df.columns) > 4 else ""

                    # Extract daily values across dates
                    for date_idx, col_idx in enumerate(date_columns):
                        if col_idx < len(df.columns):
                            value = clean_and_validate_data(
                                df.iloc[i, col_idx] if i < len(df) else 0
                            )
                            record = {
                                'Date': dates[date_idx],
                                'Source': source,
                                'METRIC1': metric1,
                                'METRIC2': metric2,
                                'Value': value
                            }
                            all_data.append(record)
        
        # Create DataFrame
        result_df = pd.DataFrame(all_data)
        
        if not result_df.empty:
            result_df = result_df.sort_values(['Date', 'Source', 'METRIC1', 'METRIC2'])
            # Remove rows with empty values to clean up the data
            result_df = result_df[result_df['Value'] != 0]
            self.logger.info(f"Total hoisting records created: {len(result_df)}")
        
        return result_df

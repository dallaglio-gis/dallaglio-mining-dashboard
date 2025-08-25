
"""
STOPING Sheet Data Processor
Based on the final working extraction script
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

class StopingProcessor:
    """Processor for STOPING sheet data extraction"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('StopingProcessor', 'logs/stoping.log')
        self.stopes_with_missing_budget = []
        self.stopes_with_partial_data = []
    
    def extract_stoping_data(self, file_path: str) -> pd.DataFrame:
        """
        Extract all stoping data including stopes with missing budget data
        Based on extract_all_stopes_final.py logic
        """
        self.logger.info("Loading STOPING sheet...")
        
        try:
            stoping_df = pd.read_excel(file_path, sheet_name='Stoping')
        except Exception as e:
            self.logger.error(f"Failed to load STOPING sheet: {e}")
            return pd.DataFrame()
        
        # Find date columns (starting from column 7)
        date_row = 15  # Row with dates
        dates, date_columns = detect_date_columns(stoping_df, date_row, 7)
        
        if not dates:
            self.logger.error("No dates found in STOPING sheet")
            return pd.DataFrame()
        
        self.logger.info(f"Found {len(dates)} dates from {dates[0]} to {dates[-1]}")
        
        # Find all stope entries
        stope_entries = self._find_stope_entries(stoping_df)
        self.logger.info(f"Found {len(stope_entries)} stope entries")
        
        # Extract data for each stope
        all_data = []
        
        for stope_idx, (start_row, stope_name) in enumerate(stope_entries):
            self.logger.info(f"Processing {stope_name}...")
            
            stope_data = self._extract_stope_data(
                stoping_df, start_row, stope_name, stope_idx, 
                stope_entries, dates, date_columns
            )
            
            if stope_data:
                all_data.extend(stope_data)
        
        # Create DataFrame
        df = pd.DataFrame(all_data)
        
        if not df.empty:
            # Sort by Date and ID
            df = df.sort_values(['Date', 'ID'])
            
            self.logger.info(f"Extraction Summary:")
            self.logger.info(f"Total stopes processed: {len(stope_entries)}")
            self.logger.info(f"Stopes with missing budget data: {len(self.stopes_with_missing_budget)}")
            self.logger.info(f"Stopes with partial data: {len(self.stopes_with_partial_data)}")
            self.logger.info(f"Total daily records created: {len(df)}")
        
        return df
    
    def _find_stope_entries(self, df: pd.DataFrame) -> List[Tuple[int, str]]:
        """Find all stope entries in the sheet"""
        stope_entries = []
        
        for i in range(len(df)):
            if len(df.columns) > 2:
                stope_name = df.iloc[i, 2] if i < len(df) else None
                
                if (pd.notna(stope_name) and 
                    'STOPE' in str(stope_name).upper() and 
                    stope_name != 'Stopes'):
                    stope_entries.append((i, stope_name))
        
        return stope_entries
    
    def _extract_stope_data(self, df: pd.DataFrame, start_row: int, stope_name: str,
                           stope_idx: int, stope_entries: List, dates: List,
                           date_columns: List) -> List[Dict]:
        """Extract data for a single stope.

        This method has been updated to robustly handle cases where some metric
        rows (e.g. Tonnes budget/actual) appear a few rows before the stope
        name rather than immediately after. For each metric (tonnes, grade,
        gold) and each data type (budget, actual), we search within a
        configurable window around the stope row and select the row that is
        closest to the stope name. This prevents mis-assignment of data when
        rows are slightly out of order, as occurred with the 6L_W18_STOPE
        section in the August 2025 daily report.
        """

        # Standardize the stope name and prepare structure
        stope_name_clean = stope_name.strip()
        stope_data = {
            'stope_name': stope_name_clean,
            'daily_data': {}
        }

        # Determine the end row for this stope (start of next stope or end of sheet)
        if stope_idx < len(stope_entries) - 1:
            end_row = stope_entries[stope_idx + 1][0]
        else:
            end_row = len(df)

        # Define a window around the stope row to search for metric rows. We look
        # backwards a few rows in case the budget/actual rows are placed before
        # the stope name (as observed for 6L_W18_STOPE). We limit the forward
        # search to avoid accidentally capturing rows belonging to the next stope.
        window_start = max(0, start_row - 5)
        window_end = min(start_row + 15, end_row)

        # Candidate storage for the nearest row index of each metric/data type
        # We'll record the row index and the absolute distance to the stope row
        metric_candidates = {
            ('tonnes', 'budget'): None,
            ('tonnes', 'actual'): None,
            ('grade', 'budget'): None,
            ('grade', 'actual'): None,
            ('gold', 'budget'): None,
            ('gold', 'actual'): None,
        }

        # Iterate through the candidate window and evaluate each row
        for row_idx in range(window_start, window_end):
            if row_idx >= len(df):
                break

            # Read potential metric and type values. We try to extract from
            # column 3 (metric) and column 4 (type). In cases where the metric
            # cell is blank (e.g. actual rows), we infer the metric by looking
            # back up to a few rows.
            col3_val = df.iloc[row_idx, 3] if len(df.columns) > 3 else None
            col4_val = df.iloc[row_idx, 4] if len(df.columns) > 4 else None

            metric_type = None
            data_type = None

            if pd.notna(col3_val) and pd.notna(col4_val):
                metric_type = str(col3_val).lower().strip()
                data_type = str(col4_val).lower().strip()
            elif pd.isna(col3_val) and pd.notna(col4_val):
                # If metric cell is blank but type is present, infer metric
                data_type = str(col4_val).lower().strip()
                # Look backwards up to 5 rows to find a non-null metric
                for prev_row in range(row_idx - 1, max(row_idx - 5, window_start) - 1, -1):
                    prev_col3 = df.iloc[prev_row, 3] if len(df.columns) > 3 else None
                    if pd.notna(prev_col3):
                        metric_type = str(prev_col3).lower().strip()
                        break

            # Skip rows without a valid metric/data type
            if not (metric_type and data_type):
                continue

            # Exclude explanatory rows or derived metrics
            lower_type = data_type.lower()
            if 'reason' in lower_type or 'variance' in lower_type:
                continue

            # Determine high-level metric category and budget/actual indicator
            category = None
            if 'tonnes' in metric_type or 'ton' in metric_type:
                category = 'tonnes'
            elif 'grade' in metric_type:
                category = 'grade'
            elif 'gold' in metric_type:
                category = 'gold'

            indicator = None
            if 'budget' in data_type:
                indicator = 'budget'
            elif 'actual' in data_type:
                indicator = 'actual'

            if not (category and indicator):
                continue

            # For non-tonnage metrics (grade/gold), only consider rows that appear
            # at or after the stope row. This prevents mis-assigning budget/actual
            # rows from a previous stope (e.g. 1L W28) to the current stope. For
            # tonnage metrics, we allow rows slightly above the stope row to
            # accommodate cases like 6L_W18_STOPE where the tonnage row appears
            # before the stope name.
            if category != 'tonnes' and row_idx < start_row:
                continue

            # Compute distance to stope row
            distance = abs(row_idx - start_row)

            # If we haven't seen this metric yet or found a closer occurrence, record it
            key = (category, indicator)
            current_candidate = metric_candidates.get(key)
            if current_candidate is None or distance < current_candidate[1]:
                metric_candidates[key] = (row_idx, distance)

        # After scanning, extract values from the selected metric rows
        # Track whether we observed any budget or actual values
        has_budget_data = False
        has_actual_data = False

        for (category, indicator), candidate in metric_candidates.items():
            if candidate is None:
                continue
            row_idx, _ = candidate
            # Extract the row's data across date columns
            for date_idx, col_idx in enumerate(date_columns):
                if col_idx < len(df.columns) and row_idx < len(df):
                    daily_val = df.iloc[row_idx, col_idx]
                    if pd.notna(daily_val) and daily_val != 0:
                        cleaned_val = clean_and_validate_data(daily_val)
                        date = dates[date_idx]
                        if date not in stope_data['daily_data']:
                            stope_data['daily_data'][date] = {}

                        key_name = f"{category}_{indicator}"
                        stope_data['daily_data'][date][key_name] = cleaned_val
                        if indicator == 'budget':
                            has_budget_data = True
                        elif indicator == 'actual':
                            has_actual_data = True

        # If no data collected, log and return empty list
        if not stope_data['daily_data']:
            self.logger.info(f"  -> No daily data found for {stope_name_clean}")
            return []

        # Determine completeness for reporting
        if has_actual_data and not has_budget_data:
            self.stopes_with_missing_budget.append(stope_name_clean)
            self.logger.info(f"  -> Missing budget data (new stope)")
        elif has_actual_data and has_budget_data:
            # Check if budgets are zero across all dates
            budget_values = []
            for date_data in stope_data['daily_data'].values():
                for key, val in date_data.items():
                    if key.endswith('_budget'):
                        budget_values.append(val)
            if budget_values and all(v == 0 for v in budget_values):
                self.stopes_with_partial_data.append(stope_name_clean)
                self.logger.info(f"  -> Partial data (zero budgets)")
            else:
                self.logger.info(f"  -> Complete data")

        # Convert collected data into records for each date
        records = []
        for date in sorted(stope_data['daily_data'].keys()):
            daily_data = stope_data['daily_data'][date]
            record = {
                'Date': date,
                'ID': stope_data['stope_name'],
                'Stoping_Actual_t': daily_data.get('tonnes_actual', ''),
                'Stoping_Actual_gpt': daily_data.get('grade_actual', ''),
                'Stoping_Actual_kgs': daily_data.get('gold_actual', ''),
                'Stoping_Budget_t': daily_data.get('tonnes_budget', ''),
                'Stoping_Budget_gpt': daily_data.get('grade_budget', ''),
                'Stoping_Budget_kgs': daily_data.get('gold_budget', '')
            }
            records.append(record)

        return records

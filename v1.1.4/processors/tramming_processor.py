
"""
TRAMMING Sheet Data Processor
Similar logic to STOPING but for tramming operations
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Tuple, List, Dict, Optional
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.common import clean_and_validate_data, setup_logger, detect_date_columns

class TrammingProcessor:
    """Processor for TRAMMING sheet data extraction"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('TrammingProcessor', 'logs/tramming.log')
    
    def extract_tramming_data(self, file_path: str) -> pd.DataFrame:
        """Extract tramming data following similar logic to stoping"""
        self.logger.info("Loading TRAMMING sheet...")
        
        try:
            tramming_df = pd.read_excel(file_path, sheet_name='Tramming')
        except Exception as e:
            self.logger.error(f"Failed to load TRAMMING sheet: {e}")
            return pd.DataFrame()
        
        # Find date columns
        date_row = 15
        dates, date_columns = detect_date_columns(tramming_df, date_row, 7)
        
        if not dates:
            self.logger.error("No dates found in TRAMMING sheet")
            return pd.DataFrame()
        
        self.logger.info(f"Found {len(dates)} dates from {dates[0]} to {dates[-1]}")
        
        # Find all tramming entries (similar to stopes but look for different patterns)
        tramming_entries = self._find_tramming_entries(tramming_df)
        self.logger.info(f"Found {len(tramming_entries)} tramming entries")
        
        # Extract data for each entry
        all_data = []
        
        for entry_idx, (start_row, entry_name) in enumerate(tramming_entries):
            self.logger.info(f"Processing {entry_name}...")
            
            entry_data = self._extract_tramming_entry_data(
                tramming_df, start_row, entry_name, entry_idx,
                tramming_entries, dates, date_columns
            )
            
            if entry_data:
                all_data.extend(entry_data)
        
        # Create DataFrame
        df = pd.DataFrame(all_data)
        
        if not df.empty:
            df = df.sort_values(['Date', 'ID'])
            self.logger.info(f"Total daily records created: {len(df)}")
        
        return df
    
    def _find_tramming_entries(self, df: pd.DataFrame) -> List[Tuple[int, str]]:
        """Find all tramming entries in the sheet.

        The original implementation relied on heuristics such as looking for
        'STOPE', 'BOX', or strings containing both 'L' and 'W' to identify
        tramming entries. However, this could miss valid entries or include
        header rows like 'BOX ID'. To make the detection more thorough, we
        instead look for rows where both the priority column (index 1) and
        the ID column (index 2) are populated. We skip obvious headers or
        summary rows based on keywords.
        """
        entries: List[Tuple[int, str]] = []

        # Start scanning from the date header row onwards (row 15) to avoid
        # picking up preamble information. Adjust if the sheet structure changes.
        start_scan_row = 15

        for i in range(start_scan_row, len(df)):
            if len(df.columns) > 2:
                col1_val = df.iloc[i, 1] if i < len(df) else None  # Priority number or grouping
                col2_val = df.iloc[i, 2] if i < len(df) else None  # Box/stope identifier

                # Both priority and ID must be present
                if pd.isna(col1_val) or pd.isna(col2_val):
                    continue

                # Convert to string for analysis
                col2_str = str(col2_val).strip()
                col1_str = str(col1_val).strip()

                # Skip rows with obvious header text or placeholders
                skip_terms = ['box id', 'priority', 'budget', 'actual', 'variance', 'reason']
                if any(term in col2_str.lower() for term in skip_terms):
                    continue

                # Ensure the priority value looks like a small number (e.g., 1, 2, 1.2)
                if not (col1_str.replace('.', '', 1).isdigit() or len(col1_str) <= 3):
                    continue

                # If we reach here, treat this as a valid tramming entry
                entries.append((i, col2_str))

        return entries
    
    def _extract_tramming_entry_data(self, df: pd.DataFrame, start_row: int, 
                                   entry_name: str, entry_idx: int, 
                                   entries: List, dates: List, 
                                   date_columns: List) -> List[Dict]:
        """Extract data for a single tramming entry.

        Similar to the stoping extraction, we select the nearest valid metric row
        for each category (tonnes, grade, gold) and data type (budget, actual)
        while avoiding mis-assignment of values from neighbouring entries.
        For tramming, the metric rows always appear at or after the entry
        name, so we restrict our search to forward rows only.
        """

        # Clean the entry name and prepare the data structure
        entry_name_clean = entry_name.strip()
        entry_data = {
            'entry_name': entry_name_clean,
            'daily_data': {}
        }

        # Determine the end row for this entry
        if entry_idx < len(entries) - 1:
            end_row = entries[entry_idx + 1][0]
        else:
            end_row = len(df)

        # Define a forward search window; we do not look backwards for tramming
        window_start = start_row
        window_end = min(start_row + 15, end_row)

        # Candidate rows for each metric and indicator
        metric_candidates = {
            ('tonnes', 'budget'): None,
            ('tonnes', 'actual'): None,
            ('grade', 'budget'): None,
            ('grade', 'actual'): None,
            ('gold', 'budget'): None,
            ('gold', 'actual'): None,
        }

        # Iterate through the candidate window
        for row_idx in range(window_start, window_end):
            if row_idx >= len(df):
                break

            # Extract potential metric and type from the row
            col3_val = df.iloc[row_idx, 3] if len(df.columns) > 3 else None  # Metric (Tonnes/Grade/Gold)
            col4_val = df.iloc[row_idx, 4] if len(df.columns) > 4 else None  # Data type (Budget/Actual)

            metric_type: Optional[str] = None
            data_type: Optional[str] = None

            # If both metric and data type columns are populated
            if pd.notna(col3_val) and pd.notna(col4_val):
                metric_type = str(col3_val).lower().strip()
                data_type = str(col4_val).lower().strip()
            elif pd.notna(col4_val):
                # Only data type column is populated; parse metric from the unit in data_type
                data_type = str(col4_val).lower().strip()
                # Attempt to infer metric based on the unit in the data type
                if 'g/t' in data_type:
                    metric_type = 'grade'
                elif 'kg' in data_type:
                    metric_type = 'gold'
                elif 't' in data_type:
                    # Ensure we are not mis-classifying g/t as tonnes
                    metric_type = 'tonnes'
            elif pd.notna(col3_val):
                # Only metric column is populated; carry forward to next rows when data type appears
                metric_type = str(col3_val).lower().strip()

            # If we still cannot determine metric or data type, skip the row
            if not data_type or not metric_type:
                continue

            # Skip explanatory or variance rows
            lower_type = data_type.lower()
            if 'reason' in lower_type or 'variance' in lower_type:
                continue

            # Determine metric category based on parsed metric_type
            category: Optional[str] = None
            if 'tonnes' in metric_type or 'ton' in metric_type:
                category = 'tonnes'
            elif 'grade' in metric_type:
                category = 'grade'
            elif 'gold' in metric_type:
                category = 'gold'

            # Determine indicator (budget/actual) from data_type
            indicator: Optional[str] = None
            if 'budget' in data_type:
                indicator = 'budget'
            elif 'actual' in data_type:
                indicator = 'actual'

            # Skip if category or indicator not identified
            if not category or not indicator:
                continue

            # Compute distance from start_row
            distance = row_idx - start_row
            # Only keep the first occurrence (smallest distance) for each metric
            key = (category, indicator)
            current_candidate = metric_candidates.get(key)
            if current_candidate is None or distance < current_candidate[1]:
                metric_candidates[key] = (row_idx, distance)

        # Extract data from the selected metric rows
        for (category, indicator), candidate in metric_candidates.items():
            if candidate is None:
                continue
            row_idx, _ = candidate
            for date_idx, col_idx in enumerate(date_columns):
                if col_idx < len(df.columns) and row_idx < len(df):
                    daily_val = df.iloc[row_idx, col_idx]
                    # Even if the value is zero, record it to ensure completeness
                    if pd.notna(daily_val):
                        cleaned_val = clean_and_validate_data(daily_val)
                        date = dates[date_idx]
                        if date not in entry_data['daily_data']:
                            entry_data['daily_data'][date] = {}
                        key_name = f"{category}_{indicator}"
                        entry_data['daily_data'][date][key_name] = cleaned_val

        if not entry_data['daily_data']:
            return []

        # Build records
        records = []
        for date in sorted(entry_data['daily_data'].keys()):
            daily_data = entry_data['daily_data'][date]
            records.append({
                'Date': date,
                'ID': entry_data['entry_name'],
                'Tramming_Actual_t': daily_data.get('tonnes_actual', ''),
                'Tramming_Actual_gpt': daily_data.get('grade_actual', ''),
                'Tramming_Actual_kgs': daily_data.get('gold_actual', ''),
                'Tramming_Budget_t': daily_data.get('tonnes_budget', ''),
                'Tramming_Budget_gpt': daily_data.get('grade_budget', ''),
                'Tramming_Budget_kgs': daily_data.get('gold_budget', '')
            })

        return records

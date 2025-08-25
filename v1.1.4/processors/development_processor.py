
"""
DEVELOPMENT Sheet Data Processor — thorough, block-aware, duplicate-safe
"""

import pandas as pd
import numpy as np
from typing import Tuple, List
import os, sys, re

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.common import clean_and_validate_data, setup_logger, detect_date_columns

class DevelopmentProcessor:
    """Processor for DEVELOPMENT sheet data extraction (robust)"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('DevelopmentProcessor', 'logs/development.log')
    
    def extract_development_data(self, file_path: str) -> pd.DataFrame:
        """Extract development data with robust entry detection, metric row scanning and duplicate aggregation"""
        self.logger.info("Loading DEVELOPMENT sheet...")
        
        try:
            df = pd.read_excel(file_path, sheet_name='Development')
            self.logger.info(f"Loaded sheet 'Development': {df.shape[0]} rows x {df.shape[1]} cols")
        except Exception as e:
            self.logger.error(f"Failed to load DEVELOPMENT sheet: {e}")
            return pd.DataFrame()
        
        # Detect date columns (row 16, starting from col 7 in provided template)
        date_row = 16
        self.logger.info(f"Detecting dates at row index {date_row}")
        dates, date_cols = detect_date_columns(df, date_row=date_row, start_col=7)
        if not dates:
            self.logger.error("No dates found in DEVELOPMENT sheet")
            return pd.DataFrame()
        self.logger.info(f"Found {len(dates)} dates: {dates[0]} .. {dates[-1]}")
        
        # Find entries (dev ends)
        entries = self._find_development_entries(df)
        self.logger.info(f"Development entries detected: {len(entries)}")
        
        # Helpers
        def norm_id(s: str) -> str:
            s = str(s)
            s = s.replace("\t", " ").replace("\n", " ")
            s = re.sub(r"\s+", " ", s).strip()
            return s
        
        def find_metric_rows(start_row: int, end_row: int) -> Tuple[int, int]:
            """Find the nearest 'Budget (m)' and 'Actual (m)' rows below start_row, within a window"""
            budget_row, actual_row = None, None
            window_end = min(end_row, start_row + 15)
            for r in range(start_row, window_end):
                label = df.iloc[r, 4] if 4 < df.shape[1] else None
                if pd.isna(label):
                    continue
                t = str(label).lower().strip()
                if budget_row is None and ("budget" in t and "(m" in t):  # '(m' tolerates '(m)' or '(m]'
                    budget_row = r
                elif actual_row is None and ("actual" in t and "(m" in t):
                    actual_row = r
                if budget_row is not None and actual_row is not None:
                    break
            return budget_row, actual_row
        
        # Extract
        records: List[dict] = []
        for idx, (start_row, dev_id_raw) in enumerate(entries):
            end_row = entries[idx+1][0] if idx < len(entries) - 1 else len(df)
            dev_id = norm_id(dev_id_raw)
            
            budget_row, actual_row = find_metric_rows(start_row, end_row)
            # If BOTH metric rows are missing, skip. Otherwise zero-fill the missing side.
            if budget_row is None and actual_row is None:
                self.logger.warning(
                    f"Skipping '{dev_id}' - no Budget/Actual rows between {start_row} and {end_row}"
                )
                continue
            
            # MTD values at col 6 (zero-fill if one side is missing)
            mtd_budget = clean_and_validate_data(df.iloc[budget_row, 6]) if budget_row is not None else 0
            mtd_actual = clean_and_validate_data(df.iloc[actual_row, 6]) if actual_row is not None else 0
            
            # Daily values (zero-fill if one side is missing)
            for d_idx, col_idx in enumerate(date_cols):
                if col_idx < df.shape[1]:
                    daily_budget = clean_and_validate_data(df.iloc[budget_row, col_idx]) if budget_row is not None else 0
                    daily_actual = clean_and_validate_data(df.iloc[actual_row, col_idx]) if actual_row is not None else 0
                    records.append({
                        "Date": dates[d_idx],
                        "Dev_ID": dev_id,
                        "Budget_Metres": daily_budget,
                        "Actual_Metres": daily_actual,
                        "MTD_Budget": mtd_budget,
                        "MTD_Actual": mtd_actual,
                    })
        
        if not records:
            return pd.DataFrame()
        
        out = pd.DataFrame.from_records(records)
        
        # Aggregate duplicates (same Dev_ID in multiple blocks): sum daily metres, keep max MTDs
        out = (
            out.groupby(["Date", "Dev_ID"], as_index=False)
               .agg({
                   "Budget_Metres": "sum",
                   "Actual_Metres": "sum",
                   "MTD_Budget": "max",
                   "MTD_Actual": "max",
               })
               .sort_values(["Date", "Dev_ID"])
        )
        
        self.logger.info(f"Final development rows: {len(out)} for {out['Dev_ID'].nunique()} ends")
        return out
    
    def _find_development_entries(self, df: pd.DataFrame) -> List[Tuple[int, str]]:
        """Scan the whole sheet and pick rows that look like development ends (robust)"""
        entries: List[Tuple[int, str]] = []
        start_row = 17
        skip_terms = {"priority", "development_ends", "tonnes", "grade", "gold", "variance", "reason"}
        
        for i in range(start_row, len(df)):
            if df.shape[1] <= 2:
                break
            name = df.iloc[i, 2]
            if pd.isna(name):
                continue
            s = str(name).strip()
            if not s or any(t in s.lower() for t in skip_terms):
                continue
            
            # Check that a Budget/Actual pair exists shortly below
            has_budget, has_actual = False, False
            window_end = min(i + 15, len(df))
            for r in range(i, window_end):
                label = df.iloc[r, 4] if 4 < df.shape[1] else None
                if pd.isna(label): 
                    continue
                t = str(label).lower().strip()
                if "budget" in t and "(m" in t:
                    has_budget = True
                elif "actual" in t and "(m" in t:
                    has_actual = True
                if has_budget or has_actual:
                    entries.append((i, s))
                    break
        
        return entries

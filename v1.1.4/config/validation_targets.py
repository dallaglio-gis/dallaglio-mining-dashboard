
"""
Validation targets and processing configurations for all sheet types
"""

# Validation targets from your previous extractions
VALIDATION_TARGETS = {
    'STOPING': {
        'tonnes_budget_target': 31746,  # Based on your previous validations
        'tonnes_actual_target': 26928,
        'gold_budget_target': 68.1,
        'gold_actual_target': 62.8,
        'grade_budget_target': 2.14,
        'grade_actual_target': 2.33,
        'tolerance': 0.02  # 2% tolerance
    },
    'TRAMMING': {
        'tonnes_budget_target': 31746.5,
        'tonnes_actual_target': 26927.7,
        'gold_budget_target': 68.1,
        'gold_actual_target': 62.8,
        'grade_budget_target': 2.163,
        'grade_actual_target': 2.349,
        'tolerance': 0.02
    },
    'DEVELOPMENT': {
        'budget_metres_target': 800,  # Approximate based on data
        'actual_metres_target': 750,
        'tolerance': 0.02
    },
    'HOISTING': {
        'tolerance': 0.02
    },
    'BENCHES': {
        'min_samples': 1700,
        'min_average_grades': 300,
        'qaqc_percentage_min': 10,
        'qaqc_percentage_max': 20,
        'tolerance': 0.02
    }
}

# Processing configurations
PROCESSING_CONFIG = {
    'date_formats': ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'],
    'numeric_precision': 4,
    'qaqc_indicators': ['FDUP', 'BLANK', 'STD', 'QC'],
    'forward_fill_columns': ['DATE', 'LEVEL', 'SECTION', 'DIRECTION'],
    'default_output_dir': 'outputs'
}

# Sheet-specific configurations
SHEET_CONFIGS = {
    'STOPING': {
        'sheet_name': 'Stoping',
        'date_row': 15,
        'data_start_col': 7,
        'output_columns': ['Date', 'ID', 'Stoping_Actual_t', 'Stoping_Actual_gpt', 
                          'Stoping_Actual_kgs', 'Stoping_Budget_t', 'Stoping_Budget_gpt', 
                          'Stoping_Budget_kgs']
    },
    'TRAMMING': {
        'sheet_name': 'Tramming',
        'date_row': 15,
        'data_start_col': 7,
        'output_columns': ['Date', 'ID', 'Tramming_Actual_t', 'Tramming_Actual_gpt',
                          'Tramming_Actual_kgs', 'Tramming_Budget_t', 'Tramming_Budget_gpt',
                          'Tramming_Budget_kgs']
    },
    'DEVELOPMENT': {
        'sheet_name': 'Development',
        'data_start_row': 18,
        'data_start_col': 7,
        'output_columns': ['Date', 'Dev_ID', 'Budget_Metres', 'Actual_Metres']
    },
    'HOISTING': {
        'sheet_name': 'Hoisting',
        'data_start_col': 7,
        'output_columns': ['Date', 'Source', 'METRIC1', 'METRIC2', 'Value']
    },
    'BENCHES': {
        'sheet_name': 'BENCHES.',
        'output_columns_raw': ['DATE', 'LEVEL', 'SECTION', 'Unnamed: 3', 'LOCATION', 
                              'DIRECTION', 'SAMPLE TYPE', 'PEG ID', 'DEPTH (m)', 
                              'FROM', 'TO', 'AU', 'SAMPLE ID', 'ID', 'is_qaqc'],
        'output_columns_avg': ['DATE', 'ID', 'LEVEL', 'SECTION', 'Column_D', 
                              'LOCATION', 'DIRECTION', 'SAMPLE_TYPE', 'PEG_ID', 
                              'AU_avg', 'AU_count', 'Sample_count']
    }
}

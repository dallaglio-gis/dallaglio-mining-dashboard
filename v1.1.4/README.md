
# Mining Daily Report Processing Dashboard

A comprehensive Streamlit application for automated mining daily report data extraction and processing.

## 🎯 Overview

This tool consolidates all your mining data extraction needs into one streamlined interface, processing all 5 key sheets from your daily reports:
- **STOPING**: Daily actual/budget tonnes, grade, gold data
- **TRAMMING**: Tramming operations with similar structure to stoping  
- **DEVELOPMENT**: Budget and actual development metres
- **HOISTING**: Complex source/metric1/metric2/value format data
- **BENCHES**: Forward fill processing with QAQC sample identification

## 🚀 Features

- ✅ **Automated Processing**: No manual data manipulation needed
- ✅ **Error Handling**: Robust processing with detailed error reporting
- ✅ **Data Validation**: Compares results against established validation targets
- ✅ **Multiple Output Formats**: CSV downloads and comprehensive reports
- ✅ **Real-time Progress**: See processing status as it happens
- ✅ **Batch Processing**: Handle multiple sheets simultaneously
- ✅ **Visualizations**: Interactive charts and graphs
- ✅ **Quality Assurance**: Comprehensive validation and error checking

## 📋 Requirements

- Python 3.8+
- Streamlit
- pandas
- numpy
- openpyxl
- plotly

## 🔧 Installation

1. Clone or download the dashboard files
2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Launch the dashboard:
```bash
streamlit run app.py
```

## 📊 Usage

1. **Upload Excel File**: Use the file uploader to select your daily report
2. **Configure Settings**: Choose which sheets to process and options
3. **Start Processing**: Click the process button and watch progress
4. **Review Results**: View extracted data, validation results, and visualizations
5. **Download Data**: Get CSV files or bulk ZIP download

## 🎯 Validation Targets

The system validates extracted data against these established targets:

### STOPING
- Budget Tonnes: 31,746t
- Actual Tonnes: 26,928t  
- Budget Gold: 68.1kg
- Actual Gold: 62.8kg
- Budget Grade: 2.14g/t
- Actual Grade: 2.33g/t

### TRAMMING  
- Similar targets to STOPING with slight variations
- Tolerance: ±2% for all metrics

### DEVELOPMENT
- Budget Metres: ~800m
- Actual Metres: ~750m

### BENCHES
- Minimum samples: 1,700
- QAQC percentage: 10-20%
- Forward fill validation

## 📁 Project Structure

```
mining_dashboard/
├── app.py                          # Main Streamlit application
├── mining_processor.py             # Master processing engine
├── requirements.txt                # Python dependencies
├── config/
│   └── validation_targets.py       # Validation targets and configs
├── processors/
│   ├── stoping_processor.py        # STOPING sheet processor
│   ├── tramming_processor.py       # TRAMMING sheet processor  
│   ├── development_processor.py    # DEVELOPMENT sheet processor
│   ├── hoisting_processor.py       # HOISTING sheet processor
│   └── benches_processor.py        # BENCHES sheet processor
├── utils/
│   └── common.py                   # Common utility functions
├── logs/                           # Processing logs
└── outputs/                        # Output CSV files
```

## 🔍 Processing Details

### STOPING & TRAMMING
- Extracts daily actual/budget data for tonnes, grade, and gold
- Handles missing budget data for new stopes
- Identifies partial data scenarios
- Creates comprehensive daily records

### DEVELOPMENT  
- Processes budget and actual development metres
- Spreads monthly budgets across daily records
- Tracks development progress

### HOISTING
- Handles complex Source/METRIC1/METRIC2/Value format
- Processes multiple metrics per source
- Creates normalized daily records

### BENCHES
- Applies forward fill processing for hierarchical data
- Identifies QAQC samples vs regular samples  
- Creates both raw data and average grades outputs
- Uses 5-component ID formation: LEVEL_SECTION_Column_D_LOCATION_DIRECTION

## 🚨 Error Handling

The system includes comprehensive error handling:
- File validation before processing
- Sheet existence checks
- Data type validation  
- Missing data handling
- Processing progress tracking
- Detailed error reporting

## 📈 Outputs

For each processed sheet, you get:
- **CSV Data File**: Clean, normalized data ready for analysis
- **Processing Report**: Detailed extraction summary
- **Validation Results**: Comparison against known targets
- **Visualizations**: Charts and graphs (optional)
- **Master Summary**: Overall processing results

## 🎛️ Advanced Options

- **Sheet Selection**: Choose which sheets to process
- **Validation Toggle**: Enable/disable validation against targets
- **Visualization Toggle**: Include/exclude data charts
- **Custom Output Directory**: Specify where files are saved
- **Batch Download**: Download all results as ZIP file

## 🐛 Troubleshooting

**File Upload Issues:**
- Ensure file is .xlsx or .xls format
- Check that all required sheets exist
- Verify file is not corrupted

**Processing Errors:**
- Check the processing logs in the logs/ directory
- Review error messages in the interface
- Ensure Excel file structure matches expectations

**Validation Failures:**
- Review validation results for specific failures
- Check if data extraction completed successfully
- Compare against known good files

## 📞 Support

This tool consolidates the working extraction logic from all your previous individual sheet processors. Each processor maintains the exact logic that was validated in your separate extractions.

For issues or questions, review the processing logs and error messages provided in the interface.

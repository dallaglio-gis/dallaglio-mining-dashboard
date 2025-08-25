
"""
Mining Daily Report Processing Dashboard
Streamlit application for automated mining data extraction
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import sys
from datetime import datetime
from typing import Dict, List
import zipfile
import tempfile
import re

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from mining_processor import MiningDataProcessor
from config.validation_targets import VALIDATION_TARGETS, SHEET_CONFIGS

# NOTE: Page config and CSS styling are set within main() to avoid import-time
# Streamlit calls when this module is imported by another app.

def _decode_bytes(x):
    # Convert raw bytes to str; keep other values unchanged
    if isinstance(x, (bytes, bytearray)):
        try:
            return x.decode("utf-8", "ignore")
        except Exception:
            return str(x)
    return x

NUMERIC_COL_HINT = re.compile(
    r"(tonnes|grade|gold|metres|mtd|budget|actual|kg|gpt|value)$",
    re.IGNORECASE
)

def sanitize_for_streamlit(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make DataFrame Arrow- & CSV-friendly:
    - Decode bytes to str
    - Normalize placeholders ('' '-' '–' '—' 'N/A') -> NaN
    - Coerce likely numeric columns to float
    - Coerce 'Date' columns to datetime
    - Handle mixed types for Arrow compatibility
    """
    if df is None or df.empty:
        return df

    out = df.copy()

    # First pass: decode all bytes in object columns
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].map(_decode_bytes)

    # Second pass: handle column types
    for col in out.columns:
        # Skip if already properly typed numeric
        if pd.api.types.is_numeric_dtype(out[col]):
            # Ensure no NaN values in numeric columns
            out[col] = out[col].fillna(0.0)
            continue
        
        # Check if this should be a numeric column
        if NUMERIC_COL_HINT.search(col) or any(hint in col.lower() for hint in ['actual', 'budget', 'mtd']):
            # For columns that should be numeric, be very aggressive
            if out[col].dtype == object:
                # Convert everything to string first, then clean
                out[col] = out[col].astype(str)
                # Replace all non-numeric placeholders
                out[col] = out[col].replace(['', ' ', '-', '–', '—', 'N/A', 'n/a', 'NA', 
                                            'na', 'None', 'none', 'nan', 'NaN'], '0')
                # Remove any remaining non-numeric characters
                out[col] = out[col].str.replace(r'[^\d\.\-]', '', regex=True)
                # Replace empty strings with '0'
                out[col] = out[col].replace('', '0')
            
            # Convert to float64 explicitly
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0.0).astype('float64')
            
        # Check if this is a date column
        elif 'date' in col.lower():
            out[col] = pd.to_datetime(out[col], errors="coerce")
            
        # For remaining object columns (text columns)
        else:
            # Ensure consistent string type
            out[col] = out[col].fillna("").astype(str)
            # Clean up nan strings
            out[col] = out[col].replace(['nan', 'None'], '')

    return out

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Always provide bytes to st.download_button."""
    return df.to_csv(index=False).encode("utf-8")

def main():
    # Streamlit page configuration and styling (set at runtime)
    try:
        st.set_page_config(
            page_title="Mining Data Processor",
            page_icon="⛏️",
            layout="wide",
            initial_sidebar_state="expanded"
        )
    except Exception:
        # If page config was already set by a parent app, ignore
        pass

    # Custom CSS for better styling
    st.markdown("""
    <style>
    .main > div {
        padding: 1rem 2rem;
    }
    .stAlert > div {
        padding: 0.5rem 1rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .success-message {
        color: #28a745;
        font-weight: bold;
    }
    .error-message {
        color: #dc3545;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("⛏️ Mining Daily Report Processing Dashboard")
    st.markdown("---")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("🔧 Processing Configuration")
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload Excel Daily Report",
            type=['xlsx', 'xls'],
            help="Select your mining daily report Excel file"
        )
        
        # Sheet selection
        st.subheader("📊 Sheet Selection")
        sheet_options = ['STOPING', 'TRAMMING', 'DEVELOPMENT', 'HOISTING', 'BENCHES']
        
        selected_sheets = []
        for sheet in sheet_options:
            if st.checkbox(f"{sheet}", value=True, key=f"sheet_{sheet}"):
                selected_sheets.append(sheet)
        
        # Processing options
        st.subheader("⚙️ Processing Options")
        
        include_validation = st.checkbox("Enable Validation Against Targets", value=True)
        include_visualization = st.checkbox("Generate Data Visualizations", value=True)
        create_summary_report = st.checkbox("Create Summary Report", value=True)
        
        # Output directory
        output_dir = st.text_input("Output Directory", value="outputs")
    
    # Main processing area
    if uploaded_file is not None:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_file_path = tmp_file.name
        
        # Initialize processor
        if 'processor' not in st.session_state:
            st.session_state.processor = MiningDataProcessor(output_dir)
        
        # Display file information
        st.subheader("📁 File Information")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Filename", uploaded_file.name)
        with col2:
            st.metric("File Size", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            st.metric("Sheets Selected", len(selected_sheets))
        
        # Validate file
        st.subheader("✅ File Validation")
        
        with st.spinner("Validating Excel file..."):
            is_valid, missing_sheets = st.session_state.processor.validate_excel_file(temp_file_path)
        
        if is_valid:
            st.success("✅ File validation successful! All required sheets found.")
        else:
            st.error(f"❌ File validation failed. Missing sheets: {missing_sheets}")
        
        # Processing section
        st.subheader("🚀 Data Processing")
        
        if st.button("🔄 Start Processing", type="primary", use_container_width=True):
            if not selected_sheets:
                st.error("Please select at least one sheet to process.")
            else:
                process_data(temp_file_path, selected_sheets, include_validation, 
                           include_visualization, create_summary_report)
        
        # Results display
        if 'processing_results' in st.session_state:
            display_results(st.session_state.processing_results, include_visualization)
        
        # Cleanup temp file
        try:
            os.unlink(temp_file_path)
        except:
            pass
    
    else:
        # Welcome screen
        display_welcome_screen()

def process_data(file_path: str, selected_sheets: List[str], include_validation: bool,
                include_visualization: bool, create_summary: bool):
    """Process the uploaded data"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.text("Initializing processing...")
        progress_bar.progress(10)
        
        # Process all selected sheets
        status_text.text("Processing sheets...")
        results = st.session_state.processor.process_all_sheets(file_path, selected_sheets)
        progress_bar.progress(80)
        
        # Store results
        st.session_state.processing_results = results
        
        status_text.text("Processing complete!")
        progress_bar.progress(100)
        
        st.success("🎉 Processing completed successfully!")
        
    except Exception as e:
        st.error(f"❌ Processing failed: {str(e)}")
        st.session_state.processing_results = {'error': str(e)}

def display_results(results: Dict, include_visualization: bool):
    """Display processing results"""
    
    if 'error' in results:
        st.error(f"Processing Error: {results['error']}")
        return
    
    st.header("📊 Processing Results")
    
    # Overall summary
    st.subheader("📈 Overall Summary")
    
    overall = results.get('overall_summary', {})
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Sheets Processed", f"{overall.get('successful_sheets', 0)}/{overall.get('total_sheets_requested', 0)}")
    with col2:
        st.metric("Total Records", overall.get('total_records', 0))
    with col3:
        st.metric("Processing Time", results.get('processing_time', 'N/A'))
    with col4:
        st.metric("Output Directory", overall.get('output_directory', 'N/A'))
    
    # Individual sheet results
    st.subheader("📋 Sheet Processing Details")
    
    for sheet_type, sheet_result in results.get('sheets_processed', {}).items():
        with st.expander(f"{sheet_type} Results", expanded=True):
            
            if sheet_result['success']:
                st.success(f"✅ {sheet_type} processed successfully")
                
                # Basic metrics
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Records Extracted", len(sheet_result['data']))
                with col2:
                    if 'output_file' in sheet_result:
                        st.text(f"Output: {os.path.basename(sheet_result['output_file'])}")
                
                # Validation results
                if sheet_result.get('validation'):
                    display_validation_results(sheet_result['validation'], sheet_type)
                
                # Data preview
                st.subheader(f"{sheet_type} Data Preview")
                df_raw = sheet_result['data']
                df = sanitize_for_streamlit(df_raw)
                
                if not df.empty:
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    # Download button
                    csv_bytes = to_csv_bytes(df)
                    st.download_button(
                        f"📥 Download {sheet_type} Data",
                        data=csv_bytes,
                        file_name=f"{sheet_type.lower()}_data.csv",
                        mime="text/csv",
                        key=f"dl_{sheet_type.lower()}"
                    )
                    
                    # Visualization
                    if include_visualization and len(df) > 0:
                        display_visualizations(df, sheet_type)
                
                # BENCHES special case (has additional average grades file)
                if sheet_type == 'BENCHES' and 'data_avg' in sheet_result:
                    st.subheader("Average Grades Data")
                    avg_df = sanitize_for_streamlit(sheet_result['data_avg'])
                    st.dataframe(avg_df.head(10), use_container_width=True)
                    
                    avg_bytes = to_csv_bytes(avg_df)
                    st.download_button(
                        "📥 Download Average Grades",
                        data=avg_bytes,
                        file_name="benches_average_grades.csv",
                        mime="text/csv",
                        key="dl_benches_avg"
                    )
            
            else:
                st.error(f"❌ {sheet_type} processing failed")
                st.error(f"Error: {sheet_result.get('error', 'Unknown error')}")
    
    # Bulk download option
    st.subheader("📦 Bulk Download")
    if st.button("📥 Download All Results", use_container_width=True):
        create_bulk_download(results)

def display_validation_results(validation: Dict, sheet_type: str):
    """Display validation results"""
    if 'details' not in validation:
        return
    
    st.subheader("🎯 Validation Results")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Tests Passed", validation['passed'], delta_color="normal")
    with col2:
        st.metric("Tests Failed", validation['failed'], delta_color="inverse")
    
    # Detailed validation
    for metric, result in validation['details'].items():
        status = "✅" if result['passed'] else "❌"
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.text(f"{status} {metric}")
        with col2:
            st.text(f"Target: {result['target']:.2f}")
        with col3:
            st.text(f"Actual: {result['actual']:.2f}")
        with col4:
            diff_pct = result.get('diff_percentage', 0) * 100
            st.text(f"Diff: {diff_pct:.1f}%")

def display_visualizations(df: pd.DataFrame, sheet_type: str):
    """Display data visualizations"""
    st.subheader(f"📊 {sheet_type} Visualizations")
    
    if sheet_type in ['STOPING', 'TRAMMING']:
        # Production charts
        if len(df.columns) >= 6:
            # Daily production trend
            if 'Date' in df.columns:
                # Work with already sanitized data - no need to re-convert
                df_clean = df.copy()
                
                daily_summary = df_clean.groupby('Date').agg({
                    df.columns[2]: 'sum',  # Actual tonnes
                    df.columns[5]: 'sum'   # Budget tonnes
                }).reset_index()
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=daily_summary['Date'], 
                                       y=daily_summary[df.columns[2]], 
                                       mode='lines+markers', 
                                       name='Actual Tonnes'))
                fig.add_trace(go.Scatter(x=daily_summary['Date'], 
                                       y=daily_summary[df.columns[5]], 
                                       mode='lines+markers', 
                                       name='Budget Tonnes'))
                
                fig.update_layout(title=f"Daily {sheet_type} Production Trend",
                                xaxis_title="Date",
                                yaxis_title="Tonnes")
                st.plotly_chart(fig, use_container_width=True)
            
            # Top producers
            if 'ID' in df.columns:
                # Already sanitized - no need to re-convert
                producer_summary = df_clean.groupby('ID')[df.columns[2]].sum().sort_values(ascending=False).head(10)
                
                fig = px.bar(x=producer_summary.index, 
                           y=producer_summary.values,
                           title=f"Top 10 {sheet_type} Producers")
                fig.update_layout(xaxis_title="Producer ID", yaxis_title="Total Tonnes")
                st.plotly_chart(fig, use_container_width=True)
    
    elif sheet_type == 'DEVELOPMENT':
        # Development progress
        if 'Date' in df.columns and len(df.columns) >= 4:
            # Already sanitized - use directly
            df_clean = df.copy()
            
            daily_dev = df_clean.groupby('Date').agg({
                'Budget_Metres': 'sum',
                'Actual_Metres': 'sum'
            }).reset_index()
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=daily_dev['Date'], 
                                   y=daily_dev['Actual_Metres'],
                                   mode='lines+markers', 
                                   name='Actual Metres'))
            fig.add_trace(go.Scatter(x=daily_dev['Date'], 
                                   y=daily_dev['Budget_Metres'],
                                   mode='lines+markers', 
                                   name='Budget Metres'))
            
            fig.update_layout(title="Daily Development Progress",
                            xaxis_title="Date",
                            yaxis_title="Metres")
            st.plotly_chart(fig, use_container_width=True)
    
    elif sheet_type == 'HOISTING':
        # Hoisting metrics
        if 'Value' in df.columns and 'Source' in df.columns:
            # Already sanitized - use directly
            df_clean = df.copy()
            source_summary = df_clean.groupby('Source')['Value'].sum().sort_values(ascending=False)
            
            fig = px.pie(values=source_summary.values, 
                        names=source_summary.index,
                        title="Hoisting by Source")
            st.plotly_chart(fig, use_container_width=True)
    
    elif sheet_type == 'BENCHES':
        # Grade distribution
        if 'AU' in df.columns:
            # Already sanitized - just filter out zeros if needed
            au_numeric = df['AU'][df['AU'] > 0]  # Only show non-zero grades
            
            if len(au_numeric) > 0:
                fig = px.histogram(x=au_numeric, 
                                 title="Gold Grade Distribution",
                                 nbins=30)
                fig.update_layout(xaxis_title="AU Grade (g/t)", yaxis_title="Frequency")
                st.plotly_chart(fig, use_container_width=True)

def create_bulk_download(results: Dict):
    """Create a ZIP file with all results"""
    zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
    
    with zipfile.ZipFile(zip_buffer.name, 'w') as zip_file:
        for sheet_type, sheet_result in results.get('sheets_processed', {}).items():
            if sheet_result['success'] and not sheet_result['data'].empty:
                clean = sanitize_for_streamlit(sheet_result['data'])
                zip_file.writestr(f"{sheet_type.lower()}_data.csv", clean.to_csv(index=False))
                
                # Add average grades for BENCHES
                if sheet_type == 'BENCHES' and 'data_avg' in sheet_result:
                    avg_clean = sanitize_for_streamlit(sheet_result['data_avg'])
                    zip_file.writestr("benches_average_grades.csv", avg_clean.to_csv(index=False))
    
    # Read the ZIP file content first
    try:
        with open(zip_buffer.name, 'rb') as f:
            zip_content = f.read()
        
        st.download_button(
            "📦 Download ZIP Archive",
            data=zip_content,
            file_name=f"mining_extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            key="dl_zip_all"
        )
    finally:
        # Clean up the temporary file
        try:
            os.unlink(zip_buffer.name)
        except (OSError, PermissionError) as e:
            # If we can't delete it immediately, it will be cleaned up by the OS eventually
            pass

def display_welcome_screen():
    """Display welcome screen with instructions"""
    
    st.markdown("""
    ## 👋 Welcome to the Mining Data Processing Dashboard!
    
    This powerful tool consolidates all your mining daily report data extraction needs into one streamlined interface.
    
    ### 🎯 What This Tool Does:
    - **Extracts data from all 5 key sheets**: STOPING, TRAMMING, DEVELOPMENT, HOISTING, and BENCHES
    - **Validates against known targets** to ensure data accuracy
    - **Provides comprehensive reporting** with detailed summaries
    - **Generates visualizations** to help you understand your data
    - **Handles complex processing** like forward fill and QAQC sample identification
    
    ### 🚀 Key Features:
    - ✅ **Automated Processing**: No manual data manipulation needed
    - ✅ **Error Handling**: Robust processing with detailed error reporting  
    - ✅ **Data Validation**: Compares results against your validation targets
    - ✅ **Multiple Output Formats**: CSV downloads and comprehensive reports
    - ✅ **Real-time Progress**: See processing status as it happens
    - ✅ **Batch Processing**: Handle multiple sheets simultaneously
    
    ### 📊 Supported Sheet Types:
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **STOPING**
        - Daily actual/budget tonnes, grade, gold
        - Handles missing budget data
        - Identifies new stopes
        
        **TRAMMING** 
        - Similar structure to STOPING
        - Tramming operations tracking
        - Budget vs actual analysis
        
        **DEVELOPMENT**
        - Budget and actual metres
        - Development progress tracking
        """)
    
    with col2:
        st.markdown("""
        **HOISTING**
        - Source/METRIC1/METRIC2/Value format
        - Complex multi-metric data
        - Daily hoisting operations
        
        **BENCHES**
        - Forward fill processing
        - QAQC sample identification
        - Average grades calculation
        - Raw data and processed outputs
        """)
    
    st.markdown("""
    ### 🔧 How to Use:
    1. **Upload your Excel file** using the file uploader in the sidebar
    2. **Select which sheets to process** (all are selected by default)
    3. **Configure processing options** (validation, visualizations, reports)
    4. **Click "Start Processing"** and watch the magic happen!
    5. **Review results** and download your processed data
    
    ### ⚡ Quick Start:
    Simply upload your daily report Excel file and click "Start Processing" with the default settings for the best experience!
    """)
    
    # Display validation targets info
    st.subheader("🎯 Validation Targets")
    st.markdown("This tool validates your extracted data against these established targets:")
    
    for sheet, targets in VALIDATION_TARGETS.items():
        with st.expander(f"{sheet} Validation Targets"):
            for key, value in targets.items():
                if key != 'tolerance':
                    st.text(f"{key}: {value}")

if __name__ == "__main__":
    main()

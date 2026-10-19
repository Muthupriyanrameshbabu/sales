"""
Data Validation and Column Mapping Module
Handles Excel/CSV file ingestion, sheet selection, column auto-detection,
flexible column mapping, missing value checking, date parsing, and validation warnings.
"""

import pandas as pd
import numpy as np
from datetime import datetime

# Standard required business fields and common synonyms for auto-mapping
STANDARD_FIELDS = {
    "Lead ID": ["lead_id", "lead id", "leadid", "id", "customer_id", "enquiry_id"],
    "Customer Name": ["customer_name", "customer name", "name", "full name", "client name", "lead name"],
    "Enquiry Date": ["enquiry_date", "enquiry date", "date", "created_date", "registration_date", "lead_date"],
    "Course Name": ["course_name", "course name", "course", "program", "program_name", "product"],
    "Lead Source": ["lead_source", "lead source", "source", "channel", "campaign", "acquisition_source"],
    "Lead Category": ["lead_category", "lead category", "category", "lead_type", "type", "segment"],
    "Call Count": ["call_count", "call count", "calls", "number_of_calls", "call_attempts"],
    "Call Status": ["call_status", "call status", "status", "call_outcome"],
    "Interest Status": ["interest_status", "interest status", "interest", "intent", "lead_interest"],
    "Registration Status": ["registration_status", "registration status", "registered", "enrollment_status", "conversion_status"],
    "Registration Date": ["registration_date", "registration date", "enrolled_date", "conversion_date"],
    "Sales Representative": ["sales_representative", "sales representative", "representative", "rep", "agent", "owner"],
    "Revenue": ["revenue", "course fee", "amount", "price", "fee", "sales_value", "payment"]
}

def load_uploaded_file(uploaded_file, sheet_name=0):
    """Loads an uploaded Excel (.xlsx) or CSV file into a pandas DataFrame."""
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif file_name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        else:
            raise ValueError("Unsupported file format. Please upload an .xlsx, .xls, or .csv file.")
        return df, None
    except Exception as e:
        return None, str(e)

def get_excel_sheets(uploaded_file):
    """Returns sheet names if the file is an Excel workbook, else []."""
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(uploaded_file)
            return xl.sheet_names
        return []
    except Exception:
        return []

def auto_detect_mapping(df_columns):
    """Attempts to automatically map DataFrame columns to standard business fields."""
    mapping = {}
    cols_lower = {str(col).strip().lower(): col for col in df_columns}
    
    for std_field, synonyms in STANDARD_FIELDS.items():
        matched = None
        for syn in synonyms:
            if syn in cols_lower:
                matched = cols_lower[syn]
                break
        if matched:
            mapping[std_field] = matched
        else:
            # Fuzzy substring matching fallback
            for col_l, original in cols_lower.items():
                if any(syn in col_l for syn in synonyms):
                    mapping[std_field] = original
                    break
            if std_field not in mapping:
                mapping[std_field] = None
    return mapping

def validate_and_clean_data(df, column_mapping):
    """
    Renames columns according to mapping, validates data types, handles missing values,
    parses dates, removes duplicate lead IDs (while preserving call logs if needed),
    and generates detailed validation warnings and quality metrics.
    """
    warnings = []
    cleaned_df = pd.DataFrame()
    
    # Rename columns based on mapping
    rename_dict = {orig: std for std, orig in column_mapping.items() if orig is not None and orig != ""}
    cleaned_df = df.rename(columns=rename_dict)
    
    # Ensure all standard fields exist (fill with NaN or defaults if missing)
    for std_field in STANDARD_FIELDS.keys():
        if std_field not in cleaned_df.columns:
            cleaned_df[std_field] = np.nan

    # 1. Lead ID validation
    if cleaned_df['Lead ID'].isna().all():
        warnings.warn_lead_id = "Lead ID column is missing or empty. Auto-generating unique Lead IDs."
        cleaned_df['Lead ID'] = [f"LD-{1000 + i}" for i in range(len(cleaned_df))]
    else:
        cleaned_df['Lead ID'] = cleaned_df['Lead ID'].astype(str).str.strip()
        null_ids = cleaned_df['Lead ID'].isin(['nan', 'None', '', 'NAT'])
        if null_ids.any():
            warnings.append(f"{null_ids.sum()} rows found with missing Lead ID; auto-assigned placeholder IDs.")
            cleaned_df.loc[null_ids, 'Lead ID'] = [f"LD-AUTO-{i}" for i in range(null_ids.sum())]

    # 2. Date parsing (Enquiry Date & Registration Date)
    for date_col in ['Enquiry Date', 'Registration Date']:
        if not cleaned_df[date_col].isna().all():
            parsed_dates = pd.to_datetime(cleaned_df[date_col], errors='coerce')
            invalid_dates = cleaned_df[date_col].notna() & parsed_dates.isna()
            if invalid_dates.sum() > 0:
                warnings.append(f"{invalid_dates.sum()} invalid date format values detected in '{date_col}' and set to NaT.")
            cleaned_df[date_col] = parsed_dates
        else:
            cleaned_df[date_col] = pd.NaT

    # 3. Categorical / Text fields cleaning
    text_fields = ['Customer Name', 'Course Name', 'Lead Source', 'Lead Category', 'Call Status', 'Interest Status', 'Registration Status', 'Sales Representative']
    for field in text_fields:
        if not cleaned_df[field].isna().all():
            cleaned_df[field] = cleaned_df[field].astype(str).str.strip()
            cleaned_df[field] = cleaned_df[field].replace(['nan', 'None', 'NAT', ''], 'Unknown')
        else:
            cleaned_df[field] = 'Unknown'

    # Normalize Registration Status values
    reg_status_map = {
        'yes': 'Registered', 'true': 'Registered', '1': 'Registered', 'registered': 'Registered', 'enrolled': 'Registered',
        'no': 'Not Registered', 'false': 'Not Registered', '0': 'Not Registered', 'not registered': 'Not Registered', 'pending': 'Not Registered'
    }
    cleaned_df['Registration Status Norm'] = cleaned_df['Registration Status'].str.lower().map(reg_status_map).fillna('Not Registered')

    # Normalize Interest Status values
    interest_map = {
        'yes': 'Interested', 'true': 'Interested', '1': 'Interested', 'interested': 'Interested', 'high': 'Interested',
        'no': 'Not Interested', 'false': 'Not Interested', '0': 'Not Interested', 'not interested': 'Not Interested', 'low': 'Not Interested', 'neutral': 'Neutral'
    }
    cleaned_df['Interest Status Norm'] = cleaned_df['Interest Status'].str.lower().map(interest_map).fillna('Neutral')

    # 4. Numeric fields cleaning (Call Count & Revenue)
    # Call Count
    if not cleaned_df['Call Count'].isna().all():
        cleaned_df['Call Count'] = pd.to_numeric(cleaned_df['Call Count'], errors='coerce').fillna(1)
        cleaned_df['Call Count'] = cleaned_df['Call Count'].apply(lambda x: max(1, int(x)))
    else:
        cleaned_df['Call Count'] = 1

    # Revenue / Course Fee
    if not cleaned_df['Revenue'].isna().all():
        cleaned_df['Revenue'] = pd.to_numeric(cleaned_df['Revenue'], errors='coerce').fillna(0.0)
    else:
        cleaned_df['Revenue'] = 0.0

    # Summary data quality stats
    quality_report = {
        "total_rows": len(cleaned_df),
        "unique_leads": cleaned_df['Lead ID'].nunique(),
        "duplicate_lead_rows": len(cleaned_df) - cleaned_df['Lead ID'].nunique(),
        "missing_enquiry_dates": cleaned_df['Enquiry Date'].isna().sum(),
        "warnings": warnings
    }

    return cleaned_df, quality_report
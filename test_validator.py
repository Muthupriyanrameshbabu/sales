"""Unit tests for data validation and column mapping."""

import pytest
import pandas as pd
from src.validator import auto_detect_mapping, validate_and_clean_data

def test_auto_detect_mapping():
    sample_columns = ["Lead_ID", "Customer Name", "Enquiry Date", "Course Name", "Source"]
    mapping = auto_detect_mapping(sample_columns)
    assert mapping["Lead ID"] == "Lead_ID"
    assert mapping["Customer Name"] == "Customer Name"
    assert mapping["Course Name"] == "Course Name"
    assert mapping["Lead Source"] == "Source"

def test_validate_and_clean_data():
    raw_data = pd.DataFrame({
        "lead_id": ["L1", "L2", "L1"], # Duplicate to test unique count logic
        "Customer Name": ["Alice", "Bob", "Alice"],
        "Enquiry Date": ["2026-01-01", "2026-01-02", "2026-01-01"],
        "Course Name": ["Python", "Data Science", "Python"],
        "Registration Status": ["Registered", "Not Registered", "Registered"]
    })
    mapping = {"Lead ID": "lead_id", "Customer Name": "Customer Name", "Enquiry Date": "Enquiry Date", "Course Name": "Course Name", "Registration Status": "Registration Status"}
    cleaned, quality = validate_and_clean_data(raw_data, mapping)
    
    assert quality["total_rows"] == 3
    assert quality["unique_leads"] == 2
    assert "Registration Status Norm" in cleaned.columns
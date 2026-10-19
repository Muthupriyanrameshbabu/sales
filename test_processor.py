"""Unit tests for KPI calculations and metric rules."""

import pytest
import pandas as pd
from src.processor import calculate_kpis, get_course_performance

def test_calculate_kpis():
    df = pd.DataFrame({
        "Lead ID": ["L1", "L2", "L3", "L4"],
        "Call Count": [2, 3, 1, 4],
        "Interest Status Norm": ["Interested", "Interested", "Neutral", "Interested"],
        "Registration Status Norm": ["Registered", "Not Registered", "Not Registered", "Registered"],
        "Revenue": [1000, 0, 0, 2000]
    })
    kpis = calculate_kpis(df)
    
    assert kpis["total_leads"] == 4
    assert kpis["total_calls"] == 10
    assert kpis["total_interested"] == 3
    assert kpis["total_registrations"] == 2
    assert kpis["lead_to_reg_rate"] == 50.0 # 2 reg / 4 leads * 100
    assert kpis["interest_to_reg_rate"] == 66.67 # 2 reg / 3 interested * 100
    assert kpis["total_revenue"] == 3000.0
    assert kpis["avg_reg_value"] == 1500.0

def test_zero_denominator_safety():
    df = pd.DataFrame({
        "Lead ID": [],
        "Call Count": [],
        "Interest Status Norm": [],
        "Registration Status Norm": [],
        "Revenue": []
    })
    kpis = calculate_kpis(df)
    assert kpis["lead_to_reg_rate"] == 0.0
    assert kpis["interest_to_reg_rate"] == 0.0

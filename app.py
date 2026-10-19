"""
Sales Performance Analysis and Reporting Dashboard
Main Streamlit Application implementing Enterprise Power BI aesthetics,
interactive data uploading, mapping, validation, KPIs, charts, and report exports.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
from datetime import datetime

from src.validator import load_uploaded_file, get_excel_sheets, auto_detect_mapping, validate_and_clean_data
from src.processor import (
    calculate_kpis, get_course_performance, get_monthly_trends,
    get_lead_category_distribution, get_lead_source_performance,
    get_sales_rep_performance, generate_insights
)

# Page Configuration
st.set_page_config(
    page_title="Sales Performance Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Professional Enterprise Styling (Power BI Inspired)
st.markdown("""
<style>
    /* Global Styling */
    .main {
        background-color: #F8F9FA;
        color: #212529;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Header styling */
    .dashboard-header {
        background: #FFFFFF;
        padding: 20px 24px;
        border-radius: 8px;
        border: 1px solid #E5E7EB;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .dashboard-title {
        font-size: 24px;
        font-weight: 700;
        color: #1E3A8A;
        margin: 0 0 4px 0;
    }
    .dashboard-subtitle {
        font-size: 14px;
        color: #6B7280;
        margin: 0;
    }
    
    /* KPI Metric Cards */
    .kpi-card {
        background: #FFFFFF;
        padding: 16px 20px;
        border-radius: 8px;
        border: 1px solid #E5E7EB;
        border-left: 4px solid #1E3A8A;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        margin-bottom: 16px;
    }
    .kpi-label {
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        color: #6B7280;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 700;
        color: #111827;
        margin: 0;
    }
    
    /* Section containers */
    .section-container {
        background: #FFFFFF;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #E5E7EB;
        margin-bottom: 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #F3F4F6;
        border-right: 1px solid #E5E7EB;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if 'raw_df' not in st.session_state:
    st.session_state.raw_df = None
if 'cleaned_df' not in st.session_state:
    st.session_state.cleaned_df = None
if 'quality_report' not in st.session_state:
    st.session_state.quality_report = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = "No dataset loaded"
if 'last_refresh' not in st.session_state:
    st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
if 'use_demo_data' not in st.session_state:
    st.session_state.use_demo_data = False

# Sidebar - Data Management & Navigation
st.sidebar.markdown("### 📊 Sales Dashboard")
st.sidebar.markdown("---")

nav_mode = st.sidebar.radio("Navigation", ["Dashboard", "Data Upload & Mapping", "Power BI Integration & Guide"])

st.sidebar.markdown("---")
st.sidebar.markdown("### Data Management")

# Demo Data Loader
sample_path = "data/sample_sales_data.xlsx"
if os.path.exists(sample_path):
    if st.sidebar.button("Load Demo Dataset (Excel)", use_container_width=True):
        try:
            df_demo = pd.read_excel(sample_path)
            mapping_demo = auto_detect_mapping(df_demo.columns)
            cleaned, q_rep = validate_and_clean_data(df_demo, mapping_demo)
            st.session_state.raw_df = df_demo
            st.session_state.cleaned_df = cleaned
            st.session_state.quality_report = q_rep
            st.session_state.file_name = "sample_sales_data.xlsx (DEMO)"
            st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.sidebar.success("Demo dataset loaded successfully!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error loading demo: {e}")

# Reset Session
if st.sidebar.button("Reset Session", use_container_width=True):
    st.session_state.raw_df = None
    st.session_state.cleaned_df = None
    st.session_state.quality_report = None
    st.session_state.file_name = "No dataset loaded"
    st.rerun()

# ---------------------------------------------------------
# VIEW 1: DATA UPLOAD & MAPPING
# ---------------------------------------------------------
if nav_mode == "Data Upload & Mapping":
    st.markdown("""
    <div class="dashboard-header">
        <h1 class="dashboard-title">Data Ingestion, Validation & Mapping</h1>
        <p class="dashboard-subtitle">Upload your sales dataset (.xlsx or .csv), inspect column types, map business fields, and validate data quality.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Upload Excel (.xlsx, .xls) or CSV file", type=["xlsx", "xls", "csv"])
    with col2:
        st.markdown("##### Download Template")
        if os.path.exists(sample_path):
            with open(sample_path, "rb") as f:
                st.download_button(
                    label="Download Excel Template",
                    data=f,
                    file_name="sales_data_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("Template file not found on disk.")

    if uploaded_file is not None:
        sheets = get_excel_sheets(uploaded_file)
        selected_sheet = 0
        if len(sheets) > 1:
            selected_sheet = st.selectbox("Select Worksheet", sheets)
        
        df_raw, err = load_uploaded_file(uploaded_file, sheet_name=selected_sheet)
        if err:
            st.error(f"Failed to parse file: {err}")
        else:
            st.session_state.raw_df = df_raw
            st.session_state.file_name = uploaded_file.name
            st.session_state.last_refresh = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.success(f"Successfully loaded **{uploaded_file.name}** ({len(df_raw):,} rows, {len(df_raw.columns)} columns).")
            
            with st.expander("🔍 Preview Raw Dataset", expanded=False):
                st.dataframe(df_raw.head(10), use_container_width=True)

            st.markdown("### 🗺️ Column Mapping Configuration")
            st.markdown("Map your dataset columns to the required business fields below:")
            
            auto_map = auto_detect_mapping(df_raw.columns)
            col_options = ["(None / Missing)"] + list(df_raw.columns)
            
            user_mapping = {}
            cols = st.columns(2)
            for i, (std_field, default_col) in enumerate(auto_map.items()):
                target_col = cols[i % 2].selectbox(
                    f"Map to **{std_field}**",
                    options=col_options,
                    index=col_options.index(default_col) if default_col in col_options else 0,
                    key=f"map_{std_field}"
                )
                user_mapping[std_field] = None if target_col == "(None / Missing)" else target_col

            if st.button("Validate and Process Dataset", type="primary", use_container_width=True):
                cleaned, q_rep = validate_and_clean_data(df_raw, user_mapping)
                st.session_state.cleaned_df = cleaned
                st.session_state.quality_report = q_rep
                st.success("Dataset validated, cleaned, and processed successfully! Navigate to 'Dashboard' to view insights.")

    # Show validation summary if cleaned dataset exists
    if st.session_state.cleaned_df is not None:
        st.markdown("---")
        st.markdown("### 📋 Data Quality & Validation Report")
        q = st.session_state.quality_report
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Rows", f"{q['total_rows']:,}")
        m2.metric("Unique Leads", f"{q['unique_leads']:,}")
        m3.metric("Duplicate Rows", f"{q['duplicate_lead_rows']:,}")
        m4.metric("Missing Enquiry Dates", f"{q['missing_enquiry_dates']:,}")

        if q['warnings']:
            st.warning("⚠️ **Validation Warnings:**")
            for w in q['warnings']:
                st.markdown(f"- {w}")
        else:
            st.success("✨ No critical validation warnings detected in the dataset.")

# ---------------------------------------------------------
# VIEW 2: DASHBOARD
# ---------------------------------------------------------
elif nav_mode == "Dashboard":
    cleaned_df = st.session_state.cleaned_df

    if cleaned_df is None:
        st.markdown("""
        <div class="dashboard-header">
            <h1 class="dashboard-title">Sales Performance Analysis and Reporting Dashboard</h1>
            <p class="dashboard-subtitle">No dataset currently loaded. Please load the <b>Demo Dataset</b> from the sidebar or upload your Excel/CSV file in <b>Data Upload & Mapping</b>.</p>
        </div>
        """, unsafe_allow_html=True)
        st.info("👈 Use the sidebar controls to load data.")
    else:
        # Header Info Card
        st.markdown(f"""
        <div class="dashboard-header">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <h1 class="dashboard-title">Sales Performance Analysis & Reporting Dashboard</h1>
                    <p class="dashboard-subtitle">Active Dataset: <b>{st.session_state.file_name}</b> | Last Refreshed: {st.session_state.last_refresh}</p>
                </div>
                <div>
                    <span style="background: #E0F2FE; color: #0369A1; padding: 6px 12px; border-radius: 4px; font-weight: 600; font-size: 13px;">Status: Active & Validated</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # INTERACTIVE FILTERS (SIDEBAR)
        # ---------------------------------------------------------
        st.sidebar.markdown("### 🔍 Dashboard Filters")
        
        # Date range filter
        min_date = cleaned_df['Enquiry Date'].min()
        max_date = cleaned_df['Enquiry Date'].max()
        
        if pd.notna(min_date) and pd.notna(max_date):
            date_range = st.sidebar.date_input(
                "Enquiry Date Range",
                value=(min_date.date(), max_date.date()),
                min_value=min_date.date(),
                max_value=max_date.date()
            )
        else:
            date_range = None

        # Course Filter
        courses = ["All"] + sorted(cleaned_df['Course Name'].unique().tolist())
        selected_course = st.sidebar.selectbox("Course Name", courses)

        # Lead Source Filter
        sources = ["All"] + sorted(cleaned_df['Lead Source'].unique().tolist())
        selected_source = st.sidebar.selectbox("Lead Source", sources)

        # Lead Category Filter
        categories = ["All"] + sorted(cleaned_df['Lead Category'].unique().tolist())
        selected_category = st.sidebar.selectbox("Lead Category", categories)

        # Registration Status Filter
        reg_statuses = ["All", "Registered", "Not Registered"]
        selected_reg_status = st.sidebar.selectbox("Registration Status", reg_statuses)

        if st.sidebar.button("Reset Filters", use_container_width=True):
            st.rerun()

        # Apply Filters to DataFrame
        filtered_df = cleaned_df.copy()
        if date_range and len(date_range) == 2:
            start_d, end_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            filtered_df = filtered_df[(filtered_df['Enquiry Date'] >= start_d) & (filtered_df['Enquiry Date'] <= end_d + pd.Timedelta(days=1))]
        if selected_course != "All":
            filtered_df = filtered_df[filtered_df['Course Name'] == selected_course]
        if selected_source != "All":
            filtered_df = filtered_df[filtered_df['Lead Source'] == selected_source]
        if selected_category != "All":
            filtered_df = filtered_df[filtered_df['Lead Category'] == selected_category]
        if selected_reg_status != "All":
            filtered_df = filtered_df[filtered_df['Registration Status Norm'] == selected_reg_status]

        # Calculate KPIs for Filtered Data
        kpis = calculate_kpis(filtered_df)

        # ---------------------------------------------------------
        # KPI CARDS
        # ---------------------------------------------------------
        st.markdown("### Key Performance Indicators")
        k1, k2, k3, k4, k5 = st.columns(5)

        with k1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">Total Calls Made</div>
                <div class="kpi-value">{kpis['total_calls']:,}</div>
            </div>
            """, unsafe_allow_html=True)

        with k2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">Customer Enquiries</div>
                <div class="kpi-value">{kpis['total_enquiries']:,}</div>
            </div>
            """, unsafe_allow_html=True)

        with k3:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">Interested Leads</div>
                <div class="kpi-value">{kpis['total_interested']:,}</div>
            </div>
            """, unsafe_allow_html=True)

        with k4:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">Successful Registrations</div>
                <div class="kpi-value">{kpis['total_registrations']:,}</div>
            </div>
            """, unsafe_allow_html=True)

        with k5:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">Conversion Rate</div>
                <div class="kpi-value">{kpis['lead_to_reg_rate']}%</div>
            </div>
            """, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # CHARTS SECTION (Row 1: Course Registrations & Monthly Trends)
        # ---------------------------------------------------------
        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown("#### Course-wise Registrations & Leads")
            course_perf = get_course_performance(filtered_df)
            if not course_perf.empty:
                fig_course = px.bar(
                    course_perf.sort_values(by='Registrations', ascending=True),
                    x='Registrations',
                    y='Course Name',
                    orientation='h',
                    text='Registrations',
                    color='Conversion_Rate (%)',
                    color_continuous_scale='Blues',
                    labels={'Registrations': 'Successful Registrations', 'Course Name': 'Course'}
                )
                fig_course.update_layout(
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=320,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="sans-serif", size=12, color="#374151")
                )
                st.plotly_chart(fig_course, use_container_width=True)
            else:
                st.info("No course data available.")
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown("#### Monthly Enquiries and Registrations")
            monthly_df = get_monthly_trends(filtered_df)
            if not monthly_df.empty:
                fig_monthly = px.line(
                    monthly_df,
                    x='Month-Year',
                    y=['Enquiries', 'Registrations'],
                    markers=True,
                    labels={'value': 'Count', 'Month-Year': 'Month', 'variable': 'Metric'}
                )
                fig_monthly.update_layout(
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=320,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="sans-serif", size=12, color="#374151"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_monthly, use_container_width=True)
            else:
                st.info("No monthly date series available.")
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------------
        # CHARTS SECTION (Row 2: Lead Category & Lead Source Performance)
        # ---------------------------------------------------------
        c3, c4 = st.columns(2)

        with c3:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown("#### Lead Category Distribution")
            cat_df = get_lead_category_distribution(filtered_df)
            if not cat_df.empty:
                fig_cat = px.bar(
                    cat_df,
                    x='Lead Category',
                    y='Leads',
                    text='Leads',
                    color='Lead Category',
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_cat.update_layout(
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=300,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    showlegend=False
                )
                st.plotly_chart(fig_cat, use_container_width=True)
            else:
                st.info("No lead category data available.")
            st.markdown('</div>', unsafe_allow_html=True)

        with c4:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown("#### Lead Source Conversion Performance")
            src_df = get_lead_source_performance(filtered_df)
            if not src_df.empty:
                fig_src = px.bar(
                    src_df,
                    x='Lead Source',
                    y='Registrations',
                    text='Registrations',
                    color='Conversion Rate (%)',
                    color_continuous_scale='Viridis'
                )
                fig_src.update_layout(
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=300,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_src, use_container_width=True)
            else:
                st.info("No lead source data available.")
            st.markdown('</div>', unsafe_allow_html=True)

        # Sales Representative Performance (If available)
        rep_df = get_sales_rep_performance(filtered_df)
        if not rep_df.empty:
            st.markdown('<div class="section-container">', unsafe_allow_html=True)
            st.markdown("#### Sales Representative Performance")
            fig_rep = px.bar(
                rep_df,
                x='Sales Representative',
                y=['Successful_Registrations', 'Leads_Managed'],
                barmode='group',
                labels={'value': 'Count', 'Sales Representative': 'Sales Rep', 'variable': 'Metric'}
            )
            fig_rep.update_layout(
                margin=dict(l=10, r=10, t=10, b=10),
                height=320,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_rep, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------------
        # INSIGHTS & OBSERVATIONS SECTION
        # ---------------------------------------------------------
        st.markdown('<div class="section-container">', unsafe_allow_html=True)
        st.markdown("#### 💡 Key Data-Driven Insights")
        insights = generate_insights(filtered_df, kpis, course_perf, get_lead_source_performance(filtered_df))
        for ins in insights:
            st.markdown(f"- {ins}")
        st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------------
        # DATA TABLE & EXPORTS SECTION
        # ---------------------------------------------------------
        st.markdown('<div class="section-container">', unsafe_allow_html=True)
        st.markdown("#### 📋 Filtered Records Data Table")
        
        search_query = st.text_input("Search records by Customer Name, Lead ID, or Course:", "")
        display_df = filtered_df.copy()
        if search_query:
            mask = display_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
            display_df = display_df[mask]

        st.dataframe(display_df, use_container_width=True)

        # Report Exports
        st.markdown("---")
        st.markdown("##### 📥 Export Cleaned Dataset & Reports")
        ex1, ex2, ex3 = st.columns(3)

        # Export Cleaned Excel
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='Cleaned_Sales_Data')
            get_course_performance(filtered_df).to_excel(writer, index=False, sheet_name='Course_Performance')
            get_lead_source_performance(filtered_df).to_excel(writer, index=False, sheet_name='Lead_Source_Summary')
        output_excel.seek(0)

        ex1.download_button(
            label="Download Power BI-Ready Excel",
            data=output_excel,
            file_name="cleaned_sales_dataset_powerbi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Export CSV
        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
        ex2.download_button(
            label="Download Cleaned CSV",
            data=csv_data,
            file_name="cleaned_sales_dataset.csv",
            mime="text/csv",
            use_container_width=True
        )

        # Export Summary Report
        summary_txt = f"""SALES PERFORMANCE SUMMARY REPORT
Generated At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Dataset: {st.session_state.file_name}

KEY PERFORMANCE INDICATORS:
- Total Leads: {kpis['total_leads']}
- Total Calls Made: {kpis['total_calls']}
- Customer Enquiries: {kpis['total_enquiries']}
- Interested Leads: {kpis['total_interested']}
- Successful Registrations: {kpis['total_registrations']}
- Lead-to-Registration Conversion Rate: {kpis['lead_to_reg_rate']}%
- Interest-to-Registration Conversion Rate: {kpis['interest_to_reg_rate']}%
- Total Revenue: {kpis['total_revenue']}
- Average Registration Value: {kpis['avg_reg_value']}

INSIGHTS:
""" + "\n".join([f"- {i.replace('**', '')}" for i in insights])
        
        ex3.download_button(
            label="Download Summary Report (TXT)",
            data=summary_txt.encode('utf-8'),
            file_name="sales_summary_report.txt",
            mime="text/plain",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# VIEW 3: POWER BI INTEGRATION & GUIDE
# ---------------------------------------------------------
elif nav_mode == "Power BI Integration & Guide":
    st.markdown("""
    <div class="dashboard-header">
        <h1 class="dashboard-title">Microsoft Power BI Integration & Implementation Guide</h1>
        <p class="dashboard-subtitle">Connect your validated sales dataset directly into Microsoft Power BI Desktop using DAX measures and enterprise reporting best practices.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    ### 1. Step-by-Step Power BI Connection Workflow
    1. **Export Cleaned Data:** Go to the **Dashboard** tab and click **"Download Power BI-Ready Excel"** to save your validated dataset.
    2. **Open Power BI Desktop:** Launch Power BI Desktop on your Windows machine.
    3. **Get Data:** Click **Home > Get Data > Excel Workbook** and select your downloaded cleaned Excel file.
    4. **Select Tables:** Select the `Cleaned_Sales_Data`, `Course_Performance`, and `Lead_Source_Summary` tables, then click **Load** (or **Transform Data** for advanced Power Query adjustments).
    5. **Refresh Data:** Whenever your source sales data updates in Excel, simply re-export or save the file to the same location and click **Refresh** in Power BI Desktop.
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    ### 2. Recommended DAX Measures for Power BI
    Use these Data Analysis Expressions (DAX) in Power BI Desktop to replicate and extend the web dashboard metrics:
    """)

    st.code("""
-- Total Unique Leads
Total Unique Leads = DISTINCTCOUNT('Cleaned_Sales_Data'[Lead ID])

-- Total Successful Registrations
Total Registrations = 
CALCULATE(
    DISTINCTCOUNT('Cleaned_Sales_Data'[Lead ID]),
    'Cleaned_Sales_Data'[Registration Status Norm] = "Registered"
)

-- Total Interested Leads
Total Interested Leads = 
CALCULATE(
    DISTINCTCOUNT('Cleaned_Sales_Data'[Lead ID]),
    'Cleaned_Sales_Data'[Interest Status Norm] = "Interested"
)

-- Lead-to-Registration Conversion Rate (%)
Lead to Reg Conversion (%) = 
DIVIDE([Total Registrations], [Total Unique Leads], 0) * 100

-- Total Calls Made
Total Calls = SUM('Cleaned_Sales_Data'[Call Count])

-- Total Revenue
Total Revenue = SUM('Cleaned_Sales_Data'[Revenue])
    """, language="dax")

    st.markdown("---")
    st.markdown("""
    ### 3. Recommended Power BI Report Layout
    - **Page 1: Executive Summary**
        - Card Visuals: Total Leads, Total Calls, Interested Leads, Successful Registrations, Conversion Rate.
        - Clustered Bar Chart: Registrations by Course Name.
        - Line Chart: Monthly Enquiries vs. Registrations over time.
    - **Page 2: Lead Generation & Source Analysis**
        - Donut Chart: Lead Category Distribution (Hot, Warm, Cold).
        - Stacked Bar Chart: Lead Source performance comparison.
    - **Page 3: Sales Representative Performance**
        - Table Visual: Sales Rep, Leads Managed, Total Calls, Successful Registrations, and Conversion %.
    """, unsafe_allow_html=True)
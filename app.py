import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter

try:
    from reportlab.lib.pagesizes import letter, legal, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

st.set_page_config(
    page_title="Sold Property Report",
    page_icon="📊",
    layout="wide"
)

# Field mapping from Close.com to report headers (for display only)
FIELD_MAPPING = {
    'display_name': 'Property Name',
    'custom.Asset_Owner': 'Owner',
    'custom.All_State': 'State',
    'custom.All_County': 'County',
    'custom.All_Asset_Surveyed_Acres': 'Acres',
    'custom.Asset_Cost_Basis': 'Cost Basis',
    'custom.Asset_Date_Purchased': 'Date Purchased',
    'primary_opportunity_status_label': 'Opportunity Status',
    'custom.Asset_Date_Sold': 'Date Sold',
    'custom.Asset_Gross_Sales_Price': 'Gross Sales Price',
    'custom.Asset_Closing_Costs': 'Closing Costs'
}

def safe_numeric_value(value, default=0):
    """Safely convert value to numeric, handling NaN/INF"""
    try:
        if pd.isna(value) or np.isinf(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default

def safe_int_value(value, default=0):
    """Safely convert value to integer, handling NaN/INF"""
    try:
        if pd.isna(value) or np.isinf(value):
            return default
        return int(value)
    except (ValueError, TypeError):
        return default

def parse_date(date_str):
    """Parse date string to datetime object"""
    if pd.isna(date_str) or date_str == '':
        return None
    try:
        return pd.to_datetime(date_str)
    except:
        return None

def get_quarter_year(date):
    """Get quarter and year from date"""
    if pd.isna(date):
        return None
    quarter = f"Q{((date.month - 1) // 3) + 1}"
    year = date.year
    return f"{quarter} {year}"

def sort_quarters_chronologically(quarters):
    """Sort quarters in chronological order (Q1 2022, Q2 2022, Q3 2022, Q4 2022, Q1 2023, etc.)"""
    def quarter_sort_key(quarter_str):
        if not quarter_str or pd.isna(quarter_str):
            return (0, 0)  # Put invalid quarters first
        try:
            parts = quarter_str.split()
            if len(parts) != 2:
                return (0, 0)
            quarter_num = int(parts[0][1:])  # Extract number from "Q1", "Q2", etc.
            year = int(parts[1])
            return (year, quarter_num)
        except (ValueError, IndexError):
            return (0, 0)
    
    return sorted(quarters, key=quarter_sort_key)

def calculate_days_until_sold(date_purchased, date_sold):
    """Calculate days between purchase and sale"""
    if pd.isna(date_purchased) or pd.isna(date_sold):
        return None
    return (date_sold - date_purchased).days

def calculate_realized_gross_profit(gross_sales_price, cost_basis, closing_costs):
    """Calculate realized gross profit"""
    gross_sales_price = safe_numeric_value(gross_sales_price, 0)
    cost_basis = safe_numeric_value(cost_basis, 0)
    closing_costs = safe_numeric_value(closing_costs, 0)
    return gross_sales_price - cost_basis - closing_costs

def calculate_realized_markup(gross_sales_price, cost_basis, closing_costs):
    """Calculate realized markup percentage"""
    gross_sales_price = safe_numeric_value(gross_sales_price, 0)
    cost_basis = safe_numeric_value(cost_basis, 0)
    closing_costs = safe_numeric_value(closing_costs, 0)
    
    total_cost = cost_basis + closing_costs
    if total_cost == 0:
        return 0
    return ((gross_sales_price / total_cost) - 1) * 100

def calculate_realized_margin(realized_gross_profit, gross_sales_price):
    """Calculate realized margin percentage"""
    realized_gross_profit = safe_numeric_value(realized_gross_profit, 0)
    gross_sales_price = safe_numeric_value(gross_sales_price, 0)
    
    if gross_sales_price == 0:
        return 0
    return (realized_gross_profit / gross_sales_price) * 100

def process_data(df):
    """Process the uploaded data and return processed data plus error report"""
    # Store original data for error tracking
    df_original = df.copy()
    total_records = len(df_original)
    
    # Initialize error tracking
    error_records = []
    
    # Work with original column names but create display versions
    df_processed = df.copy()
    
    # Convert date columns (using original field names)
    if 'custom.Asset_Date_Purchased' in df_processed.columns:
        df_processed['custom.Asset_Date_Purchased'] = df_processed['custom.Asset_Date_Purchased'].apply(parse_date)
    if 'custom.Asset_Date_Sold' in df_processed.columns:
        df_processed['custom.Asset_Date_Sold'] = df_processed['custom.Asset_Date_Sold'].apply(parse_date)
    
    # Track records that will be filtered out
    if 'primary_opportunity_status_label' in df_processed.columns:
        # Find records that are NOT "Sold"
        non_sold_mask = df_processed['primary_opportunity_status_label'] != 'Sold'
        non_sold_records = df_processed[non_sold_mask].copy()
        
        for idx, row in non_sold_records.iterrows():
            error_records.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': row.get('primary_opportunity_status_label', 'Missing'),
                'Error Type': 'Status Not "Sold"',
                'Error Detail': f'Status is "{row.get("primary_opportunity_status_label")}" instead of "Sold"',
                'Date Sold': row.get('custom.Asset_Date_Sold', 'N/A'),
                'Row Number': idx + 2  # +2 because of 0-indexing and header row
            })
        
        # Find records with null/missing status
        null_status_mask = df_processed['primary_opportunity_status_label'].isna()
        null_status_records = df_processed[null_status_mask].copy()
        
        for idx, row in null_status_records.iterrows():
            error_records.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': 'NULL/Missing',
                'Error Type': 'Missing Status',
                'Error Detail': 'Opportunity Status field is empty or null',
                'Date Sold': row.get('custom.Asset_Date_Sold', 'N/A'),
                'Row Number': idx + 2
            })
        
        # Filter for only "Sold" opportunities
        df_processed = df_processed[df_processed['primary_opportunity_status_label'] == 'Sold'].copy()
    
    # Track records with missing Date Sold
    if 'custom.Asset_Date_Sold' in df_processed.columns:
        missing_date_mask = df_processed['custom.Asset_Date_Sold'].isna()
        missing_date_records = df_processed[missing_date_mask].copy()
        
        for idx, row in missing_date_records.iterrows():
            error_records.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': row.get('primary_opportunity_status_label', 'Unknown'),
                'Error Type': 'Missing Date Sold',
                'Error Detail': 'Date Sold field is empty or null',
                'Date Sold': 'NULL/Missing',
                'Row Number': idx + 2
            })
        
        # Filter out records with missing Date Sold
        df_processed = df_processed[df_processed['custom.Asset_Date_Sold'].notna()].copy()
    
    # Add Quarter_Year column
    if 'custom.Asset_Date_Sold' in df_processed.columns:
        df_processed['Quarter_Year'] = df_processed['custom.Asset_Date_Sold'].apply(get_quarter_year)
    
    # Add calculated columns using original field names
    if all(col in df_processed.columns for col in ['custom.Asset_Date_Purchased', 'custom.Asset_Date_Sold']):
        df_processed['Days_Until_Sold'] = df_processed.apply(
            lambda row: calculate_days_until_sold(row['custom.Asset_Date_Purchased'], row['custom.Asset_Date_Sold']),
            axis=1
        )
    
    if all(col in df_processed.columns for col in ['custom.Asset_Gross_Sales_Price', 'custom.Asset_Cost_Basis', 'custom.Asset_Closing_Costs']):
        df_processed['Realized_Gross_Profit'] = df_processed.apply(
            lambda row: calculate_realized_gross_profit(
                row['custom.Asset_Gross_Sales_Price'],
                row['custom.Asset_Cost_Basis'],
                row['custom.Asset_Closing_Costs']
            ),
            axis=1
        )
        
        df_processed['Realized_Markup'] = df_processed.apply(
            lambda row: calculate_realized_markup(
                row['custom.Asset_Gross_Sales_Price'],
                row['custom.Asset_Cost_Basis'],
                row['custom.Asset_Closing_Costs']
            ),
            axis=1
        )
        
        df_processed['Realized_Margin'] = df_processed.apply(
            lambda row: calculate_realized_margin(
                row['Realized_Gross_Profit'],
                row['custom.Asset_Gross_Sales_Price']
            ),
            axis=1
        )
    
    # Create error dataframe
    error_df = pd.DataFrame(error_records) if error_records else pd.DataFrame()
    
    return df_processed, error_df, total_records

def format_currency(value):
    """Format value as currency"""
    value = safe_numeric_value(value, 0)
    return f"${value:,.2f}"

def format_percentage(value):
    """Format value as percentage"""
    value = safe_numeric_value(value, 0)
    return f"{value:.2f}%"

def create_summary_stats(df):
    """Create summary statistics for a dataframe"""
    stats = {}
    
    # Count of properties
    stats['total_properties'] = len(df)
    
    # Financial totals
    stats['total_gross_sales'] = safe_numeric_value(df['custom.Asset_Gross_Sales_Price'].sum(), 0)
    stats['total_cost_basis'] = safe_numeric_value(df['custom.Asset_Cost_Basis'].sum(), 0)
    stats['total_closing_costs'] = safe_numeric_value(df['custom.Asset_Closing_Costs'].sum(), 0)
    stats['total_gross_profit'] = safe_numeric_value(df['Realized_Gross_Profit'].sum(), 0)
    
    # Markup and Margin statistics (excluding infinite values)
    markup_data = df['Realized_Markup'].replace([np.inf, -np.inf], np.nan).dropna()
    margin_data = df['Realized_Margin'].replace([np.inf, -np.inf], np.nan).dropna()
    
    stats['average_markup'] = safe_numeric_value(markup_data.mean(), 0)
    stats['median_markup'] = safe_numeric_value(markup_data.median(), 0)
    stats['average_margin'] = safe_numeric_value(margin_data.mean(), 0)
    stats['median_margin'] = safe_numeric_value(margin_data.median(), 0)
    
    # Days to sell statistics
    days_data = df['Days_Until_Sold'].dropna()
    stats['average_days'] = safe_numeric_value(days_data.mean(), 0)
    stats['median_days'] = safe_numeric_value(days_data.median(), 0)
    stats['max_days'] = safe_numeric_value(days_data.max(), 0)
    stats['min_days'] = safe_numeric_value(days_data.min(), 0)
    
    return stats

def create_display_dataframe(df):
    """Create a display version of the dataframe with human-readable column names"""
    # Select and rename columns for display
    display_columns = {
        'display_name': 'Property Name',
        'custom.Asset_Owner': 'Owner',
        'custom.All_State': 'State',
        'custom.All_County': 'County',
        'custom.All_Asset_Surveyed_Acres': 'Acres',
        'custom.Asset_Cost_Basis': 'Cost Basis',
        'custom.Asset_Date_Purchased': 'Date Purchased',
        'custom.Asset_Date_Sold': 'Date Sold',
        'Days_Until_Sold': 'Days Until Sold',
        'custom.Asset_Gross_Sales_Price': 'Gross Sales Price',
        'custom.Asset_Closing_Costs': 'Closing Costs',
        'Realized_Gross_Profit': 'Realized Gross Profit',
        'Realized_Markup': 'Realized Markup %',
        'Realized_Margin': 'Realized Margin %'
    }
    
    # Create display dataframe with only the columns we want
    display_df = df.copy()
    
    # Select only columns that exist in the dataframe
    cols_to_keep = [col for col in display_columns.keys() if col in display_df.columns]
    display_df = display_df[cols_to_keep]
    
    # Rename columns
    display_df = display_df.rename(columns=display_columns)
    
    return display_df

def create_excel_download(df, filename):
    """Create Excel file for download with original field names"""
    output = BytesIO()
    
    # Ensure ID column is included
    cols_to_export = ['id'] + [col for col in df.columns if col != 'id']
    df_export = df[cols_to_export].copy()
    
    # Format date columns for Excel
    date_columns = ['custom.Asset_Date_Purchased', 'custom.Asset_Date_Sold']
    for col in date_columns:
        if col in df_export.columns:
            df_export[col] = df_export[col].apply(
                lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else ''
            )
    
    # Create Excel writer with xlsxwriter engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Sold Properties', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sold Properties']
        
        # Format currency columns
        currency_format = workbook.add_format({'num_format': '$#,##0.00'})
        percentage_format = workbook.add_format({'num_format': '0.00"%"'})
        
        # Apply formats to specific columns
        currency_cols = ['custom.Asset_Cost_Basis', 'custom.Asset_Gross_Sales_Price', 
                        'custom.Asset_Closing_Costs', 'Realized_Gross_Profit']
        percentage_cols = ['Realized_Markup', 'Realized_Margin']
        
        for idx, col in enumerate(df_export.columns):
            if col in currency_cols:
                worksheet.set_column(idx, idx, 18, currency_format)
            elif col in percentage_cols:
                worksheet.set_column(idx, idx, 15, percentage_format)
            else:
                worksheet.set_column(idx, idx, 15)
    
    output.seek(0)
    return output

def create_error_report_excel(error_df, total_records, processed_records, filename):
    """Create Excel file for error report"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Add summary sheet
        summary_data = {
            'Metric': ['Total Records in Upload', 'Successfully Processed Records', 'Records with Errors', 'Success Rate'],
            'Value': [
                total_records,
                processed_records,
                len(error_df),
                f"{(processed_records / total_records * 100):.2f}%" if total_records > 0 else "0%"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Add error details sheet
        if len(error_df) > 0:
            error_df.to_excel(writer, sheet_name='Error Details', index=False)
            
            # Format the error details worksheet
            workbook = writer.book
            worksheet = writer.sheets['Error Details']
            
            # Set column widths
            worksheet.set_column('A:A', 15)  # ID
            worksheet.set_column('B:B', 30)  # Property Name
            worksheet.set_column('C:C', 20)  # Owner
            worksheet.set_column('D:D', 20)  # Opportunity Status
            worksheet.set_column('E:E', 20)  # Error Type
            worksheet.set_column('F:F', 50)  # Error Detail
            worksheet.set_column('G:G', 15)  # Date Sold
            worksheet.set_column('H:H', 12)  # Row Number
        
        # Format the summary worksheet
        workbook = writer.book
        worksheet = writer.sheets['Summary']
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 20)
    
    output.seek(0)
    return output

def create_pdf_download(quarter_data_dict, filename):
    """Create PDF file for download"""
    if not REPORTLAB_AVAILABLE:
        return None
    
    output = BytesIO()
    
    # Create PDF with landscape legal page size
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(legal),
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    # Container for PDF elements
    elements = []
    
    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=30,
        alignment=1  # Center alignment
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=12,
        spaceBefore=12
    )
    
    # Title
    title_text = f"Sold Property Report - {datetime.now().strftime('%B %d, %Y')}"
    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 0.2*inch))
    
    # Iterate through quarters in chronological order
    for quarter, df in quarter_data_dict.items():
        if len(df) == 0:
            continue
            
        # Quarter heading
        elements.append(Paragraph(f"{quarter}", heading_style))
        
        # Prepare data for table
        table_data = []
        
        # Headers
        headers = ['Property', 'Owner', 'State', 'County', 'Acres', 'Cost Basis', 
                  'Date Purchased', 'Date Sold', 'Days', 'Gross Sales', 
                  'Closing Costs', 'Gross Profit', 'Markup %', 'Margin %']
        table_data.append(headers)
        
        # Data rows
        for _, row in df.iterrows():
            table_data.append([
                str(row.get('Property Name', ''))[:20],  # Truncate long names
                str(row.get('Owner', ''))[:15],
                str(row.get('State', '')),
                str(row.get('County', ''))[:15],
                f"{safe_numeric_value(row.get('Acres', 0)):.1f}",
                format_currency(row.get('Cost Basis', 0)),
                row.get('Date Purchased', ''),
                row.get('Date Sold', ''),
                f"{safe_int_value(row.get('Days Until Sold', 0))}",
                format_currency(row.get('Gross Sales Price', 0)),
                format_currency(row.get('Closing Costs', 0)),
                format_currency(row.get('Realized Gross Profit', 0)),
                format_percentage(row.get('Realized Markup %', 0)),
                format_percentage(row.get('Realized Margin %', 0))
            ])
        
        # Create table
        table = Table(table_data)
        
        # Table style
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elements.append(table)
        
        # Summary statistics
        stats = create_summary_stats(df)
        elements.append(Spacer(1, 0.2*inch))
        
        summary_text = f"""
        <b>Quarter Summary:</b> {stats['total_properties']} properties | 
        Total Sales: {format_currency(stats['total_gross_sales'])} | 
        Total Profit: {format_currency(stats['total_gross_profit'])} | 
        Avg Markup: {format_percentage(stats['average_markup'])} | 
        Avg Margin: {format_percentage(stats['average_margin'])} | 
        Avg Days to Sell: {safe_numeric_value(stats['average_days']):.0f}
        """
        
        elements.append(Paragraph(summary_text, styles['Normal']))
        elements.append(PageBreak())
    
    # Disclaimer on last page
    disclaimer_style = ParagraphStyle(
        'Disclaimer',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.grey,
        alignment=1
    )
    
    elements.append(Spacer(1, 0.5*inch))
    elements.append(Paragraph(
        "This data is sourced from our CRM and not our accounting software, based on then-available data. "
        "Final accounting data and results may vary slightly.",
        disclaimer_style
    ))
    
    # Build PDF
    doc.build(elements)
    output.seek(0)
    
    return output

def main():
    st.title("📊 Sold Property Report Generator")
    st.markdown("Upload your Close.com CSV export to generate comprehensive sold property reports")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Close.com CSV Export", type=['csv'])
    
    if uploaded_file is not None:
        try:
            # Read the CSV
            df = pd.read_csv(uploaded_file)
            
            # Process the data
            df_processed, error_df, total_records = process_data(df)
            
            # Display error summary if there are errors
            if len(error_df) > 0:
                with st.expander(f"⚠️ Import Warnings ({len(error_df)} records excluded)", expanded=False):
                    st.warning(f"**{len(error_df)} records** could not be included in the report. Common reasons:")
                    st.write("- Opportunity Status is not 'Sold'")
                    st.write("- Missing Date Sold field")
                    st.write("- Other data quality issues")
                    st.write("")
                    st.write("Download the Error Report below to see detailed information about excluded records.")
                    
                    # Show preview of errors
                    st.dataframe(error_df.head(10), use_container_width=True)
                    if len(error_df) > 10:
                        st.info(f"Showing first 10 of {len(error_df)} errors. Download the full Error Report for complete details.")
            
            # Check if we have any data to display
            if len(df_processed) == 0:
                st.error("No valid sold properties found in the uploaded file. Please check the Error Report for details.")
                
                # Still offer error report download
                if len(error_df) > 0:
                    st.subheader("Download Error Report")
                    current_date = datetime.now().strftime("%Y%m%d")
                    error_filename = f"{current_date} Import Error Report.xlsx"
                    error_file = create_error_report_excel(error_df, total_records, len(df_processed), error_filename)
                    
                    st.download_button(
                        label="Download Error Report",
                        data=error_file.getvalue(),
                        file_name=error_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=f"Download detailed report of {len(error_df)} records that could not be processed"
                    )
                return
            
            st.success(f"Successfully processed {len(df_processed)} sold properties from {total_records} total records")
            
            # Get unique values for filters
            all_quarters = sorted(df_processed['Quarter_Year'].dropna().unique().tolist(), 
                                 key=lambda x: sort_quarters_chronologically([x])[0] if x else (0, 0))
            all_owners = sorted(df_processed['custom.Asset_Owner'].dropna().unique().tolist())
            
            # Initialize session state for filters if not exists
            if 'selected_quarters' not in st.session_state:
                st.session_state.selected_quarters = all_quarters
            if 'selected_owners' not in st.session_state:
                st.session_state.selected_owners = all_owners
            
            # Filters
            st.sidebar.header("Filters")
            
            # Quarter filter with Select All/None buttons
            st.sidebar.subheader("Select Quarters")
            
            col1, col2 = st.sidebar.columns(2)
            with col1:
                if st.button("Select All Quarters", key="quarters_all_btn"):
                    st.session_state.selected_quarters = all_quarters
            with col2:
                if st.button("Select None", key="quarters_none_btn"):
                    st.session_state.selected_quarters = []
            
            selected_quarters = st.sidebar.multiselect(
                "Quarters",
                options=all_quarters,
                default=st.session_state.selected_quarters,
                key='selected_quarters'
            )
            
            # Owner filter with Select All/None buttons
            st.sidebar.subheader("Select Owners")
            
            col1, col2 = st.sidebar.columns(2)
            with col1:
                if st.button("Select All Owners", key="owners_all_btn"):
                    st.session_state.selected_owners = all_owners
            with col2:
                if st.button("Select None", key="owners_none_btn"):
                    st.session_state.selected_owners = []
            
            selected_owners = st.sidebar.multiselect(
                "Owners",
                options=all_owners,
                default=st.session_state.selected_owners,
                key='selected_owners'
            )
            
            # Apply filters
            filtered_df = df_processed[
                (df_processed['Quarter_Year'].isin(selected_quarters)) &
                (df_processed['custom.Asset_Owner'].isin(selected_owners))
            ].copy()
            
            # Create display version (with renamed columns)
            filtered_df_display = create_display_dataframe(filtered_df)
            filtered_df_display['Quarter_Year'] = filtered_df['Quarter_Year'].values
            
            # Keep original version for Excel export
            filtered_df_original = filtered_df.copy()
            
            if len(filtered_df_display) == 0:
                st.warning("No properties match the selected filters. Please adjust your filter criteria.")
            else:
                st.subheader(f"Sold Properties Report ({len(filtered_df_display)} properties)")
                
                # Display data by quarter in chronological order
                sorted_selected_quarters = sort_quarters_chronologically(selected_quarters)
                
                for quarter in sorted_selected_quarters:
                    quarter_data = filtered_df_display[filtered_df_display['Quarter_Year'] == quarter].copy()
                    
                    if len(quarter_data) == 0:
                        continue
                    
                    st.markdown(f"### {quarter}")
                    
                    # Remove Quarter_Year column for display
                    display_df = quarter_data.drop(columns=['Quarter_Year'])
                    
                    # Format currency and percentage columns for display
                    currency_cols = ['Cost Basis', 'Gross Sales Price', 'Closing Costs', 'Realized Gross Profit']
                    percentage_cols = ['Realized Markup %', 'Realized Margin %']
                    
                    for col in currency_cols:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].apply(lambda x: format_currency(x) if pd.notna(x) else '')
                    
                    for col in percentage_cols:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].apply(lambda x: format_percentage(x) if pd.notna(x) else '')
                    
                    # Format dates
                    if 'Date Purchased' in display_df.columns:
                        display_df['Date Purchased'] = display_df['Date Purchased'].apply(
                            lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else ''
                        )
                    if 'Date Sold' in display_df.columns:
                        display_df['Date Sold'] = display_df['Date Sold'].apply(
                            lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else ''
                        )
                    
                    st.dataframe(display_df, use_container_width=True)
                    
                    # Summary statistics for quarter
                    stats = create_summary_stats(filtered_df[filtered_df['Quarter_Year'] == quarter])
                    
                    # First row: Total Properties (centered)
                    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
                    with col2:
                        st.metric("Total Properties", stats['total_properties'])
                    
                    # Second row: Financial totals
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Gross Sales", format_currency(stats['total_gross_sales']))
                    with col2:
                        st.metric("Total Cost Basis", format_currency(stats['total_cost_basis']))
                    with col3:
                        st.metric("Total Closing Costs", format_currency(stats['total_closing_costs']))
                    with col4:
                        st.metric("Total Gross Profit", format_currency(stats['total_gross_profit']))
                    
                    # Third row: Markup and Margin percentages
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Average Markup", format_percentage(stats['average_markup']))
                    with col2:
                        st.metric("Median Markup", format_percentage(stats['median_markup']))
                    with col3:
                        st.metric("Average Margin", format_percentage(stats['average_margin']))
                    with col4:
                        st.metric("Median Margin", format_percentage(stats['median_margin']))
                    
                    # Fourth row: Days to sell metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Average Days to Sell", f"{safe_numeric_value(stats['average_days']):.0f}")
                    with col2:
                        st.metric("Median Days to Sell", f"{safe_numeric_value(stats['median_days']):.0f}")
                    with col3:
                        st.metric("Max Days to Sell", f"{safe_numeric_value(stats['max_days']):.0f}")
                    with col4:
                        st.metric("Min Days to Sell", f"{safe_numeric_value(stats['min_days']):.0f}")
                    
                    st.divider()
                
                # Overall summary
                if len(selected_quarters) > 1:
                    st.markdown("### Overall Summary")
                    overall_stats = create_summary_stats(filtered_df)
                    
                    # First row: Total Properties (centered)
                    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
                    with col2:
                        st.metric("Total Properties", overall_stats['total_properties'])
                    
                    # Second row: Financial totals
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Gross Sales", format_currency(overall_stats['total_gross_sales']))
                    with col2:
                        st.metric("Total Cost Basis", format_currency(overall_stats['total_cost_basis']))
                    with col3:
                        st.metric("Total Closing Costs", format_currency(overall_stats['total_closing_costs']))
                    with col4:
                        st.metric("Total Gross Profit", format_currency(overall_stats['total_gross_profit']))
                    
                    # Third row: Markup and Margin percentages
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Average Markup", format_percentage(overall_stats['average_markup']))
                    with col2:
                        st.metric("Median Markup", format_percentage(overall_stats['median_markup']))
                    with col3:
                        st.metric("Average Margin", format_percentage(overall_stats['average_margin']))
                    with col4:
                        st.metric("Median Margin", format_percentage(overall_stats['median_margin']))
                    
                    # Fourth row: Days to sell metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Average Days to Sell", f"{safe_numeric_value(overall_stats['average_days']):.0f}")
                    with col2:
                        st.metric("Median Days to Sell", f"{safe_numeric_value(overall_stats['median_days']):.0f}")
                    with col3:
                        st.metric("Max Days to Sell", f"{safe_numeric_value(overall_stats['max_days']):.0f}")
                    with col4:
                        st.metric("Min Days to Sell", f"{safe_numeric_value(overall_stats['min_days']):.0f}")
                
                # Download section
                st.subheader("Download Reports")
                
                # Prepare data for PDF (organized by quarter in chronological order) - using display data
                quarter_data_dict = {}
                sorted_selected_quarters = sort_quarters_chronologically(selected_quarters)
                for quarter in sorted_selected_quarters:
                    quarter_data_dict[quarter] = filtered_df_display[filtered_df_display['Quarter_Year'] == quarter].copy()
                
                # Generate filenames
                current_date = datetime.now().strftime("%Y%m%d")
                excel_filename = f"{current_date} Sold Property Report.xlsx"
                pdf_filename = f"{current_date} Sold Property Report.pdf"
                error_filename = f"{current_date} Import Error Report.xlsx"
                
                # Create columns for downloads
                if len(error_df) > 0:
                    col1, col2, col3 = st.columns(3)
                else:
                    col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Excel Report (Original Field Names)**")
                    # Create Excel file using original field names and including ID
                    excel_file = create_excel_download(filtered_df_original, excel_filename)
                    
                    st.download_button(
                        label="Download Excel Report",
                        data=excel_file.getvalue(),
                        file_name=excel_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Excel file with original Close.com field names and ID column for data correction"
                    )
                
                with col2:
                    st.write("**PDF Report (Landscape Legal)**")
                    if REPORTLAB_AVAILABLE:
                        # Create PDF file using display data
                        pdf_file = create_pdf_download(quarter_data_dict, pdf_filename)
                        
                        if pdf_file:
                            st.download_button(
                                label="Download PDF Report",
                                data=pdf_file.getvalue(),
                                file_name=pdf_filename,
                                mime="application/pdf"
                            )
                    else:
                        st.warning("PDF generation requires reportlab. Run: pip install reportlab")
                        st.info("Excel download is still available above.")
                
                # Error report download (only show if there are errors)
                if len(error_df) > 0:
                    with col3:
                        st.write("**Error Report**")
                        # Create error report file
                        error_file = create_error_report_excel(error_df, total_records, len(df_processed), error_filename)
                        
                        st.download_button(
                            label="Download Error Report",
                            data=error_file.getvalue(),
                            file_name=error_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help=f"Download detailed report of {len(error_df)} records that could not be processed"
                        )
                
                # Disclaimer
                st.markdown("---")
                st.markdown("**Disclaimer:** This data is sourced from our CRM and not our accounting software, based on then-available data. Final accounting data and results may vary slightly.")
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.write("Please make sure you've uploaded a valid CSV export from Close.com with all fields included.")
    
    else:
        st.info("Please upload your Close.com CSV export to generate the sold property report")

if __name__ == "__main__":
    main()

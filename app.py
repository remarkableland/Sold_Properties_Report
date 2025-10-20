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
        
        # Filter to only sold properties
        df_processed = df_processed[df_processed['primary_opportunity_status_label'] == 'Sold'].copy()
    
    # Track records with missing or invalid Date Sold
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
                'Error Detail': 'Date Sold field is empty, null, or could not be parsed',
                'Date Sold': 'Invalid/Missing',
                'Row Number': idx + 2
            })
    
    # Convert numeric columns with safe handling (using original field names)
    numeric_columns = ['custom.All_Asset_Surveyed_Acres', 'custom.Asset_Cost_Basis', 
                      'custom.Asset_Gross_Sales_Price', 'custom.Asset_Closing_Costs']
    for col in numeric_columns:
        if col in df_processed.columns:
            df_processed[col] = df_processed[col].apply(lambda x: safe_numeric_value(x, 0))
    
    # Calculate derived fields with safe handling (using original field names)
    df_processed['Days_Until_Sold'] = df_processed.apply(
        lambda row: calculate_days_until_sold(row.get('custom.Asset_Date_Purchased'), row.get('custom.Asset_Date_Sold')), axis=1
    )
    
    df_processed['Realized_Gross_Profit'] = df_processed.apply(
        lambda row: calculate_realized_gross_profit(
            row.get('custom.Asset_Gross_Sales_Price', 0), 
            row.get('custom.Asset_Cost_Basis', 0), 
            row.get('custom.Asset_Closing_Costs', 0)
        ), axis=1
    )
    
    df_processed['Realized_Markup'] = df_processed.apply(
        lambda row: calculate_realized_markup(
            row.get('custom.Asset_Gross_Sales_Price', 0), 
            row.get('custom.Asset_Cost_Basis', 0), 
            row.get('custom.Asset_Closing_Costs', 0)
        ), axis=1
    )
    
    df_processed['Realized_Margin'] = df_processed.apply(
        lambda row: calculate_realized_margin(
            row.get('Realized_Gross_Profit', 0), 
            row.get('custom.Asset_Gross_Sales_Price', 0)
        ), axis=1
    )
    
    # Add quarter-year for filtering
    df_processed['Quarter_Year'] = df_processed['custom.Asset_Date_Sold'].apply(get_quarter_year)
    
    # Create display versions for UI (mapped names)
    df_display = df_processed.copy()
    columns_to_rename = {k: v for k, v in FIELD_MAPPING.items() if k in df_display.columns}
    df_display = df_display.rename(columns=columns_to_rename)
    
    # Rename calculated fields for display
    df_display = df_display.rename(columns={
        'Days_Until_Sold': 'Days Until Sold',
        'Realized_Gross_Profit': 'Realized Gross Profit',
        'Realized_Markup': 'Realized Markup',
        'Realized_Margin': 'Realized Margin'
    })
    
    # Create error report DataFrame
    error_df = pd.DataFrame(error_records)
    
    return df_processed, df_display, error_df, total_records

def create_excel_download(df_original, filename):
    """Create Excel file for download with original Close.com field names and ID column"""
    output = BytesIO()
    
    # Create a workbook with NaN/INF handling enabled
    workbook = xlsxwriter.Workbook(output, {
        'in_memory': True,
        'nan_inf_to_errors': True  # This option handles NaN/INF values
    })
    worksheet = workbook.add_worksheet('Sold Properties')
    
    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1
    })
    
    currency_format = workbook.add_format({
        'num_format': '$#,##0',
        'border': 1
    })
    
    percentage_format = workbook.add_format({
        'num_format': '0%',
        'border': 1
    })
    
    date_format = workbook.add_format({
        'num_format': 'mm/dd/yyyy',
        'border': 1
    })
    
    number_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })
    
    # Highlight formats for missing/zero values
    currency_highlight_format = workbook.add_format({
        'num_format': '$#,##0',
        'border': 1,
        'bg_color': '#FFFF99'  # Yellow background
    })
    
    percentage_highlight_format = workbook.add_format({
        'num_format': '0%',
        'border': 1,
        'bg_color': '#FFFF99'  # Yellow background
    })
    
    date_highlight_format = workbook.add_format({
        'num_format': 'mm/dd/yyyy',
        'border': 1,
        'bg_color': '#FFFF99'  # Yellow background
    })
    
    number_highlight_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1,
        'bg_color': '#FFFF99'  # Yellow background
    })
    
    text_highlight_format = workbook.add_format({
        'border': 1,
        'bg_color': '#FFFF99'  # Yellow background
    })
    
    # Headers using original Close.com field names, including ID
    headers = ['id', 'display_name', 'custom.Asset_Owner', 'custom.All_State', 'custom.All_County', 
               'custom.All_Asset_Surveyed_Acres', 'custom.Asset_Cost_Basis', 'custom.Asset_Date_Purchased',
               'primary_opportunity_status_label', 'Days_Until_Sold', 'custom.Asset_Date_Sold', 
               'custom.Asset_Gross_Sales_Price', 'custom.Asset_Closing_Costs',
               'Realized_Gross_Profit', 'Realized_Markup', 'Realized_Margin']
    
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    
    # Write data with safe value handling
    for row, (_, data) in enumerate(df_original.iterrows(), 1):
        # ID field - highlight if empty
        lead_id = str(data.get('id', ''))
        if lead_id == '' or lead_id == 'nan':
            worksheet.write(row, 0, '', text_highlight_format)
        else:
            worksheet.write(row, 0, lead_id)
        
        # Property Name - highlight if empty
        prop_name = str(data.get('display_name', ''))
        if prop_name == '' or prop_name == 'nan':
            worksheet.write(row, 1, '', text_highlight_format)
        else:
            worksheet.write(row, 1, prop_name)
        
        # Owner - highlight if empty
        owner = str(data.get('custom.Asset_Owner', ''))
        if owner == '' or owner == 'nan':
            worksheet.write(row, 2, '', text_highlight_format)
        else:
            worksheet.write(row, 2, owner)
        
        # State - highlight if empty
        state = str(data.get('custom.All_State', ''))
        if state == '' or state == 'nan':
            worksheet.write(row, 3, '', text_highlight_format)
        else:
            worksheet.write(row, 3, state)
        
        # County - highlight if empty
        county = str(data.get('custom.All_County', ''))
        if county == '' or county == 'nan':
            worksheet.write(row, 4, '', text_highlight_format)
        else:
            worksheet.write(row, 4, county)
        
        # Handle Acres with safe numeric conversion
        acres = safe_numeric_value(data.get('custom.All_Asset_Surveyed_Acres', 0))
        if acres == 0:
            worksheet.write(row, 5, 0, number_highlight_format)
        else:
            worksheet.write(row, 5, acres, number_format)
        
        # Handle Cost Basis with safe numeric conversion
        cost_basis = safe_numeric_value(data.get('custom.Asset_Cost_Basis', 0))
        if cost_basis == 0:
            worksheet.write(row, 6, 0, currency_highlight_format)
        else:
            worksheet.write(row, 6, cost_basis, currency_format)
        
        # Handle Date Purchased with null check and highlighting
        date_purchased = data.get('custom.Asset_Date_Purchased')
        if pd.isna(date_purchased) or date_purchased == '':
            worksheet.write(row, 7, '', date_highlight_format)
        else:
            worksheet.write(row, 7, date_purchased, date_format)
            
        # Opportunity Status - highlight if empty
        opp_status = str(data.get('primary_opportunity_status_label', ''))
        if opp_status == '' or opp_status == 'nan':
            worksheet.write(row, 8, '', text_highlight_format)
        else:
            worksheet.write(row, 8, opp_status)
        
        # Handle Days Until Sold with safe conversion
        days_sold = safe_int_value(data.get('Days_Until_Sold', 0))
        if days_sold == 0:
            worksheet.write(row, 9, 0, number_highlight_format)
        else:
            worksheet.write(row, 9, days_sold, number_format)
        
        # Handle Date Sold with null check and highlighting
        date_sold = data.get('custom.Asset_Date_Sold')
        if pd.isna(date_sold) or date_sold == '':
            worksheet.write(row, 10, '', date_highlight_format)
        else:
            worksheet.write(row, 10, date_sold, date_format)
        
        # Handle Gross Sales Price with safe numeric conversion
        gross_sales = safe_numeric_value(data.get('custom.Asset_Gross_Sales_Price', 0))
        if gross_sales == 0:
            worksheet.write(row, 11, 0, currency_highlight_format)
        else:
            worksheet.write(row, 11, gross_sales, currency_format)
        
        # Handle Closing Costs with safe numeric conversion
        closing_costs = safe_numeric_value(data.get('custom.Asset_Closing_Costs', 0))
        if closing_costs == 0:
            worksheet.write(row, 12, 0, currency_highlight_format)
        else:
            worksheet.write(row, 12, closing_costs, currency_format)
        
        # Handle Realized Gross Profit with safe numeric conversion
        gross_profit = safe_numeric_value(data.get('Realized_Gross_Profit', 0))
        if gross_profit == 0:
            worksheet.write(row, 13, 0, currency_highlight_format)
        else:
            worksheet.write(row, 13, gross_profit, currency_format)
        
        # Handle Realized Markup with safe numeric conversion
        markup = safe_numeric_value(data.get('Realized_Markup', 0))
        if markup == 0:
            worksheet.write(row, 14, 0, percentage_highlight_format)
        else:
            worksheet.write(row, 14, markup / 100, percentage_format)
        
        # Handle Realized Margin with safe numeric conversion
        margin = safe_numeric_value(data.get('Realized_Margin', 0))
        if margin == 0:
            worksheet.write(row, 15, 0, percentage_highlight_format)
        else:
            worksheet.write(row, 15, margin / 100, percentage_format)
    
    # Auto-adjust column widths
    for col in range(len(headers)):
        if col == 0:  # ID column
            worksheet.set_column(col, col, 12)
        elif col == 1:  # Property name
            worksheet.set_column(col, col, 25)
        else:
            worksheet.set_column(col, col, 15)
    
    workbook.close()
    output.seek(0)
    
    return output

def create_error_report_excel(error_df, total_records, processed_records, filename):
    """Create Excel error report for records that didn't import"""
    output = BytesIO()
    
    # Create a workbook with NaN/INF handling enabled
    workbook = xlsxwriter.Workbook(output, {
        'in_memory': True,
        'nan_inf_to_errors': True
    })
    
    # Summary worksheet
    summary_worksheet = workbook.add_worksheet('Import Summary')
    
    # Error details worksheet
    error_worksheet = workbook.add_worksheet('Error Details')
    
    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D32F2F',
        'font_color': 'white',
        'border': 1
    })
    
    summary_header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#1976D2',
        'font_color': 'white',
        'border': 1
    })
    
    error_format = workbook.add_format({
        'bg_color': '#FFEBEE',
        'border': 1
    })
    
    success_format = workbook.add_format({
        'bg_color': '#E8F5E8',
        'border': 1
    })
    
    # Summary worksheet content
    summary_worksheet.write(0, 0, 'Import Summary Report', summary_header_format)
    summary_worksheet.write(2, 0, 'Metric', summary_header_format)
    summary_worksheet.write(2, 1, 'Value', summary_header_format)
    
    summary_worksheet.write(3, 0, 'Total Records in CSV', success_format)
    summary_worksheet.write(3, 1, total_records, success_format)
    
    summary_worksheet.write(4, 0, 'Successfully Processed', success_format)
    summary_worksheet.write(4, 1, processed_records, success_format)
    
    summary_worksheet.write(5, 0, 'Records with Errors', error_format)
    summary_worksheet.write(5, 1, len(error_df), error_format)
    
    success_rate = (processed_records / total_records * 100) if total_records > 0 else 0
    summary_worksheet.write(6, 0, 'Success Rate', summary_header_format)
    summary_worksheet.write(6, 1, f"{success_rate:.1f}%", summary_header_format)
    
    # Error breakdown by type
    if len(error_df) > 0:
        summary_worksheet.write(8, 0, 'Error Breakdown by Type', summary_header_format)
        summary_worksheet.write(9, 0, 'Error Type', summary_header_format)
        summary_worksheet.write(9, 1, 'Count', summary_header_format)
        
        error_counts = error_df['Error Type'].value_counts()
        for idx, (error_type, count) in enumerate(error_counts.items()):
            summary_worksheet.write(10 + idx, 0, error_type, error_format)
            summary_worksheet.write(10 + idx, 1, count, error_format)
    
    # Set column widths for summary
    summary_worksheet.set_column(0, 0, 25)
    summary_worksheet.set_column(1, 1, 15)
    
    # Error details worksheet
    if len(error_df) > 0:
        headers = ['Row Number', 'ID', 'Property Name', 'Owner', 'Opportunity Status', 'Error Type', 'Error Detail', 'Date Sold']
        
        # Write headers
        for col, header in enumerate(headers):
            error_worksheet.write(0, col, header, header_format)
        
        # Write error data with safe handling
        for row, (_, data) in enumerate(error_df.iterrows(), 1):
            error_worksheet.write(row, 0, safe_int_value(data.get('Row Number', 0)), error_format)
            error_worksheet.write(row, 1, str(data.get('ID', '')), error_format)
            error_worksheet.write(row, 2, str(data.get('Property Name', '')), error_format)
            error_worksheet.write(row, 3, str(data.get('Owner', '')), error_format)
            error_worksheet.write(row, 4, str(data.get('Opportunity Status', '')), error_format)
            error_worksheet.write(row, 5, str(data.get('Error Type', '')), error_format)
            error_worksheet.write(row, 6, str(data.get('Error Detail', '')), error_format)
            error_worksheet.write(row, 7, str(data.get('Date Sold', '')), error_format)
        
        # Set column widths for error details
        error_worksheet.set_column(0, 0, 10)  # Row Number
        error_worksheet.set_column(1, 1, 12)  # ID
        error_worksheet.set_column(2, 2, 25)  # Property Name
        error_worksheet.set_column(3, 3, 20)  # Owner
        error_worksheet.set_column(4, 4, 15)  # Opportunity Status
        error_worksheet.set_column(5, 5, 20)  # Error Type
        error_worksheet.set_column(6, 6, 40)  # Error Detail
        error_worksheet.set_column(7, 7, 15)  # Date Sold
    else:
        error_worksheet.write(0, 0, 'No errors found - all records processed successfully!', success_format)
    
    workbook.close()
    output.seek(0)
    
    return output

def create_pdf_download(df_dict, filename):
    """Create PDF file organized by quarter"""
    if not REPORTLAB_AVAILABLE:
        st.error("PDF generation requires reportlab. Please install it: pip install reportlab")
        return None

    buffer = BytesIO()
    
    # Create the PDF document with landscape legal page size
    doc = SimpleDocTemplate(buffer, pagesize=landscape(legal),
                          topMargin=0.5*inch, bottomMargin=0.5*inch,
                          leftMargin=0.5*inch, rightMargin=0.5*inch)
    
    story = []
    
    # Get styles
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        alignment=1,  # Center alignment
        textColor=colors.darkblue
    )
    
    quarter_style = ParagraphStyle(
        'QuarterTitle',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=8,
        spaceBefore=12,
        textColor=colors.darkblue
    )
    
    summary_style = ParagraphStyle(
        'SummaryStyle',
        parent=styles['Normal'],
        fontSize=10,
        spaceAfter=6
    )
    
    disclaimer_style = ParagraphStyle(
        'DisclaimerStyle',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.grey,
        alignment=1,  # Center alignment
        spaceAfter=12
    )
    
    # Cell text style for wrapping
    cell_style = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontSize=8,
        leading=9
    )
    
    # Title
    story.append(Paragraph("Remarkable Land LLC - Sold Properties Report", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", summary_style))
    story.append(Spacer(1, 12))
    
    # Process each quarter (in chronological order)
    for quarter in sort_quarters_chronologically(df_dict.keys()):
        quarter_data = df_dict[quarter]
        
        if len(quarter_data) == 0:
            continue
            
        # Quarter title
        story.append(Paragraph(f"{quarter}", quarter_style))
        
        # Prepare table data with wrapped headers (using display names for PDF)
        table_headers = ['Property\nName', 'Owner', 'State', 'County', 'Acres', 'Cost\nBasis',
                        'Date\nPurchased', 'Days\nUntil Sold', 'Date\nSold', 'Gross Sales\nPrice',
                        'Closing\nCosts', 'Realized Gross\nProfit', 'Realized\nMarkup', 'Realized\nMargin']
        
        table_data = [table_headers]
        
        for _, row in quarter_data.iterrows():
            # Create wrapped property name
            prop_name = str(row.get('Property Name', ''))
            if len(prop_name) > 25:
                prop_name_para = Paragraph(prop_name, cell_style)
            else:
                prop_name_para = prop_name
            
            # Safe formatting for all values (using display column names)
            formatted_row = [
                prop_name_para,  # Property name with wrapping
                str(row.get('Owner', ''))[:18],  # Allow longer owner names
                str(row.get('State', '')),
                str(row.get('County', ''))[:15],  # Allow longer county names
                f"{safe_numeric_value(row.get('Acres', 0)):.1f}",
                f"${safe_numeric_value(row.get('Cost Basis', 0)):,.0f}",
                row.get('Date Purchased').strftime('%m/%d/%Y') if pd.notna(row.get('Date Purchased')) else '',
                f"{safe_int_value(row.get('Days Until Sold', 0))}",
                row.get('Date Sold').strftime('%m/%d/%Y') if pd.notna(row.get('Date Sold')) else '',
                f"${safe_numeric_value(row.get('Gross Sales Price', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Closing Costs', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Realized Gross Profit', 0)):,.0f}",
                f"{safe_numeric_value(row.get('Realized Markup', 0)):.0f}%",
                f"{safe_numeric_value(row.get('Realized Margin', 0)):.0f}%"
            ]
            table_data.append(formatted_row)
        
        # Create table with wider columns for landscape legal (14" x 8.5")
        # Total usable width: approximately 13" (14" - 1" margins)
        col_widths = [2.0*inch, 1.2*inch, 0.4*inch, 0.9*inch, 0.5*inch, 0.9*inch, 
                     0.9*inch, 0.6*inch, 0.9*inch, 1.1*inch, 0.9*inch, 1.1*inch, 0.6*inch, 0.6*inch]
        
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Table style with better formatting for landscape legal
        table.setStyle(TableStyle([
            # Header formatting
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),  # Larger header font
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            # Data formatting
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),  # Larger data font
            ('ALIGN', (0, 1), (1, -1), 'LEFT'),  # Property Name and Owner left-aligned
            ('ALIGN', (2, 1), (3, -1), 'CENTER'),  # State and County centered
            ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),  # All numbers right-aligned
            ('GRID', (0, 0), (-1, -1), 0.8, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 12))
        
        # Quarter summary statistics (using display column names)
        stats = create_summary_stats(quarter_data)
        
        summary_data = [
            ['Metric', 'Value', 'Metric', 'Value'],
            ['Properties Sold', f"{stats['total_properties']}", 'Total Cost Basis', f"${stats['total_cost_basis']:,.0f}"],
            ['Total Gross Sales', f"${stats['total_gross_sales']:,.0f}", 'Total Gross Profit', f"${stats['total_gross_profit']:,.0f}"],
            ['Average Markup', f"{stats['average_markup']:.0f}%", 'Median Markup', f"{stats['median_markup']:.0f}%"],
            ['Average Margin', f"{stats['average_margin']:.0f}%", 'Median Margin', f"{stats['median_margin']:.0f}%"],
            ['Average Days to Sell', f"{stats['average_days']:.0f}", 'Median Days to Sell', f"{stats['median_days']:.0f}"]
        ]
        
        summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(summary_table)
        
        # Add page break between quarters (except for the last one)
        sorted_quarters = sort_quarters_chronologically(df_dict.keys())
        if quarter != sorted_quarters[-1]:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 12))
    
    # Overall summary if multiple quarters
    if len(df_dict) > 1:
        # Combine all data for overall summary
        all_data = pd.concat(df_dict.values(), ignore_index=True)
        overall_stats = create_summary_stats(all_data)
        
        story.append(Paragraph("Overall Summary", quarter_style))
        
        # Create a 4-column layout similar to web version
        overall_summary_data = [
            # Row 1: Total Properties (centered across all columns)
            ['Total Properties', f"{overall_stats['total_properties']}", '', ''],
            # Row 2: Financial totals headers
            ['Total Gross Sales', 'Total Cost Basis', 'Total Closing Costs', 'Total Gross Profit'],
            # Row 3: Financial totals values
            [f"${overall_stats['total_gross_sales']:,.0f}", f"${overall_stats['total_cost_basis']:,.0f}", f"${overall_stats['total_closing_costs']:,.0f}", f"${overall_stats['total_gross_profit']:,.0f}"],
            # Row 4: Markup and Margin headers
            ['Average Markup', 'Median Markup', 'Average Margin', 'Median Margin'],
            # Row 5: Markup and Margin values
            [f"{overall_stats['average_markup']:.0f}%", f"{overall_stats['median_markup']:.0f}%", f"{overall_stats['average_margin']:.0f}%", f"{overall_stats['median_margin']:.0f}%"],
            # Row 6: Days to sell headers
            ['Average Days to Sell', 'Median Days to Sell', 'Max Days to Sell', 'Min Days to Sell'],
            # Row 7: Days to sell values
            [f"{overall_stats['average_days']:.0f}", f"{overall_stats['median_days']:.0f}", f"{overall_stats['max_days']:.0f}", f"{overall_stats['min_days']:.0f}"]
        ]
        
        overall_summary_table = Table(overall_summary_data, colWidths=[1.75*inch, 1.75*inch, 1.75*inch, 1.75*inch])
        overall_summary_table.setStyle(TableStyle([
            # Header row for Total Properties
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('SPAN', (0, 0), (1, 0)),  # Span Total Properties across first two columns
            
            # Subheader rows (metric names)
            ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
            ('BACKGROUND', (0, 3), (-1, 3), colors.lightgrey),
            ('BACKGROUND', (0, 5), (-1, 5), colors.lightgrey),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
            ('FONTNAME', (0, 5), (-1, 5), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 9),
            ('FONTSIZE', (0, 3), (-1, 3), 9),
            ('FONTSIZE', (0, 5), (-1, 5), 9),
            ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
            ('ALIGN', (0, 3), (-1, 3), 'CENTER'),
            ('ALIGN', (0, 5), (-1, 5), 'CENTER'),
            
            # Data rows
            ('FONTNAME', (0, 2), (-1, 2), 'Helvetica'),
            ('FONTNAME', (0, 4), (-1, 4), 'Helvetica'),
            ('FONTNAME', (0, 6), (-1, 6), 'Helvetica'),
            ('FONTSIZE', (0, 2), (-1, 2), 10),
            ('FONTSIZE', (0, 4), (-1, 4), 10),
            ('FONTSIZE', (0, 6), (-1, 6), 10),
            ('ALIGN', (0, 2), (-1, 2), 'CENTER'),
            ('ALIGN', (0, 4), (-1, 4), 'CENTER'),
            ('ALIGN', (0, 6), (-1, 6), 'CENTER'),
            
            # Grid and spacing
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        story.append(overall_summary_table)
        story.append(Spacer(1, 12))
    
    # Add disclaimer
    story.append(Paragraph("Disclaimer: This data is sourced from our CRM and not our accounting software, based on then-available data. Final accounting data and results may vary slightly.", disclaimer_style))
    
    # Build PDF
    try:
        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

def format_currency(value):
    """Format value as currency with safe handling"""
    safe_value = safe_numeric_value(value, 0)
    return f"${safe_value:,.0f}"

def format_percentage(value):
    """Format value as percentage with safe handling"""
    safe_value = safe_numeric_value(value, 0)
    return f"{safe_value:.0f}%"

def create_summary_stats(df):
    """Create summary statistics with safe numeric handling - works with display column names"""
    if len(df) == 0:
        return {}
    
    # Safe aggregations - handle both original and display column names
    cost_basis_col = 'Cost Basis' if 'Cost Basis' in df.columns else 'custom.Asset_Cost_Basis'
    gross_sales_col = 'Gross Sales Price' if 'Gross Sales Price' in df.columns else 'custom.Asset_Gross_Sales_Price'
    closing_costs_col = 'Closing Costs' if 'Closing Costs' in df.columns else 'custom.Asset_Closing_Costs'
    gross_profit_col = 'Realized Gross Profit' if 'Realized Gross Profit' in df.columns else 'Realized_Gross_Profit'
    markup_col = 'Realized Markup' if 'Realized Markup' in df.columns else 'Realized_Markup'
    margin_col = 'Realized Margin' if 'Realized Margin' in df.columns else 'Realized_Margin'
    days_col = 'Days Until Sold' if 'Days Until Sold' in df.columns else 'Days_Until_Sold'
    
    cost_basis_sum = df[cost_basis_col].apply(lambda x: safe_numeric_value(x, 0)).sum()
    gross_sales_sum = df[gross_sales_col].apply(lambda x: safe_numeric_value(x, 0)).sum()
    closing_costs_sum = df[closing_costs_col].apply(lambda x: safe_numeric_value(x, 0)).sum()
    gross_profit_sum = df[gross_profit_col].apply(lambda x: safe_numeric_value(x, 0)).sum()
    
    # Safe statistical calculations
    markup_values = df[markup_col].apply(lambda x: safe_numeric_value(x, 0))
    margin_values = df[margin_col].apply(lambda x: safe_numeric_value(x, 0))
    days_values = df[days_col].apply(lambda x: safe_numeric_value(x, 0))
    
    return {
        'total_properties': len(df),
        'total_cost_basis': cost_basis_sum,
        'total_gross_sales': gross_sales_sum,
        'total_closing_costs': closing_costs_sum,
        'total_gross_profit': gross_profit_sum,
        'average_markup': markup_values.mean(),
        'median_markup': markup_values.median(),
        'max_markup': markup_values.max(),
        'min_markup': markup_values.min(),
        'average_margin': margin_values.mean(),
        'median_margin': margin_values.median(),
        'average_days': days_values.mean(),
        'median_days': days_values.median(),
        'max_days': days_values.max(),
        'min_days': days_values.min()
    }

def main():
    st.title("Sold Property Report")
    st.markdown("Generate quarterly reports for sold properties from Close.com CRM data")
    
    # Instructions
    with st.expander("Instructions", expanded=True):
        st.markdown("""
        **How to use this report:**
        1. **Export from Close.com:** Go to your Close.com CRM and export properties with "Remarkable - Sold" status
        2. **Export All Fields:** Make sure to select "All Fields" when exporting
        3. **Upload CSV:** Upload the exported CSV file below
        4. **Select Filters:** Choose which quarters and owners to include in your report
        5. **Generate Report:** View the report online or download as Excel/PDF
        
        **Note:** The Excel download will use original Close.com field names and include the ID field for easier data correction.
        """)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload your Close.com CSV export (Remarkable - Sold status, All Fields)",
        type=['csv'],
        help="Export properties with 'Remarkable - Sold' status from Close.com with all fields selected"
    )
    
    if uploaded_file is not None:
        try:
            # Load and process data
            df = pd.read_csv(uploaded_file)
            df_processed, df_display, error_df, total_records = process_data(df)
            
            if len(df_processed) == 0:
                st.warning("No sold properties found in the uploaded data.")
                if len(error_df) > 0:
                    st.error(f"Found {len(error_df)} records with errors that prevented processing.")
                return
            
            # Display import summary
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Records", total_records)
            with col2:
                st.metric("Successfully Processed", len(df_processed))
            with col3:
                st.metric("Records with Errors", len(error_df))
            with col4:
                success_rate = (len(df_processed) / total_records * 100) if total_records > 0 else 0
                st.metric("Success Rate", f"{success_rate:.1f}%")
            
            # Show error summary if there are errors
            if len(error_df) > 0:
                st.warning(f"Warning: {len(error_df)} records could not be processed. Download the error report below for details.")
                
                # Show error breakdown
                with st.expander("View Error Summary"):
                    error_counts = error_df['Error Type'].value_counts()
                    st.write("**Error Breakdown:**")
                    for error_type, count in error_counts.items():
                        st.write(f"• {error_type}: {count} records")
            else:
                st.success(f"All {total_records} records processed successfully!")
            
            # Filters
            st.subheader("Report Filters")
            
            # Initialize session state
            if 'selected_quarters' not in st.session_state:
                st.session_state.selected_quarters = []
            if 'selected_owners' not in st.session_state:
                st.session_state.selected_owners = []
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Select Calendar Quarters:**")
                available_quarters = sort_quarters_chronologically([q for q in df_display['Quarter_Year'].unique() if pd.notna(q)])
                
                # Quarter selection controls
                quarter_col1, quarter_col2 = st.columns(2)
                with quarter_col1:
                    if st.button("Select All Quarters", key="btn_select_all_quarters"):
                        st.session_state.selected_quarters = available_quarters.copy()
                with quarter_col2:
                    if st.button("Select None Quarters", key="btn_select_none_quarters"):
                        st.session_state.selected_quarters = []
                
                # Multiselect for quarters
                selected_quarters = st.multiselect(
                    "Choose quarters:",
                    options=available_quarters,
                    default=st.session_state.selected_quarters,
                    key="multiselect_quarters"
                )
                # Update session state with current selection
                st.session_state.selected_quarters = selected_quarters
            
            with col2:
                st.write("**Select Owners:**")
                available_owners = sorted([o for o in df_display['Owner'].unique() if pd.notna(o) and o != ''])
                
                # Owner selection controls
                owner_col1, owner_col2 = st.columns(2)
                with owner_col1:
                    if st.button("Select All Owners", key="btn_select_all_owners"):
                        st.session_state.selected_owners = available_owners.copy()
                with owner_col2:
                    if st.button("Select None Owners", key="btn_select_none_owners"):
                        st.session_state.selected_owners = []
                
                # Multiselect for owners
                selected_owners = st.multiselect(
                    "Choose owners:",
                    options=available_owners,
                    default=st.session_state.selected_owners,
                    key="multiselect_owners"
                )
                # Update session state with current selection
                st.session_state.selected_owners = selected_owners
            
            # Filter data based on selections (for both display and original data)
            filtered_df_display = df_display[
                (df_display['Quarter_Year'].isin(selected_quarters)) &
                (df_display['Owner'].isin(selected_owners))
            ].copy()
            
            # Also filter the original data for Excel export
            filtered_df_original = df_processed[
                (df_processed['Quarter_Year'].isin(selected_quarters)) &
                (df_processed['custom.Asset_Owner'].isin(selected_owners))
            ].copy()
            
            if len(filtered_df_display) == 0:
                st.warning("No properties match your selected filters.")
                return
            
            # Display results by quarter (in chronological order)
            st.subheader("Sold Properties Report")
            
            # Sort selected quarters chronologically
            sorted_selected_quarters = sort_quarters_chronologically(selected_quarters)
            
            for quarter in sorted_selected_quarters:
                quarter_data = filtered_df_display[filtered_df_display['Quarter_Year'] == quarter].copy()
                if len(quarter_data) == 0:
                    continue
                
                st.markdown(f"### {quarter}")
                
                # Prepare display data
                display_columns = [
                    'Property Name', 'Owner', 'State', 'County', 'Acres', 'Cost Basis',
                    'Date Purchased', 'Opportunity Status', 'Days Until Sold',
                    'Date Sold', 'Gross Sales Price', 'Closing Costs', 'Realized Gross Profit', 
                    'Realized Markup', 'Realized Margin'
                ]
                
                display_df = quarter_data[display_columns].copy()
                
                # Format for display with safe handling
                if 'Cost Basis' in display_df.columns:
                    display_df['Cost Basis'] = display_df['Cost Basis'].apply(format_currency)
                if 'Gross Sales Price' in display_df.columns:
                    display_df['Gross Sales Price'] = display_df['Gross Sales Price'].apply(format_currency)
                if 'Closing Costs' in display_df.columns:
                    display_df['Closing Costs'] = display_df['Closing Costs'].apply(format_currency)
                if 'Realized Gross Profit' in display_df.columns:
                    display_df['Realized Gross Profit'] = display_df['Realized Gross Profit'].apply(format_currency)
                if 'Realized Markup' in display_df.columns:
                    display_df['Realized Markup'] = display_df['Realized Markup'].apply(format_percentage)
                if 'Realized Margin' in display_df.columns:
                    display_df['Realized Margin'] = display_df['Realized Margin'].apply(format_percentage)
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
                stats = create_summary_stats(quarter_data)
                
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
                overall_stats = create_summary_stats(filtered_df_display)
                
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

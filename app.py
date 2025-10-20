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
        
        # Filter to only "Sold" records
        df_processed = df_processed[df_processed['primary_opportunity_status_label'] == 'Sold'].copy()
    
    # Track records with missing sold date
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
        
        # Remove records with missing sold date
        df_processed = df_processed[~missing_date_mask].copy()
    
    # Calculate derived fields (using original field names)
    df_processed['Days_Until_Sold'] = df_processed.apply(
        lambda x: calculate_days_until_sold(x['custom.Asset_Date_Purchased'], x['custom.Asset_Date_Sold']),
        axis=1
    )
    
    df_processed['Realized_Gross_Profit'] = df_processed.apply(
        lambda x: calculate_realized_gross_profit(
            x.get('custom.Asset_Gross_Sales_Price', 0),
            x.get('custom.Asset_Cost_Basis', 0),
            x.get('custom.Asset_Closing_Costs', 0)
        ),
        axis=1
    )
    
    df_processed['Realized_Markup'] = df_processed.apply(
        lambda x: calculate_realized_markup(
            x.get('custom.Asset_Gross_Sales_Price', 0),
            x.get('custom.Asset_Cost_Basis', 0),
            x.get('custom.Asset_Closing_Costs', 0)
        ),
        axis=1
    )
    
    df_processed['Realized_Margin'] = df_processed.apply(
        lambda x: calculate_realized_margin(
            x.get('Realized_Gross_Profit', 0),
            x.get('custom.Asset_Gross_Sales_Price', 0)
        ),
        axis=1
    )
    
    # Add quarter/year for grouping
    df_processed['Quarter_Year'] = df_processed['custom.Asset_Date_Sold'].apply(get_quarter_year)
    
    # Create error dataframe
    error_df = pd.DataFrame(error_records)
    
    # Return both the processed data and error report
    return df_processed, error_df, total_records

def create_summary_stats(df):
    """Create summary statistics for a dataframe"""
    return {
        'total_properties': len(df),
        'total_gross_sales': safe_numeric_value(df['custom.Asset_Gross_Sales_Price'].sum()),
        'total_cost_basis': safe_numeric_value(df['custom.Asset_Cost_Basis'].sum()),
        'total_closing_costs': safe_numeric_value(df['custom.Asset_Closing_Costs'].sum()),
        'total_gross_profit': safe_numeric_value(df['Realized_Gross_Profit'].sum()),
        'average_markup': safe_numeric_value(df['Realized_Markup'].mean()),
        'median_markup': safe_numeric_value(df['Realized_Markup'].median()),
        'average_margin': safe_numeric_value(df['Realized_Margin'].mean()),
        'median_margin': safe_numeric_value(df['Realized_Margin'].median()),
        'average_days': safe_numeric_value(df['Days_Until_Sold'].mean()),
        'median_days': safe_numeric_value(df['Days_Until_Sold'].median()),
        'max_days': safe_numeric_value(df['Days_Until_Sold'].max()),
        'min_days': safe_numeric_value(df['Days_Until_Sold'].min())
    }

def format_currency(value):
    """Format number as currency"""
    value = safe_numeric_value(value, 0)
    return f"${value:,.0f}"

def format_percentage(value):
    """Format number as percentage"""
    value = safe_numeric_value(value, 0)
    return f"{value:.1f}%"

def create_excel_download(df, filename):
    """Create Excel file for download with original Close.com field names and ID"""
    output = BytesIO()
    
    # Define the desired column order
    column_order = [
        'id',  # Keep ID first for data correction purposes
        'display_name',
        'custom.Asset_Owner',
        'custom.All_State',
        'custom.All_County',
        'custom.All_Asset_Surveyed_Acres',
        'custom.Asset_Cost_Basis',
        'custom.Asset_Date_Purchased',
        'custom.Asset_Date_Sold',
        'custom.Asset_Gross_Sales_Price',
        'custom.Asset_Closing_Costs',
        'Days_Until_Sold',
        'Realized_Gross_Profit',
        'Realized_Markup',
        'Realized_Margin',
        'Quarter_Year'
    ]
    
    # Filter to only include columns that exist in the dataframe
    existing_columns = [col for col in column_order if col in df.columns]
    
    # Reorder the dataframe
    df_export = df[existing_columns].copy()
    
    # Sort by Date Sold (most recent first)
    if 'custom.Asset_Date_Sold' in df_export.columns:
        df_export = df_export.sort_values('custom.Asset_Date_Sold', ascending=False)
    
    # Convert dates to strings for Excel
    date_columns = ['custom.Asset_Date_Purchased', 'custom.Asset_Date_Sold']
    for col in date_columns:
        if col in df_export.columns:
            df_export[col] = df_export[col].apply(
                lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else ''
            )
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Sold Properties')
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sold Properties']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#366092',
            'font_color': 'white',
            'border': 1
        })
        
        currency_format = workbook.add_format({'num_format': '$#,##0'})
        percentage_format = workbook.add_format({'num_format': '0.0%'})
        number_format = workbook.add_format({'num_format': '#,##0'})
        
        # Write headers with formatting
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Apply number formatting to appropriate columns
        for idx, col in enumerate(df_export.columns):
            if col in ['custom.Asset_Cost_Basis', 'custom.Asset_Gross_Sales_Price', 
                      'custom.Asset_Closing_Costs', 'Realized_Gross_Profit']:
                worksheet.set_column(idx, idx, 15, currency_format)
            elif col in ['Realized_Markup', 'Realized_Margin']:
                # Convert percentage values (already in 0-100 format) to decimal for Excel percentage format
                for row_num in range(1, len(df_export) + 1):
                    cell_value = df_export.iloc[row_num - 1][col]
                    if pd.notna(cell_value):
                        worksheet.write(row_num, idx, safe_numeric_value(cell_value) / 100, percentage_format)
            elif col in ['Days_Until_Sold', 'custom.All_Asset_Surveyed_Acres']:
                worksheet.set_column(idx, idx, 12, number_format)
            elif col == 'id':
                worksheet.set_column(idx, idx, 25)  # Wider column for ID
            else:
                worksheet.set_column(idx, idx, 18)
    
    output.seek(0)
    return output

def create_error_report_excel(error_df, total_records, processed_records, filename):
    """Create Excel file for error report"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write error details
        error_df.to_excel(writer, index=False, sheet_name='Errors', startrow=4)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Errors']
        
        # Define formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'bg_color': '#C00000',
            'font_color': 'white',
            'align': 'center'
        })
        
        stats_label_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFC7CE'
        })
        
        stats_value_format = workbook.add_format({
            'bg_color': '#FFF0F0'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#C00000',
            'font_color': 'white',
            'border': 1
        })
        
        # Write title and statistics
        worksheet.merge_range('A1:H1', 'Import Error Report', title_format)
        
        worksheet.write('A2', 'Total Records:', stats_label_format)
        worksheet.write('B2', total_records, stats_value_format)
        
        worksheet.write('A3', 'Successfully Processed:', stats_label_format)
        worksheet.write('B3', processed_records, stats_value_format)
        
        worksheet.write('A4', 'Records with Errors:', stats_label_format)
        worksheet.write('B4', len(error_df), stats_value_format)
        
        # Format error table headers
        for col_num, value in enumerate(error_df.columns.values):
            worksheet.write(4, col_num, value, header_format)
        
        # Set column widths
        worksheet.set_column('A:A', 25)  # ID
        worksheet.set_column('B:B', 30)  # Property Name
        worksheet.set_column('C:C', 20)  # Owner
        worksheet.set_column('D:D', 20)  # Opportunity Status
        worksheet.set_column('E:E', 25)  # Error Type
        worksheet.set_column('F:F', 50)  # Error Detail
        worksheet.set_column('G:G', 15)  # Date Sold
        worksheet.set_column('H:H', 12)  # Row Number
    
    output.seek(0)
    return output

def create_pdf_download(quarter_data_dict, filename):
    """Create PDF file for download with quarters in chronological order (landscape legal size)"""
    if not REPORTLAB_AVAILABLE:
        return None
    
    output = BytesIO()
    
    # Use landscape legal paper size
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(legal),
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        alignment=1  # Center
    )
    
    quarter_style = ParagraphStyle(
        'QuarterHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=8,
        spaceBefore=12
    )
    
    # Main title
    story.append(Paragraph("Sold Property Report", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%m/%d/%Y')}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Get quarters in chronological order
    sorted_quarters = sort_quarters_chronologically(list(quarter_data_dict.keys()))
    
    # Process each quarter
    for quarter in sorted_quarters:
        df = quarter_data_dict[quarter]
        
        if len(df) == 0:
            continue
        
        # Quarter header
        story.append(Paragraph(f"{quarter}", quarter_style))
        
        # Prepare table data with display names
        table_data = [[
            'Property Name', 'Owner', 'State', 'County', 'Acres',
            'Cost Basis', 'Date Purchased', 'Date Sold',
            'Gross Sales', 'Closing Costs', 'Gross Profit',
            'Markup %', 'Margin %', 'Days to Sell'
        ]]
        
        # Sort by Date Sold (most recent first within quarter)
        df_sorted = df.sort_values('Date Sold', ascending=False)
        
        for _, row in df_sorted.iterrows():
            table_data.append([
                str(row.get('Property Name', ''))[:30],  # Truncate long names
                str(row.get('Owner', ''))[:20],
                str(row.get('State', '')),
                str(row.get('County', ''))[:15],
                f"{safe_numeric_value(row.get('Acres', 0)):.1f}",
                f"${safe_numeric_value(row.get('Cost Basis', 0)):,.0f}",
                row.get('Date Sold', pd.NaT).strftime('%m/%d/%Y') if pd.notna(row.get('Date Purchased')) else '',
                row.get('Date Sold', pd.NaT).strftime('%m/%d/%Y') if pd.notna(row.get('Date Sold')) else '',
                f"${safe_numeric_value(row.get('Gross Sales Price', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Closing Costs', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Realized Gross Profit', 0)):,.0f}",
                f"{safe_numeric_value(row.get('Realized Markup', 0)):.1f}%",
                f"{safe_numeric_value(row.get('Realized Margin', 0)):.1f}%",
                f"{safe_int_value(row.get('Days Until Sold', 0))}"
            ])
        
        # Create table with adjusted column widths for legal landscape
        col_widths = [1.5*inch, 1.2*inch, 0.4*inch, 0.9*inch, 0.5*inch,
                     0.8*inch, 0.75*inch, 0.75*inch, 0.8*inch, 0.75*inch,
                     0.8*inch, 0.6*inch, 0.6*inch, 0.6*inch]
        
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Table styling
        table.setStyle(TableStyle([
            # Header row
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # State
            ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),  # Numeric columns
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F0F0F0')])
        ]))
        
        story.append(table)
        
        # Quarter summary
        stats = create_summary_stats(df)
        summary_text = f"""
        <b>Quarter Summary:</b> {stats['total_properties']} properties | 
        Gross Sales: {format_currency(stats['total_gross_sales'])} | 
        Gross Profit: {format_currency(stats['total_gross_profit'])} | 
        Avg Markup: {format_percentage(stats['average_markup'])} | 
        Avg Margin: {format_percentage(stats['average_margin'])} | 
        Avg Days: {safe_numeric_value(stats['average_days']):.0f}
        """
        
        summary_para = Paragraph(summary_text, styles['Normal'])
        story.append(Spacer(1, 0.1*inch))
        story.append(summary_para)
        story.append(PageBreak())
    
    # Remove last page break
    if story and isinstance(story[-1], PageBreak):
        story.pop()
    
    # Add disclaimer at the end
    story.append(Spacer(1, 0.2*inch))
    disclaimer = Paragraph(
        "<b>Disclaimer:</b> This data is sourced from our CRM and not our accounting software, "
        "based on then-available data. Final accounting data and results may vary slightly.",
        styles['Normal']
    )
    story.append(disclaimer)
    
    # Build PDF
    try:
        doc.build(story)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

def main():
    st.title("📊 Sold Property Report Generator")
    st.markdown("Upload your Close.com CSV export to generate comprehensive sold property reports")
    
    # Initialize session state for filters if not exists
    if 'selected_quarters' not in st.session_state:
        st.session_state.selected_quarters = []
    if 'selected_owners' not in st.session_state:
        st.session_state.selected_owners = []
    if 'available_quarters' not in st.session_state:
        st.session_state.available_quarters = []
    if 'available_owners' not in st.session_state:
        st.session_state.available_owners = []
    
    uploaded_file = st.file_uploader("Choose your Close.com CSV export", type=['csv'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            
            # Process the data
            df_processed, error_df, total_records = process_data(df)
            
            # Display processing summary
            st.subheader("📊 Processing Summary")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Records", total_records)
            with col2:
                st.metric("Successfully Processed", len(df_processed), 
                         delta=f"{(len(df_processed)/total_records*100):.1f}%")
            with col3:
                st.metric("Records with Errors", len(error_df),
                         delta=f"{(len(error_df)/total_records*100):.1f}%" if len(error_df) > 0 else "0%",
                         delta_color="inverse")
            with col4:
                success_rate = (len(df_processed) / total_records * 100) if total_records > 0 else 0
                st.metric("Success Rate", f"{success_rate:.1f}%")
            
            if len(df_processed) == total_records:
                st.success(f"All {total_records} records processed successfully!")
            
            # Show error details if any
            if len(error_df) > 0:
                with st.expander(f"⚠️ View {len(error_df)} Records with Errors", expanded=False):
                    st.dataframe(error_df, use_container_width=True)
            
            st.divider()
            
            # Create display version with mapped column names
            df_display = df_processed.copy()
            
            # Rename columns for display
            rename_dict = {}
            for old_name, new_name in FIELD_MAPPING.items():
                if old_name in df_display.columns:
                    rename_dict[old_name] = new_name
            
            # Also rename calculated fields
            rename_dict.update({
                'Days_Until_Sold': 'Days Until Sold',
                'Realized_Gross_Profit': 'Realized Gross Profit',
                'Realized_Markup': 'Realized Markup',
                'Realized_Margin': 'Realized Margin',
                'Quarter_Year': 'Quarter_Year'  # Keep this name for filtering
            })
            
            df_display = df_display.rename(columns=rename_dict)
            
            # Keep original version for Excel export (with ID)
            df_original = df_processed.copy()
            
            # Get unique quarters and owners - update session state available options
            st.session_state.available_quarters = sort_quarters_chronologically(df_display['Quarter_Year'].unique().tolist())
            st.session_state.available_owners = sorted(df_display['Owner'].unique().tolist())
            
            # Report Filters Section
            st.subheader("📋 Report Filters")
            
            # Create two columns for the filters
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Select Calendar Quarters:**")
                
                # Buttons for Select All / Select None Quarters
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("Select All Quarters", key="select_all_quarters"):
                        st.session_state.selected_quarters = st.session_state.available_quarters.copy()
                        st.rerun()
                with btn_col2:
                    if st.button("Select None Quarters", key="select_none_quarters"):
                        st.session_state.selected_quarters = []
                        st.rerun()
                
                # Quarter checkboxes
                for quarter in st.session_state.available_quarters:
                    is_checked = quarter in st.session_state.selected_quarters
                    if st.checkbox(quarter, value=is_checked, key=f"quarter_{quarter}"):
                        if quarter not in st.session_state.selected_quarters:
                            st.session_state.selected_quarters.append(quarter)
                    else:
                        if quarter in st.session_state.selected_quarters:
                            st.session_state.selected_quarters.remove(quarter)
            
            with col2:
                st.write("**Select Owners:**")
                
                # Buttons for Select All / Select None Owners
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("Select All Owners", key="select_all_owners"):
                        st.session_state.selected_owners = st.session_state.available_owners.copy()
                        st.rerun()
                with btn_col2:
                    if st.button("Select None Owners", key="select_none_owners"):
                        st.session_state.selected_owners = []
                        st.rerun()
                
                # Owner checkboxes
                for owner in st.session_state.available_owners:
                    is_checked = owner in st.session_state.selected_owners
                    if st.checkbox(owner, value=is_checked, key=f"owner_{owner}"):
                        if owner not in st.session_state.selected_owners:
                            st.session_state.selected_owners.append(owner)
                    else:
                        if owner in st.session_state.selected_owners:
                            st.session_state.selected_owners.remove(owner)
            
            st.divider()
            
            # Get selected filters from session state
            selected_quarters = st.session_state.selected_quarters
            selected_owners = st.session_state.selected_owners
            
            # Apply filters
            if not selected_quarters:
                st.warning("Please select at least one quarter to view data")
                return
            
            if not selected_owners:
                st.warning("Please select at least one owner to view data")
                return
            
            # Filter both display and original dataframes
            filtered_df_display = df_display[
                (df_display['Quarter_Year'].isin(selected_quarters)) &
                (df_display['Owner'].isin(selected_owners))
            ].copy()
            
            filtered_df_original = df_original[
                (df_original['Quarter_Year'].isin(selected_quarters)) &
                (df_original['custom.Asset_Owner'].isin(selected_owners))
            ].copy()
            
            if len(filtered_df_display) == 0:
                st.warning("No properties match the selected filters")
                return
            
            # Display data by quarter
            st.subheader(f"📈 Property Details ({len(filtered_df_display)} properties)")
            
            # Process each selected quarter in chronological order
            sorted_selected_quarters = sort_quarters_chronologically(selected_quarters)
            for quarter in sorted_selected_quarters:
                quarter_data = filtered_df_display[filtered_df_display['Quarter_Year'] == quarter].copy()
                
                if len(quarter_data) == 0:
                    continue
                
                with st.expander(f"**{quarter}** ({len(quarter_data)} properties)", expanded=True):
                    # Sort by Date Sold (most recent first)
                    quarter_data = quarter_data.sort_values('Date Sold', ascending=False)
                    
                    # Create display dataframe with selected columns
                    display_columns = [
                        'Property Name', 'Owner', 'State', 'County', 'Acres',
                        'Cost Basis', 'Date Purchased', 'Date Sold',
                        'Gross Sales Price', 'Closing Costs', 'Realized Gross Profit',
                        'Realized Markup', 'Realized Margin', 'Days Until Sold'
                    ]
                    
                    display_df = quarter_data[[col for col in display_columns if col in quarter_data.columns]].copy()
                    
                    # Format currency columns
                    currency_columns = ['Cost Basis', 'Gross Sales Price', 'Closing Costs', 'Realized Gross Profit']
                    for col in currency_columns:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].apply(lambda x: f"${x:,.0f}" if pd.notna(x) else "")
                    
                    # Format percentage columns
                    percentage_columns = ['Realized Markup', 'Realized Margin']
                    for col in percentage_columns:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "")
                    
                    # Format numeric columns
                    if 'Acres' in display_df.columns:
                        display_df['Acres'] = display_df['Acres'].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "")
                    if 'Days Until Sold' in display_df.columns:
                        display_df['Days Until Sold'] = display_df['Days Until Sold'].apply(lambda x: f"{x:.0f}" if pd.notna(x) else "")
                    
                    # Format date columns
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

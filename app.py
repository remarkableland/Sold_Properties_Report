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

# Field mapping from Close.com to report headers
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
    closing_costs = closing_costs if not pd.isna(closing_costs) else 0
    return gross_sales_price - cost_basis - closing_costs

def calculate_realized_markup(gross_sales_price, cost_basis, closing_costs):
    """Calculate realized markup percentage"""
    closing_costs = closing_costs if not pd.isna(closing_costs) else 0
    total_cost = cost_basis + closing_costs
    if total_cost == 0:
        return 0
    return ((gross_sales_price / total_cost) - 1) * 100

def calculate_realized_margin(realized_gross_profit, gross_sales_price):
    """Calculate realized margin percentage"""
    if gross_sales_price == 0:
        return 0
    return (realized_gross_profit / gross_sales_price) * 100

def process_data(df):
    """Process the uploaded data"""
    # Rename columns based on mapping
    df_processed = df.copy()
    
    # Rename columns that exist in the dataframe
    columns_to_rename = {k: v for k, v in FIELD_MAPPING.items() if k in df_processed.columns}
    df_processed = df_processed.rename(columns=columns_to_rename)
    
    # Convert date columns
    if 'Date Purchased' in df_processed.columns:
        df_processed['Date Purchased'] = df_processed['Date Purchased'].apply(parse_date)
    if 'Date Sold' in df_processed.columns:
        df_processed['Date Sold'] = df_processed['Date Sold'].apply(parse_date)
    
    # Filter only sold properties
    if 'Opportunity Status' in df_processed.columns:
        df_processed = df_processed[df_processed['Opportunity Status'] == 'Sold'].copy()
    
    # Convert numeric columns
    numeric_columns = ['Acres', 'Cost Basis', 'Gross Sales Price', 'Closing Costs']
    for col in numeric_columns:
        if col in df_processed.columns:
            df_processed[col] = pd.to_numeric(df_processed[col], errors='coerce').fillna(0)
    
    # Calculate derived fields
    df_processed['Days Until Sold'] = df_processed.apply(
        lambda row: calculate_days_until_sold(row.get('Date Purchased'), row.get('Date Sold')), axis=1
    )
    
    df_processed['Realized Gross Profit'] = df_processed.apply(
        lambda row: calculate_realized_gross_profit(
            row.get('Gross Sales Price', 0), 
            row.get('Cost Basis', 0), 
            row.get('Closing Costs', 0)
        ), axis=1
    )
    
    df_processed['Realized Markup'] = df_processed.apply(
        lambda row: calculate_realized_markup(
            row.get('Gross Sales Price', 0), 
            row.get('Cost Basis', 0), 
            row.get('Closing Costs', 0)
        ), axis=1
    )
    
    df_processed['Realized Margin'] = df_processed.apply(
        lambda row: calculate_realized_margin(
            row.get('Realized Gross Profit', 0), 
            row.get('Gross Sales Price', 0)
        ), axis=1
    )
    
    # Add quarter-year for filtering
    df_processed['Quarter_Year'] = df_processed['Date Sold'].apply(get_quarter_year)
    
    return df_processed

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
    
    # Title
    story.append(Paragraph("Sold Properties Report", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", summary_style))
    story.append(Spacer(1, 12))
    
    # Process each quarter
    for quarter in sorted(df_dict.keys()):
        quarter_data = df_dict[quarter]
        
        if len(quarter_data) == 0:
            continue
            
        # Quarter title
        story.append(Paragraph(f"{quarter}", quarter_style))
        
        # Prepare table data
        table_headers = ['Property Name', 'Owner', 'State', 'County', 'Acres', 'Cost Basis',
                        'Date Purchased', 'Days Until Sold', 'Date Sold', 'Gross Sales Price',
                        'Closing Costs', 'Realized Gross Profit', 'Realized Markup', 'Realized Margin']
        
        table_data = [table_headers]
        
        for _, row in quarter_data.iterrows():
            formatted_row = [
                str(row.get('Property Name', ''))[:25],  # Truncate long names
                str(row.get('Owner', ''))[:15],
                str(row.get('State', '')),
                str(row.get('County', ''))[:12],
                f"{row.get('Acres', 0):.1f}",
                f"${row.get('Cost Basis', 0):,.0f}",
                row.get('Date Purchased').strftime('%m/%d/%Y') if pd.notna(row.get('Date Purchased')) else '',
                f"{row.get('Days Until Sold', 0):.0f}" if pd.notna(row.get('Days Until Sold')) else '',
                row.get('Date Sold').strftime('%m/%d/%Y') if pd.notna(row.get('Date Sold')) else '',
                f"${row.get('Gross Sales Price', 0):,.0f}",
                f"${row.get('Closing Costs', 0):,.0f}",
                f"${row.get('Realized Gross Profit', 0):,.0f}",
                f"{row.get('Realized Markup', 0):.0f}%",
                f"{row.get('Realized Margin', 0):.0f}%"
            ]
            table_data.append(formatted_row)
        
        # Create table with appropriate column widths for landscape legal
        col_widths = [1.2*inch, 0.8*inch, 0.4*inch, 0.7*inch, 0.4*inch, 0.7*inch, 
                     0.7*inch, 0.4*inch, 0.7*inch, 0.8*inch, 0.6*inch, 0.8*inch, 0.5*inch, 0.5*inch]
        
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Table style
        table.setStyle(TableStyle([
            # Header formatting
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            
            # Data formatting
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('ALIGN', (0, 1), (1, -1), 'LEFT'),  # Property Name and Owner left-aligned
            ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),  # Numbers right-aligned
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 12))
        
        # Quarter summary statistics
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
        if quarter != sorted(df_dict.keys())[-1]:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 12))
    
    # Overall summary if multiple quarters
    if len(df_dict) > 1:
        # Combine all data for overall summary
        all_data = pd.concat(df_dict.values(), ignore_index=True)
        overall_stats = create_summary_stats(all_data)
        
        story.append(Paragraph("Overall Summary", quarter_style))
        
        overall_summary_data = [
            ['Metric', 'Value', 'Metric', 'Value'],
            ['Total Properties', f"{overall_stats['total_properties']}", 'Total Cost Basis', f"${overall_stats['total_cost_basis']:,.0f}"],
            ['Total Gross Sales', f"${overall_stats['total_gross_sales']:,.0f}", 'Total Gross Profit', f"${overall_stats['total_gross_profit']:,.0f}"],
            ['Average Markup', f"{overall_stats['average_markup']:.0f}%", 'Max Markup', f"{overall_stats['max_markup']:.0f}%"],
            ['Average Margin', f"{overall_stats['average_margin']:.0f}%", 'Median Margin', f"{overall_stats['median_margin']:.0f}%"],
            ['Average Days to Sell', f"{overall_stats['average_days']:.0f}", 'Max Days to Sell', f"{overall_stats['max_days']:.0f}"]
        ]
        
        overall_summary_table = Table(overall_summary_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
        overall_summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
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
        
        story.append(overall_summary_table)
        story.append(Spacer(1, 12))
    
    # Disclaimer
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
    """Format value as currency"""
    if pd.isna(value):
        return "$0"
    return f"${value:,.0f}"

def format_percentage(value):
    """Format value as percentage"""
    if pd.isna(value):
        return "0%"
    return f"{value:.0f}%"

def create_summary_stats(df):
    """Create summary statistics"""
    if len(df) == 0:
        return {}
    
    return {
        'total_properties': len(df),
        'total_cost_basis': df['Cost Basis'].sum(),
        'total_gross_sales': df['Gross Sales Price'].sum(),
        'total_gross_profit': df['Realized Gross Profit'].sum(),
        'average_markup': df['Realized Markup'].mean(),
        'median_markup': df['Realized Markup'].median(),
        'max_markup': df['Realized Markup'].max(),
        'min_markup': df['Realized Markup'].min(),
        'average_margin': df['Realized Margin'].mean(),
        'median_margin': df['Realized Margin'].median(),
        'average_days': df['Days Until Sold'].mean(),
        'median_days': df['Days Until Sold'].median(),
        'max_days': df['Days Until Sold'].max(),
        'min_days': df['Days Until Sold'].min()
    }

def create_excel_download(df, filename):
    """Create Excel file for download"""
    output = BytesIO()
    
    # Create a workbook and worksheet
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
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
    
    # Write headers
    headers = ['Property Name', 'Owner', 'State', 'County', 'Acres', 'Cost Basis', 'Date Purchased',
               'Opportunity Status', 'Days Until Sold', 'Date Sold', 'Gross Sales Price', 'Closing Costs',
               'Realized Gross Profit', 'Realized Markup', 'Realized Margin']
    
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    
    # Write data
    for row, (_, data) in enumerate(df.iterrows(), 1):
        # Property Name - highlight if empty
        prop_name = data.get('Property Name', '')
        if prop_name == '' or pd.isna(prop_name):
            worksheet.write(row, 0, prop_name, text_highlight_format)
        else:
            worksheet.write(row, 0, prop_name)
        
        # Owner - highlight if empty
        owner = data.get('Owner', '')
        if owner == '' or pd.isna(owner):
            worksheet.write(row, 1, owner, text_highlight_format)
        else:
            worksheet.write(row, 1, owner)
        
        # State - highlight if empty
        state = data.get('State', '')
        if state == '' or pd.isna(state):
            worksheet.write(row, 2, state, text_highlight_format)
        else:
            worksheet.write(row, 2, state)
        
        # County - highlight if empty
        county = data.get('County', '')
        if county == '' or pd.isna(county):
            worksheet.write(row, 3, county, text_highlight_format)
        else:
            worksheet.write(row, 3, county)
        
        # Handle Acres with null/inf check and highlighting
        acres = data.get('Acres', 0)
        if pd.notna(acres) and np.isfinite(acres) and acres != 0:
            worksheet.write(row, 4, acres, number_format)
        else:
            worksheet.write(row, 4, 0, number_highlight_format)
        
        # Handle Cost Basis with null/inf check and highlighting
        cost_basis = data.get('Cost Basis', 0)
        if pd.notna(cost_basis) and np.isfinite(cost_basis) and cost_basis != 0:
            worksheet.write(row, 5, cost_basis, currency_format)
        else:
            worksheet.write(row, 5, 0, currency_highlight_format)
        
        # Handle Date Purchased with null check and highlighting
        date_purchased = data.get('Date Purchased')
        if pd.notna(date_purchased) and date_purchased != '':
            worksheet.write(row, 6, date_purchased, date_format)
        else:
            worksheet.write(row, 6, '', date_highlight_format)
            
        # Opportunity Status - highlight if empty
        opp_status = data.get('Opportunity Status', '')
        if opp_status == '' or pd.isna(opp_status):
            worksheet.write(row, 7, opp_status, text_highlight_format)
        else:
            worksheet.write(row, 7, opp_status)
        
        # Handle Days Until Sold with null/inf check and highlighting
        days_sold = data.get('Days Until Sold', 0)
        if pd.notna(days_sold) and np.isfinite(days_sold) and days_sold != 0:
            worksheet.write(row, 8, int(days_sold), number_format)
        else:
            worksheet.write(row, 8, 0, number_highlight_format)
        
        # Handle Date Sold with null check and highlighting
        date_sold = data.get('Date Sold')
        if pd.notna(date_sold) and date_sold != '':
            worksheet.write(row, 9, date_sold, date_format)
        else:
            worksheet.write(row, 9, '', date_highlight_format)
        
        # Handle Gross Sales Price with null/inf check and highlighting
        gross_sales = data.get('Gross Sales Price', 0)
        if pd.notna(gross_sales) and np.isfinite(gross_sales) and gross_sales != 0:
            worksheet.write(row, 10, gross_sales, currency_format)
        else:
            worksheet.write(row, 10, 0, currency_highlight_format)
        
        # Handle Closing Costs with null/inf check and highlighting
        closing_costs = data.get('Closing Costs', 0)
        if pd.notna(closing_costs) and np.isfinite(closing_costs) and closing_costs != 0:
            worksheet.write(row, 11, closing_costs, currency_format)
        else:
            worksheet.write(row, 11, 0, currency_highlight_format)
        
        # Handle Realized Gross Profit with null/inf check and highlighting
        gross_profit = data.get('Realized Gross Profit', 0)
        if pd.notna(gross_profit) and np.isfinite(gross_profit) and gross_profit != 0:
            worksheet.write(row, 12, gross_profit, currency_format)
        else:
            worksheet.write(row, 12, 0, currency_highlight_format)
        
        # Handle Realized Markup with null/inf check and highlighting
        markup = data.get('Realized Markup', 0)
        if pd.notna(markup) and np.isfinite(markup) and markup != 0:
            worksheet.write(row, 13, markup / 100, percentage_format)
        else:
            worksheet.write(row, 13, 0, percentage_highlight_format)
        
        # Handle Realized Margin with null/inf check and highlighting
        margin = data.get('Realized Margin', 0)
        if pd.notna(margin) and np.isfinite(margin) and margin != 0:
            worksheet.write(row, 14, margin / 100, percentage_format)
        else:
            worksheet.write(row, 14, 0, percentage_highlight_format)
    
    # Auto-adjust column widths
    for col in range(len(headers)):
        worksheet.set_column(col, col, 15)
    
    workbook.close()
    output.seek(0)
    
    return output

def main():
    st.title("📊 Sold Property Report")
    st.markdown("Generate quarterly reports for sold properties from Close.com CRM data")
    
    # Instructions
    with st.expander("📋 Instructions", expanded=True):
        st.markdown("""
        **How to use this report:**
        1. **Export from Close.com:** Go to your Close.com CRM and export properties with "Remarkable - Sold" status
        2. **Export All Fields:** Make sure to select "All Fields" when exporting
        3. **Upload CSV:** Upload the exported CSV file below
        4. **Select Filters:** Choose which quarters and owners to include in your report
        5. **Generate Report:** View the report online or download as Excel
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
            df_processed = process_data(df)
            
            if len(df_processed) == 0:
                st.warning("No sold properties found in the uploaded data.")
                return
            
            st.success(f"✅ Loaded {len(df_processed)} sold properties")
            
            # Filters
            st.subheader("📊 Report Filters")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Select Calendar Quarters:**")
                available_quarters = sort_quarters_chronologically([q for q in df_processed['Quarter_Year'].unique() if pd.notna(q)])
                
                # Quarter selection controls
                quarter_col1, quarter_col2 = st.columns(2)
                with quarter_col1:
                    select_all_quarters = st.button("Select All Quarters", key="btn_select_all_quarters")
                with quarter_col2:
                    select_none_quarters = st.button("Select None Quarters", key="btn_select_none_quarters")
                
                # Initialize session state for quarter selections if not exists
                if 'quarter_selections' not in st.session_state:
                    st.session_state.quarter_selections = {q: False for q in available_quarters}
                
                # Handle button clicks
                if select_all_quarters:
                    for q in available_quarters:
                        st.session_state.quarter_selections[q] = True
                        
                if select_none_quarters:
                    for q in available_quarters:
                        st.session_state.quarter_selections[q] = False
                
                # Display checkboxes and collect selected quarters
                selected_quarters = []
                for quarter in available_quarters:
                    # Use session state value if it exists, otherwise default to False
                    current_value = st.session_state.quarter_selections.get(quarter, False)
                    if st.checkbox(f"{quarter}", value=current_value, key=f"cb_quarter_{quarter}"):
                        selected_quarters.append(quarter)
                        st.session_state.quarter_selections[quarter] = True
                    else:
                        st.session_state.quarter_selections[quarter] = False
            
            with col2:
                st.write("**Select Owners:**")
                available_owners = sorted([o for o in df_processed['Owner'].unique() if pd.notna(o) and o != ''])
                
                # Owner selection controls
                owner_col1, owner_col2 = st.columns(2)
                with owner_col1:
                    select_all_owners = st.button("Select All Owners", key="btn_select_all_owners")
                with owner_col2:
                    select_none_owners = st.button("Select None Owners", key="btn_select_none_owners")
                
                # Initialize session state for owner selections if not exists
                if 'owner_selections' not in st.session_state:
                    st.session_state.owner_selections = {o: False for o in available_owners}
                
                # Handle button clicks
                if select_all_owners:
                    for o in available_owners:
                        st.session_state.owner_selections[o] = True
                        
                if select_none_owners:
                    for o in available_owners:
                        st.session_state.owner_selections[o] = False
                
                # Display checkboxes and collect selected owners
                selected_owners = []
                for owner in available_owners:
                    # Use session state value if it exists, otherwise default to False
                    current_value = st.session_state.owner_selections.get(owner, False)
                    if st.checkbox(f"{owner}", value=current_value, key=f"cb_owner_{owner}"):
                        selected_owners.append(owner)
                        st.session_state.owner_selections[owner] = True
                    else:
                        st.session_state.owner_selections[owner] = False
            
            # Filter data based on selections
            filtered_df = df_processed[
                (df_processed['Quarter_Year'].isin(selected_quarters)) &
                (df_processed['Owner'].isin(selected_owners))
            ].copy()
            
            if len(filtered_df) == 0:
                st.warning("No properties match your selected filters.")
                return
            
            # Display results by quarter
            st.subheader("📈 Sold Properties Report")
            
            for quarter in selected_quarters:
                quarter_data = filtered_df[filtered_df['Quarter_Year'] == quarter].copy()
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
                
                # Format for display
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
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Properties Sold", stats['total_properties'])
                    st.metric("Total Cost Basis", format_currency(stats['total_cost_basis']))
                with col2:
                    st.metric("Total Gross Sales", format_currency(stats['total_gross_sales']))
                    st.metric("Total Gross Profit", format_currency(stats['total_gross_profit']))
                with col3:
                    st.metric("Average Markup", format_percentage(stats['average_markup']))
                    st.metric("Median Markup", format_percentage(stats['median_markup']))
                with col4:
                    st.metric("Average Margin", format_percentage(stats['average_margin']))
                    st.metric("Median Margin", format_percentage(stats['median_margin']))
                
                # Additional row for days to sell
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Average Days to Sell", f"{stats['average_days']:.0f}")
                with col2:
                    st.metric("Median Days to Sell", f"{stats['median_days']:.0f}")
                with col3:
                    st.metric("Max Days to Sell", f"{stats['max_days']:.0f}")
                with col4:
                    st.metric("Min Days to Sell", f"{stats['min_days']:.0f}")
                
                st.divider()
            
            # Overall summary
            if len(selected_quarters) > 1:
                st.markdown("### Overall Summary")
                overall_stats = create_summary_stats(filtered_df)
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Properties", overall_stats['total_properties'])
                    st.metric("Total Cost Basis", format_currency(overall_stats['total_cost_basis']))
                with col2:
                    st.metric("Total Gross Sales", format_currency(overall_stats['total_gross_sales']))
                    st.metric("Total Gross Profit", format_currency(overall_stats['total_gross_profit']))
                with col3:
                    st.metric("Average Markup", format_percentage(overall_stats['average_markup']))
                    st.metric("Max Markup", format_percentage(overall_stats['max_markup']))
                with col4:
                    st.metric("Average Margin", format_percentage(overall_stats['average_margin']))
                    st.metric("Median Margin", format_percentage(overall_stats['median_margin']))
                
                # Additional row for days to sell
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Average Days to Sell", f"{overall_stats['average_days']:.0f}")
                with col2:
                    st.metric("Max Days to Sell", f"{overall_stats['max_days']:.0f}")
                with col3:
                    st.metric("Min Days to Sell", f"{overall_stats['min_days']:.0f}")
                with col4:
                    st.metric("Median Days to Sell", f"{overall_stats['median_days']:.0f}")
            
            # Download section
            st.subheader("📥 Download Reports")
            
            # Prepare data for PDF (organized by quarter)
            quarter_data_dict = {}
            for quarter in selected_quarters:
                quarter_data_dict[quarter] = filtered_df[filtered_df['Quarter_Year'] == quarter].copy()
            
            # Generate filenames
            quarters_str = "_".join(selected_quarters).replace(" ", "")
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"sold_properties_{quarters_str}_{current_time}.xlsx"
            pdf_filename = f"sold_properties_{quarters_str}_{current_time}.pdf"
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Excel Report**")
                # Create Excel file
                excel_file = create_excel_download(filtered_df, excel_filename)
                
                st.download_button(
                    label="📄 Download Excel Report",
                    data=excel_file.getvalue(),
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                st.write("**PDF Report (Landscape Legal)**")
                if REPORTLAB_AVAILABLE:
                    # Create PDF file
                    pdf_file = create_pdf_download(quarter_data_dict, pdf_filename)
                    
                    if pdf_file:
                        st.download_button(
                            label="📋 Download PDF Report",
                            data=pdf_file.getvalue(),
                            file_name=pdf_filename,
                            mime="application/pdf"
                        )
                else:
                    st.warning("PDF generation requires reportlab. Run: pip install reportlab")
                    st.info("Excel download is still available above.")
            
            # Disclaimer
            st.markdown("---")
            st.markdown("**Disclaimer:** This data is sourced from our CRM and not our accounting software, based on then-available data. Final accounting data and results may vary slightly.")
            
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.write("Please make sure you've uploaded a valid CSV export from Close.com with all fields included.")
    
    else:
        st.info("👆 Please upload your Close.com CSV export to generate the sold property report")

if __name__ == "__main__":
    main()

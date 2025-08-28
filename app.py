import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter

st.set_page_config(
    page_title="Sold Property Report",
    page_icon="📊",
    layout="wide"
)

# Field mapping from Close.com to report headers
FIELD_MAPPING = {
    'display_name': 'Property Name',
    'custom.All_State': 'State',
    'custom.All_County': 'County',
    'custom.All_Asset_Surveyed_Acres': 'Acres',
    'custom.Asset_Cost_Basis': 'Cost Basis',
    'custom.Asset_Date_Purchased': 'Date Purchased',
    'primary_opportunity_status_label': 'Opportunity Status',
    'custom.Asset_Date_Sold': 'Date Sold',
    'custom.Asset_Gross_Sales_Price': 'Gross Sales Price',
    'custom.Asset_Closing_Costs': 'Closing Costs',
    'custom.Asset_Owner': 'Owner'
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
    
    # Write headers
    headers = ['Property Name', 'State', 'County', 'Acres', 'Cost Basis', 'Date Purchased',
               'Opportunity Status', 'Days Until Sold', 'Date Sold', 'Gross Sales Price',
               'Realized Gross Profit', 'Realized Markup']
    
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    
    # Write data
    for row, (_, data) in enumerate(df.iterrows(), 1):
        worksheet.write(row, 0, data.get('Property Name', ''))
        worksheet.write(row, 1, data.get('State', ''))
        worksheet.write(row, 2, data.get('County', ''))
        
        # Handle Acres with null/inf check
        acres = data.get('Acres', 0)
        if pd.notna(acres) and np.isfinite(acres):
            worksheet.write(row, 3, acres, number_format)
        else:
            worksheet.write(row, 3, 0, number_format)
        
        # Handle Cost Basis with null/inf check
        cost_basis = data.get('Cost Basis', 0)
        if pd.notna(cost_basis) and np.isfinite(cost_basis):
            worksheet.write(row, 4, cost_basis, currency_format)
        else:
            worksheet.write(row, 4, 0, currency_format)
        
        # Handle Date Purchased with null check
        date_purchased = data.get('Date Purchased')
        if pd.notna(date_purchased) and date_purchased != '':
            worksheet.write(row, 5, date_purchased, date_format)
        else:
            worksheet.write(row, 5, '')
            
        worksheet.write(row, 6, data.get('Opportunity Status', ''))
        
        # Handle Days Until Sold with null/inf check
        days_sold = data.get('Days Until Sold', 0)
        if pd.notna(days_sold) and np.isfinite(days_sold):
            worksheet.write(row, 7, int(days_sold), number_format)
        else:
            worksheet.write(row, 7, 0, number_format)
        
        # Handle Date Sold with null check
        date_sold = data.get('Date Sold')
        if pd.notna(date_sold) and date_sold != '':
            worksheet.write(row, 8, date_sold, date_format)
        else:
            worksheet.write(row, 8, '')
        
        # Handle Gross Sales Price with null/inf check
        gross_sales = data.get('Gross Sales Price', 0)
        if pd.notna(gross_sales) and np.isfinite(gross_sales):
            worksheet.write(row, 9, gross_sales, currency_format)
        else:
            worksheet.write(row, 9, 0, currency_format)
        
        # Handle Realized Gross Profit with null/inf check
        gross_profit = data.get('Realized Gross Profit', 0)
        if pd.notna(gross_profit) and np.isfinite(gross_profit):
            worksheet.write(row, 10, gross_profit, currency_format)
        else:
            worksheet.write(row, 10, 0, currency_format)
        
        # Handle Realized Markup with null/inf check
        markup = data.get('Realized Markup', 0)
        if pd.notna(markup) and np.isfinite(markup):
            worksheet.write(row, 11, markup / 100, percentage_format)
        else:
            worksheet.write(row, 11, 0, percentage_format)
    
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
                available_quarters = sorted([q for q in df_processed['Quarter_Year'].unique() if pd.notna(q)])
                selected_quarters = []
                
                for quarter in available_quarters:
                    if st.checkbox(f"{quarter}", value=True, key=f"quarter_{quarter}"):
                        selected_quarters.append(quarter)
            
            with col2:
                st.write("**Select Owners:**")
                available_owners = sorted([o for o in df_processed['Owner'].unique() if pd.notna(o) and o != ''])
                selected_owners = []
                
                for owner in available_owners:
                    if st.checkbox(f"{owner}", value=True, key=f"owner_{owner}"):
                        selected_owners.append(owner)
            
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
                    'Property Name', 'State', 'County', 'Acres', 'Cost Basis',
                    'Date Purchased', 'Opportunity Status', 'Days Until Sold',
                    'Date Sold', 'Gross Sales Price', 'Realized Gross Profit', 'Realized Markup'
                ]
                
                display_df = quarter_data[display_columns].copy()
                
                # Format for display
                if 'Cost Basis' in display_df.columns:
                    display_df['Cost Basis'] = display_df['Cost Basis'].apply(format_currency)
                if 'Gross Sales Price' in display_df.columns:
                    display_df['Gross Sales Price'] = display_df['Gross Sales Price'].apply(format_currency)
                if 'Realized Gross Profit' in display_df.columns:
                    display_df['Realized Gross Profit'] = display_df['Realized Gross Profit'].apply(format_currency)
                if 'Realized Markup' in display_df.columns:
                    display_df['Realized Markup'] = display_df['Realized Markup'].apply(format_percentage)
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
                    st.metric("Average Days to Sell", f"{stats['average_days']:.0f}")
                    st.metric("Median Days to Sell", f"{stats['median_days']:.0f}")
                
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
                    st.metric("Average Days to Sell", f"{overall_stats['average_days']:.0f}")
                    st.metric("Max Days to Sell", f"{overall_stats['max_days']:.0f}")
            
            # Download section
            st.subheader("📥 Download Report")
            
            # Generate filename
            quarters_str = "_".join(selected_quarters).replace(" ", "")
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"sold_properties_{quarters_str}_{current_time}.xlsx"
            
            # Create Excel file
            excel_file = create_excel_download(filtered_df, filename)
            
            st.download_button(
                label="📄 Download Excel Report",
                data=excel_file.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
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

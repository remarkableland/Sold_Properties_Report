import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import xlsxwriter
from typing import Dict, Tuple, Optional, Iterable, List, Tuple as Tup

try:
    from reportlab.lib.pagesizes import legal, landscape
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

# ---------- Display mapping ----------
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

# ---------- Helpers ----------
def safe_numeric_value(value, default=0.0) -> float:
    try:
        if pd.isna(value) or np.isinf(value):
            return float(default)
        return float(value)
    except (ValueError, TypeError):
        return float(default)

def safe_int_value(value, default=0) -> int:
    try:
        if pd.isna(value) or np.isinf(value):
            return int(default)
        return int(value)
    except (ValueError, TypeError):
        return int(default)

def get_quarter_year(dt: pd.Timestamp) -> Optional[str]:
    if pd.isna(dt):
        return None
    q = 1 + (dt.month - 1) // 3
    return f"Q{q} {dt.year}"

def sort_quarters_chronologically(qs):
    def key(q):
        if not q or pd.isna(q):
            return (0, 0)
        try:
            qpart, ypart = q.split()
            return (int(ypart), int(qpart.replace('Q', '')))
        except Exception:
            return (0, 0)
    return sorted([q for q in qs if pd.notna(q)], key=key)

def format_currency(v) -> str:
    return f"${safe_numeric_value(v):,.0f}"

def format_percentage(v) -> str:
    return f"{safe_numeric_value(v):.0f}%"

# ---------- XIRR (date-aware IRR) ----------
def _npv_at_rate(rate: float, cfs: List[Tup[pd.Timestamp, float]]) -> float:
    """NPV with actual day count; rate is annual (decimal)."""
    if rate <= -0.999999999:
        # avoid division by zero or negative base ^ power issues
        return np.inf
    t0 = cfs[0][0]
    npv = 0.0
    for dt, amt in cfs:
        years = (dt - t0).days / 365.0
        npv += amt / ((1.0 + rate) ** years)
    return npv

def xirr(cashflows: List[Tup[pd.Timestamp, float]],
         lo: float = -0.9999,
         hi: float = 10.0,
         tol: float = 1e-7,
         max_iter: int = 200) -> Optional[float]:
    """
    Find r where NPV(r) ~= 0 using bisection.
    Returns annual rate (decimal), or None if not solvable.
    """
    # Require at least one positive and one negative flow
    vals = [cf[1] for cf in cashflows]
    if not (any(v > 0 for v in vals) and any(v < 0 for v in vals)):
        return None

    # Sort by date, anchor to first date
    cfs = sorted(cashflows, key=lambda x: x[0])

    # Ensure sign change on [lo, hi]
    f_lo = _npv_at_rate(lo, cfs)
    f_hi = _npv_at_rate(hi, cfs)
    # Try to expand if needed
    expand = 0
    while np.sign(f_lo) == np.sign(f_hi) and expand < 5:
        hi *= 2
        f_hi = _npv_at_rate(hi, cfs)
        expand += 1

    if np.isnan(f_lo) or np.isnan(f_hi) or np.sign(f_lo) == np.sign(f_hi):
        return None

    for _ in range(max_iter):
        mid = (lo + hi) / 2.0
        f_mid = _npv_at_rate(mid, cfs)
        if abs(f_mid) < tol:
            return mid
        if np.sign(f_mid) == np.sign(f_lo):
            lo, f_lo = mid, f_mid
        else:
            hi, f_hi = mid, f_mid
    return mid  # best effort

def compute_row_xirr(purchase_dt, sale_dt, cost_basis, gross_sales, closing_costs) -> Optional[float]:
    """
    Returns IRR as PERCENT (e.g., 18.5 for 18.5%),
    or None if not computable.
    """
    if pd.isna(purchase_dt) or pd.isna(sale_dt):
        return None
    cost_basis = safe_numeric_value(cost_basis, 0.0)
    gross_sales = safe_numeric_value(gross_sales, 0.0)
    closing_costs = safe_numeric_value(closing_costs, 0.0)

    net_proceeds = gross_sales - closing_costs
    if cost_basis <= 0 or net_proceeds <= 0:
        return None

    cf = [
        (pd.to_datetime(purchase_dt), -cost_basis),
        (pd.to_datetime(sale_dt), net_proceeds)
    ]
    r = xirr(cf)
    if r is None or np.isinf(r) or np.isnan(r):
        return None
    return r * 100.0  # store as percent, to match your other % fields

# ---------- Core processing ----------
@st.cache_data(show_spinner=False)
def process_data(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, int]:
    df = df_raw.copy()
    total_records = len(df)

    # Parse dates (vectorized, tolerant)
    if 'custom.Asset_Date_Purchased' in df.columns:
        df['custom.Asset_Date_Purchased'] = pd.to_datetime(df['custom.Asset_Date_Purchased'], errors='coerce', utc=False)
    if 'custom.Asset_Date_Sold' in df.columns:
        df['custom.Asset_Date_Sold'] = pd.to_datetime(df['custom.Asset_Date_Sold'], errors='coerce', utc=False)

    # Build error report
    error_rows = []

    if 'primary_opportunity_status_label' in df.columns:
        mask_non_sold = df['primary_opportunity_status_label'].notna() & (df['primary_opportunity_status_label'] != 'Sold')
        for idx, row in df[mask_non_sold].iterrows():
            error_rows.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': row.get('primary_opportunity_status_label', 'Missing'),
                'Error Type': 'Status Not "Sold"',
                'Error Detail': f'Status is "{row.get("primary_opportunity_status_label")}" instead of "Sold"',
                'Date Sold': row.get('custom.Asset_Date_Sold', 'N/A'),
                'Row Number': idx + 2
            })

        mask_null_status = df['primary_opportunity_status_label'].isna()
        for idx, row in df[mask_null_status].iterrows():
            error_rows.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': 'NULL/Missing',
                'Error Type': 'Missing Status',
                'Error Detail': 'Opportunity Status field is empty or null',
                'Date Sold': row.get('custom.Asset_Date_Sold', 'N/A'),
                'Row Number': idx + 2
            })

        # Keep only Sold
        df = df[df['primary_opportunity_status_label'] == 'Sold'].copy()

    # Missing Date Sold among the Sold set
    if 'custom.Asset_Date_Sold' in df.columns:
        mask_missing_date = df['custom.Asset_Date_Sold'].isna()
        for idx, row in df[mask_missing_date].iterrows():
            error_rows.append({
                'ID': row.get('id', 'Unknown'),
                'Property Name': row.get('display_name', 'Unknown'),
                'Owner': row.get('custom.Asset_Owner', 'Unknown'),
                'Opportunity Status': row.get('primary_opportunity_status_label', 'Unknown'),
                'Error Type': 'Missing Date Sold',
                'Error Detail': 'Date Sold field is empty, null, or could not be parsed',
                'Date Sold': 'Invalid/Missing',
                'Row Number': idx + 2
            })

    # Numeric coercions
    for c in ['custom.All_Asset_Surveyed_Acres',
              'custom.Asset_Cost_Basis',
              'custom.Asset_Gross_Sales_Price',
              'custom.Asset_Closing_Costs']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    # Derived fields
    date_purch = df.get('custom.Asset_Date_Purchased', pd.Series([pd.NaT]*len(df)))
    date_sold  = df.get('custom.Asset_Date_Sold', pd.Series([pd.NaT]*len(df)))

    df['Days_Until_Sold'] = (date_sold - date_purch).dt.days
    df.loc[df['Days_Until_Sold'].isna(), 'Days_Until_Sold'] = np.nan

    gross = df.get('custom.Asset_Gross_Sales_Price', pd.Series([0.0]*len(df)))
    basis = df.get('custom.Asset_Cost_Basis', pd.Series([0.0]*len(df)))
    closec = df.get('custom.Asset_Closing_Costs', pd.Series([0.0]*len(df)))

    df['Realized_Gross_Profit'] = gross - basis - closec

    total_cost = basis + closec
    df['Realized_Markup'] = np.where(total_cost > 0, (gross / total_cost - 1.0) * 100.0, 0.0)
    df['Realized_Margin'] = np.where(gross > 0, (df['Realized_Gross_Profit'] / gross) * 100.0, 0.0)

    # Realized IRR (percent)
    df['Realized_IRR'] = [
        compute_row_xirr(dp, ds, cb, gs, cc)
        for dp, ds, cb, gs, cc in zip(
            df.get('custom.Asset_Date_Purchased', pd.Series([pd.NaT]*len(df))),
            df.get('custom.Asset_Date_Sold', pd.Series([pd.NaT]*len(df))),
            df.get('custom.Asset_Cost_Basis', pd.Series([0.0]*len(df))),
            df.get('custom.Asset_Gross_Sales_Price', pd.Series([0.0]*len(df))),
            df.get('custom.Asset_Closing_Costs', pd.Series([0.0]*len(df))),
        )
    ]

    # Quarter/Year
    df['Quarter_Year'] = pd.to_datetime(df.get('custom.Asset_Date_Sold')).apply(get_quarter_year)

    # Display copy
    df_display = df.rename(columns={k: v for k, v in FIELD_MAPPING.items() if k in df.columns})
    df_display = df_display.rename(columns={
        'Days_Until_Sold': 'Days Until Sold',
        'Realized_Gross_Profit': 'Realized Gross Profit',
        'Realized_Markup': 'Realized Markup',
        'Realized_Margin': 'Realized Margin',
        'Realized_IRR': 'Realized IRR'  # keep as percent value
    })

    error_df = pd.DataFrame(error_rows)
    return df, df_display, error_df, total_records

def create_summary_stats(df: pd.DataFrame) -> Dict[str, float]:
    if len(df) == 0:
        return {k: 0 for k in [
            'total_properties','total_cost_basis','total_gross_sales','total_closing_costs','total_gross_profit',
            'average_markup','median_markup','max_markup','min_markup',
            'average_margin','median_margin',
            'average_days','median_days','max_days','min_days',
            'average_irr','median_irr','max_irr','min_irr'
        ]}
    # Resolve either display or original column names
    cost_basis_col = 'Cost Basis' if 'Cost Basis' in df.columns else 'custom.Asset_Cost_Basis'
    gross_sales_col = 'Gross Sales Price' if 'Gross Sales Price' in df.columns else 'custom.Asset_Gross_Sales_Price'
    closing_costs_col = 'Closing Costs' if 'Closing Costs' in df.columns else 'custom.Asset_Closing_Costs'
    gross_profit_col = 'Realized Gross Profit' if 'Realized Gross Profit' in df.columns else 'Realized_Gross_Profit'
    markup_col = 'Realized Markup' if 'Realized Markup' in df.columns else 'Realized_Markup'
    margin_col = 'Realized Margin' if 'Realized Margin' in df.columns else 'Realized_Margin'
    irr_col = 'Realized IRR' if 'Realized IRR' in df.columns else 'Realized_IRR'
    days_col = 'Days Until Sold' if 'Days Until Sold' in df.columns else 'Days_Until_Sold'

    numeric = lambda s: pd.to_numeric(s, errors='coerce')

    cost_basis_sum = numeric(df[cost_basis_col]).fillna(0).sum()
    gross_sales_sum = numeric(df[gross_sales_col]).fillna(0).sum()
    closing_costs_sum = numeric(df[closing_costs_col]).fillna(0).sum()
    gross_profit_sum = numeric(df[gross_profit_col]).fillna(0).sum()

    markup_vals = numeric(df[markup_col]).dropna()
    margin_vals = numeric(df[margin_col]).dropna()
    irr_vals = numeric(df[irr_col]).dropna()  # already in percent units
    days_vals = numeric(df[days_col]).dropna()

    return {
        'total_properties': int(len(df)),
        'total_cost_basis': float(cost_basis_sum),
        'total_gross_sales': float(gross_sales_sum),
        'total_closing_costs': float(closing_costs_sum),
        'total_gross_profit': float(gross_profit_sum),
        'average_markup': float(markup_vals.mean()) if len(markup_vals) else 0.0,
        'median_markup': float(markup_vals.median()) if len(markup_vals) else 0.0,
        'max_markup': float(markup_vals.max()) if len(markup_vals) else 0.0,
        'min_markup': float(markup_vals.min()) if len(markup_vals) else 0.0,
        'average_margin': float(margin_vals.mean()) if len(margin_vals) else 0.0,
        'median_margin': float(margin_vals.median()) if len(margin_vals) else 0.0,
        'average_days': float(days_vals.mean()) if len(days_vals) else 0.0,
        'median_days': float(days_vals.median()) if len(days_vals) else 0.0,
        'max_days': float(days_vals.max()) if len(days_vals) else 0.0,
        'min_days': float(days_vals.min()) if len(days_vals) else 0.0,
        'average_irr': float(irr_vals.mean()) if len(irr_vals) else 0.0,
        'median_irr': float(irr_vals.median()) if len(irr_vals) else 0.0,
        'max_irr': float(irr_vals.max()) if len(irr_vals) else 0.0,
        'min_irr': float(irr_vals.min()) if len(irr_vals) else 0.0,
    }

# ---------- Excel builders (unchanged output fields unless you want IRR there too) ----------
def create_excel_download(df_original: pd.DataFrame, filename: str) -> BytesIO:
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})
    ws = wb.add_worksheet('Sold Properties')

    header_fmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
    currency_fmt = wb.add_format({'num_format': '$#,##0', 'border': 1})
    pct_fmt = wb.add_format({'num_format': '0%', 'border': 1})
    date_fmt = wb.add_format({'num_format': 'mm/dd/yyyy', 'border': 1})
    number_fmt = wb.add_format({'num_format': '#,##0', 'border': 1})

    currency_hi = wb.add_format({'num_format': '$#,##0', 'border': 1, 'bg_color': '#FFFF99'})
    pct_hi = wb.add_format({'num_format': '0%', 'border': 1, 'bg_color': '#FFFF99'})
    date_hi = wb.add_format({'num_format': 'mm/dd/yyyy', 'border': 1, 'bg_color': '#FFFF99'})
    number_hi = wb.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF99'})
    text_hi = wb.add_format({'border': 1, 'bg_color': '#FFFF99'})

    headers = ['id','display_name','custom.Asset_Owner','custom.All_State','custom.All_County',
               'custom.All_Asset_Surveyed_Acres','custom.Asset_Cost_Basis','custom.Asset_Date_Purchased',
               'primary_opportunity_status_label','Days_Until_Sold','custom.Asset_Date_Sold',
               'custom.Asset_Gross_Sales_Price','custom.Asset_Closing_Costs',
               'Realized_Gross_Profit','Realized_Markup','Realized_Margin',
               'Realized_IRR']  # include IRR in Excel as well, since it’s useful

    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    for row, (_, data) in enumerate(df_original.iterrows(), start=1):
        # ID
        lead_id = str(data.get('id', '') or '')
        ws.write(row, 0, '' if lead_id in ('', 'nan') else lead_id if lead_id else '', text_hi if lead_id in ('', 'nan') else None)

        # Property Name
        pname = str(data.get('display_name', '') or '')
        ws.write(row, 1, '' if pname in ('', 'nan') else pname, text_hi if pname in ('', 'nan') else None)

        # Owner
        owner = str(data.get('custom.Asset_Owner', '') or '')
        ws.write(row, 2, '' if owner in ('', 'nan') else owner, text_hi if owner in ('', 'nan') else None)

        # State
        state = str(data.get('custom.All_State', '') or '')
        ws.write(row, 3, '' if state in ('', 'nan') else state, text_hi if state in ('', 'nan') else None)

        # County
        county = str(data.get('custom.All_County', '') or '')
        ws.write(row, 4, '' if county in ('', 'nan') else county, text_hi if county in ('', 'nan') else None)

        # Acres
        acres = safe_numeric_value(data.get('custom.All_Asset_Surveyed_Acres', 0))
        ws.write_number(row, 5, acres, number_fmt if acres else number_hi)

        # Cost Basis
        basis = safe_numeric_value(data.get('custom.Asset_Cost_Basis', 0))
        ws.write_number(row, 6, basis, currency_fmt if basis else currency_hi)

        # Date Purchased
        dp = data.get('custom.Asset_Date_Purchased')
        if pd.isna(dp):
            ws.write(row, 7, '', date_hi)
        else:
            ws.write_datetime(row, 7, pd.to_datetime(dp).to_pydatetime(), date_fmt)

        # Opportunity Status
        status = str(data.get('primary_opportunity_status_label', '') or '')
        ws.write(row, 8, '' if status in ('', 'nan') else status, text_hi if status in ('', 'nan') else None)

        # Days Until Sold
        days = data.get('Days_Until_Sold')
        if pd.isna(days) or days == 0:
            ws.write_number(row, 9, 0, number_hi)
        else:
            ws.write_number(row, 9, safe_int_value(days), number_fmt)

        # Date Sold
        ds = data.get('custom.Asset_Date_Sold')
        if pd.isna(ds):
            ws.write(row, 10, '', date_hi)
        else:
            ws.write_datetime(row, 10, pd.to_datetime(ds).to_pydatetime(), date_fmt)

        # Gross Sales
        gross = safe_numeric_value(data.get('custom.Asset_Gross_Sales_Price', 0))
        ws.write_number(row, 11, gross, currency_fmt if gross else currency_hi)

        # Closing Costs
        cc = safe_numeric_value(data.get('custom.Asset_Closing_Costs', 0))
        ws.write_number(row, 12, cc, currency_fmt if cc else currency_hi)

        # Gross Profit
        gp = safe_numeric_value(data.get('Realized_Gross_Profit', 0))
        ws.write_number(row, 13, gp, currency_fmt if gp else currency_hi)

        # Markup (percent expects 0–1)
        mu = safe_numeric_value(data.get('Realized_Markup', 0)) / 100.0
        ws.write_number(row, 14, mu, pct_fmt if mu else pct_hi)

        # Margin (percent expects 0–1)
        mar = safe_numeric_value(data.get('Realized_Margin', 0)) / 100.0
        ws.write_number(row, 15, mar, pct_fmt if mar else pct_hi)

        # IRR (percent expects 0–1)
        irr = data.get('Realized_IRR')
        if pd.isna(irr) or irr is None:
            ws.write_number(row, 16, 0.0, pct_hi)
        else:
            ws.write_number(row, 16, safe_numeric_value(irr)/100.0, pct_fmt)

    # Column sizes
    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 25)
    for c in range(2, len(headers)):
        ws.set_column(c, c, 15)

    wb.close()
    output.seek(0)
    return output

def create_error_report_excel(error_df: pd.DataFrame, total_records: int, processed_records: int, filename: str) -> BytesIO:
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})
    ws_summary = wb.add_worksheet('Import Summary')
    ws_error = wb.add_worksheet('Error Details')

    hdr = wb.add_format({'bold': True, 'bg_color': '#1976D2', 'font_color': 'white', 'border': 1})
    hdr_err = wb.add_format({'bold': True, 'bg_color': '#D32F2F', 'font_color': 'white', 'border': 1})
    ok_fmt = wb.add_format({'bg_color': '#E8F5E8', 'border': 1})
    warn_fmt = wb.add_format({'bg_color': '#FFEBEE', 'border': 1})

    # Summary
    ws_summary.write(0, 0, 'Import Summary Report', hdr)
    ws_summary.write(2, 0, 'Metric', hdr)
    ws_summary.write(2, 1, 'Value', hdr)
    ws_summary.write(3, 0, 'Total Records in CSV', ok_fmt)
    ws_summary.write(3, 1, total_records, ok_fmt)
    ws_summary.write(4, 0, 'Successfully Processed', ok_fmt)
    ws_summary.write(4, 1, processed_records, ok_fmt)
    ws_summary.write(5, 0, 'Records with Errors', warn_fmt)
    ws_summary.write(5, 1, len(error_df), warn_fmt)

    success_rate = (processed_records / total_records * 100) if total_records else 0.0
    ws_summary.write(6, 0, 'Success Rate', hdr)
    ws_summary.write(6, 1, f"{success_rate:.1f}%", hdr)
    ws_summary.set_column(0, 0, 25)
    ws_summary.set_column(1, 1, 15)

    # Breakdown
    if len(error_df) > 0:
        ws_summary.write(8, 0, 'Error Breakdown by Type', hdr)
        ws_summary.write(9, 0, 'Error Type', hdr)
        ws_summary.write(9, 1, 'Count', hdr)
        counts = error_df['Error Type'].value_counts()
        for i, (etype, cnt) in enumerate(counts.items(), start=10):
            ws_summary.write(i, 0, etype, warn_fmt)
            ws_summary.write(i, 1, int(cnt), warn_fmt)

        # Details sheet
        headers = ['Row Number','ID','Property Name','Owner','Opportunity Status','Error Type','Error Detail','Date Sold']
        for c, h in enumerate(headers):
            ws_error.write(0, c, h, hdr_err)
        for r, (_, data) in enumerate(error_df.iterrows(), start=1):
            ws_error.write(r, 0, safe_int_value(data.get('Row Number', 0)), warn_fmt)
            ws_error.write(r, 1, str(data.get('ID', '')), warn_fmt)
            ws_error.write(r, 2, str(data.get('Property Name', '')), warn_fmt)
            ws_error.write(r, 3, str(data.get('Owner', '')), warn_fmt)
            ws_error.write(r, 4, str(data.get('Opportunity Status', '')), warn_fmt)
            ws_error.write(r, 5, str(data.get('Error Type', '')), warn_fmt)
            ws_error.write(r, 6, str(data.get('Error Detail', '')), warn_fmt)
            ws_error.write(r, 7, str(data.get('Date Sold', '')), warn_fmt)

        ws_error.set_column(0, 0, 10)
        ws_error.set_column(1, 1, 12)
        ws_error.set_column(2, 2, 25)
        ws_error.set_column(3, 3, 20)
        ws_error.set_column(4, 4, 18)
        ws_error.set_column(5, 5, 20)
        ws_error.set_column(6, 6, 42)
        ws_error.set_column(7, 7, 15)
    else:
        ws_error.write(0, 0, 'No errors found - all records processed successfully!', ok_fmt)

    wb.close()
    output.seek(0)
    return output

# ---------- PDF builder (labels still "Days Held" from your last change) ----------
def create_pdf_download(df_dict: Dict[str, pd.DataFrame], filename: str) -> Optional[BytesIO]:
    if not REPORTLAB_AVAILABLE:
        st.error("PDF generation requires reportlab. Please install it: pip install reportlab")
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(legal),
                            topMargin=0.5*inch, bottomMargin=0.5*inch,
                            leftMargin=0.5*inch, rightMargin=0.5*inch)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=12, alignment=1, textColor=colors.darkblue)
    quarter_style = ParagraphStyle('QuarterTitle', parent=styles['Heading2'], fontSize=14, spaceAfter=8, spaceBefore=12, textColor=colors.darkblue)
    summary_style = ParagraphStyle('SummaryStyle', parent=styles['Normal'], fontSize=10, spaceAfter=6)
    disclaimer_style = ParagraphStyle('Disclaimer', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=1, spaceAfter=12)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=9)

    story.append(Paragraph("Remarkable Land LLC - Sold Properties Report", title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", summary_style))
    story.append(Spacer(1, 12))

    sorted_qs = sort_quarters_chronologically(df_dict.keys())
    for quarter in sorted_qs:
        qdf = df_dict[quarter]
        if len(qdf) == 0:
            continue

        story.append(Paragraph(f"{quarter}", quarter_style))

        headers = ['Property\nName','Owner','State','County','Acres','Cost\nBasis',
                   'Date\nPurchased','Days\nHeld','Date\nSold','Gross Sales\nPrice',
                   'Closing\nCosts','Realized Gross\nProfit','Realized\nMarkup','Realized\nMargin','Realized\nIRR']
        table_data = [headers]

        for _, row in qdf.iterrows():
            prop_name = str(row.get('Property Name', '') or '')
            prop_para = Paragraph(prop_name, cell_style) if len(prop_name) > 25 else prop_name

            def dstr(x):
                return pd.to_datetime(x).strftime('%m/%d/%Y') if pd.notna(x) else ''
            r = [
                prop_para,
                str(row.get('Owner', '') or '')[:18],
                str(row.get('State', '') or ''),
                str(row.get('County', '') or '')[:15],
                f"{safe_numeric_value(row.get('Acres', 0)):.1f}",
                f"${safe_numeric_value(row.get('Cost Basis', 0)):,.0f}",
                dstr(row.get('Date Purchased')),
                f"{safe_int_value(row.get('Days Until Sold', 0)) if pd.notna(row.get('Days Until Sold')) else ''}",
                dstr(row.get('Date Sold')),
                f"${safe_numeric_value(row.get('Gross Sales Price', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Closing Costs', 0)):,.0f}",
                f"${safe_numeric_value(row.get('Realized Gross Profit', 0)):,.0f}",
                f"{safe_numeric_value(row.get('Realized Markup', 0)):.0f}%",
                f"{safe_numeric_value(row.get('Realized Margin', 0)):.0f}%",
                (f"{safe_numeric_value(row.get('Realized IRR', np.nan)):.0f}%" if pd.notna(row.get('Realized IRR', np.nan)) else '')
            ]
            table_data.append(r)

        # Slightly squeeze columns to add IRR
        col_widths = [1.9*inch,1.1*inch,0.4*inch,0.85*inch,0.5*inch,0.85*inch,
                      0.85*inch,0.6*inch,0.85*inch,1.05*inch,0.85*inch,1.05*inch,0.6*inch,0.6*inch,0.6*inch]
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 9),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('TOPPADDING', (0,0), (-1,0), 8),

            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('ALIGN', (0,1), (1,-1), 'LEFT'),
            ('ALIGN', (2,1), (3,-1), 'CENTER'),
            ('ALIGN', (4,1), (-1,-1), 'RIGHT'),
            ('GRID', (0,0), (-1,-1), 0.8, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,1), (-1,-1), 6),
            ('BOTTOMPADDING', (0,1), (-1,-1), 6),

            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.lightgrey]),
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

        stats = create_summary_stats(qdf)
        sdata = [
            ['Metric','Value','Metric','Value'],
            ['Properties Sold', f"{stats['total_properties']}", 'Total Cost Basis', f"${stats['total_cost_basis']:,.0f}"],
            ['Total Gross Sales', f"${stats['total_gross_sales']:,.0f}", 'Total Gross Profit', f"${stats['total_gross_profit']:,.0f}"],
            ['Average Markup', f"{stats['average_markup']:.0f}%", 'Median Markup', f"{stats['median_markup']:.0f}%"],
            ['Average Margin', f"{stats['average_margin']:.0f}%", 'Median Margin', f"{stats['median_margin']:.0f}%"],
            ['Average IRR', f"{stats['average_irr']:.0f}%", 'Median IRR', f"{stats['median_irr']:.0f}%"],
            ['Average Days Held', f"{stats['average_days']:.0f}", 'Median Days Held', f"{stats['median_days']:.0f}"]
        ]
        stab = Table(sdata, colWidths=[1.5*inch,1.5*inch,1.5*inch,1.5*inch])
        stab.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (1,1), (1,-1), 'RIGHT'),
            ('ALIGN', (3,1), (3,-1), 'RIGHT'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(stab)

        if quarter != sorted_qs[-1]:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 12))

    # Overall summary if multiple quarters
    if len([q for q in sorted_qs if q]) > 1:
        all_data = pd.concat(df_dict.values(), ignore_index=True)
        overall = create_summary_stats(all_data)
        story.append(Paragraph("Overall Summary", quarter_style))

        overall_data = [
            ['Total Properties', f"{overall['total_properties']}", '', ''],
            ['Total Gross Sales','Total Cost Basis','Total Closing Costs','Total Gross Profit'],
            [f"${overall['total_gross_sales']:,.0f}", f"${overall['total_cost_basis']:,.0f}",
             f"${overall['total_closing_costs']:,.0f}", f"${overall['total_gross_profit']:,.0f}"],
            ['Average Markup','Median Markup','Average Margin','Median Margin'],
            [f"{overall['average_markup']:.0f}%", f"{overall['median_markup']:.0f}%",
             f"{overall['average_margin']:.0f}%", f"{overall['median_margin']:.0f}%"],
            ['Average IRR','Median IRR','Max Days Held','Min Days Held'],
            [f"{overall['average_irr']:.0f}%", f"{overall['median_irr']:.0f}%",
             f"{overall['max_days']:.0f}", f"{overall['min_days']:.0f}"]
        ]
        ostab = Table(overall_data, colWidths=[1.75*inch]*4)
        ostab.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkgreen),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 11),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('SPAN', (0,0), (1,0)),

            ('BACKGROUND', (0,1), (-1,1), colors.lightgrey),
            ('BACKGROUND', (0,3), (-1,3), colors.lightgrey),
            ('BACKGROUND', (0,5), (-1,5), colors.lightgrey),
            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
            ('FONTNAME', (0,3), (-1,3), 'Helvetica-Bold'),
            ('FONTNAME', (0,5), (-1,5), 'Helvetica-Bold'),
            ('FONTSIZE', (0,1), (-1,1), 9),
            ('FONTSIZE', (0,3), (-1,3), 9),
            ('FONTSIZE', (0,5), (-1,5), 9),
            ('ALIGN', (0,1), (-1,1), 'CENTER'),
            ('ALIGN', (0,3), (-1,3), 'CENTER'),
            ('ALIGN', (0,5), (-1,5), 'CENTER'),

            ('FONTNAME', (0,2), (-1,2), 'Helvetica'),
            ('FONTNAME', (0,4), (-1,4), 'Helvetica'),
            ('FONTNAME', (0,6), (-1,6), 'Helvetica'),
            ('FONTSIZE', (0,2), (-1,2), 10),
            ('FONTSIZE', (0,4), (-1,4), 10),
            ('FONTSIZE', (0,6), (-1,6), 10),
            ('ALIGN', (0,2), (-1,2), 'CENTER'),
            ('ALIGN', (0,4), (-1,4), 'CENTER'),
            ('ALIGN', (0,6), (-1,6), 'CENTER'),

            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        story.append(ostab)
        story.append(Spacer(1, 12))

    story.append(Paragraph("Disclaimer: This data is sourced from our CRM and not our accounting software, based on then-available data. Final accounting data and results may vary slightly.", disclaimer_style))

    try:
        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

# ---------- UI ----------
def main():
    st.title("Sold Property Report")
    st.markdown("Generate quarterly reports for sold properties from Close.com CRM data")

    with st.expander("Instructions", expanded=True):
        st.markdown("""
        **How to use this report:**
        1. **Export from Close.com:** Export properties with **Remarkable – Sold** status  
        2. **Export All Fields**  
        3. **Upload CSV** below  
        4. **Pick filters** (empty = all)  
        5. **Download Excel/PDF**
        """)

    file = st.file_uploader(
        "Upload your Close.com CSV export (Remarkable – Sold, All Fields)",
        type=['csv'],
        help="Export properties with 'Remarkable – Sold' status from Close.com with all fields selected"
    )

    if file is None:
        st.info("Upload your CSV to generate the report.")
        return

    try:
        df_raw = pd.read_csv(file)
    except Exception as e:
        st.error(f"Could not read CSV: {e}")
        return

    df_proc, df_disp, error_df, total_records = process_data(df_raw)

    if len(df_proc) == 0:
        st.warning("No sold properties found in the uploaded data.")
        if len(error_df) > 0:
            st.error(f"Found {len(error_df)} records with errors that prevented processing.")
        return

    # Summary header
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records", total_records)
    c2.metric("Successfully Processed", len(df_proc))
    c3.metric("Records with Errors", len(error_df))
    success_rate = (len(df_proc) / total_records * 100) if total_records else 0.0
    c4.metric("Success Rate", f"{success_rate:.1f}%")

    if len(error_df) > 0:
        st.warning(f"Warning: {len(error_df)} records could not be processed. Download the error report below for details.")
        with st.expander("View Error Summary"):
            counts = error_df['Error Type'].value_counts()
            st.write("**Error Breakdown:**")
            for et, cnt in counts.items():
                st.write(f"• {et}: {cnt} records")

    st.subheader("Report Filters")

    # Quarters/Owners (empty = all)
    all_quarters = sort_quarters_chronologically(df_disp['Quarter_Year'].dropna().unique())
    all_owners = sorted([o for o in df_disp['Owner'].dropna().unique() if o != ''])

    c1, c2 = st.columns(2)
    with c1:
        selected_quarters = st.multiselect("Choose quarters (empty = all):", options=all_quarters)
    with c2:
        selected_owners = st.multiselect("Choose owners (empty = all):", options=all_owners)

    # Filtering logic: empty means "no filter"
    f = df_disp.copy()
    if selected_quarters:
        f = f[f['Quarter_Year'].isin(selected_quarters)]
    if selected_owners:
        f = f[f['Owner'].isin(selected_owners)]

    if len(f) == 0:
        st.warning("No properties match your selected filters.")
        return

    st.subheader("Sold Properties Report")

    # Display by quarter in chrono order
    quarters_to_show = sort_quarters_chronologically(f['Quarter_Year'].dropna().unique())

    for q in quarters_to_show:
        qdf = f[f['Quarter_Year'] == q].copy()
        if len(qdf) == 0:
            continue

        st.markdown(f"### {q}")

        display_cols = [
            'Property Name','Owner','State','County','Acres','Cost Basis',
            'Date Purchased','Opportunity Status','Days Until Sold',
            'Date Sold','Gross Sales Price','Closing Costs','Realized Gross Profit',
            'Realized Markup','Realized Margin','Realized IRR'
        ]
        qview = qdf[display_cols].copy()

        # Pretty formatting
        for col in ('Cost Basis','Gross Sales Price','Closing Costs','Realized Gross Profit'):
            if col in qview.columns:
                qview[col] = qview[col].apply(format_currency)
        for col in ('Realized Markup','Realized Margin','Realized IRR'):
            if col in qview.columns:
                qview[col] = qview[col].apply(format_percentage)
        if 'Date Purchased' in qview.columns:
            qview['Date Purchased'] = qdf['Date Purchased'].apply(lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else '')
        if 'Date Sold' in qview.columns:
            qview['Date Sold'] = qdf['Date Sold'].apply(lambda x: x.strftime('%m/%d/%Y') if pd.notna(x) else '')

        st.dataframe(qview, use_container_width=True)

        stats = create_summary_stats(qdf)
        c1, c2 = st.columns([1,1])
        with c2:
            st.metric("Total Properties", stats['total_properties'])

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Gross Sales", format_currency(stats['total_gross_sales']))
        c2.metric("Total Cost Basis", format_currency(stats['total_cost_basis']))
        c3.metric("Total Closing Costs", format_currency(stats['total_closing_costs']))
        c4.metric("Total Gross Profit", format_currency(stats['total_gross_profit']))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Average Markup", format_percentage(stats['average_markup']))
        c2.metric("Median Markup", format_percentage(stats['median_markup']))
        c3.metric("Average Margin", format_percentage(stats['average_margin']))
        c4.metric("Median Margin", format_percentage(stats['median_margin']))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Average IRR", format_percentage(stats['average_irr']))
        c2.metric("Median IRR", format_percentage(stats['median_irr']))
        c3.metric("Max Days Held", f"{safe_numeric_value(stats['max_days']):.0f}")
        c4.metric("Min Days Held", f"{safe_numeric_value(stats['min_days']):.0f}")

        st.divider()

    # Overall summary if showing multiple quarters
    if len(quarters_to_show) > 1:
        st.markdown("### Overall Summary")
        overall = create_summary_stats(f)

        c1, c2 = st.columns([1,1])
        with c2:
            st.metric("Total Properties", overall['total_properties'])

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Gross Sales", format_currency(overall['total_gross_sales']))
        c2.metric("Total Cost Basis", format_currency(overall['total_cost_basis']))
        c3.metric("Total Closing Costs", format_currency(overall['total_closing_costs']))
        c4.metric("Total Gross Profit", format_currency(overall['total_gross_profit']))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Average Markup", format_percentage(overall['average_markup']))
        c2.metric("Median Markup", format_percentage(overall['median_markup']))
        c3.metric("Average Margin", format_percentage(overall['average_margin']))
        c4.metric("Median Margin", format_percentage(overall['median_margin']))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Average IRR", format_percentage(overall['average_irr']))
        c2.metric("Median IRR", format_percentage(overall['median_irr']))
        c3.metric("Max Days Held", f"{safe_numeric_value(overall['max_days']):.0f}")
        c4.metric("Min Days Held", f"{safe_numeric_value(overall['min_days']):.0f}")

    # Downloads
    st.subheader("Download Reports")
    filtered_original = df_proc.copy()
    if selected_quarters:
        filtered_original = filtered_original[filtered_original['Quarter_Year'].isin(selected_quarters)]
    if selected_owners:
        filtered_original = filtered_original[filtered_original['custom.Asset_Owner'].isin(selected_owners)]

    quarter_data_dict = {q: f[f['Quarter_Year'] == q].copy() for q in quarters_to_show}

    today = datetime.now().strftime("%Y%m%d")
    excel_name = f"{today} Sold Property Report.xlsx"
    pdf_name = f"{today} Sold Property Report.pdf"
    err_name = f"{today} Import Error Report.xlsx"

    if len(error_df) > 0:
        d1, d2, d3 = st.columns(3)
    else:
        d1, d2 = st.columns(2)

    with d1:
        st.write("**Excel Report (Original Field Names)**")
        xfile = create_excel_download(filtered_original, excel_name)
        st.download_button(
            "Download Excel Report",
            data=xfile.getvalue(),
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Excel with original Close.com field names (now includes Realized_IRR)"
        )

    with d2:
        st.write("**PDF Report (Landscape Legal)**")
        if REPORTLAB_AVAILABLE:
            pfile = create_pdf_download(quarter_data_dict, pdf_name)
            if pfile:
                st.download_button(
                    "Download PDF Report",
                    data=pfile.getvalue(),
                    file_name=pdf_name,
                    mime="application/pdf"
                )
        else:
            st.warning("PDF generation requires reportlab. Run: pip install reportlab")
            st.info("Excel download is still available above.")

    if len(error_df) > 0:
        with d3:
            st.write("**Error Report**")
            efile = create_error_report_excel(error_df, total_records, len(df_proc), err_name)
            st.download_button(
                "Download Error Report",
                data=efile.getvalue(),
                file_name=err_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=f"Detailed report of {len(error_df)} records that could not be processed"
            )

    st.markdown("---")
    st.markdown("**Disclaimer:** This data is sourced from our CRM and not our accounting software, based on then-available data. Final accounting data and results may vary slightly.")

if __name__ == "__main__":
    main()

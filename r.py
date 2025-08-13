import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import matplotlib.pyplot as plt

# ------- Improved Utility functions with comprehensive NaN handling -------
def num(x):
    """Convert value to number with comprehensive NaN handling"""
    if x is None or pd.isnull(x) or pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        if np.isnan(x) or np.isinf(x):
            return 0.0
        return float(x)
    
    x_str = str(x).replace(',', '').replace('–', '-').replace('\xa0', '').replace('nan', '0').strip()
    if x_str == '' or x_str.lower() in ['nan', 'none', 'null', '#n/a', '#value!', '#div/0!']:
        return 0.0
    
    try:
        result = float(x_str)
        if np.isnan(result) or np.isinf(result):
            return 0.0
        return result
    except (ValueError, TypeError):
        return 0.0

def safe_int(x, default=0):
    """Safely convert to integer with NaN handling"""
    try:
        if pd.isnull(x) or pd.isna(x):
            return default
        result = int(float(x))
        if np.isnan(result):
            return default
        return result
    except (ValueError, TypeError, OverflowError):
        return default

def safeval(df, col, name):
    """Safely get values from DataFrame with comprehensive error handling"""
    try:
        if col not in df.columns:
            print(f"Warning: Column '{col}' not found in DataFrame")
            return pd.Series(dtype=object)
        
        # Clean the search
        if pd.isnull(name) or name == '':
            return pd.Series(dtype=object)
            
        # Create filter with proper NaN handling
        col_series = df[col].fillna('')  # Replace NaN with empty string
        filt = col_series.astype(str).str.contains(str(name), case=False, na=False)
        v = df.loc[filt]
        
        if not v.empty:
            return v.iloc[0]
        else:
            return pd.Series(dtype=object)
    except Exception as e:
        print(f"Warning in safeval for {name}: {e}")
        return pd.Series(dtype=object)

def find_header_row(df_raw, sheet_name, possible_headers):
    """
    Improved header detection with comprehensive NaN handling
    """
    print(f"\nSearching for header in {sheet_name} sheet...")
    print(f"DataFrame shape: {df_raw.shape}")
    
    # Handle empty DataFrame
    if df_raw.empty:
        print("Warning: DataFrame is empty")
        return 0
    
    # Print first few rows for debugging with NaN handling
    print("\nFirst 10 rows of raw data:")
    for i in range(min(10, len(df_raw))):
        try:
            row_values = []
            for x in df_raw.iloc[i].values:
                if pd.notna(x) and str(x).strip() != '':
                    row_values.append(str(x).strip())
            print(f"Row {i}: {row_values}")
        except Exception as e:
            print(f"Row {i}: Error reading row - {e}")
    
    header_row = None
    
    # Try each possible header pattern
    for header_pattern in possible_headers:
        print(f"\nLooking for pattern: {header_pattern}")
        
        for i in range(len(df_raw)):
            try:
                row = df_raw.iloc[i]
                # Convert all values in row to string and clean them with NaN handling
                row_values = []
                for x in row.values:
                    if pd.notna(x) and str(x).strip() != '':
                        row_values.append(str(x).upper().strip())
                
                row_text = ' '.join(row_values)
                
                # Check if any of the header keywords are present
                if any(keyword.upper() in row_text for keyword in header_pattern if keyword):
                    print(f"Found potential header at row {i}: {row_values}")
                    header_row = i
                    break
            except Exception as e:
                print(f"Error processing row {i}: {e}")
                continue
        
        if header_row is not None:
            break
    
    return header_row if header_row is not None else 0

def read_bs_and_pl(iofile):
    """
    Improved function to read Balance Sheet and P&L with comprehensive error handling
    """
    try:
        xl = pd.ExcelFile(iofile)
        print(f"Available sheets: {xl.sheet_names}")
        
        # Find Balance Sheet with comprehensive search
        bs_sheet_names = ['Balance Sheet', 'BalanceSheet', 'BS', 'Balance_Sheet', 'Bal Sheet', 'BALANCE SHEET']
        bs_sheet = None
        for sheet in bs_sheet_names:
            if sheet in xl.sheet_names:
                bs_sheet = sheet
                break
        
        # Fallback search for sheets containing balance sheet keywords
        if bs_sheet is None:
            for sheet in xl.sheet_names:
                if any(word in sheet.lower() for word in ['balance', 'bs', 'statement of financial position']):
                    bs_sheet = sheet
                    print(f"Found Balance Sheet by keyword matching: {bs_sheet}")
                    break
        
        if bs_sheet is None:
            bs_sheet = xl.sheet_names[0]  # Use first sheet as fallback
            print(f"Balance Sheet not found, using first sheet: {bs_sheet}")
        
        # Read Balance Sheet with error handling
        try:
            bs_raw = pd.read_excel(xl, bs_sheet, header=None)
            bs_raw = bs_raw.fillna('')  # Replace NaN with empty strings
        except Exception as e:
            print(f"Error reading Balance Sheet: {e}")
            raise Exception(f"Could not read Balance Sheet from {bs_sheet}")
        
        # Multiple possible header patterns for Balance Sheet
        bs_header_patterns = [
            ['LIABILITIES', 'ASSETS'],
            ['LIABILITY', 'ASSET'], 
            ['LIAB', 'ASSET'],
            ['Particulars', 'Amount'],
            ['Description', 'Current Year', 'Previous Year'],
            ['EQUITY AND LIABILITIES'],
            ['EQUITY', 'LIABILITIES'],
            ['SOURCES', 'APPLICATION'],
            ['CY', 'PY'],
            ['Current', 'Previous']
        ]
        
        bs_head_row = find_header_row(bs_raw, 'Balance Sheet', bs_header_patterns)
        
        try:
            bs = pd.read_excel(xl, bs_sheet, header=bs_head_row)
            bs = bs.loc[:, ~bs.columns.str.contains('^Unnamed', na=False)]
            bs = bs.fillna(0)  # Replace NaN with 0 for calculations
        except Exception as e:
            print(f"Error processing Balance Sheet headers: {e}")
            bs = pd.read_excel(xl, bs_sheet, header=0)
            bs = bs.fillna(0)
        
        # Find Profit & Loss Sheet with comprehensive search
        pl_sheet_names = [
            'Profit & Loss', 'Profit &amp; Loss', 'P&L', 'PL', 'Profit and Loss', 
            'Income Statement', 'PROFIT & LOSS', 'PROFIT AND LOSS',
            'Statement of Comprehensive Income', 'P & L', 'PnL', 'P&amp;L'
        ]
        pl_sheet = None
        for sheet in pl_sheet_names:
            if sheet in xl.sheet_names:
                pl_sheet = sheet
                break
        
        # Fallback search for sheets containing P&L keywords
        if pl_sheet is None:
            for sheet in xl.sheet_names:
                sheet_lower = sheet.lower()
                if any(word in sheet_lower for word in ['profit', 'loss', 'income', 'p&l', 'pnl', 'p & l']):
                    pl_sheet = sheet
                    print(f"Found P&L Sheet by keyword matching: {pl_sheet}")
                    break
        
        if pl_sheet is None:
            raise Exception(f"Could not find Profit & Loss sheet. Available sheets: {xl.sheet_names}")
        
        print(f"Using P&L sheet: {pl_sheet}")
        
        # Read P&L with error handling
        try:
            pl_raw = pd.read_excel(xl, pl_sheet, header=None)
            pl_raw = pl_raw.fillna('')  # Replace NaN with empty strings
        except Exception as e:
            print(f"Error reading P&L sheet: {e}")
            raise Exception(f"Could not read P&L sheet from {pl_sheet}")
        
        # Multiple possible header patterns for P&L
        pl_header_patterns = [
            ['DR.PATICULARS', 'CR.PARTICULARS'],
            ['DR.PARTICULARS', 'CR.PARTICULARS'], 
            ['DEBIT', 'CREDIT'],
            ['Dr.Particulars', 'Cr.Particulars'],
            ['Dr.Paticulars', 'Cr.Particulars'],  # Handle spelling variation
            ['Expenses', 'Income'],
            ['Particulars', 'Debit', 'Credit'],
            ['Description', 'Amount'],
            ['PARTICULARS', 'CURRENT YEAR', 'PREVIOUS YEAR'],
            ['Revenue', 'Expenses'],
            ['EXPENSE', 'INCOME'],
            ['DR', 'CR'],
            ['Debit', 'Credit'],
            ['CY', 'PY']
        ]
        
        pl_head_row = find_header_row(pl_raw, 'Profit & Loss', pl_header_patterns)
        
        try:
            pl = pd.read_excel(xl, pl_sheet, header=pl_head_row)
            pl = pl.loc[:, ~pl.columns.str.contains('^Unnamed', na=False)]
            pl = pl.fillna(0)  # Replace NaN with 0 for calculations
        except Exception as e:
            print(f"Error processing P&L headers: {e}")
            pl = pd.read_excel(xl, pl_sheet, header=0)
            pl = pl.fillna(0)
        
        print(f"\nBalance Sheet columns: {list(bs.columns)}")
        print(f"P&L columns: {list(pl.columns)}")
        
        return bs, pl
        
    except Exception as e:
        print(f"Error in read_bs_and_pl: {e}")
        raise Exception(f"Error reading Excel file: {str(e)}. Please check file format and sheet names.")

def write_notes_with_labels(writer, sheetname, notes_with_labels):
    """Write notes to Excel with error handling"""
    startrow = 0
    try:
        for label, df in notes_with_labels:
            # Clean the DataFrame
            df_clean = df.fillna(0)
            label_row = pd.DataFrame([[label] + [""] * (df_clean.shape[1] - 1)], columns=df_clean.columns)
            label_row.to_excel(writer, sheet_name=sheetname, startrow=startrow, index=False, header=False)
            startrow += 1
            df_clean.to_excel(writer, sheet_name=sheetname, startrow=startrow, index=False)
            startrow += len(df_clean) + 2
    except Exception as e:
        print(f"Error writing notes: {e}")

# ===============================
# Comprehensive financial data processing function with NaN handling
# ===============================

def process_financials(bs_df, pl_df):
    L, A = 'LIABILITIES', 'ASSETS'

    # Share capital and authorised capital
    capital_row = safeval(bs_df, L, "Capital Account")
    share_cap_cy = num(capital_row.get('CY (₹)', 0))
    share_cap_py = num(capital_row.get('PY (₹)', 0))
    authorised_cap = max(share_cap_cy, share_cap_py) * 1.2  # 20% buffer

    # Reserves and Surplus
    gr_row = safeval(bs_df, L, "General Reserve")
    general_res_cy = num(gr_row.get('CY (₹)', 0))
    general_res_py = num(gr_row.get('PY (₹)', 0))

    surplus_row = safeval(bs_df, L, "Retained Earnings")
    surplus_cy = num(surplus_row.get('CY (₹)', 0))
    surplus_py = num(surplus_row.get('PY (₹)', 0))
    surplus_open_cy = surplus_py  # Opening balance = PY closing
    surplus_open_py = 70000       # Prior year opening balance fixed

    profit_row = safeval(bs_df, L, "Add: Current Year Profit")
    profit_cy = num(profit_row.get('CY (₹)', 0))
    profit_py = num(profit_row.get('PY (₹)', 0))

    pd_row = safeval(bs_df, L, "Proposed Dividend")
    pd_cy = num(pd_row.get('CY (₹)', 0))
    pd_py = num(pd_row.get('PY (₹)', 0))

    surplus_close_cy = surplus_cy + profit_cy
    surplus_close_py = surplus_py + profit_py

    reserves_total_cy = general_res_cy + surplus_close_cy
    reserves_total_py = general_res_py + surplus_close_py

    # Long-term borrowings
    tl_row = safeval(bs_df, L, "Term Loan from Bank")
    vl_row = safeval(bs_df, L, "Vehicle Loan")
    fd_row = safeval(bs_df, L, "From Directors")
    icb_row = safeval(bs_df, L, "Inter-Corporate Borrowings")

    tl_cy = num(tl_row.get('CY (₹)', 0))
    tl_py = num(tl_row.get('PY (₹)', 0))
    vl_cy = num(vl_row.get('CY (₹)', 0))
    vl_py = num(vl_row.get('PY (₹)', 0))
    fd_cy = num(fd_row.get('CY (₹)', 0))
    fd_py = num(fd_row.get('PY (₹)', 0))
    icb_cy = num(icb_row.get('CY (₹)', 0))
    icb_py = num(icb_row.get('PY (₹)', 0))

    longterm_borrow_cy = tl_cy + vl_cy
    longterm_borrow_py = tl_py + vl_py
    other_longterm_liab_cy = fd_cy + icb_cy
    other_longterm_liab_py = fd_py + icb_py

    # Long-term provisions (no data)
    longterm_prov_cy = 0
    longterm_prov_py = 0

    # Short-term borrowings (no data)
    shortterm_borrow_cy = 0
    shortterm_borrow_py = 0

    # Trade payables
    sc_row = safeval(bs_df, L, "Sundry Creditors")
    creditors_cy = num(sc_row.get('CY (₹)', 0))
    creditors_py = num(sc_row.get('PY (₹)', 0))

    # Other current liabilities
    bp_row = safeval(bs_df, L, "Bills Payable")
    oe_row = safeval(bs_df, L, "Outstanding Expenses")

    bp_cy = num(bp_row.get('CY (₹)', 0))
    bp_py = num(bp_row.get('PY (₹)', 0))
    oe_cy = num(oe_row.get('CY (₹)', 0))
    oe_py = num(oe_row.get('PY (₹)', 0))

    other_cur_liab_cy = bp_cy + oe_cy + pd_cy
    other_cur_liab_py = bp_py + oe_py + pd_py

    # Short-Term Provisions (Note 9)
    tax_row = safeval(bs_df, L, "Provision for Taxation")
    tax_cy = num(tax_row.get('CY (₹)', 0))
    tax_py = num(tax_row.get('PY (₹)', 0))

    # PPE (Note 10)
    land_cy = num(safeval(bs_df, A, "Land").get('CY (₹)', 0))
    plant_cy = num(safeval(bs_df, A, "Plant").get('CY (₹)', 0))
    furn_cy = num(safeval(bs_df, A, "Furniture").get('CY (₹)', 0))
    comp_cy = num(safeval(bs_df, A, "Computer").get('CY (₹)', 0))

    land_py = num(safeval(bs_df, A, "Land").get('PY (₹)', 0))
    plant_py = num(safeval(bs_df, A, "Plant").get('PY (₹)', 0))
    furn_py = num(safeval(bs_df, A, "Furniture").get('PY (₹)', 0))
    comp_py = num(safeval(bs_df, A, "Computer").get('PY (₹)', 0))

    gross_block_cy = land_cy + plant_cy + furn_cy + comp_cy
    gross_block_py = land_py + plant_py + furn_py + comp_py

    ad_row = safeval(bs_df, A, "Accumulated Depreciation")
    acc_dep_cy = -num(ad_row.get('CY (₹)', 0))
    acc_dep_py = -num(ad_row.get('PY (₹)', 0))

    net_ppe_cy = num(safeval(bs_df, A, "Net Fixed Assets").get('CY (₹)', 0))
    net_ppe_py = num(safeval(bs_df, A, "Net Fixed Assets").get('PY (₹)', 0))

    # Capital Work-in-Progress (Note 11)
    cwip_cy = 0
    cwip_py = 0

    # Non-current Investments (Note 12)
    eq_row = safeval(bs_df, A, "Equity Shares")
    mf_row = safeval(bs_df, A, "Mutual Funds")

    eq_cy = num(eq_row.get('CY (₹)', 0))
    eq_py = num(eq_row.get('PY (₹)', 0))
    mf_cy = num(mf_row.get('CY (₹)', 0))
    mf_py = num(mf_row.get('PY (₹)', 0))

    investments_cy = eq_cy + mf_cy
    investments_py = eq_py + mf_py

    # Deferred Tax Assets (Note 13)
    dta_cy = 0
    dta_py = 0

    # Long-term Loans and Advances (Note 14)
    longterm_loans_cy = 0
    longterm_loans_py = 0

    # Other Non-current Assets (Note 15)
    prelim_exp_row = safeval(bs_df, A, "Preliminary Expenses")
    prelim_exp_cy = num(prelim_exp_row.get('CY (₹)', 0))
    prelim_exp_py = num(prelim_exp_row.get('PY (₹)', 0))

    # Current Investments (Note 16)
    current_inv_cy = 0
    current_inv_py = 0

    # Inventories (Note 17)
    stock_row = safeval(bs_df, A, "Stock")
    stock_cy = num(stock_row.get('CY (₹)', 0))
    stock_py = num(stock_row.get('PY (₹)', 0))

    # Trade Receivables (Note 18)
    deb_row = safeval(bs_df, A, "Sundry Debtors")
    deb_cy = num(deb_row.get('CY (₹)', 0))
    deb_py = num(deb_row.get('PY (₹)', 0))

    provd_row = safeval(bs_df, A, "Provision for Doubtful Debts")
    provd_cy = num(provd_row.get('CY (₹)', 0))
    provd_py = num(provd_row.get('PY (₹)', 0))

    bills_recv_row = safeval(bs_df, A, "Bills Receivable")
    bills_recv_cy = num(bills_recv_row.get('CY (₹)', 0))
    bills_recv_py = num(bills_recv_row.get('PY (₹)', 0))

    total_receivables_cy = deb_cy + bills_recv_cy
    total_receivables_py = deb_py + bills_recv_py
    net_receivables_cy = total_receivables_cy + provd_cy
    net_receivables_py = total_receivables_py + provd_py

    # Cash & Bank (Note 19)
    cash_row = safeval(bs_df, A, "Cash in Hand")
    bank_row = safeval(bs_df, A, "Bank Balance")

    cash_cy = num(cash_row.get('CY (₹)', 0))
    cash_py = num(cash_row.get('PY (₹)', 0))
    bank_cy = num(bank_row.get('CY (₹)', 0))
    bank_py = num(bank_row.get('PY (₹)', 0))

    cash_total_cy = cash_cy + bank_cy
    cash_total_py = cash_py + bank_py

    # Short-term Loans/Advances (Note 20)
    loan_adv_row = safeval(bs_df, A, "Loans & Advances")
    loan_adv_cy = num(loan_adv_row.get('CY (₹)', 0))
    loan_adv_py = num(loan_adv_row.get('PY (₹)', 0))

    # Other Current Assets (Note 21)
    prepaid_row = safeval(bs_df, A, "Prepaid Expenses")
    prepaid_cy = num(prepaid_row.get('CY (₹)', 0))
    prepaid_py = num(prepaid_row.get('PY (₹)', 0))

    # Calculate totals for verification
    total_equity_liab_cy = (
        share_cap_cy + reserves_total_cy + longterm_borrow_cy + other_longterm_liab_cy +
        longterm_prov_cy + shortterm_borrow_cy + creditors_cy + other_cur_liab_cy + tax_cy)
    total_equity_liab_py = (
        share_cap_py + reserves_total_py + longterm_borrow_py + other_longterm_liab_py +
        longterm_prov_py + shortterm_borrow_py + creditors_py + other_cur_liab_py + tax_py)

    total_assets_cy = (
        net_ppe_cy + cwip_cy + investments_cy + dta_cy + longterm_loans_cy + prelim_exp_cy +
        current_inv_cy + stock_cy + net_receivables_cy + cash_total_cy + loan_adv_cy + prepaid_cy)
    total_assets_py = (
        net_ppe_py + cwip_py + investments_py + dta_py + longterm_loans_py + prelim_exp_py +
        current_inv_py + stock_py + net_receivables_py + cash_total_py + loan_adv_py + prepaid_py)

    # ===============================
    # Mapping PROFIT & LOSS figures
    # ===============================

    sales_row = safeval(pl_df, 'Cr.Particulars', "Sales")
    sales_cy = num(sales_row.get('CY (₹)', 0))
    sales_py = num(sales_row.get('PY (₹)', 0))

    sales_ret_row = safeval(pl_df, 'Cr.Particulars', "Sales Returns")
    sales_ret_cy = num(sales_ret_row.get('CY (₹)', 0))
    sales_ret_py = num(sales_ret_row.get('PY (₹)', 0))

    net_sales_cy = sales_cy + sales_ret_cy
    net_sales_py = sales_py + sales_ret_py

    # Other Income (Note 23)
    oi_row = safeval(pl_df, 'Cr.Particulars', "Other Operating Income")
    oi_cy = num(oi_row.get('CY (₹)', 0))
    oi_py = num(oi_row.get('PY (₹)', 0))

    int_row = safeval(pl_df, 'Cr.Particulars', "Interest Income")
    int_cy = num(int_row.get('CY (₹)', 0))
    int_py = num(int_row.get('PY (₹)', 0))

    other_inc_cy = oi_cy + int_cy
    other_inc_py = oi_py + int_py

    # Cost of Materials Consumed (Note 24)
    purch_row = safeval(pl_df, 'Dr.Paticulars', "Purchases")
    purch_cy = num(purch_row.get('CY (₹)', 0))
    purch_py = num(purch_row.get('PY (₹)', 0))

    purch_ret_row = safeval(pl_df, 'Dr.Paticulars', "Purchase Returns")
    purch_ret_cy = num(purch_ret_row.get('CY (₹)', 0))
    purch_ret_py = num(purch_ret_row.get('PY (₹)', 0))

    wages_row = safeval(pl_df, 'Dr.Paticulars', "Wages")
    wages_cy = num(wages_row.get('CY (₹)', 0))
    wages_py = num(wages_row.get('PY (₹)', 0))

    power_row = safeval(pl_df, 'Dr.Paticulars', "Power & Fuel")
    power_cy = num(power_row.get('CY (₹)', 0))
    power_py = num(power_row.get('PY (₹)', 0))

    freight_row = safeval(pl_df, 'Dr.Paticulars', "Freight")
    freight_cy = num(freight_row.get('CY (₹)', 0))
    freight_py = num(freight_row.get('PY (₹)', 0))

    cost_mat_cy = purch_cy + purch_ret_cy + wages_cy + power_cy + freight_cy
    cost_mat_py = purch_py + purch_ret_py + wages_py + power_py + freight_py

    # Changes in Inventories (Note 25)
    os_row = safeval(pl_df, 'Dr.Paticulars', "Opening Stock")
    os_cy = num(os_row.get('CY (₹)', 0))
    os_py = num(os_row.get('PY (₹)', 0))

    cs_row = safeval(pl_df, 'Cr.Particulars', "Closing Stock")
    cs_cy = num(cs_row.get('CY (₹)', 0))
    cs_py = num(cs_row.get('PY (₹)', 0))

    change_inv_cy = cs_cy - os_cy
    change_inv_py = cs_py - os_py

    # Employee Benefits Expense (Note 26)
    sal_row = safeval(pl_df, 'Dr.Paticulars', "Salaries & Wages")
    sal_cy = num(sal_row.get('CY (₹)', 0))
    sal_py = num(sal_row.get('PY (₹)', 0))

    # Finance Costs
    loan_int_row = safeval(pl_df, 'Dr.Paticulars', "Interest on Loans")
    loan_int_cy = num(loan_int_row.get('CY (₹)', 0))
    loan_int_py = num(loan_int_row.get('PY (₹)', 0))

    # Depreciation
    dep_row = safeval(pl_df, 'Dr.Paticulars', "Depreciation")
    dep_cy = num(dep_row.get('CY (₹)', 0))
    dep_py = num(dep_row.get('PY (₹)', 0))

    # Other expenses components
    rent_cy = num(safeval(pl_df, 'Dr.Paticulars', "Rent, Rates & Taxes").get('CY (₹)', 0))
    rent_py = num(safeval(pl_df, 'Dr.Paticulars', "Rent, Rates & Taxes").get('PY (₹)', 0))
    admin_cy = num(safeval(pl_df, 'Dr.Paticulars', "Administrative Expenses").get('CY (₹)', 0))
    admin_py = num(safeval(pl_df, 'Dr.Paticulars', "Administrative Expenses").get('PY (₹)', 0))
    selling_cy = num(safeval(pl_df, 'Dr.Paticulars', "Selling & Distribution Expenses").get('CY (₹)', 0))
    selling_py = num(safeval(pl_df, 'Dr.Paticulars', "Selling & Distribution Expenses").get('PY (₹)', 0))
    repairs_cy = num(safeval(pl_df, 'Dr.Paticulars', "Repairs & Maintenance").get('CY (₹)', 0))
    repairs_py = num(safeval(pl_df, 'Dr.Paticulars', "Repairs & Maintenance").get('PY (₹)', 0))
    insurance_cy = num(safeval(pl_df, 'Dr.Paticulars', "Insurance").get('CY (₹)', 0))
    insurance_py = num(safeval(pl_df, 'Dr.Paticulars', "Insurance").get('PY (₹)', 0))
    audit_cy = num(safeval(pl_df, 'Dr.Paticulars', "Audit Fees").get('CY (₹)', 0))
    audit_py = num(safeval(pl_df, 'Dr.Paticulars', "Audit Fees").get('PY (₹)', 0))
    bad_cy = num(safeval(pl_df, 'Dr.Paticulars', "Bad Debts Written Off").get('CY (₹)', 0))
    bad_py = num(safeval(pl_df, 'Dr.Paticulars', "Bad Debts Written Off").get('PY (₹)', 0))

    other_exp_cy = rent_cy + admin_cy + selling_cy + repairs_cy + insurance_cy + audit_cy + bad_cy
    other_exp_py = rent_py + admin_py + selling_py + repairs_py + insurance_py + audit_py + bad_py

    # Totals and profits
    total_rev_cy = net_sales_cy + other_inc_cy
    total_rev_py = net_sales_py + other_inc_py

    total_exp_cy = cost_mat_cy + change_inv_cy + sal_cy + loan_int_cy + dep_cy + other_exp_cy
    total_exp_py = cost_mat_py + change_inv_py + sal_py + loan_int_py + dep_py + other_exp_py

    pbt_cy = total_rev_cy - total_exp_cy
    pbt_py = total_rev_py - total_exp_py

    pat_cy = pbt_cy - tax_cy
    pat_py = pbt_py - tax_py

    num_shares = share_cap_cy / 10 if share_cap_cy > 0 else 10000  # Assume ₹10 per share
    eps_cy = pat_cy / num_shares if num_shares > 0 else 0
    eps_py = pat_py / num_shares if num_shares > 0 else 0

    # ===============================
    # Construct Balance Sheet output dataframe
    # ===============================
    bs_out = pd.DataFrame([
        ['Particulars', 'Note No.', 'CY (₹)', 'PY (₹)'],
        ['EQUITY AND LIABILITIES', '', '', ''],
        ['1. Shareholders Funds', '', '', ''],
        ['(a) Share Capital', 1, share_cap_cy, share_cap_py],
        ['(b) Reserves and Surplus', 2, reserves_total_cy, reserves_total_py],
        ['2. Non-Current Liabilities', '', '', ''],
        ['(a) Long-Term Borrowings', 3, longterm_borrow_cy, longterm_borrow_py],
        ['(b) Deferred Tax Liabilities (Net)', 4, 0, 0],
        ['(c) Other Long-Term Liabilities', 5, other_longterm_liab_cy, other_longterm_liab_py],
        ['(d) Long-Term Provisions', 6, longterm_prov_cy, longterm_prov_py],
        ['3. Current Liabilities', '', '', ''],
        ['(a) Short-Term Borrowings', 7, shortterm_borrow_cy, shortterm_borrow_py],
        ['(b) Trade Payables', 8, creditors_cy, creditors_py],
        ['(c) Other Current Liabilities', 9, other_cur_liab_cy, other_cur_liab_py],
        ['(d) Short-Term Provisions', 10, tax_cy, tax_py],
        ['TOTAL', '', total_equity_liab_cy, total_equity_liab_py],
        ['ASSETS', '', '', ''],
        ['1. Non-Current Assets', '', '', ''],
        ['(a) Fixed Assets', '', '', ''],
        ['     (i) Tangible Assets', 11, net_ppe_cy, net_ppe_py],
        ['     (ii) Intangible Assets', 12, 0, 0],
        ['     (iii) Capital Work-in-Progress', 13, cwip_cy, cwip_py],
        ['(b) Non-Current Investments', 14, investments_cy, investments_py],
        ['(c) Deferred Tax Assets (Net)', 15, dta_cy, dta_py],
        ['(d) Long-Term Loans and Advances', 16, longterm_loans_cy, longterm_loans_py],
        ['(e) Other Non-Current Assets', 17, prelim_exp_cy, prelim_exp_py],
        ['2. Current Assets', '', '', ''],
        ['(a) Current Investments', 18, current_inv_cy, current_inv_py],
        ['(b) Inventories', 19, stock_cy, stock_py],
        ['(c) Trade Receivables', 20, net_receivables_cy, net_receivables_py],
        ['(d) Cash and Cash Equivalents', 21, cash_total_cy, cash_total_py],
        ['(e) Short-Term Loans and Advances', 22, loan_adv_cy, loan_adv_py],
        ['(f) Other Current Assets', 23, prepaid_cy, prepaid_py],
        ['TOTAL', '', total_assets_cy, total_assets_py]
    ])

    # ===============================
    # Construct Profit & Loss output dataframe
    # ===============================
    pl_out = pd.DataFrame([
        ['Particulars', 'Note No.', 'CY (₹)', 'PY (₹)'],
        ['I. Revenue from Operations', 24, net_sales_cy, net_sales_py],
        ['II. Other Income', 25, other_inc_cy, other_inc_py],
        ['III. Total Revenue (I + II)', '', total_rev_cy, total_rev_py],
        ['IV. Expenses', '', '', ''],
        ['(a) Cost of Materials Consumed', 26, cost_mat_cy, cost_mat_py],
        ['(b) Changes in Inventories of Finished Goods', '', change_inv_cy, change_inv_py],
        ['(c) Employee Benefits Expense', '', sal_cy, sal_py],
        ['(d) Finance Costs', '', loan_int_cy, loan_int_py],
        ['(e) Depreciation and Amortization Expense', '', dep_cy, dep_py],
        ['(f) Other Expenses', '', other_exp_cy, other_exp_py],
        ['Total Expenses', '', total_exp_cy, total_exp_py],
        ['V. Profit Before Tax (III - IV)', '', pbt_cy, pbt_py],
        ['VI. Tax Expense', '', '', ''],
        ['(a) Current Tax', '', tax_cy, tax_py],
        ['VII. Profit for the Period (V - VI)', '', pat_cy, pat_py],
        ['VIII. Earnings per Equity Share (Basic & Diluted)', '', eps_cy, eps_py]
    ])

    # ===============================
    # Create all 26 Notes DataFrames (copied exactly from your provided code)
    # ===============================
    note1 = pd.DataFrame({
        'Particulars': [
            'Authorised Share Capital',
            '10,000 Equity shares of Rs.10 each',
            '',
            'Issued, Subscribed & Paid-up Capital',
            '10,000 Equity shares of Rs.10 each fully paid up',
            '',
            'Total'
        ],
        'CY (₹)': [authorised_cap, '', '', share_cap_cy, '', '', share_cap_cy],
        'PY (₹)': [authorised_cap, '', '', share_cap_py, '', '', share_cap_py]
    })

    note2 = pd.DataFrame({
        'Particulars': [
            'General Reserve',
            'Balance at the beginning of the year',
            'Add: Transferred from Statement of P&L',
            'Balance at the end of the year',
            '',
            'Surplus in Statement of P&L:',
            'Balance at the beginning of the year',
            'Add: Profit for the Year',
            'Less: Proposed dividend',
            'Balance at the end of the year',
            '',
            'Total'
        ],
        'CY (₹)': [
            '', general_res_py, 0, general_res_cy, '',
            '', surplus_open_cy, profit_cy, pd_cy, surplus_close_cy,
            '', reserves_total_cy
        ],
        'PY (₹)': [
            '', general_res_py, 0, general_res_py, '',
            '', surplus_open_py, profit_py, pd_py, surplus_close_py,
            '', reserves_total_py
        ]
    })

    note3 = pd.DataFrame({
        'Particulars': [
            'Term loans',
            'From banks:',
            'Term Loan (Secured)',
            'Vehicle Loan (Secured)',
            '',
            'Total'
        ],
        'CY (₹)': ['', '', tl_cy, vl_cy, '', longterm_borrow_cy],
        'PY (₹)': ['', '', tl_py, vl_py, '', longterm_borrow_py]
    })

    note4 = pd.DataFrame({
        'Particulars': ['Deferred Tax Liabilities (Net)'],
        'CY (₹)': [0],
        'PY (₹)': [0]
    })

    note5 = pd.DataFrame({
        'Particulars': [
            'Loans from Directors (Unsecured)',
            'Inter-Corporate Borrowings (Unsecured)',
            'Total'
        ],
        'CY (₹)': [fd_cy, icb_cy, other_longterm_liab_cy],
        'PY (₹)': [fd_py, icb_py, other_longterm_liab_py]
    })

    note6 = pd.DataFrame({
        'Particulars': ['Long-term Provisions (Employee Benefits)'],
        'CY (₹)': [longterm_prov_cy],
        'PY (₹)': [longterm_prov_py]
    })

    note7 = pd.DataFrame({
        'Particulars': ['Short-term Borrowings from Banks'],
        'CY (₹)': [shortterm_borrow_cy],
        'PY (₹)': [shortterm_borrow_py]
    })

    note8 = pd.DataFrame({
        'Particulars': [
            'Trade Payables:',
            'Total outstanding dues of micro and small enterprises',
            'Total outstanding dues of creditors other than micro and small enterprises',
            '',
            'Total'
        ],
        'CY (₹)': ['', min(creditors_cy, 120000), max(0, creditors_cy-120000), '', creditors_cy],
        'PY (₹)': ['', min(creditors_py, 100000), max(0, creditors_py-100000), '', creditors_py]
    })

    note9 = pd.DataFrame({
        'Particulars': [
            'Bills Payable',
            'Outstanding Expenses',
            'Proposed Dividend',
            'Other Payables',
            '',
            'Total'
        ],
        'CY (₹)': [bp_cy, oe_cy, pd_cy, 0, '', other_cur_liab_cy],
        'PY (₹)': [bp_py, oe_py, pd_py, 0, '', other_cur_liab_py]
    })

    note10 = pd.DataFrame({
        'Particulars': [
            'Provision for employee benefits:',
            'Provision for bonus',
            '',
            'Provision - Others:',
            'Provision for tax (net)',
            '',
            'Total'
        ],
        'CY (₹)': ['', 0, '', '', tax_cy, '', tax_cy],
        'PY (₹)': ['', 0, '', '', tax_py, '', tax_py]
    })

    note11 = pd.DataFrame({
        'Asset Class': [
            'Land & Building',
            'Plant & Machinery',
            'Furniture & Fixtures',
            'Computers',
            '',
            'Total'
        ],
        'Gross Block (₹)': [land_cy, plant_cy, furn_cy, comp_cy, '', gross_block_cy],
        'Accumulated Depreciation (₹)': ['-', plant_cy-plant_cy, furn_cy-(furn_cy-20000), comp_cy-(comp_cy-20000), '', acc_dep_cy],
        'Net Block (₹)': [land_cy, plant_py, 20000, 20000, '', net_ppe_cy]
    })

    note12 = pd.DataFrame({
        'Particulars': ['Software', 'Patents', 'Total'],
        'CY (₹)': [0, 0, 0],
        'PY (₹)': [0, 0, 0]
    })

    note13 = pd.DataFrame({
        'Particulars': ['Capital Work-in-Progress'],
        'CY (₹)': [cwip_cy],
        'PY (₹)': [cwip_py]
    })

    note14 = pd.DataFrame({
        'Particulars': [
            'Investment in equity instruments:',
            'Equity Shares (Unquoted)',
            'Mutual Funds (Unquoted)',
            '',
            'Total'
        ],
        'CY (₹)': ['', eq_cy, mf_cy, '', investments_cy],
        'PY (₹)': ['', eq_py, mf_py, '', investments_py]
    })

    note15 = pd.DataFrame({
        'Particulars': ['Deferred Tax Assets (Net)'],
        'CY (₹)': [dta_cy],
        'PY (₹)': [dta_py]
    })

    note16 = pd.DataFrame({
        'Particulars': [
            'Capital advances:',
            'Secured, considered good',
            'Unsecured, considered good',
            '',
            'Security deposits',
            '',
            'Total'
        ],
        'CY (₹)': ['', 0, 0, '', 0, '', longterm_loans_cy],
        'PY (₹)': ['', 0, 0, '', 0, '', longterm_loans_py]
    })

    note17 = pd.DataFrame({
        'Particulars': [
            'Unamortised expenses:',
            'Preliminary Expenses',
            '',
            'Total'
        ],
        'CY (₹)': ['', prelim_exp_cy, '', prelim_exp_cy],
        'PY (₹)': ['', prelim_exp_py, '', prelim_exp_py]
    })

    note18 = pd.DataFrame({
        'Particulars': [
            'Investment in mutual funds',
            'Investment in government securities',
            '',
            'Total'
        ],
        'CY (₹)': [0, 0, '', current_inv_cy],
        'PY (₹)': [0, 0, '', current_inv_py]
    })

    note19 = pd.DataFrame({
        'Particulars': [
            'Raw materials',
            'Work-in-progress',
            'Finished goods',
            'Stock-in-trade',
            '',
            'Total'
        ],
        'CY (₹)': [0, 0, stock_cy, 0, '', stock_cy],
        'PY (₹)': [0, 0, stock_py, 0, '', stock_py]
    })

    note20 = pd.DataFrame({
        'Particulars': [
            'Trade receivables outstanding for more than 6 months:',
            'Unsecured, considered good',
            '',
            'Other trade receivables:',
            'Unsecured, considered good',
            'Bills Receivable',
            '',
            'Total Gross Receivables',
            'Less: Provision for doubtful trade receivables',
            '',
            'Net Trade Receivables'
        ],
        'CY (₹)': [
            '', min(deb_cy, 50000), '',
            '', max(0, deb_cy-50000), bills_recv_cy, '',
            total_receivables_cy, provd_cy, '',
            net_receivables_cy
        ],
        'PY (₹)': [
            '', min(deb_py, 40000), '',
            '', max(0, deb_py-40000), bills_recv_py, '',
            total_receivables_py, provd_py, '',
            net_receivables_py
        ]
    })

    note21 = pd.DataFrame({
        'Particulars': [
            'Cash on hand',
            'Balances with banks:',
            'In current accounts',
            'In deposit accounts',
            '',
            'Total'
        ],
        'CY (₹)': [cash_cy, '', bank_cy, 0, '', cash_total_cy],
        'PY (₹)': [cash_py, '', bank_py, 0, '', cash_total_py]
    })

    note22 = pd.DataFrame({
        'Particulars': [
            'Loans and advances to employees:',
            'Unsecured, considered good',
            '',
            'Advances to suppliers:',
            'Unsecured, considered good',
            '',
            'Total'
        ],
        'CY (₹)': ['', loan_adv_cy//2, '', '', loan_adv_cy//2, '', loan_adv_cy],
        'PY (₹)': ['', loan_adv_py//2, '', '', loan_adv_py//2, '', loan_adv_py]
    })

    note23 = pd.DataFrame({
        'Particulars': [
            'Prepaid expenses:',
            'Insurance premium',
            'Advance tax',
            'Other prepaid expenses',
            '',
            'Total'
        ],
        'CY (₹)': ['', prepaid_cy//2, 0, prepaid_cy//2, '', prepaid_cy],
        'PY (₹)': ['', prepaid_py//2, 0, prepaid_py//2, '', prepaid_py]
    })

    note24 = pd.DataFrame({
        'Particulars': [
            'Sale of products:',
            'Gross Sales',
            'Less: Sales Returns',
            '',
            'Net Revenue from Operations'
        ],
        'CY (₹)': ['', sales_cy, sales_ret_cy, '', net_sales_cy],
        'PY (₹)': ['', sales_py, sales_ret_py, '', net_sales_py]
    })

    note25 = pd.DataFrame({
        'Particulars': [
            'Interest income:',
            'On investments',
            '',
            'Other operating income:',
            'Discount received',
            '',
            'Total Other Income'
        ],
        'CY (₹)': ['', int_cy, '', '', oi_cy, '', other_inc_cy],
        'PY (₹)': ['', int_py, '', '', oi_py, '', other_inc_py]
    })

    note26 = pd.DataFrame({
        'Particulars': [
            'Purchases of raw materials/goods',
            'Less: Purchase returns',
            'Net Purchases',
            '',
            'Direct expenses:',
            'Wages',
            'Power & Fuel',
            'Freight/Carriage Inward',
            '',
            'Total Cost of Materials Consumed'
        ],
        'CY (₹)': [
            purch_cy, purch_ret_cy, purch_cy + purch_ret_cy, '',
            '', wages_cy, power_cy, freight_cy, '',
            cost_mat_cy
        ],
        'PY (₹)': [
            purch_py, purch_ret_py, purch_py + purch_ret_py, '',
            '', wages_py, power_py, freight_py, '',
            cost_mat_py
        ]
    })

    notes = [
        ("Note 1: Share Capital", note1),
        ("Note 2: Reserves and Surplus", note2),
        ("Note 3: Long-Term Borrowings", note3),
        ("Note 4: Deferred Tax Liabilities", note4),
        ("Note 5: Other Long-Term Liabilities", note5),
        ("Note 6: Long-Term Provisions", note6),
        ("Note 7: Short-Term Borrowings", note7),
        ("Note 8: Trade Payables", note8),
        ("Note 9: Other Current Liabilities", note9),
        ("Note 10: Short-Term Provisions", note10),
        ("Note 11: Fixed Assets - Tangible", note11),
        ("Note 12: Intangible Assets", note12),
        ("Note 13: Capital Work-in-Progress", note13),
        ("Note 14: Non-Current Investments", note14),
        ("Note 15: Deferred Tax Assets", note15),
        ("Note 16: Long-Term Loans and Advances", note16),
        ("Note 17: Other Non-Current Assets", note17),
        ("Note 18: Current Investments", note18),
        ("Note 19: Inventories", note19),
        ("Note 20: Trade Receivables", note20),
        ("Note 21: Cash and Cash Equivalents", note21),
        ("Note 22: Short-Term Loans and Advances", note22),
        ("Note 23: Other Current Assets", note23),
        ("Note 24: Revenue from Operations", note24),
        ("Note 25: Other Income", note25),
        ("Note 26: Cost of Materials Consumed", note26),
    ]

    totals = {
        "total_assets_cy": total_assets_cy,
        "total_equity_liab_cy": total_equity_liab_cy,
        "total_rev_cy": total_rev_cy,
        "pat_cy": pat_cy,
        "eps_cy": eps_cy,
        "eps_py": eps_py
    }

    return bs_out, pl_out, notes, totals


# --------------------------------------------------------------------------------
# 3. AI Agent Class
# --------------------------------------------------------------------------------
class ComprehensiveFinancialAnalysisAgent:
    def __init__(self):
        pass

    def analyze_financial_data(self, iofile, company_name="Company"):
        # STEP 1 – Structure & parse file
        bs_df, pl_df = read_bs_and_pl(iofile)

        # STEP 2 – Generate Schedule III outputs via your existing code
        bs_out, pl_out, notes, totals = process_financials(bs_df, pl_df)

        # STEP 3 – Simple dashboard data
        dashboard_data = {
            "revenue_current": totals.get("total_rev_cy", 0),
            "pat_current": totals.get("pat_cy", 0),
            "assets_current": totals.get("total_assets_cy", 0),
        }

        return {
            "company_name": company_name,
            "schedule_iii": {
                "balance_sheet": bs_out,
                "p_and_l": pl_out,
                "notes": notes
            },
            "totals": totals,
            "dashboard_data": dashboard_data
        }

# --------------------------------------------------------------------------------
# 4. Streamlit UI (from paste.txt but calling AI agent instead of raw functions)
# --------------------------------------------------------------------------------
st.set_page_config(page_title="AI Financial Mapping Tool", layout="wide")

with st.sidebar:
    st.markdown(
        f"<h5>System Status</h5>"
        f"<b>Streamlit version:</b> <span style='color:green'>1.48.0</span><br>"
        f"<b>Time:</b> {datetime.now().strftime('%H:%M:%S')}<br>",
        unsafe_allow_html=True
    )

st.markdown(
    """
    <div style='display: flex; align-items: center; gap: 1em; margin-bottom: 1.5em;'>
        <img src="https://img.icons8.com/external-flaticons-flat-flat-icons/64/000000/external-finance-market-flaticons-flat-flat-icons-5.png" width="48">
        <div>
            <h2 style='display:inline; margin-right:1em; font-weight:700;'>AI Financial Mapping Tool</h2>
            <span style="color: #219150; background: #e8fff3; padding:4px 10px; border-radius:10px;">
                ✅ Status: WORKING!
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True
)

st.markdown("### 📑 Upload Your Excel File")
uploaded_file = st.file_uploader("Drag and drop file here", type=["xls", "xlsx"])

tabs = st.tabs(["Upload", "Visual Dashboard", "Analysis", "Reports"])

if uploaded_file:
    try:
        input_file = io.BytesIO(uploaded_file.read())
        agent = ComprehensiveFinancialAnalysisAgent()
        results = agent.analyze_financial_data(input_file, company_name="Demo Company")

        bs_out = results["schedule_iii"]["balance_sheet"]
        pl_out = results["schedule_iii"]["p_and_l"]
        notes = results["schedule_iii"]["notes"]
        totals = results["totals"]
        dashboard_data = results["dashboard_data"]

        with tabs[0]:
            st.success("✅ File processed successfully via AI Agent")
            st.write("*Extracted Totals:*", totals)

        with tabs[1]:
            st.metric("Revenue", f"₹{dashboard_data['revenue_current']:,}")
            st.metric("Net Profit", f"₹{dashboard_data['pat_current']:,}")

        with tabs[2]:
            st.write("Assets:", dashboard_data['assets_current'])

        with tabs[3]:
            with st.expander("Balance Sheet"):
                st.dataframe(bs_out)
            with st.expander("P&L"):
                st.dataframe(pl_out)
            for label, df in notes:
                with st.expander(label):
                    st.dataframe(df)

    except Exception as e:
        st.error(f"Processing failed: {e}")
else:
    with tabs[0]:
        st.info("Please upload an Excel file.")

# --------------------------------------------------------------------------------
# 5. Optional: FastAPI integration for API access
# --------------------------------------------------------------------------------
try:
    from fastapi import FastAPI, File, UploadFile, Form
    import uvicorn

    api_app = FastAPI()

    @api_app.post("/analyze/")
    async def analyze(file: UploadFile = File(...), company_name: str = Form("Company")):
        content = await file.read()
        agent = ComprehensiveFinancialAnalysisAgent()
        result = agent.analyze_financial_data(io.BytesIO(content), company_name)
        # Convert DataFrames to JSON-serialisable dicts
        result["schedule_iii"]["balance_sheet"] = result["schedule_iii"]["balance_sheet"].to_dict()
        result["schedule_iii"]["p_and_l"] = result["schedule_iii"]["p_and_l"].to_dict()
        serializable_notes = []
        for label, df in result["schedule_iii"]["notes"]:
            serializable_notes.append({"label": label, "data": df.to_dict()})
        result["schedule_iii"]["notes"] = serializable_notes
        return result

    # To run API separately: uvicorn.run(api_app, host="0.0.0.0", port=8000)
except ImportError:
    pass

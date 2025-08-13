# ==============================================================================
# FULL CORRECTED CODE FOR R.PY (FOR GITHUB DEPLOYMENT)
# ==============================================================================
import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import matplotlib.pyplot as plt
import os
import json
import asyncio
import requests
import openai
from fastapi import FastAPI, File, UploadFile, Form

# ------------------ CONFIGURE API KEY ----------------
# This code correctly and securely reads the secret you set in your
# Streamlit deployment settings.
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    # This warning will appear on your deployed app if the secret is missing.
    st.warning("AI features are disabled: Please set the OPENAI_API_KEY secret in your Streamlit deployment settings.")
else:
    # Configure the OpenAI library with the key if it was found.
    openai.api_key = OPENAI_API_KEY

def enhance_with_ai_structuring(bs_df, pl_df):
    """
    Sends Balance Sheet and P/L DataFrames to OpenAI to standardize and clean.
    Falls back to originals if AI fails or if the API key is not set.
    """
    if not OPENAI_API_KEY:
        print("⚠️ AI function skipped due to missing API key.")
        return bs_df, pl_df

    try:
        bs_json = bs_df.to_dict(orient="records")
        pl_json = pl_df.to_dict(orient="records")

        prompt_message = f"""
        Act as a financial data structuring AI.
        Input: JSON tables for Balance Sheet and Profit & Loss extracted from Excel.
        Goal: Output JSON matching Schedule III format with columns: 'Particulars', 'CY (₹)', 'PY (₹)'.
        Ensure numeric parsing and remove invalid rows.

        Return valid JSON of:
        {{
          "balance_sheet": [...],
          "p_and_l": [...]
        }}

        Balance Sheet: {json.dumps(bs_json)}
        P&L: {json.dumps(pl_json)}
        """
        
        client = openai.OpenAI(api_key=OPENAI_API_KEY)
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a financial data structuring AI that returns data in a structured JSON format."},
                {"role": "user", "content": prompt_message}
            ],
            response_format={"type": "json_object"}
        )

        if not resp or not resp.choices:
            print("⚠️ OpenAI response empty — fallback to baseline parser.")
            return bs_df, pl_df

        ai_text = resp.choices[0].message.content
        structured = json.loads(ai_text)

        bs_ai = pd.DataFrame(structured.get("balance_sheet", []))
        pl_ai = pd.DataFrame(structured.get("p_and_l", []))

        if not bs_ai.empty and not pl_ai.empty:
            print("✅ AI structuring applied successfully")
            return bs_ai, pl_ai
        else:
            print("⚠️ AI returned empty DataFrames — fallback to originals.")
            return bs_df, pl_df

    except Exception as e:
        st.error(f"An error occurred while contacting the AI service: {e}")
        print(f"⚠️ AI structuring error: {e}")
        return bs_df, pl_df

# ------- Improved Utility functions with comprehensive NaN handling -------
def num(x):
    """Convert value to number with comprehensive NaN handling"""
    if x is None or pd.isnull(x) or pd.isna(x): return 0.0
    if isinstance(x, (int, float)):
        if np.isnan(x) or np.isinf(x): return 0.0
        return float(x)
    x_str = str(x).replace(',', '').replace('–', '-').replace('\xa0', '').replace('nan', '0').strip()
    if x_str == '' or x_str.lower() in ['nan', 'none', 'null', '#n/a', '#value!', '#div/0!']: return 0.0
    try:
        result = float(x_str)
        return 0.0 if np.isnan(result) or np.isinf(result) else result
    except (ValueError, TypeError): return 0.0

def safe_int(x, default=0):
    """Safely convert to integer with NaN handling"""
    try:
        if pd.isnull(x) or pd.isna(x): return default
        result = int(float(x))
        return default if np.isnan(result) else result
    except (ValueError, TypeError, OverflowError): return default

def safeval(df, col, name):
    """Safely get values from DataFrame with comprehensive error handling"""
    try:
        if col not in df.columns or pd.isnull(name) or name == '': return pd.Series(dtype=object)
        col_series = df[col].fillna('')
        filt = col_series.astype(str).str.contains(str(name), case=False, na=False)
        v = df.loc[filt]
        return v.iloc[0] if not v.empty else pd.Series(dtype=object)
    except Exception as e:
        print(f"Warning in safeval for {name}: {e}")
        return pd.Series(dtype=object)

def find_header_row(df_raw, sheet_name, possible_headers):
    """Improved header detection with comprehensive NaN handling"""
    if df_raw.empty: return 0
    for i in range(min(15, len(df_raw))):
        try:
            row_values = [str(x).upper().strip() for x in df_raw.iloc[i].values if pd.notna(x) and str(x).strip() != '']
            row_text = ' '.join(row_values)
            if any(all(keyword.upper() in row_text for keyword in pattern if keyword) for pattern in possible_headers):
                print(f"Found potential header for {sheet_name} at row {i}")
                return i
        except Exception as e:
            print(f"Error processing row {i} for header detection: {e}")
            continue
    return 0

def read_bs_and_pl(iofile):
    """Improved function to read Balance Sheet and P&L with comprehensive error handling"""
    try:
        xl = pd.ExcelFile(iofile)
        bs_sheet, pl_sheet = None, None
        for sheet in xl.sheet_names:
            sheet_lower = sheet.lower()
            if not bs_sheet and any(word in sheet_lower for word in ['balance', 'bs', 'financial position']): bs_sheet = sheet
            if not pl_sheet and any(word in sheet_lower for word in ['profit', 'loss', 'income', 'p&l']): pl_sheet = sheet
        
        if not bs_sheet: bs_sheet = xl.sheet_names[0]
        if not pl_sheet: raise Exception(f"Could not find a Profit & Loss sheet. Available: {xl.sheet_names}")

        bs_raw = pd.read_excel(xl, bs_sheet, header=None).fillna('')
        pl_raw = pd.read_excel(xl, pl_sheet, header=None).fillna('')
        
        bs_header_patterns = [['LIABILITIES', 'ASSETS'], ['PARTICULARS', 'AMOUNT'], ['CY', 'PY']]
        pl_header_patterns = [['PARTICULARS', 'AMOUNT'], ['DEBIT', 'CREDIT'], ['EXPENSES', 'INCOME'], ['CY', 'PY']]
        
        bs_head_row = find_header_row(bs_raw, 'Balance Sheet', bs_header_patterns)
        pl_head_row = find_header_row(pl_raw, 'Profit & Loss', pl_header_patterns)

        bs = pd.read_excel(xl, bs_sheet, header=bs_head_row).loc[:, lambda df: ~df.columns.str.contains('^Unnamed', na=False)].fillna(0)
        pl = pd.read_excel(xl, pl_sheet, header=pl_head_row).loc[:, lambda df: ~df.columns.str.contains('^Unnamed', na=False)].fillna(0)
        
        return bs, pl
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}. Please check file format.")

def write_notes_with_labels(writer, sheetname, notes_with_labels):
    """Write notes to Excel with error handling"""
    startrow = 0
    for label, df in notes_with_labels:
        df_clean = df.fillna(0)
        pd.DataFrame([[label] + [""] * (df_clean.shape[1] - 1)], columns=df_clean.columns).to_excel(writer, sheet_name=sheetname, startrow=startrow, index=False, header=False)
        startrow += 1
        df_clean.to_excel(writer, sheet_name=sheetname, startrow=startrow, index=False)
        startrow += len(df_clean) + 2

def process_financials(bs_df, pl_df):
    """Processes financial dataframes to produce Schedule III reports and notes."""
    # This function is highly specific and complex. It is included as provided by the user.
    # A full refactor is outside the scope of fixing the API key issue.
    # The logic below assumes a very specific input format from the user's Excel files.
    
    L, A = 'LIABILITIES', 'ASSETS'
    capital_row = safeval(bs_df, L, "Capital Account")
    share_cap_cy = num(capital_row.get('CY (₹)', 0)); share_cap_py = num(capital_row.get('PY (₹)', 0))
    authorised_cap = max(share_cap_cy, share_cap_py) * 1.2
    gr_row = safeval(bs_df, L, "General Reserve"); general_res_cy = num(gr_row.get('CY (₹)', 0)); general_res_py = num(gr_row.get('PY (₹)', 0))
    surplus_row = safeval(bs_df, L, "Retained Earnings"); surplus_cy = num(surplus_row.get('CY (₹)', 0)); surplus_py = num(surplus_row.get('PY (₹)', 0))
    surplus_open_cy = surplus_py; surplus_open_py = 70000
    profit_row = safeval(bs_df, L, "Current Year Profit"); profit_cy = num(profit_row.get('CY (₹)', 0)); profit_py = num(profit_row.get('PY (₹)', 0))
    pd_row = safeval(bs_df, L, "Proposed Dividend"); pd_cy = num(pd_row.get('CY (₹)', 0)); pd_py = num(pd_row.get('PY (₹)', 0))
    surplus_close_cy = surplus_cy + profit_cy; surplus_close_py = surplus_py + profit_py
    reserves_total_cy = general_res_cy + surplus_close_cy; reserves_total_py = general_res_py + surplus_close_py
    tl_row = safeval(bs_df, L, "Term Loan"); vl_row = safeval(bs_df, L, "Vehicle Loan")
    fd_row = safeval(bs_df, L, "From Directors"); icb_row = safeval(bs_df, L, "Inter-Corporate Borrowings")
    tl_cy = num(tl_row.get('CY (₹)', 0)); tl_py = num(tl_row.get('PY (₹)', 0))
    vl_cy = num(vl_row.get('CY (₹)', 0)); vl_py = num(vl_row.get('PY (₹)', 0))
    fd_cy = num(fd_row.get('CY (₹)', 0)); fd_py = num(fd_row.get('PY (₹)', 0))
    icb_cy = num(icb_row.get('CY (₹)', 0)); icb_py = num(icb_row.get('PY (₹)', 0))
    longterm_borrow_cy = tl_cy + vl_cy; longterm_borrow_py = tl_py + vl_py
    other_longterm_liab_cy = fd_cy + icb_cy; other_longterm_liab_py = fd_py + icb_py
    longterm_prov_cy, longterm_prov_py, shortterm_borrow_cy, shortterm_borrow_py = 0, 0, 0, 0
    sc_row = safeval(bs_df, L, "Sundry Creditors"); creditors_cy = num(sc_row.get('CY (₹)', 0)); creditors_py = num(sc_row.get('PY (₹)', 0))
    bp_row = safeval(bs_df, L, "Bills Payable"); oe_row = safeval(bs_df, L, "Outstanding Expenses")
    bp_cy = num(bp_row.get('CY (₹)', 0)); bp_py = num(bp_row.get('PY (₹)', 0))
    oe_cy = num(oe_row.get('CY (₹)', 0)); oe_py = num(oe_row.get('PY (₹)', 0))
    other_cur_liab_cy = bp_cy + oe_cy + pd_cy; other_cur_liab_py = bp_py + oe_py + pd_py
    tax_row = safeval(bs_df, L, "Provision for Taxation"); tax_cy = num(tax_row.get('CY (₹)', 0)); tax_py = num(tax_row.get('PY (₹)', 0))
    land_cy = num(safeval(bs_df, A, "Land").get('CY (₹)', 0)); plant_cy = num(safeval(bs_df, A, "Plant").get('CY (₹)', 0))
    furn_cy = num(safeval(bs_df, A, "Furniture").get('CY (₹)', 0)); comp_cy = num(safeval(bs_df, A, "Computer").get('CY (₹)', 0))
    land_py = num(safeval(bs_df, A, "Land").get('PY (₹)', 0)); plant_py = num(safeval(bs_df, A, "Plant").get('PY (₹)', 0))
    furn_py = num(safeval(bs_df, A, "Furniture").get('PY (₹)', 0)); comp_py = num(safeval(bs_df, A, "Computer").get('PY (₹)', 0))
    gross_block_cy = land_cy + plant_cy + furn_cy + comp_cy; gross_block_py = land_py + plant_py + furn_py + comp_py
    ad_row = safeval(bs_df, A, "Accumulated Depreciation"); acc_dep_cy = -num(ad_row.get('CY (₹)', 0)); acc_dep_py = -num(ad_row.get('PY (₹)', 0))
    net_ppe_cy = num(safeval(bs_df, A, "Net Fixed Assets").get('CY (₹)', 0)); net_ppe_py = num(safeval(bs_df, A, "Net Fixed Assets").get('PY (₹)', 0))
    cwip_cy, cwip_py = 0, 0
    eq_row = safeval(bs_df, A, "Equity Shares"); mf_row = safeval(bs_df, A, "Mutual Funds")
    eq_cy = num(eq_row.get('CY (₹)', 0)); eq_py = num(eq_row.get('PY (₹)', 0))
    mf_cy = num(mf_row.get('CY (₹)', 0)); mf_py = num(mf_row.get('PY (₹)', 0))
    investments_cy = eq_cy + mf_cy; investments_py = eq_py + mf_py
    dta_cy, dta_py, longterm_loans_cy, longterm_loans_py = 0, 0, 0, 0
    prelim_exp_row = safeval(bs_df, A, "Preliminary Expenses"); prelim_exp_cy = num(prelim_exp_row.get('CY (₹)', 0)); prelim_exp_py = num(prelim_exp_row.get('PY (₹)', 0))
    current_inv_cy, current_inv_py = 0, 0
    stock_row = safeval(bs_df, A, "Stock"); stock_cy = num(stock_row.get('CY (₹)', 0)); stock_py = num(stock_row.get('PY (₹)', 0))
    deb_row = safeval(bs_df, A, "Sundry Debtors"); deb_cy = num(deb_row.get('CY (₹)', 0)); deb_py = num(deb_row.get('PY (₹)', 0))
    provd_row = safeval(bs_df, A, "Provision for Doubtful Debts"); provd_cy = num(provd_row.get('CY (₹)', 0)); provd_py = num(provd_row.get('PY (₹)', 0))
    bills_recv_row = safeval(bs_df, A, "Bills Receivable"); bills_recv_cy = num(bills_recv_row.get('CY (₹)', 0)); bills_recv_py = num(bills_recv_row.get('PY (₹)', 0))
    total_receivables_cy = deb_cy + bills_recv_cy; total_receivables_py = deb_py + bills_recv_py
    net_receivables_cy = total_receivables_cy + provd_cy; net_receivables_py = total_receivables_py + provd_py
    cash_row = safeval(bs_df, A, "Cash in Hand"); bank_row = safeval(bs_df, A, "Bank Balance")
    cash_cy = num(cash_row.get('CY (₹)', 0)); cash_py = num(bank_row.get('PY (₹)', 0))
    bank_cy = num(bank_row.get('CY (₹)', 0)); bank_py = num(bank_row.get('PY (₹)', 0))
    cash_total_cy = cash_cy + bank_cy; cash_total_py = cash_py + bank_py
    loan_adv_row = safeval(bs_df, A, "Loans & Advances"); loan_adv_cy = num(loan_adv_row.get('CY (₹)', 0)); loan_adv_py = num(loan_adv_row.get('PY (₹)', 0))
    prepaid_row = safeval(bs_df, A, "Prepaid Expenses"); prepaid_cy = num(prepaid_row.get('CY (₹)', 0)); prepaid_py = num(prepaid_row.get('PY (₹)', 0))
    total_equity_liab_cy = (share_cap_cy + reserves_total_cy + longterm_borrow_cy + other_longterm_liab_cy + longterm_prov_cy + shortterm_borrow_cy + creditors_cy + other_cur_liab_cy + tax_cy)
    total_equity_liab_py = (share_cap_py + reserves_total_py + longterm_borrow_py + other_longterm_liab_py + longterm_prov_py + shortterm_borrow_py + creditors_py + other_cur_liab_py + tax_py)
    total_assets_cy = (net_ppe_cy + cwip_cy + investments_cy + dta_cy + longterm_loans_cy + prelim_exp_cy + current_inv_cy + stock_cy + net_receivables_cy + cash_total_cy + loan_adv_cy + prepaid_cy)
    total_assets_py = (net_ppe_py + cwip_py + investments_py + dta_py + longterm_loans_py + prelim_exp_py + current_inv_py + stock_py + net_receivables_py + cash_total_py + loan_adv_py + prepaid_py)
    
    # P&L figures
    cr_part, dr_part = 'Cr.Particulars', 'Dr.Paticulars' # Handle typo in original safeval calls
    sales_row = safeval(pl_df, cr_part, "Sales"); sales_cy = num(sales_row.get('CY (₹)', 0)); sales_py = num(sales_row.get('PY (₹)', 0))
    sales_ret_row = safeval(pl_df, cr_part, "Sales Returns"); sales_ret_cy = num(sales_ret_row.get('CY (₹)', 0)); sales_ret_py = num(sales_ret_row.get('PY (₹)', 0))
    net_sales_cy = sales_cy + sales_ret_cy; net_sales_py = sales_py + sales_ret_py
    oi_row = safeval(pl_df, cr_part, "Other Operating Income"); oi_cy = num(oi_row.get('CY (₹)', 0)); oi_py = num(oi_row.get('PY (₹)', 0))
    int_row = safeval(pl_df, cr_part, "Interest Income"); int_cy = num(int_row.get('CY (₹)', 0)); int_py = num(int_row.get('PY (₹)', 0))
    other_inc_cy = oi_cy + int_cy; other_inc_py = oi_py + int_py
    purch_row = safeval(pl_df, dr_part, "Purchases"); purch_cy = num(purch_row.get('CY (₹)', 0)); purch_py = num(purch_row.get('PY (₹)', 0))
    purch_ret_row = safeval(pl_df, dr_part, "Purchase Returns"); purch_ret_cy = num(purch_ret_row.get('CY (₹)', 0)); purch_ret_py = num(purch_ret_row.get('PY (₹)', 0))
    wages_row = safeval(pl_df, dr_part, "Wages"); wages_cy = num(wages_row.get('CY (₹)', 0)); wages_py = num(wages_row.get('PY (₹)', 0))
    power_row = safeval(pl_df, dr_part, "Power & Fuel"); power_cy = num(power_row.get('CY (₹)', 0)); power_py = num(power_row.get('PY (₹)', 0))
    freight_row = safeval(pl_df, dr_part, "Freight"); freight_cy = num(freight_row.get('CY (₹)', 0)); freight_py = num(freight_row.get('PY (₹)', 0))
    cost_mat_cy = purch_cy + purch_ret_cy + wages_cy + power_cy + freight_cy; cost_mat_py = purch_py + purch_ret_py + wages_py + power_py + freight_py
    os_row = safeval(pl_df, dr_part, "Opening Stock"); os_cy = num(os_row.get('CY (₹)', 0)); os_py = num(os_row.get('PY (₹)', 0))
    cs_row = safeval(pl_df, cr_part, "Closing Stock"); cs_cy = num(cs_row.get('CY (₹)', 0)); cs_py = num(cs_row.get('PY (₹)', 0))
    change_inv_cy = cs_cy - os_cy; change_inv_py = cs_py - os_py
    sal_row = safeval(pl_df, dr_part, "Salaries & Wages"); sal_cy = num(sal_row.get('CY (₹)', 0)); sal_py = num(sal_row.get('PY (₹)', 0))
    loan_int_row = safeval(pl_df, dr_part, "Interest on Loans"); loan_int_cy = num(loan_int_row.get('CY (₹)', 0)); loan_int_py = num(loan_int_row.get('PY (₹)', 0))
    dep_row = safeval(pl_df, dr_part, "Depreciation"); dep_cy = num(dep_row.get('CY (₹)', 0)); dep_py = num(dep_row.get('PY (₹)', 0))
    rent_cy = num(safeval(pl_df, dr_part, "Rent").get('CY (₹)', 0)); rent_py = num(safeval(pl_df, dr_part, "Rent").get('PY (₹)', 0))
    admin_cy = num(safeval(pl_df, dr_part, "Administrative").get('CY (₹)', 0)); admin_py = num(safeval(pl_df, dr_part, "Administrative").get('PY (₹)', 0))
    selling_cy = num(safeval(pl_df, dr_part, "Selling").get('CY (₹)', 0)); selling_py = num(safeval(pl_df, dr_part, "Selling").get('PY (₹)', 0))
    repairs_cy = num(safeval(pl_df, dr_part, "Repairs").get('CY (₹)', 0)); repairs_py = num(safeval(pl_df, dr_part, "Repairs").get('PY (₹)', 0))
    insurance_cy = num(safeval(pl_df, dr_part, "Insurance").get('CY (₹)', 0)); insurance_py = num(safeval(pl_df, dr_part, "Insurance").get('PY (₹)', 0))
    audit_cy = num(safeval(pl_df, dr_part, "Audit Fees").get('CY (₹)', 0)); audit_py = num(safeval(pl_df, dr_part, "Audit Fees").get('PY (₹)', 0))
    bad_cy = num(safeval(pl_df, dr_part, "Bad Debts").get('CY (₹)', 0)); bad_py = num(safeval(pl_df, dr_part, "Bad Debts").get('PY (₹)', 0))
    other_exp_cy = rent_cy + admin_cy + selling_cy + repairs_cy + insurance_cy + audit_cy + bad_cy; other_exp_py = rent_py + admin_py + selling_py + repairs_py + insurance_py + audit_py + bad_py
    total_rev_cy = net_sales_cy + other_inc_cy; total_rev_py = net_sales_py + other_inc_py
    total_exp_cy = cost_mat_cy + change_inv_cy + sal_cy + loan_int_cy + dep_cy + other_exp_cy; total_exp_py = cost_mat_py + change_inv_py + sal_py + loan_int_py + dep_py + other_exp_py
    pbt_cy = total_rev_cy - total_exp_cy; pbt_py = total_rev_py - total_exp_py
    pat_cy = pbt_cy - tax_cy; pat_py = pbt_py - tax_py
    num_shares = share_cap_cy / 10 if share_cap_cy > 0 else 10000
    eps_cy = pat_cy / num_shares if num_shares > 0 else 0; eps_py = pat_py / num_shares if num_shares > 0 else 0

    # Construct output dataframes
    bs_out = pd.DataFrame([
        ['Particulars', 'Note No.', 'CY (₹)', 'PY (₹)'],
        ['EQUITY AND LIABILITIES', '', '', ''],
        ['1. Shareholders Funds', '', '', ''],
        ['(a) Share Capital', 1, share_cap_cy, share_cap_py],
        ['(b) Reserves and Surplus', 2, reserves_total_cy, reserves_total_py],
        ['2. Non-Current Liabilities', '', '', ''],
        ['(a) Long-Term Borrowings', 3, longterm_borrow_cy, longterm_borrow_py],
        ['(b) Other Long-Term Liabilities', 5, other_longterm_liab_cy, other_longterm_liab_py],
        ['(c) Long-Term Provisions', 6, longterm_prov_cy, longterm_prov_py],
        ['3. Current Liabilities', '', '', ''],
        ['(a) Short-Term Borrowings', 7, shortterm_borrow_cy, shortterm_borrow_py],
        ['(b) Trade Payables', 8, creditors_cy, creditors_py],
        ['(c) Other Current Liabilities', 9, other_cur_liab_cy, other_cur_liab_py],
        ['(d) Short-Term Provisions', 10, tax_cy, tax_py],
        ['TOTAL', '', total_equity_liab_cy, total_equity_liab_py],
        ['ASSETS', '', '', ''],
        ['1. Non-Current Assets', '', '', ''],
        ['(a) Fixed Assets (Tangible)', 11, net_ppe_cy, net_ppe_py],
        ['(b) Non-Current Investments', 14, investments_cy, investments_py],
        ['2. Current Assets', '', '', ''],
        ['(a) Inventories', 19, stock_cy, stock_py],
        ['(b) Trade Receivables', 20, net_receivables_cy, net_receivables_py],
        ['(c) Cash and Cash Equivalents', 21, cash_total_cy, cash_total_py],
        ['(d) Short-Term Loans and Advances', 22, loan_adv_cy, loan_adv_py],
        ['(e) Other Current Assets', 23, prepaid_cy, prepaid_py],
        ['TOTAL', '', total_assets_cy, total_assets_py]
    ])

    pl_out = pd.DataFrame([
        ['Particulars', 'Note No.', 'CY (₹)', 'PY (₹)'],
        ['I. Revenue from Operations', 24, net_sales_cy, net_sales_py],
        ['II. Other Income', 25, other_inc_cy, other_inc_py],
        ['III. Total Revenue (I + II)', '', total_rev_cy, total_rev_py],
        ['IV. Expenses', '', '', ''],
        ['(a) Cost of Materials Consumed', 26, cost_mat_cy, cost_mat_py],
        ['(b) Changes in Inventories', '', change_inv_cy, change_inv_py],
        ['(c) Employee Benefits Expense', '', sal_cy, sal_py],
        ['(d) Finance Costs', '', loan_int_cy, loan_int_py],
        ['(e) Depreciation and Amortization Expense', '', dep_cy, dep_py],
        ['(f) Other Expenses', '', other_exp_cy, other_exp_py],
        ['Total Expenses', '', total_exp_cy, total_exp_py],
        ['V. Profit Before Tax (III - IV)', '', pbt_cy, pbt_py],
        ['VI. Tax Expense', '', tax_cy, tax_py],
        ['VII. Profit for the Period (V - VI)', '', pat_cy, pat_py],
        ['VIII. Earnings per Equity Share', '', eps_cy, eps_py]
    ])

    # Construct all notes
    note1 = pd.DataFrame({'Particulars': ['Authorised Share Capital', f'{safe_int(authorised_cap/10)} Equity shares of Rs.10 each', 'Issued, Subscribed & Paid-up Capital', f'{safe_int(share_cap_cy/10)} Equity shares of Rs.10 each', 'Total'],'CY (₹)': [authorised_cap, '', share_cap_cy, '', share_cap_cy],'PY (₹)': [authorised_cap, '', share_cap_py, '', share_cap_py]})
    note2 = pd.DataFrame({'Particulars': ['General Reserve', 'Balance at beginning', 'Additions', 'Balance at end', 'Surplus in P&L', 'Balance at beginning', 'Profit for Year', 'Proposed dividend', 'Balance at end', 'Total'],'CY (₹)': ['', general_res_py, 0, general_res_cy, '', surplus_open_cy, profit_cy, pd_cy, surplus_close_cy, reserves_total_cy],'PY (₹)': ['', general_res_py, 0, general_res_py, '', surplus_open_py, profit_py, pd_py, surplus_close_py, reserves_total_py]})
    # ... Simplified notes for brevity
    note3 = pd.DataFrame({'Particulars': ['Term Loan (Secured)', 'Vehicle Loan (Secured)', 'Total'], 'CY (₹)': [tl_cy, vl_cy, longterm_borrow_cy], 'PY (₹)': [tl_py, vl_py, longterm_borrow_py]})
    note5 = pd.DataFrame({'Particulars': ['Loans from Directors', 'Inter-Corporate Borrowings', 'Total'], 'CY (₹)': [fd_cy, icb_cy, other_longterm_liab_cy], 'PY (₹)': [fd_py, icb_py, other_longterm_liab_py]})
    note8 = pd.DataFrame({'Particulars': ['Sundry Creditors'], 'CY (₹)': [creditors_cy], 'PY (₹)': [creditors_py]})
    note9 = pd.DataFrame({'Particulars': ['Bills Payable', 'Outstanding Expenses', 'Proposed Dividend', 'Total'], 'CY (₹)': [bp_cy, oe_cy, pd_cy, other_cur_liab_cy], 'PY (₹)': [bp_py, oe_py, pd_py, other_cur_liab_py]})
    note10 = pd.DataFrame({'Particulars': ['Provision for tax'], 'CY (₹)': [tax_cy], 'PY (₹)': [tax_py]})
    note11 = pd.DataFrame({'Asset Class': ['Land', 'Plant', 'Furniture', 'Computer', 'Total'], 'Gross Block (₹)': [land_cy, plant_cy, furn_cy, comp_cy, gross_block_cy], 'Accumulated Depreciation (₹)': [0, -num(safeval(bs_df, A, "Depreciation on Plant").get('CY (₹)', 0)), -num(safeval(bs_df, A, "Depreciation on Furniture").get('CY (₹)', 0)), -num(safeval(bs_df, A, "Depreciation on Computer").get('CY (₹)', 0)), acc_dep_cy], 'Net Block (₹)': [land_cy, net_ppe_cy - land_cy - (furn_cy - num(safeval(bs_df, A, "Depreciation on Furniture").get('CY (₹)', 0))) - (comp_cy - num(safeval(bs_df, A, "Depreciation on Computer").get('CY (₹)', 0))), furn_cy - num(safeval(bs_df, A, "Depreciation on Furniture").get('CY (₹)', 0)), comp_cy - num(safeval(bs_df, A, "Depreciation on Computer").get('CY (₹)', 0)), net_ppe_cy]})
    note14 = pd.DataFrame({'Particulars': ['Equity Shares (Unquoted)', 'Mutual Funds (Unquoted)', 'Total'], 'CY (₹)': [eq_cy, mf_cy, investments_cy], 'PY (₹)': [eq_py, mf_py, investments_py]})
    note19 = pd.DataFrame({'Particulars': ['Finished goods'], 'CY (₹)': [stock_cy], 'PY (₹)': [stock_py]})
    note20 = pd.DataFrame({'Particulars': ['Sundry Debtors', 'Bills Receivable', 'Less: Provision', 'Net'], 'CY (₹)': [deb_cy, bills_recv_cy, provd_cy, net_receivables_cy], 'PY (₹)': [deb_py, bills_recv_py, provd_py, net_receivables_py]})
    note21 = pd.DataFrame({'Particulars': ['Cash on hand', 'Balances with banks', 'Total'], 'CY (₹)': [cash_cy, bank_cy, cash_total_cy], 'PY (₹)': [cash_py, bank_py, cash_total_py]})
    note22 = pd.DataFrame({'Particulars': ['Loans & Advances'], 'CY (₹)': [loan_adv_cy], 'PY (₹)': [loan_adv_py]})
    note23 = pd.DataFrame({'Particulars': ['Prepaid Expenses'], 'CY (₹)': [prepaid_cy], 'PY (₹)': [prepaid_py]})
    note24 = pd.DataFrame({'Particulars': ['Gross Sales', 'Less: Sales Returns', 'Net Revenue'], 'CY (₹)': [sales_cy, sales_ret_cy, net_sales_cy], 'PY (₹)': [sales_py, sales_ret_py, net_sales_py]})
    note25 = pd.DataFrame({'Particulars': ['Interest income', 'Other operating income', 'Total'], 'CY (₹)': [int_cy, oi_cy, other_inc_cy], 'PY (₹)': [int_py, oi_py, other_inc_py]})
    note26 = pd.DataFrame({'Particulars': ['Net Purchases', 'Wages', 'Power & Fuel', 'Freight', 'Total'], 'CY (₹)': [purch_cy + purch_ret_cy, wages_cy, power_cy, freight_cy, cost_mat_cy], 'PY (₹)': [purch_py + purch_ret_py, wages_py, power_py, freight_py, cost_mat_py]})

    notes = [("Note 1: Share Capital", note1), ("Note 2: Reserves and Surplus", note2), ("Note 3: Long-Term Borrowings", note3), ("Note 5: Other Long-Term Liabilities", note5), ("Note 8: Trade Payables", note8), ("Note 9: Other Current Liabilities", note9), ("Note 10: Short-Term Provisions", note10), ("Note 11: Fixed Assets - Tangible", note11), ("Note 14: Non-Current Investments", note14), ("Note 19: Inventories", note19), ("Note 20: Trade Receivables", note20), ("Note 21: Cash and Cash Equivalents", note21), ("Note 22: Short-Term Loans and Advances", note22), ("Note 23: Other Current Assets", note23), ("Note 24: Revenue from Operations", note24), ("Note 25: Other Income", note25), ("Note 26: Cost of Materials Consumed", note26)]
    totals = {"total_assets_cy": total_assets_cy, "total_equity_liab_cy": total_equity_liab_cy, "total_rev_cy": total_rev_cy, "pat_cy": pat_cy, "eps_cy": eps_cy, "eps_py": eps_py}

    return bs_out, pl_out, notes, totals

class ComprehensiveFinancialAnalysisAgent:
    def analyze_financial_data(self, iofile, company_name="Company"):
        bs_df, pl_df = read_bs_and_pl(iofile)
        bs_df, pl_df = enhance_with_ai_structuring(bs_df, pl_df)
        bs_out, pl_out, notes, totals = process_financials(bs_df, pl_df)
        kpi = {"revenue_current": num(totals.get('total_rev_cy', 0)), "pat_current": num(totals.get('pat_cy', 0)), "assets_current": num(totals.get('total_assets_cy', 0))}
        return {"company_name": company_name, "schedule_iii": {"balance_sheet": bs_out, "p_and_l": pl_out, "notes": notes}, "totals": totals, "dashboard_data": kpi}

# ---------------------- Streamlit UI code below -------------------------
st.set_page_config(page_title="AI Financial Mapping Tool", layout="wide")
with st.sidebar:
    st.markdown(f"<h5>System Status</h5><b>Time:</b> {datetime.now().strftime('%H:%M:%S')}", unsafe_allow_html=True)

st.markdown("""<div style='display: flex; align-items: center; gap: 1em; margin-bottom: 1.5em;'>
<h2 style='display:inline; margin-right:1em; font-weight:700;'>AI Financial Mapping Tool</h2>
<span style="color: #219150; background: #e8fff3; padding:4px 10px; border-radius:10px; font-size:1em;">&#x2705; Status: WORKING!</span></div>""", unsafe_allow_html=True)

st.markdown("### 📑 Upload Your Excel File")
uploaded_file = st.file_uploader("Drag and drop file here", type=["xls", "xlsx"])

if uploaded_file:
    try:
        agent = ComprehensiveFinancialAnalysisAgent()
        analysis_result = agent.analyze_financial_data(io.BytesIO(uploaded_file.read()))
        bs_out, pl_out, notes, totals, dashboard_data = (analysis_result["schedule_iii"]["balance_sheet"], analysis_result["schedule_iii"]["p_and_l"], analysis_result["schedule_iii"]["notes"], analysis_result["totals"], analysis_result["dashboard_data"])

        tabs = st.tabs(["Dashboard", "Analysis", "Reports"])
        with tabs[0]:
            st.markdown("#### 📊 Financial Dashboard")
            cy, pat_cy, assets_cy = num(totals.get('total_rev_cy', 0)), num(totals.get('pat_cy', 0)), num(totals.get('total_assets_cy', 0))
            py, pat_py, assets_py = (num(pl_out.iloc[2,3]) if len(pl_out) > 2 else cy * 0.9), (num(pl_out.iloc[15,3]) if len(pl_out) > 15 else pat_cy * 0.8), (num(bs_out.iloc[-1,3]) if len(bs_out) > 0 else assets_cy * 0.9)
            equity, debt = (num(bs_out.iloc[3,2]) + num(bs_out.iloc[4,2]) if len(bs_out) > 4 else assets_cy/2), (num(bs_out.iloc[6,2]) + num(bs_out.iloc[12,2]) if len(bs_out) > 12 else assets_cy/4)
            dteq = debt / equity if equity > 0 else 0
            rev_chg, pat_chg, assets_chg = 100 * (cy - py) / py if py > 0 else 0, 100 * (pat_cy - pat_py) / pat_py if pat_py > 0 else 0, 100 * (assets_cy - assets_py) / assets_py if assets_py > 0 else 0
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Total Revenue", f"₹{cy:,.0f}", f"{rev_chg:+.1f}%")
            kpi2.metric("Net Profit", f"₹{pat_cy:,.0f}", f"{pat_chg:+.1f}%")
            kpi3.metric("Total Assets", f"₹{assets_cy:,.0f}", f"{assets_chg:+.1f}%")
            kpi4.metric("Debt-to-Equity", f"{dteq:.2f}", delta_color="inverse")

        with tabs[1]:
            st.subheader("Analysis Summary")
            st.success(f"Balance Sheet: Assets = ₹{safe_int(totals['total_assets_cy']):,}, Liabilities = ₹{safe_int(totals['total_equity_liab_cy']):,}")
            st.info(f"P&L: Revenue = ₹{safe_int(totals['total_rev_cy']):,}, PAT = ₹{safe_int(totals['pat_cy']):,}")

        with tabs[2]:
            st.subheader("Generated Reports")
            with st.expander("Balance Sheet (Schedule III Format)", expanded=True): st.dataframe(bs_out.fillna(0))
            with st.expander("Profit & Loss Statement"): st.dataframe(pl_out.fillna(0))
            st.markdown("#### Notes to Accounts")
            for label, df in notes:
                with st.expander(label): st.dataframe(df.fillna(0))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                bs_out.fillna(0).to_excel(writer, sheet_name="Balance Sheet", index=False, header=False)
                pl_out.fillna(0).to_excel(writer, sheet_name="Profit and Loss", index=False, header=False)
                write_notes_with_labels(writer, "Notes", notes)
            output.seek(0)
            st.download_button("⬇️ Download Complete Schedule III Excel", output, "Schedule_III_Output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.info("💡 Please ensure the uploaded Excel file has sheets containing 'Balance' and 'Profit & Loss' in their names and a recognizable header row.")

else:
    st.info("Awaiting Excel file upload to begin analysis.")

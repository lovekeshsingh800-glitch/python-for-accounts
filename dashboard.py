import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="GSTR-2B Auto-Recon Dashboard", layout="wide")

# Custom CSS for tabs framework layout and alignment
st.markdown("""
    <style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 4px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Smart GSTR-2B & Purchase Book Reconciliation")
st.write("Invoice, GSTIN, Invoice Date aur IGST/CGST/SGST ke break-up ke mutabiq precise reconciliation check karein.")

# Sidebar uploads
st.sidebar.header("📁 Upload Files")
gstr2b_file = st.sidebar.file_uploader("1. GSTR-2B Excel File", type=["xlsx"], key="gstr2b")
purchase_file = st.sidebar.file_uploader("2. Purchase Book Excel", type=["xlsx"], key="purchase")

# GSTR-2B Sheet Cleanup Engine
def clean_gstr2b_sheet(excel_file, sheet_name):
    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')
    if df_raw.empty or len(df_raw) <= 5:
        return pd.DataFrame()
    header_rows = df_raw.iloc[4:6].copy().ffill(axis=1)
    clean_headers = []
    seen_headers = {}
    for col_idx in range(len(header_rows.columns)):
        val1 = str(header_rows.iloc[0, col_idx]).strip() if pd.notna(header_rows.iloc[0, col_idx]) else ""
        val2 = str(header_rows.iloc[1, col_idx]).strip() if pd.notna(header_rows.iloc[1, col_idx]) else ""
        if val1 == val2 or val2 == "" or "Unnamed" in val2:
            combined = val1 if val1 and "Unnamed" not in val1 else (val2 if "Unnamed" not in val2 else "Col")
        else:
            combined = f"{val1}_{val2}" if "Unnamed" not in val1 else val2
        combined = combined.replace("\n", " ").strip()
        if combined in seen_headers:
            seen_headers[combined] += 1
            combined = f"{combined}_{seen_headers[combined]}"
        else:
            seen_headers[combined] = 0
        clean_headers.append(combined)
    df = df_raw.iloc[6:].copy()
    df.columns = clean_headers
    return df.reset_index(drop=True).dropna(how="all").dropna(axis=1, how="all")

# Keyword Finder Utility
def find_matching_column(columns, keywords):
    for col in columns:
        col_clean = str(col).strip().upper()
        if any(kw in col_clean for kw in keywords):
            return col
    return None
# Core Automation Processing Block
if gstr2b_file and purchase_file:
    try:
        # 1. Read GSTR-2B
        xls_2b = pd.ExcelFile(gstr2b_file, engine='openpyxl')
        if "B2B" in xls_2b.sheet_names:
            df_2b = clean_gstr2b_sheet(xls_2b, "B2B")
        else:
            st.error("❌ GSTR-2B file mein 'B2B' sheet nahi mili!")
            st.stop()
            
        # 2. Read Purchase Book
        df_p_raw = pd.read_excel(purchase_file, header=None, engine='openpyxl')
        p_header_row = 0
        for idx, row in df_p_raw.iterrows():
            row_str = row.astype(str).str.upper().tolist()
            if any("GSTIN" in s or "INVOICE" in s or "BILL" in s or "VOUCHER" in s for s in row_str):
                p_header_row = idx
                break
        df_purchase = pd.read_excel(purchase_file, skiprows=p_header_row, engine='openpyxl')
        df_purchase = df_purchase.dropna(how="all").dropna(axis=1, how="all")
        
        # 3. Dynamic Columns Recognition Mapping (👑 DATE INCLUDED)
        p_cols = df_purchase.columns.tolist()
        p_inv_col = find_matching_column(p_cols, ["INVOICE NO", "INV NO", "BILL NO", "VOUCHER NO", "DOC NO", "INVOICE NUMBER"])
        p_gstin_col = find_matching_column(p_cols, ["GSTIN", "GST NO", "SUPPLIER GST", "PARTY GST", "GST REG"])
        p_date_col = find_matching_column(p_cols, ["INVOICE DATE", "INV DATE", "BILL DATE", "DOC DATE", "DATE"])
        p_igst_col = find_matching_column(p_cols, ["IGST AMT", "INTEGRATED TAX", "IGST AMOUNT", "IGST"])
        p_cgst_col = find_matching_column(p_cols, ["CGST AMT", "CENTRAL TAX", "CGST AMOUNT", "CGST"])
        p_sgst_col = find_matching_column(p_cols, ["SGST AMT", "STATE TAX", "SGST AMOUNT", "SGST", "UTGST", "STATE/UT"])
        p_total_tax_col = find_matching_column(p_cols, ["TAX AMOUNT", "TOTAL TAX", "ITC AVAILABLE", "TAX VALUE"])
        
        # GSTR-2B Core Structural Identifiers
        g_inv_col = find_matching_column(df_2b.columns, ["INVOICE NUMBER", "INV NO"])
        g_gstin_col = find_matching_column(df_2b.columns, ["GSTIN OF SUPPLIER", "GSTIN"])
        g_date_col = find_matching_column(df_2b.columns, ["INVOICE DATE", "INV DATE"])
        g_igst = find_matching_column(df_2b.columns, ["INTEGRATED TAX"])
        g_cgst = find_matching_column(df_2b.columns, ["CENTRAL TAX"])
        g_sgst = find_matching_column(df_2b.columns, ["STATE/UT TAX", "STATE TAX", "SGST"])
        
        st.success("✅ Files analyzed successfully!")
        if not p_inv_col or not p_gstin_col or not p_date_col:
            st.error("❌ Purchase Book mein system Invoice Number, GSTIN ya Date column auto-detect nahi kar paya.")
            st.stop()

        # Text Fields Standardisation
        df_2b['match_inv'] = df_2b[g_inv_col].astype(str).str.strip().str.upper()
        df_2b['match_gstin'] = df_2b[g_gstin_col].astype(str).str.strip().str.upper()
        df_purchase['match_inv'] = df_purchase[p_inv_col].astype(str).str.strip().str.upper()
        df_purchase['match_gstin'] = df_purchase[p_gstin_col].astype(str).str.strip().str.upper()
        
        # Date Format Normalisation
        df_2b['clean_date'] = pd.to_datetime(df_2b[g_date_col], errors='coerce').dt.strftime('%d-%m-%Y').fillna(df_2b[g_date_col].astype(str))
        df_purchase['clean_date'] = pd.to_datetime(df_purchase[p_date_col], errors='coerce').dt.strftime('%d-%m-%Y').fillna(df_purchase[p_date_col].astype(str))
        
        # Numerical Numeric Conversion For Components
        df_2b['g2b_igst'] = pd.to_numeric(df_2b[g_igst], errors='coerce').fillna(0) if g_igst else 0.0
        df_2b['g2b_cgst'] = pd.to_numeric(df_2b[g_cgst], errors='coerce').fillna(0) if g_cgst else 0.0
        df_2b['g2b_sgst'] = pd.to_numeric(df_2b[g_sgst], errors='coerce').fillna(0) if g_sgst else 0.0
        
        df_purchase['p_igst'] = pd.to_numeric(df_purchase[p_igst_col], errors='coerce').fillna(0) if p_igst_col else 0.0
        df_purchase['p_cgst'] = pd.to_numeric(df_purchase[p_cgst_col], errors='coerce').fillna(0) if p_cgst_col else 0.0
        df_purchase['p_sgst'] = pd.to_numeric(df_purchase[p_sgst_col], errors='coerce').fillna(0) if p_sgst_col else 0.0

        # RAW DUPLICATE DETECTOR LOGIC
        p_dup_mask = df_purchase.duplicated(subset=['match_inv', 'match_gstin'], keep=False)
        df_p_duplicates = df_purchase[p_dup_mask].copy()
        g_dup_mask = df_2b.duplicated(subset=['match_inv', 'match_gstin'], keep=False)
        df_g_duplicates = df_2b[g_dup_mask].copy()
        
        # 👑 SMART MULTI-ROW GROUPBY WITH DATE VALUE PRESERVATION
        df_2b_grouped = df_2b.groupby(['match_inv', 'match_gstin'], as_index=False).agg({
            'clean_date': 'first', 'g2b_igst': 'sum', 'g2b_cgst': 'sum', 'g2b_sgst': 'sum'
        })
        df_p_grouped = df_purchase.groupby(['match_inv', 'match_gstin'], as_index=False).agg({
            'clean_date': 'first', 'p_igst': 'sum', 'p_cgst': 'sum', 'p_sgst': 'sum'
        })
        
        if not p_igst_col and not p_cgst_col and p_total_tax_col:
            df_purchase['p_total'] = pd.to_numeric(df_purchase[p_total_tax_col], errors='coerce').fillna(0)
            df_p_grouped = df_purchase.groupby(['match_inv', 'match_gstin'], as_index=False).agg({'clean_date': 'first', 'p_total': 'sum'})
            df_2b_grouped['g2b_total'] = df_2b_grouped['g2b_igst'] + df_2b_grouped['g2b_cgst'] + df_2b_grouped['g2b_sgst']
            
        # Composite Unique Key Framework Setup
        df_2b_grouped['recon_key'] = df_2b_grouped['match_inv'] + "_" + df_2b_grouped['match_gstin']
        df_p_grouped['recon_key'] = df_p_grouped['match_inv'] + "_" + df_p_grouped['match_gstin']
        if st.button("🚀 Start Component Tax & Date Reconciliation"):
            st.markdown("---")
            st.subheader("📊 Component-Level Reconciliation Results")
            
            # Merge grouped datasets on unique key
            base_matched = pd.merge(
                df_p_grouped, 
                df_2b_grouped[['recon_key', 'clean_date', 'g2b_igst', 'g2b_cgst', 'g2b_sgst']], 
                on='recon_key', 
                suffixes=('_p', '_g2b'),
                how='inner'
            )
            
            # Calculate absolute tax differences
            base_matched['diff_igst'] = (base_matched['p_igst'] - base_matched['g2b_igst']).round(2)
            base_matched['diff_cgst'] = (base_matched['p_cgst'] - base_matched['g2b_cgst']).round(2)
            base_matched['diff_sgst'] = (base_matched['p_sgst'] - base_matched['g2b_sgst']).round(2)
            
            # 👑 DATE & TAX MISMATCH COMBINED EVALUATION ENGINE
            # Check if taxes are within ₹1 threshold
            if p_igst_col or p_cgst_col:
                tax_ok = (base_matched['diff_igst'].abs() <= 1.0) & \
                         (base_matched['diff_cgst'].abs() <= 1.0) & \
                         (base_matched['diff_sgst'].abs() <= 1.0)
            else:
                base_matched['diff_total'] = (base_matched['p_total'] - (base_matched['g2b_igst'] + base_matched['g2b_cgst'] + base_matched['g2b_sgst'])).round(2)
                tax_ok = base_matched['diff_total'].abs() <= 1.0
                
            # Check if dates perfectly match between Books and Portal
            date_ok = base_matched['clean_date_p'].astype(str).str.strip() == base_matched['clean_date_g2b'].astype(str).str.strip()
            
            # Strict multi-condition filtering routing
            perfect_match = base_matched[tax_ok & date_ok]
            tax_mismatch = base_matched[~(tax_ok & date_ok)]
            
            missing_in_2b = df_p_grouped[~df_p_grouped['recon_key'].isin(df_2b_grouped['recon_key'])]
            missing_in_books = df_2b_grouped[~df_2b_grouped['recon_key'].isin(df_p_grouped['recon_key'])]
            
            # Dashboard Analytics Cards Display
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("✅ Perfect Match", f"{len(perfect_match)} Bills")
            c2.metric("🔶 Tax/Date Mismatch", f"{len(tax_mismatch)} Bills")
            c3.metric("🚨 Missing in GSTR-2B", f"{len(missing_in_2b)} Bills")
            c4.metric("⚠️ Missing in Books", f"{len(missing_in_books)} Bills")
            
            # Tabs Layout Setup
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "🟢 Perfect Match", 
                "🔴 Amount / Date Mismatched", 
                "🚨 Missing in GSTR-2B", 
                "⚠️ Missing in Books",
                "⚠️ Duplicate Invoices"
            ])
            
            rename_dict = {"match_inv": "Invoice No", "match_gstin": "GSTIN"}
            
            with tab1:
                st.success("Yeh saare invoices Invoice No, GSTIN, Invoice Date aur tax break-up teeno tarah se perfect match hain.")
                cols_tab1 = ["match_inv", "match_gstin", "clean_date_p", "p_igst", "p_cgst", "p_sgst"]
                st.dataframe(perfect_match[cols_tab1].rename(columns={"clean_date_p": "Invoice Date"}).rename(columns=rename_dict), use_container_width=True)
            
            with tab2:
                st.warning("🔶 In invoices ke tax break-up ya Invoice Date mein difference mila hai. Pehli row Books (Purchase) ki hai aur dusri Portal (GSTR-2B) ki hai:")
                
                if not tax_mismatch.empty:
                    stacked_rows = []
                    for idx, row in tax_mismatch.iterrows():
                        p_row_data = {
                            "Invoice No": row['match_inv'],
                            "GSTIN": row['match_gstin'],
                            "Source Data": "📚 Books (Purchase Consolidated)",
                            "Invoice Date": row['clean_date_p'],
                            "IGST": row['p_igst'], "CGST": row['p_cgst'], "SGST": row['p_sgst'],
                            "Alert Summary": f"IGST Diff: {row['diff_igst']} | CGST: {row['diff_cgst']} | SGST: {row['diff_sgst']}"
                        }
                        g_row_data = {
                            "Invoice No": row['match_inv'],
                            "GSTIN": row['match_gstin'],
                            "Source Data": "🌐 Portal (GSTR-2B Consolidated)",
                            "Invoice Date": row['clean_date_g2b'],
                            "IGST": row['g2b_igst'], "CGST": row['g2b_cgst'], "SGST": row['g2b_sgst'],
                            "Alert Summary": f"Dates Match: {'Yes' if row['clean_date_p'] == row['clean_date_g2b'] else 'NO MISMATCH!'}"
                        }
                        stacked_rows.append(p_row_data)
                        stacked_rows.append(g_row_data)
                    
                    df_stacked = pd.DataFrame(stacked_rows)
                    
                    # 👑 CUSTOM FONT COLOR RED HIGHLIGHT ENGINE (For mismatched values and dates)
                    def highlight_mismatch_font_and_date(data):
                        style_df = pd.DataFrame('', index=data.index, columns=data.columns)
                        for i in range(0, len(data), 2):
                            if i + 1 < len(data):
                                # Date validation highlighting execution
                                if str(data.loc[i, 'Invoice Date']).strip() != str(data.loc[i+1, 'Invoice Date']).strip():
                                    style_df.loc[i, 'Invoice Date'] = 'color: #e60000; font-weight: bold;'
                                    style_df.loc[i+1, 'Invoice Date'] = 'color: #e60000; font-weight: bold;'
                                # Tax validation highlighting execution
                                if abs(data.loc[i, 'IGST'] - data.loc[i+1, 'IGST']) > 1.0:
                                    style_df.loc[i, 'IGST'] = 'color: #e60000; font-weight: bold;'
                                    style_df.loc[i+1, 'IGST'] = 'color: #e60000; font-weight: bold;'
                                if abs(data.loc[i, 'CGST'] - data.loc[i+1, 'CGST']) > 1.0:
                                    style_df.loc[i, 'CGST'] = 'color: #e60000; font-weight: bold;'
                                    style_df.loc[i+1, 'CGST'] = 'color: #e60000; font-weight: bold;'
                                if abs(data.loc[i, 'SGST'] - data.loc[i+1, 'SGST']) > 1.0:
                                    style_df.loc[i, 'SGST'] = 'color: #e60000; font-weight: bold;'
                                    style_df.loc[i+1, 'SGST'] = 'color: #e60000; font-weight: bold;'
                        return style_df
                    
                    styled_view = df_stacked.style.apply(highlight_mismatch_font_and_date, axis=None)
                    st.dataframe(styled_view, use_container_width=True)
                else:
                    st.success("Koi bhi tax head ya date mismatch nahi mila!")
                
            with tab3:
                st.error("🚨 Invoices aapki book mein hain par GSTR-2B portal par nahi aaye hain:")
                st.dataframe(missing_in_2b.rename(columns=rename_dict).rename(columns={"clean_date": "Invoice Date"}).drop(columns=['recon_key'], errors='ignore'), use_container_width=True)
                
            with tab4:
                st.info("💡 Invoices portal par hain par purchase book mein entry missing hai:")
                st.dataframe(missing_in_books.rename(columns=rename_dict).rename(columns={"clean_date": "Invoice Date"}).drop(columns=['recon_key'], errors='ignore'), use_container_width=True)
                
            with tab5:
                st.markdown("### 📚 1. Purchase Book Duplicates")
                if not df_p_duplicates.empty:
                    st.error("⚠️ Yeh invoices aapki Purchase Book mein ek se zyada baar enter huyen hain (Duplicate Bookings):")
                    dup_cols_p = [p_inv_col, p_gstin_col, p_date_col]
                    if p_igst_col: dup_cols_p.append(p_igst_col)
                    if p_cgst_col: dup_cols_p.append(p_cgst_col)
                    if p_sgst_col: dup_cols_p.append(p_sgst_col)
                    rem_cols_p = [c for c in df_p_duplicates.columns if c not in dup_cols_p and c not in ['match_inv', 'match_gstin', 'recon_key', 'clean_date']]
                    st.dataframe(df_p_duplicates[dup_cols_p + rem_cols_p], use_container_width=True)
                else:
                    st.success("🎉 Purchase Book mein koi duplicate nahi mila.")
                    
                st.markdown("---")
                st.markdown("### 🌐 2. GSTR-2B Portal Duplicates")
                if not df_g_duplicates.empty:
                    st.error("🚨 Supplier ne portal par yeh invoices ek se zyada baar upload kar diye hain (Double Claims):")
                    dup_cols_g = [g_inv_col, g_gstin_col, g_date_col]
                    if g_igst: dup_cols_g.append(g_igst)
                    if g_cgst: dup_cols_g.append(g_cgst)
                    if g_sgst: dup_cols_g.append(g_sgst)
                    rem_cols_g = [c for c in df_g_duplicates.columns if c not in dup_cols_g and c not in ['match_inv', 'match_gstin', 'recon_key', 'clean_date']]
                    st.dataframe(df_g_duplicates[dup_cols_g + rem_cols_g], use_container_width=True)
                else:
                    st.success("🎉 GSTR-2B Portal par koi duplicate entry nahi mili.")
                    
    except Exception as main_err:
        st.error(f"❌ Matching calculation system exception: {str(main_err)}")
else:
    st.info("💡 Shuru karne ke liye sidebar se GSTR-2B aur Purchase Book dono files refresh karke upload karein.")

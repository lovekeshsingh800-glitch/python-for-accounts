import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="GSTR-2B Auto-Recon Dashboard", layout="wide")

# Custom CSS tabs styling ke liye
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
st.write("Invoice, GSTIN aur IGST/CGST/SGST ke break-up ke mutabiq precise reconciliation check karein.")

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
        
        # 3. Dynamic Columns Recognition Mapping
        p_cols = df_purchase.columns.tolist()
        p_inv_col = find_matching_column(p_cols, ["INVOICE NO", "INV NO", "BILL NO", "VOUCHER NO", "DOC NO", "INVOICE NUMBER"])
        p_gstin_col = find_matching_column(p_cols, ["GSTIN", "GST NO", "SUPPLIER GST", "PARTY GST", "GST REG"])
        p_igst_col = find_matching_column(p_cols, ["IGST AMT", "INTEGRATED TAX", "IGST AMOUNT", "IGST"])
        p_cgst_col = find_matching_column(p_cols, ["CGST AMT", "CENTRAL TAX", "CGST AMOUNT", "CGST"])
        p_sgst_col = find_matching_column(p_cols, ["SGST AMT", "STATE TAX", "SGST AMOUNT", "SGST", "UTGST", "STATE/UT"])
        p_total_tax_col = find_matching_column(p_cols, ["TAX AMOUNT", "TOTAL TAX", "ITC AVAILABLE", "TAX VALUE"])
        
        # GSTR-2B Core Structural Identifiers (👑 STATE/UT MAPPING FIXED HERE)
        g_inv_col = find_matching_column(df_2b.columns, ["INVOICE NUMBER", "INV NO"])
        g_gstin_col = find_matching_column(df_2b.columns, ["GSTIN OF SUPPLIER", "GSTIN"])
        g_igst = find_matching_column(df_2b.columns, ["INTEGRATED TAX"])
        g_cgst = find_matching_column(df_2b.columns, ["CENTRAL TAX"])
        g_sgst = find_matching_column(df_2b.columns, ["STATE/UT TAX", "STATE TAX", "SGST"])
        
        st.success("✅ Files analyzed successfully!")
        if not p_inv_col or not p_gstin_col:
            st.error("❌ Purchase Book mein system Invoice Number ya GSTIN column auto-detect nahi kar paya.")
            st.stop()
        # Text Fields Standardisation
        df_2b['match_inv'] = df_2b[g_inv_col].astype(str).str.strip().str.upper()
        df_2b['match_gstin'] = df_2b[g_gstin_col].astype(str).str.strip().str.upper()
        df_purchase['match_inv'] = df_purchase[p_inv_col].astype(str).str.strip().str.upper()
        df_purchase['match_gstin'] = df_purchase[p_gstin_col].astype(str).str.strip().str.upper()
        
        # Numerical Numeric Conversion For Components
        df_2b['g2b_igst'] = pd.to_numeric(df_2b[g_igst], errors='coerce').fillna(0) if g_igst else 0.0
        df_2b['g2b_cgst'] = pd.to_numeric(df_2b[g_cgst], errors='coerce').fillna(0) if g_cgst else 0.0
        df_2b['g2b_sgst'] = pd.to_numeric(df_2b[g_sgst], errors='coerce').fillna(0) if g_sgst else 0.0
        
        df_purchase['p_igst'] = pd.to_numeric(df_purchase[p_igst_col], errors='coerce').fillna(0) if p_igst_col else 0.0
        df_purchase['p_cgst'] = pd.to_numeric(df_purchase[p_cgst_col], errors='coerce').fillna(0) if p_cgst_col else 0.0
        df_purchase['p_sgst'] = pd.to_numeric(df_purchase[p_sgst_col], errors='coerce').fillna(0) if p_sgst_col else 0.0
        
        if not p_igst_col and not p_cgst_col and p_total_tax_col:
            df_purchase['p_total'] = pd.to_numeric(df_purchase[p_total_tax_col], errors='coerce').fillna(0)
            df_2b['g2b_total'] = df_2b['g2b_igst'] + df_2b['g2b_cgst'] + df_2b['g2b_sgst']
            
        # Composite Key Preparation
        df_2b['recon_key'] = df_2b['match_inv'] + "_" + df_2b['match_gstin']
        df_purchase['recon_key'] = df_purchase['match_inv'] + "_" + df_purchase['match_gstin']
        
        if st.button("🚀 Start Component Tax Reconciliation"):
            st.markdown("---")
            st.subheader("📊 Component-Level Reconciliation Results")
            
            g2b_cols_to_merge = ['recon_key', 'g2b_igst', 'g2b_cgst', 'g2b_sgst']
            g_taxable = find_matching_column(df_2b.columns, ["TAXABLE VALUE"])
            if g_taxable: g2b_cols_to_merge.append(g_taxable)
            
            base_matched = pd.merge(df_purchase, df_2b[g2b_cols_to_merge], on='recon_key', how='inner')
            
            # Differences Calculation
            base_matched['diff_igst'] = (base_matched['p_igst'] - base_matched['g2b_igst']).round(2)
            base_matched['diff_cgst'] = (base_matched['p_cgst'] - base_matched['g2b_cgst']).round(2)
            base_matched['diff_sgst'] = (base_matched['p_sgst'] - base_matched['g2b_sgst']).round(2)
            
            if p_igst_col or p_cgst_col:
                is_perfect = (base_matched['diff_igst'].abs() <= 1.0) & (base_matched['diff_cgst'].abs() <= 1.0) & (base_matched['diff_sgst'].abs() <= 1.0)
            else:
                base_matched['diff_total'] = (base_matched['p_total'] - base_matched['g2b_total']).round(2)
                is_perfect = base_matched['diff_total'].abs() <= 1.0
                
            perfect_match = base_matched[is_perfect]
            tax_mismatch = base_matched[~is_perfect]
            missing_in_2b = df_purchase[~df_purchase['recon_key'].isin(df_2b['recon_key'])]
            missing_in_books = df_2b[~df_2b['recon_key'].isin(df_purchase['recon_key'])]
            
            # Dashboard Analytics Cards
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("✅ Perfect Match", f"{len(perfect_match)} Bills")
            c2.metric("🔶 Tax Head Mismatch", f"{len(tax_mismatch)} Bills")
            c3.metric("🚨 Missing in GSTR-2B", f"{len(missing_in_2b)} Bills")
            c4.metric("⚠️ Missing in Books", f"{len(missing_in_books)} Bills")
            
            # Tabs Routing
            tab1, tab2, tab3, tab4 = st.tabs([
                "🟢 Perfect Match", 
                "🔴 Amount Mismatched", 
                "🚨 Missing in GSTR-2B", 
                "⚠️ Missing in Books"
            ])
            
            display_headers = []
            if p_inv_col: display_headers.append(p_inv_col)
            if p_gstin_col: display_headers.append(p_gstin_col)
            if g_taxable and g_taxable in base_matched.columns: display_headers.append(g_taxable)
            display_headers.extend(['p_igst', 'p_cgst', 'p_sgst', 'g2b_igst', 'g2b_cgst', 'g2b_sgst', 'diff_igst', 'diff_cgst', 'diff_sgst'])
            
            final_cols = [c for c in display_headers if c in base_matched.columns]
            extra_cols = [c for c in base_matched.columns if c not in final_cols and c not in ['recon_key', 'match_inv', 'match_gstin']]
            ordered_df_view = base_matched[final_cols + extra_cols]
            
            with tab1:
                st.success("Yeh saare invoices structural components aur ₹1 threshold margin ke mutabiq sahi match hain.")
                st.dataframe(ordered_df_view.loc[perfect_match.index], use_container_width=True)
            
            with tab2:
                st.warning("🔶 In invoices ke tax values mein mismatch mila hai. Pehli row Books (Purchase) ki hai aur dusri Portal (GSTR-2B) ki hai:")
                
                if not tax_mismatch.empty:
                    stacked_rows = []
                    for idx, row in tax_mismatch.iterrows():
                        p_row_data = {
                            "Invoice No": row[p_inv_col] if p_inv_col else "",
                            "GSTIN": row[p_gstin_col] if p_gstin_col else "",
                            "Source Data": "📚 Books (Purchase)",
                            "IGST": row['p_igst'],
                            "CGST": row['p_cgst'],
                            "SGST": row['p_sgst'],
                            "Difference Alert": f"IGST Diff: {row['diff_igst']} | CGST: {row['diff_cgst']} | SGST: {row['diff_sgst']}"
                        }
                        g_row_data = {
                            "Invoice No": row[p_inv_col] if p_inv_col else "",
                            "GSTIN": row[p_gstin_col] if p_gstin_col else "",
                            "Source Data": "🌐 Portal (GSTR-2B)",
                            "IGST": row['g2b_igst'],
                            "CGST": row['g2b_cgst'],
                            "SGST": row['g2b_sgst'],
                            "Difference Alert": f"IGST Diff: {row['diff_igst']} | CGST: {row['diff_cgst']} | SGST: {row['diff_sgst']}"
                        }
                        stacked_rows.append(p_row_data)
                        stacked_rows.append(g_row_data)
                    
                    df_stacked = pd.DataFrame(stacked_rows)
                    
                    # 👑 FONT COLOR RED HIGHLIGHT ENGINE WITH NO BACKGROUND
                    def highlight_mismatch_font(data):
                        style_df = pd.DataFrame('', index=data.index, columns=data.columns)
                        for i in range(0, len(data), 2):
                            if i + 1 < len(data):
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
                    
                    styled_view = df_stacked.style.apply(highlight_mismatch_font, axis=None)
                    st.dataframe(styled_view, use_container_width=True)
                else:
                    st.success("Koi bhi tax head mismatch nahi mila!")
                
            with tab3:
                st.error("🚨 Invoices aapki book mein hain par GSTR-2B portal par nahi aaye hain:")
                st.dataframe(missing_in_2b.drop(columns=['recon_key', 'match_inv', 'match_gstin'], errors='ignore'), use_container_width=True)
                
            with tab4:
                st.info("💡 Invoices portal par hain par purchase book mein entry missing hai:")
                st.dataframe(missing_in_books.drop(columns=['recon_key', 'match_inv', 'match_gstin'], errors='ignore'), use_container_width=True)
                
    except Exception as main_err:
        st.error(f"❌ Matching calculation system exception: {str(main_err)}")
else:
    st.info("💡 Shuru karne ke liye sidebar se GSTR-2B aur Purchase Book dono files refresh karke upload karein.")

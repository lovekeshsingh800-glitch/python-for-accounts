import streamlit as st
import pandas as pd
import datetime
import os
import io
import plotly.graph_objects as go
from streamlit_option_menu import option_menu

# Premium Page Title Config
st.set_page_config(page_title="Imprest Management System", layout="wide")

# --- STRICT DYNAMIC TARGET OVERWRITE ENGINE (COLORFUL SIDEBAR & MENUS) ---
st.markdown("""
<style>
    /* Side panel background aur heading controls */
    div[data-testid="stSidebar"] {
        background-color: #141722 !important;
    }
    div[data-testid="stSidebar"] h2 {
        color: #ffffff !important;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-weight: 600;
        margin-bottom: 5px;
    }
    /* Scope filters block headers colorful transformation */
    div[data-testid="stSidebar"] h3 {
        color: #38bdf8 !important;
        font-size: 16px !important;
        margin-top: 25px !important;
        font-weight: bold !important;
    }
    /* Scope selector apply button style modification */
    div[data-testid="stSidebar"] div.stButton button {
        background-color: rgba(56, 189, 248, 0.08) !important;
        border: 1px solid #38bdf8 !important;
        color: #38bdf8 !important;
        text-align: center !important;
        justify-content: center !important;
        width: 100% !important;
        border-radius: 6px !important;
    }
    div[data-testid="stSidebar"] div.stButton button:hover {
        background-color: #38bdf8 !important;
        color: #000000 !important;
    }
    /* Orange Theme Download & Backup Buttons CSS */
    div.stDownloadButton button {
        background-color: rgba(255, 170, 0, 0.1) !important;
        border: 1px solid #ffaa00 !important;
        color: #ffaa00 !important;
        font-weight: bold !important;
        width: 100% !important;
        border-radius: 6px !important;
        padding: 10px !important;
    }
    div.stDownloadButton button:hover {
        background-color: #ffaa00 !important;
        color: #000000 !important;
    }
    /* Clean Divider Line */
    hr {
        margin-top: 5px !important;
        margin-bottom: 15px !important;
        border-color: rgba(255, 255, 255, 0.1) !important;
    }
</style>
""", unsafe_allow_html=True)

# Navigation State Initialize
if "current_page" not in st.session_state:
    st.session_state.current_page = "📊 Dashboard Overview"

EXCEL_FILE = "imprest_database.xlsx"
# --- LIVE EXCEL DIRECT READ/WRITE LOGIC WITH PERMISSION ERROR CAPTURE ---
def load_data_from_excel_live():
    default_names = []
    default_cats = ["Office Supplies", "Local Travel", "Refreshments", "Internet Bill", "Repairs"]

    if not os.path.exists(EXCEL_FILE):
        default_df = pd.DataFrame({
            "Date": [datetime.date(2026, 7, 1), datetime.date(2026, 7, 2)],
            "Name": ["mohan", "mohan"],
            "Imprest Received (₹)": [50000.00, 10000.00],
            "Expense Category": ["Office Supplies", "Local Travel"],
            "Description": ["Starting Fund", "Initial Cash"],
            "Amount Spent (₹)": [1500.00, 450.00]
        })
        default_df["Date"] = pd.to_datetime(default_df["Date"]).dt.date
        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                default_df.to_excel(writer, sheet_name="Transactions", index=False)
                pd.DataFrame({"Allowed_Names": ["mohan"]}).to_excel(writer, sheet_name="Master_Names", index=False)
                pd.DataFrame({"Allowed_Categories": default_cats}).to_excel(writer, sheet_name="Master_Categories", index=False)
        except PermissionError:
            st.error("⚠️ **Active Excel File Lock:** System file 'imprest_database.xlsx' abhi background mein khuli hui hai. Please close the excel file and refresh this page.")
            st.stop()

        st.session_state.allowed_names = ["mohan"]
        st.session_state.allowed_categories = default_cats
        return default_df
    else:
        try:
            xl = pd.ExcelFile(EXCEL_FILE)
            if "Transactions" in xl.sheet_names:
                df = pd.read_excel(EXCEL_FILE, sheet_name="Transactions")
            else:
                df = pd.read_excel(EXCEL_FILE)
            try:
                if "Master_Names" in xl.sheet_names:
                    df_names = pd.read_excel(EXCEL_FILE, sheet_name="Master_Names")
                    saved_names = df_names["Allowed_Names"].dropna().astype(str).tolist()
                else:
                    saved_names = sorted(list(set(df["Name"].dropna().astype(str).tolist()))) if not df.empty else ["mohan"]
            except Exception:
                saved_names = ["mohan"]

            try:
                if "Master_Categories" in xl.sheet_names:
                    df_cats = pd.read_excel(EXCEL_FILE, sheet_name="Master_Categories")
                    saved_categories = df_cats["Allowed_Categories"].dropna().astype(str).tolist()
                else:
                    saved_categories = sorted(list(set(default_cats + df["Expense Category"].dropna().astype(str).tolist()))) if not df.empty else default_cats
            except Exception:
                saved_categories = default_cats

        except PermissionError:
            st.error("⚠️ **Active Excel File Lock:** System file 'imprest_database.xlsx' abhi background mein khuli hui hai. Please close the excel file and refresh this page.")
            st.stop()

        if "allowed_names" not in st.session_state:
            st.session_state.allowed_names = sorted(list(set(saved_names)))
        if "allowed_categories" not in st.session_state:
            st.session_state.allowed_categories = sorted(list(set(saved_categories)))

        if not df.empty:
            df["Date"] = pd.to_datetime(df["Date"]).dt.date
            df = df.sort_values(by="Date").reset_index(drop=True)
            return df
        else:
            return pd.DataFrame(columns=["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)"])
def save_data_to_excel_live(df_to_write):
    df_copy = df_to_write.copy()
    if "_source_index" in df_copy.columns:
        df_copy = df_copy.drop(columns=["_source_index"])

    if not df_copy.empty:
        df_copy["Date"] = pd.to_datetime(df_copy["Date"]).dt.date
        df_copy = df_copy.sort_values(by="Date").reset_index(drop=True)
        df_copy["Date"] = df_copy["Date"].astype(str)
        df_copy["Imprest Received (₹)"] = df_copy["Imprest Received (₹)"].apply(lambda x: float(x))
        df_copy["Amount Spent (₹)"] = df_copy["Amount Spent (₹)"].apply(lambda x: float(x))

    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_copy.to_excel(writer, sheet_name="Transactions", index=False)
            pd.DataFrame({"Allowed_Names": st.session_state.allowed_names}).to_excel(writer, sheet_name="Master_Names", index=False)
            pd.DataFrame({"Allowed_Categories": st.session_state.allowed_categories}).to_excel(writer, sheet_name="Master_Categories", index=False)
        return True
    except PermissionError:
        st.error("⚠️ **Active Excel File Lock:** Streamlit data save nahi kar paa raha kyunki 'imprest_database.xlsx' Microsoft Excel mein open hai. Please close the excel file and try again.")
        return False

if "running_master_df" not in st.session_state:
    st.session_state.running_master_df = load_data_from_excel_live()
else:
    load_data_from_excel_live()

if "active_years" not in st.session_state:
    st.session_state.active_years = []
if "active_months" not in st.session_state:
    st.session_state.active_months = []
if "active_names" not in st.session_state:
    st.session_state.active_names = []

# =========================================================================
# --- SIDEBAR PANEL (SOLID ORANGE HIGHLIGHTED STREAMLIT OPTION MENU) ---
# =========================================================================
with st.sidebar:
    st.markdown("<h2>🖥️ Main Menu</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    selected_menu = option_menu(
        menu_title=None,
        options=["Dashboard Overview", "Transaction Entries", "Master Configurations", "Bulk File Imports", "Database Backups"],
        icons=["bar-chart-fill", "pencil-square", "gear-fill", "cloud-arrow-up-fill", "shield-lock-fill"],
        menu_icon=None,
        default_index=0,
        styles={
            "container": {"padding": "0px", "background-color": "transparent"},
            "icon": {"color": "#e0e6ed", "font-size": "15px"}, 
            "nav-link": {
                "font-size": "15px", 
                "text-align": "left", 
                "margin": "4px 0px", 
                "color": "#a0aec0",
                "padding": "12px 16px",
                "background-color": "transparent"
            },
            "nav-link-selected": {
                "background-color": "#ffaa00", 
                "color": "#000000", 
                "font-weight": "bold",
                "box-shadow": "0px 4px 10px rgba(255, 170, 0, 0.2)"
            }
        }
    )
    
    st.session_state.current_page = f"{'📊 ' if selected_menu == 'Dashboard Overview' else '📝 ' if selected_menu == 'Transaction Entries' else '⚙️ ' if selected_menu == 'Master Configurations' else '📥 ' if selected_menu == 'Bulk File Imports' else '🛠️ '}{selected_menu}"
    st.markdown("### 🔍 Scope Filter Criteria</h3>", unsafe_allow_html=True)
    df_filter_core = st.session_state.running_master_df.copy()
    if not df_filter_core.empty:
        df_filter_core["Date"] = pd.to_datetime(df_filter_core["Date"])
        all_years = sorted(df_filter_core["Date"].dt.year.dropna().unique().tolist())
        all_months = df_filter_core["Date"].dt.strftime("%B").dropna().unique().tolist()
    else:
        all_years = [datetime.date.today().year]
        all_months = []

    selected_years = st.multiselect("Target Years:", options=[str(y) for y in all_years], default=st.session_state.active_years)
    selected_months = st.multiselect("Target Months:", options=all_months, default=st.session_state.active_months)
    selected_names = st.multiselect("Target Accounts:", options=st.session_state.allowed_names, default=st.session_state.active_names)

    if st.button("Apply Scope Selections", use_container_width=True):
        st.session_state.active_years = selected_years
        st.session_state.active_months = selected_months
        st.session_state.active_names = selected_names
        st.rerun()

# =========================================================================
# --- CONDITIONAL INTERFACE LOGIC MANAGEMENT BY SYSTEM ACTIVE TABS ---
# =========================================================================

# ----------------- MODULE: TRANSACTION VOUCHER INSERTS -----------------
if st.session_state.current_page == "📝 Transaction Entries":
    st.header("📝 Live Voucher Recording Framework")
    with st.container(border=True):
        col_input1, col_input2, col_input3 = st.columns(3)
        with col_input1:
            chosen_date = st.date_input("Transaction Date:", datetime.date.today())
            chosen_name = st.selectbox("Select Account Name:", st.session_state.allowed_names, index=0, key="exp_insert_name")
        with col_input2:
            raw_allocated_amt = st.number_input("Imprest Allotment (₹)", step=500.0, value=0.0, key="exp_insert_rec")
            chosen_cat = st.selectbox("Ledger Category:", st.session_state.allowed_categories)
        with col_input3:
            chosen_desc = st.text_input("Voucher Description:", value="Operational Expense", key="exp_insert_desc")
            raw_chosen_spent = st.number_input("Expense Amount (₹)", step=100.0, value=0.0, key="exp_insert_spent")

        if st.button("Submit Voucher to Excel Database", use_container_width=True, key="exp_insert_btn"):
            new_entry = {
                "Date": chosen_date, "Name": chosen_name, "Imprest Received (₹)": float(raw_allocated_amt),
                "Expense Category": chosen_cat, "Description": chosen_desc, "Amount Spent (₹)": float(raw_chosen_spent)
            }
            updated_df_via_form = pd.concat([st.session_state.running_master_df, pd.DataFrame([new_entry])], ignore_index=True)
            if save_data_to_excel_live(updated_df_via_form):
                st.session_state.running_master_df = load_data_from_excel_live()
                st.success("Voucher registered and saved inside Excel workbook records successfully!")
# ----------------- MODULE: MASTER LEDGER MANAGEMENT -----------------
elif st.session_state.current_page == "⚙️ Master Configurations":
    st.header("⚙️ Master Registration controls")
    setup_col1, setup_col2 = st.columns(2)
    with setup_col1:
        with st.container(border=True):
            st.markdown("### Register New Member Profile")
            new_custom_name = st.text_input("Enter Full Name:").strip()
            if st.button("Register Name", use_container_width=True):
                if new_custom_name and new_custom_name not in st.session_state.allowed_names:
                    st.session_state.allowed_names.append(new_custom_name)
                    save_data_to_excel_live(st.session_state.running_master_df)
                    st.success(f"'{new_custom_name}' created in registry database.")
                    st.rerun()

            st.markdown("---")
            name_to_manage = st.selectbox("Select Target Registry Item:", st.session_state.allowed_names)
            edit_name_input = st.text_input("Type replacement context text:").strip()
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if st.button("Rename Target Record", use_container_width=True):
                    if edit_name_input and edit_name_input not in st.session_state.allowed_names:
                        idx = st.session_state.allowed_names.index(name_to_manage)
                        st.session_state.allowed_names[idx] = edit_name_input
                        st.session_state.running_master_df["Name"] = st.session_state.running_master_df["Name"].replace(name_to_manage, edit_name_input)
                        save_data_to_excel_live(st.session_state.running_master_df)
                        st.rerun()
            with btn_col2:
                if st.button("Purge Profile Data", use_container_width=True):
                    if len(st.session_state.allowed_names) > 1:
                        st.session_state.running_master_df = st.session_state.running_master_df[st.session_state.running_master_df["Name"] != name_to_manage].reset_index(drop=True)
                        st.session_state.allowed_names.remove(name_to_manage)
                        save_data_to_excel_live(st.session_state.running_master_df)
                        st.rerun()
    with setup_col2:
        with st.container(border=True):
            st.markdown("### Ledger Cost Categories Definitions")
            new_custom_cat = st.text_input("Enter Category Type Context Label:").strip()
            if st.button("Register New Ledger", use_container_width=True):
                if new_custom_cat and new_custom_cat not in st.session_state.allowed_categories:
                    st.session_state.allowed_categories.append(new_custom_cat)
                    save_data_to_excel_live(st.session_state.running_master_df)
                    st.success("New accounting ledger category mapped.")
                    st.rerun()
# ----------------- MODULE: BULK SHEET STREAM FILE LOGIC -----------------
elif st.session_state.current_page == "📥 Bulk File Imports":
    st.header("📥 Bulk Auditing File Pipelines Processing")
    io_col1, io_col2 = st.columns(2)
    with io_col1:
        with st.container(border=True):
            st.write("**Step 1: Download Empty Database Schema Template**")
            template_cols = ["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)"]
            template_df = pd.DataFrame(columns=template_cols)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False, sheet_name="Template")
            buffer.seek(0)
            st.download_button(
                label="⬇️ Export Excel Blank Schema", data=buffer,
                file_name="imprest_import_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
            )
    with io_col2:
        with st.container(border=True):
            st.write("**Step 2: Upload Populated Sheet File**")
            uploaded_file = st.file_uploader("Choose Excel File Workbook Asset", type=["xlsx"], label_visibility="collapsed")
            if uploaded_file is not None:
                try:
                    imported_df = pd.read_excel(uploaded_file)
                    required_cols = ["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)"]
                    if all(col in imported_df.columns for col in required_cols):
                        imported_df["Date"] = pd.to_datetime(imported_df["Date"]).dt.date
                        imported_df["Imprest Received (₹)"] = pd.to_numeric(imported_df["Imprest Received (₹)"]).fillna(0.0)
                        imported_df["Amount Spent (₹)"] = pd.to_numeric(imported_df["Amount Spent (₹)"]).fillna(0.0)
                        imported_df["Name"] = imported_df["Name"].astype(str).str.strip()
                        imported_df["Expense Category"] = imported_df["Expense Category"].astype(str).str.strip()
                        imported_df["Description"] = imported_df["Description"].fillna("").astype(str).str.strip()
                        
                        if st.button("🚀 Process & Merge Import Entries", use_container_width=True):
                            current_db = st.session_state.running_master_df.copy()
                            current_db["Date"] = pd.to_datetime(current_db["Date"]).dt.date
                            match_keys = ["Date", "Name", "Expense Category", "Description"]
                            
                            current_db_unique = current_db.groupby(match_keys, as_index=False).agg({"Imprest Received (₹)": "sum", "Amount Spent (₹)": "sum"})
                            imported_df_unique = imported_df.groupby(match_keys, as_index=False).agg({"Imprest Received (₹)": "sum", "Amount Spent (₹)": "sum"})
                            
                            current_db_unique.set_index(match_keys, inplace=True)
                            imported_df_unique.set_index(match_keys, inplace=True)
                            current_db_unique.update(imported_df_unique)
                            
                            new_entries = imported_df_unique[~imported_df_unique.index.isin(current_db_unique.index)]
                            final_merged_df = pd.concat([current_db_unique, new_entries]).reset_index()
                            
                            for imported_name in final_merged_df["Name"].unique():
                                if imported_name and imported_name not in st.session_state.allowed_names: st.session_state.allowed_names.append(imported_name)
                            for imported_cat in final_merged_df["Expense Category"].unique():
                                if imported_cat and imported_cat not in st.session_state.allowed_categories: st.session_state.allowed_categories.append(imported_cat)
                            
                            if save_data_to_excel_live(final_merged_df):
                                st.session_state.running_master_df = load_data_from_excel_live()
                                st.success("✅ Bulk database records verified and processed into ledger!")
                                st.rerun()
                except Exception as e: st.error(f"Engine parsing failure: {str(e)}")

# ----------------- MODULE: REMOTE USER PERSONAL PC DOWNLOAD BACKUP -----------------
elif st.session_state.current_page == "🛠️ Database Backups":
    st.header("🛠️ Secure Remote Database Backups")
    with st.container(border=True):
        st.write("### 📥 Download Database Backup to Your PC")
        st.write("Aapke computer par Python nahi hai tab bhi koi dikkat nahi hai. Neeche diye gaye orange button par click karte hi yeh system poore live server database ki ek clean backup file direct aapke computer ke **'Downloads'** folder mein save kar dega.")
        
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                excel_bytes = f.read()
                
            current_timestamp = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            st.write("")
            st.download_button(
                label="💾 Download Database Backup to Personal PC",
                data=excel_bytes,
                file_name=f"imprest_database_backup_{current_timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="browser_backup_trigger"
            )
        else: st.error("System error: Core database file path missing on server environment.")
# ----------------- MODULE: THE COMPREHENSIVE FINANCIAL OVERVIEW -----------------
elif st.session_state.current_page == "📊 Dashboard Overview":
    st.session_state.disp_names_str = ", ".join(st.session_state.active_names) if st.session_state.active_names else "All Users"
    st.session_state.disp_years_str = ", ".join(st.session_state.active_years) if st.session_state.active_years else "All Years"
    st.session_state.disp_months_str = ", ".join(st.session_state.active_months) if st.session_state.active_months else "All Months"

    st.subheader(f"User Summary Overview (Timeline Scope: {st.session_state.disp_years_str} - {st.session_state.disp_months_str})")
    calc_df_master = st.session_state.running_master_df.copy()
    if not calc_df_master.empty:
        calc_df_master["Date"] = pd.to_datetime(calc_df_master["Date"])
        calc_df_master["Imprest Received (₹)"] = pd.to_numeric(calc_df_master["Imprest Received (₹)"]).fillna(0.0)
        calc_df_master["Amount Spent (₹)"] = pd.to_numeric(calc_df_master["Amount Spent (₹)"]).fillna(0.0)

    target_users_cards = st.session_state.active_names if st.session_state.active_names else st.session_state.allowed_names

    if target_users_cards and not calc_df_master.empty:
        cols_metrics = st.columns(len(target_users_cards))
        for i, name_user in enumerate(target_users_cards):
            with cols_metrics[i]:
                user_full_log = calc_df_master[calc_df_master["Name"] == name_user].sort_values(by="Date")
                u_opening, u_received_this_scope, u_spent_this_scope = 0.0, 0.0, 0.0
                scope_df = user_full_log.copy()
                if st.session_state.active_years: scope_df = scope_df[scope_df["Date"].dt.year.astype(str).isin(st.session_state.active_years)]
                if st.session_state.active_months: scope_df = scope_df[scope_df["Date"].dt.strftime("%B").isin(st.session_state.active_months)]

                if not scope_df.empty and (st.session_state.active_years or st.session_state.active_months):
                    min_scope_date = scope_df["Date"].min()
                    historical_logs = user_full_log[user_full_log["Date"] < min_scope_date]
                    u_opening = float(historical_logs["Imprest Received (₹)"].sum() - historical_logs["Amount Spent (₹)"].sum())
                    u_received_this_scope = float(scope_df["Imprest Received (₹)"].sum())
                    u_spent_this_scope = float(scope_df["Amount Spent (₹)"].sum())
                else:
                    u_received_this_scope = float(user_full_log["Imprest Received (₹)"].sum())
                    u_spent_this_scope = float(user_full_log["Amount Spent (₹)"].sum())

                u_final_closing = u_opening + u_received_this_scope - u_spent_this_scope
                with st.container(border=True):
                    st.write(f"### User: {name_user}")
                    l_col, r_col = st.columns(2)
                    with l_col:
                        st.write("Opening Balance:")
                        st.write("Received Funds:")
                        st.write("Amount Spent:")
                    with r_col:
                        st.write(f"**₹{u_opening:,.2f}**")
                        st.write(f"**₹{u_received_this_scope:,.2f}**")
                        st.write(f"**₹{u_spent_this_scope:,.2f}**")
                    st.write("") 
                    if u_final_closing >= 0: st.success(f"Net Balance: ₹{u_final_closing:,.2f}")
                    else: st.error(f"Overdrawn: ₹{u_final_closing:,.2f}")
    st.markdown("---")

    master_working_df = st.session_state.running_master_df.copy()
    if not master_working_df.empty:
        master_working_df["Date"] = pd.to_datetime(master_working_df["Date"])
        master_working_df = master_working_df.sort_values(by="Date").reset_index(drop=True)
    master_working_df["_source_index"] = master_working_df.index
    master_working_df["Imprest Received (₹)"] = pd.to_numeric(master_working_df["Imprest Received (₹)"]).fillna(0.0)
    master_working_df["Amount Spent (₹)"] = pd.to_numeric(master_working_df["Amount Spent (₹)"]).fillna(0.0)

    user_timeline = master_working_df.copy()
    if st.session_state.active_names: user_timeline = user_timeline[user_timeline["Name"].isin(st.session_state.active_names)]
    opening_balance, closing_balance, scope_received, scope_spent = 0.0, 0.0, 0.0, 0.0
    filtered_df = user_timeline.copy()
    if not filtered_df.empty:
        if st.session_state.active_names: filtered_df = filtered_df[filtered_df["Name"].isin(st.session_state.active_names)]
        if st.session_state.active_years: filtered_df = filtered_df[filtered_df["Date"].dt.year.astype(str).isin(st.session_state.active_years)]
        if st.session_state.active_months: filtered_df = filtered_df[filtered_df["Date"].dt.strftime("%B").isin(st.session_state.active_months)]

    if not filtered_df.empty and (st.session_state.active_years or st.session_state.active_months):
        min_scope_date = filtered_df["Date"].min()
        prev_records = user_timeline[user_timeline["Date"] < min_scope_date]
        opening_balance = float(prev_records["Imprest Received (₹)"].sum() - prev_records["Amount Spent (₹)"].sum())
        scope_received = float(filtered_df["Imprest Received (₹)"].sum())
        scope_spent = float(filtered_df["Amount Spent (₹)"].sum())
        closing_balance = opening_balance + scope_received - scope_spent
    else:
        scope_received = float(user_timeline["Imprest Received (₹)"].sum()) if not user_timeline.empty else 0.0
        scope_spent = float(user_timeline["Amount Spent (₹)"].sum()) if not user_timeline.empty else 0.0
        closing_balance = scope_received - scope_spent

    st.subheader(f"Scope Metrics ({st.session_state.disp_names_str})")
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1: st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.1); padding: 15px; border-radius: 6px;'>Scope Opening Balance<br><span style='color:#38bdf8; font-size:24px; font-weight:bold;'>₹{opening_balance:,.2f}</span></div>", unsafe_allow_html=True)
    with col_m2: st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.1); padding: 15px; border-radius: 6px;'>Scope Received Funds<br><span style='color:#22c55e; font-size:24px; font-weight:bold;'>₹{scope_received:,.2f}</span></div>", unsafe_allow_html=True)
    with col_m3: st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.1); padding: 15px; border-radius: 6px;'>Scope Expenses<br><span style='color:#ef4444; font-size:24px; font-weight:bold;'>₹{scope_spent:,.2f}</span></div>", unsafe_allow_html=True)
    with col_m4: 
        c_color = "#22c55e" if closing_balance >= 0 else "#ef4444"
        st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.1); padding: 15px; border-radius: 6px;'>Scope Closing Balance<br><span style='color:{c_color}; font-size:24px; font-weight:bold;'>₹{closing_balance:,.2f}</span></div>", unsafe_allow_html=True)
    st.markdown("---")

    st.subheader("Auditable Transaction Log Statement")
    column_config = {
        "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY", required=True),
        "Name": st.column_config.SelectboxColumn("Account Name", options=st.session_state.allowed_names, required=True),
        "Imprest Received (₹)": st.column_config.NumberColumn("Imprest Received (₹)", format="₹%d", required=True),
        "Expense Category": st.column_config.SelectboxColumn("Ledger Category", options=st.session_state.allowed_categories, required=True),
        "Description": st.column_config.TextColumn("Voucher Description", required=False),
        "Amount Spent (₹)": st.column_config.NumberColumn("Amount Spent (₹)", format="₹%d", required=True),
    }
    if not filtered_df.empty:
        display_df = filtered_df[["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)", "_source_index"]].copy()
        display_df["Date"] = pd.to_datetime(display_df["Date"]).dt.date
        display_df = display_df.sort_values(by="Date").reset_index(drop=True)
    else: display_df = pd.DataFrame(columns=["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)", "_source_index"])
    edited_df = st.data_editor(display_df, column_config=column_config, num_rows="dynamic", use_container_width=True, key="data_editor_widget")
    if not edited_df.equals(display_df):
        updated_master = st.session_state.running_master_df.copy()
        if not updated_master.empty: updated_master["Date"] = pd.to_datetime(updated_master["Date"]).dt.date
        original_source_ids = set(display_df["_source_index"].dropna().astype(int).tolist())
        current_source_ids = set(edited_df["_source_index"].dropna().astype(int).tolist())
        deleted_ids = original_source_ids - current_source_ids
        if deleted_ids: updated_master = updated_master.drop(list(deleted_ids)).reset_index(drop=True)

        for idx in edited_df.index:
            row_data = edited_df.loc[idx].to_dict()
            source_id = row_data.get("_source_index")
            if pd.isna(row_data["Date"]) or row_data["Date"] is None: row_data["Date"] = datetime.date.today()
            else: row_data["Date"] = pd.to_datetime(row_data["Date"]).date()

            clean_row = {
                "Date": row_data["Date"], 
                "Name": row_data["Name"], 
                "Imprest Received (₹)": float(row_data["Imprest Received (₹)"]) if pd.notna(row_data["Imprest Received (₹)"]) else 0.0,
                "Expense Category": row_data["Expense Category"], 
                "Description": row_data["Description"] if pd.notna(row_data["Description"]) else "", 
                "Amount Spent (₹)": float(row_data["Amount Spent (₹)"]) if pd.notna(row_data["Amount Spent (₹)"]) else 0.0
            }
            if pd.notna(source_id) and int(source_id) in updated_master.index:
                for col in clean_row: updated_master.at[int(source_id), col] = clean_row[col]
            else:
                if pd.notna(clean_row["Name"]) and pd.notna(clean_row["Expense Category"]): updated_master = pd.concat([updated_master, pd.DataFrame([clean_row])], ignore_index=True)
        if not updated_master.empty:
            updated_master["Date"] = pd.to_datetime(updated_master["Date"]).dt.date
            updated_master = updated_master.sort_values(by="Date").reset_index(drop=True)
        if save_data_to_excel_live(updated_master):
            st.session_state.running_master_df = load_data_from_excel_live()
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Expense Category Ledger Metrics Summary")
    if not filtered_df.empty:
        summary_cat_df = filtered_df.groupby("Expense Category").agg({"Imprest Received (₹)": "sum", "Amount Spent (₹)": "sum"}).reset_index()
        summary_cat_df = summary_cat_df.sort_values(by="Amount Spent (₹)", ascending=False).reset_index(drop=True)
        
        # --- FIXED STYLER BYPASS ENGINE: Removes Arrow/Proto format failures on cloud servers ---
        st.dataframe(summary_cat_df, use_container_width=True)
        
        total_inflow_calc = float(summary_cat_df["Imprest Received (₹)"].sum())
        total_outflow_calc = float(summary_cat_df["Amount Spent (₹)"].sum())
        
        col_t1, col_t2 = st.columns(2)
        with col_t1: st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.05); padding: 12px; border-radius: 6px; background-color: rgba(34,197,94,0.02);'>📊 Scope Category Total Inflow<br><span style='color:#22c55e; font-size:22px; font-weight:bold;'>₹{total_inflow_calc:,.2f}</span></div>", unsafe_allow_html=True)
        with col_t2: st.markdown(f"<div style='border: 1px solid rgba(255,255,255,0.05); padding: 12px; border-radius: 6px; background-color: rgba(239,68,68,0.02);'>📉 Scope Category Total Outflow<br><span style='color:#ef4444; font-size:22px; font-weight:bold;'>₹{total_outflow_calc:,.2f}</span></div>", unsafe_allow_html=True)
        
        export_buffer = io.BytesIO()
        with pd.ExcelWriter(export_buffer, engine='openpyxl') as report_writer:
            summary_cat_df.to_excel(report_writer, index=False, sheet_name="Category_Summaries")
        export_buffer.seek(0)
        
        st.write("")
        st.download_button(
            label="📥 Download Sorted Ledger Report (Excel)",
            data=export_buffer,
            file_name="Sorted_Ledger_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        g_col1, g_col2 = st.columns(2)
        with g_col1:
            fig1 = go.Figure()
            daily_spent = filtered_df.groupby(filtered_df["Date"].dt.date)["Amount Spent (₹)"].sum().reset_index()
            fig1.add_trace(go.Scatter(x=daily_spent["Date"], y=daily_spent["Amount Spent (₹)"], fill='tozeroy', mode='lines+markers', line=dict(color='#ff9900'), fillcolor='rgba(255,153,0,0.1)'))
            fig1.update_layout(title="✨ Daily Expense Flow Trend", plot_bgcolor='#0a0c10', paper_bgcolor='#0a0c10', font=dict(color='#e0e6ed'))
            st.plotly_chart(fig1, use_container_width=True)
        with g_col2:
            fig2 = go.Figure()
            user_funds = filtered_df.groupby("Name")[["Imprest Received (₹)", "Amount Spent (₹)"]].sum().reset_index()
            fig2.add_trace(go.Bar(x=user_funds["Name"], y=user_funds["Imprest Received (₹)"], name="Inflow (Allocated)", marker_color='#22c55e'))
            fig2.add_trace(go.Bar(x=user_funds["Name"], y=-user_funds["Amount Spent (₹)"], name="Outflow (Expenses)", marker_color='#ef4444'))
            fig2.update_layout(title="⚡ User Structural Inflow vs Outflow Matrix", barmode='relative', plot_bgcolor='#0a0c10', paper_bgcolor='#0a0c10', font=dict(color='#e0e6ed'))
            st.plotly_chart(fig2, use_container_width=True)

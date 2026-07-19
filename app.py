import streamlit as st
import pandas as pd
import datetime
import os
import plotly.graph_objects as go


# Premium Page Title Config
st.set_page_config(page_title="Imprest Management System", layout="wide")
st.title("Imprest & Expense Management Ledger")
st.markdown("---")

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
            st.error("⚠️ **Active Excel File Lock:** System file 'imprest_database.xlsx' abhi background mein khuli hui hai. **Please close the excel file** and refresh this page.")
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
            st.error("⚠️ **Active Excel File Lock:** System file 'imprest_database.xlsx' abhi background mein khuli hui hai. **Please close the excel file** and refresh this page.")
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
        st.error("⚠️ **Active Excel File Lock:** Streamlit data save nahi kar paa raha kyunki 'imprest_database.xlsx' Microsoft Excel mein open hai. **Please close the excel file** and try again.")
        return False

if "running_master_df" not in st.session_state:
    st.session_state.running_master_df = load_data_from_excel_live()
else:
    load_data_from_excel_live()
# --- OPTION MANAGEMENT SECTION ---
st.subheader("Master Settings & Dropdown Controls")
setup_col1, setup_col2 = st.columns(2)

with setup_col1:
    with st.popover("Manage Master Names List"):
        st.write("**Add New Member**")
        new_custom_name = st.text_input("Enter Full Name:").strip()
        if st.button("Register Name"):
            if new_custom_name and new_custom_name not in st.session_state.allowed_names:
                st.session_state.allowed_names.append(new_custom_name)
                save_data_to_excel_live(st.session_state.running_master_df)
                st.success(f"'{new_custom_name}' registered successfully.")
                st.rerun()

        st.markdown("---")
        st.write("**Modify or Remove Name**")
        name_to_manage = st.selectbox("Select Target Name:", st.session_state.allowed_names)
        edit_name_input = st.text_input("Type new text to rename:").strip()

        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("Rename Name"):
                if edit_name_input and edit_name_input not in st.session_state.allowed_names:
                    idx = st.session_state.allowed_names.index(name_to_manage)
                    st.session_state.allowed_names[idx] = edit_name_input
                    st.session_state.running_master_df["Name"] = st.session_state.running_master_df["Name"].replace(name_to_manage, edit_name_input)
                    save_data_to_excel_live(st.session_state.running_master_df)
                    st.success("Global database records renamed successfully!")
                    st.rerun()
        with btn_col2:
            if st.button("Delete Name"):
                if len(st.session_state.allowed_names) > 1:
                    st.session_state.running_master_df = st.session_state.running_master_df[st.session_state.running_master_df["Name"] != name_to_manage].reset_index(drop=True)
                    st.session_state.allowed_names.remove(name_to_manage)
                    if "active_names" in st.session_state and name_to_manage in st.session_state.active_names:
                        st.session_state.active_names.remove(name_to_manage)
                    save_data_to_excel_live(st.session_state.running_master_df)
                    st.warning(f"'{name_to_manage}' has been permanently deleted.")
                    st.rerun()
                else:
                    st.error("Kam se kam ek account rehna zaroori hai!")

with setup_col2:
    with st.popover("Manage Expense Categories"):
        st.write("**Add New Ledger Category**")
        new_custom_cat = st.text_input("Enter Category Type:").strip()
        if st.button("Register Category"):
            if new_custom_cat and new_custom_cat not in st.session_state.allowed_categories:
                st.session_state.allowed_categories.append(new_custom_cat)
                save_data_to_excel_live(st.session_state.running_master_df)
                st.success(f"'{new_custom_cat}' registered successfully.")
                st.rerun()

st.markdown("---")
if "active_years" not in st.session_state:
    st.session_state.active_years = []
if "active_months" not in st.session_state:
    st.session_state.active_months = []
if "active_names" not in st.session_state:
    st.session_state.active_names = []

# --- PREMIUM USER-WISE IMPREST SUMMARY PANEL ---
disp_y_str = ", ".join(st.session_state.active_years) if st.session_state.active_years else "All Years"
disp_m_str = ", ".join(st.session_state.active_months) if st.session_state.active_months else "All Months"
st.subheader(f"User Summary Overview (Timeline Scope: {disp_y_str} - {disp_m_str})")

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
            if st.session_state.active_years:
                scope_df = scope_df[scope_df["Date"].dt.year.astype(str).isin(st.session_state.active_years)]
            if st.session_state.active_months:
                scope_df = scope_df[scope_df["Date"].dt.strftime("%B").isin(st.session_state.active_months)]

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
                if u_final_closing >= 0:
                    st.success(f"Net Balance: ₹{u_final_closing:,.2f}")
                else:
                    st.error(f"Overdrawn: ₹{u_final_closing:,.2f}")
st.markdown("---")

# --- LIVE SUBMISSION INPUT FORM ---
st.subheader("Transaction Entry Management")
with st.expander("Record New Transaction / Voucher Entry", expanded=False):
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

    if st.button("Submit Voucher to Excel", use_container_width=True, key="exp_insert_btn"):
        new_entry = {
            "Date": chosen_date, "Name": chosen_name, "Imprest Received (₹)": float(raw_allocated_amt),
            "Expense Category": chosen_cat, "Description": chosen_desc, "Amount Spent (₹)": float(raw_chosen_spent)
        }
        updated_df_via_form = pd.concat([st.session_state.running_master_df, pd.DataFrame([new_entry])], ignore_index=True)
        if save_data_to_excel_live(updated_df_via_form):
            st.session_state.running_master_df = load_data_from_excel_live()
            st.toast("Voucher saved to Excel database successfully.")
            st.rerun()

st.markdown("---")
st.session_state.disp_names_str = ", ".join(st.session_state.active_names) if st.session_state.active_names else "All Users"
st.session_state.disp_years_str = ", ".join(st.session_state.active_years) if st.session_state.active_years else "All Years"
st.session_state.disp_months_str = ", ".join(st.session_state.active_months) if st.session_state.active_months else "All Months"

# --- SIDEBAR FORM LOGIC ---
with st.sidebar.form(key="filter_form"):
    st.header("Scope Filter Controls")
    df_filter_core = st.session_state.running_master_df.copy()

    if not df_filter_core.empty:
        df_filter_core["Date"] = pd.to_datetime(df_filter_core["Date"])
        all_years = sorted(df_filter_core["Date"].dt.year.dropna().unique().tolist())
        all_months = df_filter_core["Date"].dt.strftime("%B").dropna().unique().tolist()
    else:
        all_years = [datetime.date.today().year]
        all_months = []

    selected_years = st.multiselect("Select Target Years:", options=[str(y) for y in all_years], default=st.session_state.active_years)
    selected_months = st.multiselect("Select Target Months:", options=all_months, default=st.session_state.active_months)
    selected_names = st.multiselect("Select Target Accounts:", options=st.session_state.allowed_names, default=st.session_state.active_names)

    submit_filters = st.form_submit_button(label="Apply Scope Selections", use_container_width=True)

if submit_filters:
    st.session_state.active_years = selected_years
    st.session_state.active_months = selected_months
    st.session_state.active_names = selected_names
    st.rerun()

# --- TIMELINE LEDGER CALCULATIONS ---
master_working_df = st.session_state.running_master_df.copy()
if not master_working_df.empty:
    master_working_df["Date"] = pd.to_datetime(master_working_df["Date"])
    master_working_df = master_working_df.sort_values(by="Date").reset_index(drop=True)
master_working_df["_source_index"] = master_working_df.index
master_working_df["Imprest Received (₹)"] = pd.to_numeric(master_working_df["Imprest Received (₹)"]).fillna(0.0)
master_working_df["Amount Spent (₹)"] = pd.to_numeric(master_working_df["Amount Spent (₹)"]).fillna(0.0)

user_timeline = master_working_df.copy()
if st.session_state.active_names:
    user_timeline = user_timeline[user_timeline["Name"].isin(st.session_state.active_names)]

opening_balance, closing_balance, scope_received, scope_spent = 0.0, 0.0, 0.0, 0.0

filtered_df = user_timeline.copy()
if not filtered_df.empty:
    if st.session_state.active_years:
        filtered_df = filtered_df[filtered_df["Date"].dt.year.astype(str).isin(st.session_state.active_years)]
    if st.session_state.active_months:
        filtered_df = filtered_df[filtered_df["Date"].dt.strftime("%B").isin(st.session_state.active_months)]

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

st.subheader(f"Scope Metrics ({st.session_state.disp_names_str} | Year: {st.session_state.disp_years_str} | Month: {st.session_state.disp_months_str})")
col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.metric(label="Scope Opening Balance", value=f"₹{opening_balance:,.2f}")
with col_m2:
    st.metric(label="Scope Received Funds", value=f"₹{scope_received:,.2f}")
with col_m3:
    st.metric(label="Scope Expenses", value=f"₹{scope_spent:,.2f}")
with col_m4:
    st.metric(label="Scope Closing Balance", value=f"₹{closing_balance:,.2f}")

st.markdown("---")
# --- INTERACTIVE DATA TABLE ---
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
else:
    display_df = pd.DataFrame(columns=["Date", "Name", "Imprest Received (₹)", "Expense Category", "Description", "Amount Spent (₹)", "_source_index"])

edited_df = st.data_editor(display_df, column_config=column_config, num_rows="dynamic", use_container_width=True, key="data_editor_widget")

if not edited_df.equals(display_df):
    updated_master = st.session_state.running_master_df.copy()
    if not updated_master.empty:
        updated_master["Date"] = pd.to_datetime(updated_master["Date"]).dt.date

    original_source_ids = set(display_df["_source_index"].dropna().astype(int).tolist())
    current_source_ids = set(edited_df["_source_index"].dropna().astype(int).tolist())
    deleted_ids = original_source_ids - current_source_ids

    if deleted_ids:
        updated_master = updated_master.drop(list(deleted_ids)).reset_index(drop=True)

    for idx in edited_df.index:
        row_data = edited_df.loc[idx].to_dict()
        source_id = row_data.get("_source_index")

        if pd.isna(row_data["Date"]) or row_data["Date"] is None:
            row_data["Date"] = datetime.date.today()
        else:
            row_data["Date"] = pd.to_datetime(row_data["Date"]).date()

        clean_row = {
            "Date": row_data["Date"], 
            "Name": row_data["Name"], 
            "Imprest Received (₹)": float(row_data["Imprest Received (₹)"]) if pd.notna(row_data["Imprest Received (₹)"]) else 0.0,
            "Expense Category": row_data["Expense Category"], 
            "Description": row_data["Description"] if pd.notna(row_data["Description"]) else "", 
            "Amount Spent (₹)": float(row_data["Amount Spent (₹)"]) if pd.notna(row_data["Amount Spent (₹)"]) else 0.0
        }

        if pd.notna(source_id) and int(source_id) in updated_master.index:
            for col in clean_row: 
                updated_master.at[int(source_id), col] = clean_row[col]
        else:
            if pd.notna(clean_row["Name"]) and pd.notna(clean_row["Expense Category"]):
                updated_master = pd.concat([updated_master, pd.DataFrame([clean_row])], ignore_index=True)

    if not updated_master.empty:
        updated_master["Date"] = pd.to_datetime(updated_master["Date"]).dt.date
        updated_master = updated_master.sort_values(by="Date").reset_index(drop=True)

    if save_data_to_excel_live(updated_master):
        st.session_state.running_master_df = load_data_from_excel_live()
        st.toast("Excel database updated directly from layout grid.")
        st.rerun()

# --- DATA PROCESSOR FOR ANALYTICS ---
master_working_df = st.session_state.running_master_df.copy()
master_working_df["Date"] = pd.to_datetime(master_working_df["Date"])
master_working_df = master_working_df.sort_values(by="Date").reset_index(drop=True)
master_working_df["_source_index"] = master_working_df.index

filtered_df = master_working_df.copy()
if st.session_state.active_names: filtered_df = filtered_df[filtered_df["Name"].isin(st.session_state.active_names)]
if st.session_state.active_years: filtered_df = filtered_df[filtered_df["Date"].dt.year.astype(str).isin(st.session_state.active_years)]
if st.session_state.active_months: filtered_df = filtered_df[filtered_df["Date"].dt.strftime("%B").isin(st.session_state.active_months)]

st.markdown("---")

# --- NEW: EXPENSE CATEGORY SUMMARY TABLE OVERVIEW ---
st.subheader("📋 Expense Category Ledger Metrics Summary")
if not filtered_df.empty:
    summary_cat_df = filtered_df.groupby("Expense Category").agg(
        Total_Allocated_Inflow=("Imprest Received (₹)", "sum"),
        Total_Amount_Spent=("Amount Spent (₹)", "sum")
    ).reset_index()
    
    summary_cat_df["Net Sub-Balance (₹)"] = summary_cat_df["Total_Allocated_Inflow"] - summary_cat_df["Total_Amount_Spent"]
    
    # Clean Column Formatting for Premium Readability
    st.dataframe(
        summary_cat_df.style.format({
            "Total_Allocated_Inflow": "₹{:,.2f}",
            "Total_Amount_Spent": "₹{:,.2f}",
            "Net Sub-Balance (₹)": "₹{:,.2f}"
        }),
        use_container_width=True
    )
else:
    st.info("No active records matched for categories in selected scope scope.")

st.markdown("---")

# --- GLOWING NEON PLOTLY THEME CONFIG ---
neon_layout = dict(
    plot_bgcolor='#0a0c10',   
    paper_bgcolor='#0a0c10',  
    font=dict(color='#e0e6ed', family="Courier New, monospace", size=11),
    margin=dict(l=50, r=30, t=40, b=40),
    showlegend=True,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
    xaxis=dict(showgrid=True, gridcolor='#1a1f2c', gridwidth=1, zeroline=False, linecolor='#2d3748'),
    yaxis=dict(showgrid=True, gridcolor='#1a1f2c', gridwidth=1, zeroline=True, zerolinecolor='#2d3748', linecolor='#2d3748')
)

st.subheader("📊 Premium Visual Analytics Insights")

# ROW 1: USER ANALYTICS
g_col1, g_col2 = st.columns(2)

with g_col1:
    fig1 = go.Figure()
    if not filtered_df.empty:
        daily_spent = filtered_df.groupby("Date")["Amount Spent (₹)"].sum().reset_index()
        fig1.add_trace(go.Scatter(
            x=daily_spent["Date"], y=daily_spent["Amount Spent (₹)"],
            fill='tozeroy', mode='lines+markers',
            line=dict(color='#00ffff', width=2.5, shape='linear'), 
            fillcolor='rgba(0, 255, 255, 0.12)', 
            marker=dict(size=6, color='#00ffff', symbol='circle'),
            name='Expense Flow'
        ))
    fig1.update_layout(title="✨ Daily Expense Flow Timeline", **neon_layout)
    st.plotly_chart(fig1, use_container_width=True)

with g_col2:
    fig2 = go.Figure()
    if not filtered_df.empty:
        user_funds = filtered_df.groupby("Name")[["Imprest Received (₹)", "Amount Spent (₹)"]].sum().reset_index()
        fig2.add_trace(go.Bar(
            x=user_funds["Name"], y=user_funds["Imprest Received (₹)"],
            name="Funds Inflow", marker_color='#38bdf8',
            marker_line=dict(color='#00d2ff', width=1)
        ))
        fig2.add_trace(go.Bar(
            x=user_funds["Name"], y=-user_funds["Amount Spent (₹)"],
            name="Funds Outflow", marker_color='#f43f5e',
            marker_line=dict(color='#ff4a73', width=1)
        ))
    fig2.update_layout(title="⚡ Inflow vs Outflow Structural Matrix", barmode='relative', **neon_layout)
    st.plotly_chart(fig2, use_container_width=True)

# ROW 2: LEDGER CATEGORY ANALYTICS
st.write("")
g_col3, g_col4 = st.columns(2)

with g_col3:
    fig3 = go.Figure()
    if not filtered_df.empty:
        cat_spent = filtered_df.groupby("Expense Category")["Amount Spent (₹)"].sum().reset_index()
        cat_spent = cat_spent[cat_spent["Amount Spent (₹)"] > 0]
        
        fig3.add_trace(go.Pie(
            labels=cat_spent["Expense Category"], 
            values=cat_spent["Amount Spent (₹)"],
            hole=0.4,
            hoverinfo="label+percent+value",
            textinfo="label+percent",
            marker=dict(colors=['#00ffff', '#a855f7', '#38bdf8', '#f43f5e', '#eab308', '#10b981'])
        ))
    fig3.update_layout(title="🔮 Ledger Category Expense Distribution", **neon_layout)
    st.plotly_chart(fig3, use_container_width=True)

with g_col4:
    fig4 = go.Figure()
    if not filtered_df.empty:
        cat_funds = filtered_df.groupby("Expense Category")[["Imprest Received (₹)", "Amount Spent (₹)"]].sum().reset_index()
        
        fig4.add_trace(go.Bar(
            x=cat_funds["Expense Category"], y=cat_funds["Imprest Received (₹)"],
            name="Imprest Inflow", marker_color='#a855f7',
            marker_line=dict(color='#c084fc', width=1)
        ))
        fig4.add_trace(go.Bar(
            x=cat_funds["Expense Category"], y=cat_funds["Amount Spent (₹)"],
            name="Expenses Outflow", marker_color='#eab308',
            marker_line=dict(color='#fde047', width=1)
        ))
    fig4.update_layout(title="🎯 Category-wise Inflow & Outflow Dynamics", barmode='group', **neon_layout)
    st.plotly_chart(fig4, use_container_width=True)

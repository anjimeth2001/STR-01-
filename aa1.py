import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

#st.title("📊 Modified POST Data (Progress Time + GBD/GBS Count + Project + TT_CODE)")

# --- File uploaders ---
post_file = st.file_uploader("Upload POST Excel file", type=["xlsx"], key="post_file")
tubs_file = st.file_uploader("Upload TUBS Excel file", type=["xlsx"], key="tubs_file")
demand_file = st.file_uploader("Upload Demand Excel file (for Project)", type=["xlsx"], key="demand_file")
beam_balance_file = st.file_uploader("Upload BeamBalance Excel file", type=["xlsx"], key="beam_balance_file")

def extract_number(value, decimals):
    if pd.isna(value):
        return None
    m = re.search(r"[-+]?\d*\.\d+|\d+", str(value))
    if not m:
        return None
    return round(float(m.group()), decimals)

# --- GBD/GBS count extractor ---
def count_gbd_gbs(value):
    if pd.isna(value):
        return 0
    values = re.split(r"[ ,;/]+", str(value))
    codes = [v for v in values if "GBD" in v or "GBS" in v]
    return len(codes)

# --- Main processing ---
if post_file is not None:
    # Read POST
    xls_post = pd.ExcelFile(post_file)
    post_sheet = st.selectbox("Select sheet from POST file", xls_post.sheet_names, index=0)
    post_df = pd.read_excel(xls_post, sheet_name=post_sheet)

    # Strip spaces
    post_df.columns = post_df.columns.str.strip()

    # Keep original Progress Time (removed hardcoded overwrite)
    # if "Progress Time" in post_df.columns:
    #     post_df["Progress Time"] = "6/23/2025 4:20:56 AM"

    # Add GBD/GBS count column (3rd column)
    if "Demand" in post_df.columns:
        post_df["GBD_GBS_Count"] = post_df["Demand"].apply(count_gbd_gbs)
        cols = list(post_df.columns)
        cols.insert(2, cols.pop(cols.index("GBD_GBS_Count")))
        post_df = post_df[cols]

    # --- Merge Project from Demand (FIXED: No duplicates) ---
    if demand_file is not None:
        demand_df = pd.read_excel(demand_file)
        demand_df.columns = demand_df.columns.str.strip()

        if "GRE Prod Order" in demand_df.columns and "Project" in demand_df.columns:
            # Remove duplicates from demand data before merging
            demand_unique = demand_df[["GRE Prod Order", "Project"]].drop_duplicates(subset=["GRE Prod Order"])
            
            post_df = post_df.merge(
                demand_unique,
                how="left",
                left_on="Production Order",
                right_on="GRE Prod Order"
            )
            post_df.drop(columns=["GRE Prod Order"], inplace=True, errors="ignore")

            # move Project to 4th col
            if "Project" in post_df.columns:
                cols = list(post_df.columns)
                cols.insert(3, cols.pop(cols.index("Project")))
                post_df = post_df[cols]

    # Clean numeric columns
    for col, dec in [("Beam Issue To PO", 2), ("Weft Issue To PO", 2), ("Action Qty Befor Post", 3)]:
        if col in post_df.columns:
            post_df[col] = post_df[col].apply(lambda x: extract_number(x, dec))

    # --- Merge TT_CODE from TUBS ---
    if tubs_file is not None:
        xls_tubs = pd.ExcelFile(tubs_file)
        tubs_sheet = st.selectbox("Select sheet from TUBS file", xls_tubs.sheet_names, index=0)
        tubs_df = pd.read_excel(xls_tubs, sheet_name=tubs_sheet)
        tubs_df.columns = tubs_df.columns.str.strip()

        tubs_grouped = (
            tubs_df.groupby("PRORDER", dropna=True)["TT_CODE"]
            .apply(lambda s: ",".join(sorted(set(map(str, s)))))
            .reset_index()
        )

        post_df = post_df.merge(tubs_grouped, how="left",
                                left_on="Production Order", right_on="PRORDER")
        post_df.drop(columns=["PRORDER"], inplace=True, errors="ignore")
        post_df["TT_CODE"] = post_df["TT_CODE"].fillna("Not Found")

        # move TT_CODE to last
        if "TT_CODE" in post_df.columns:
            cols = list(post_df.columns)
            cols.append(cols.pop(cols.index("TT_CODE")))
            post_df = post_df[cols]

    # --- Merge IT from BeamBalance ---
    if beam_balance_file is not None:
        xls_beam = pd.ExcelFile(beam_balance_file)
        beam_sheet = st.selectbox("Select sheet from BeamBalance file", xls_beam.sheet_names, index=0)
        beam_df = pd.read_excel(xls_beam, sheet_name=beam_sheet)
        beam_df.columns = beam_df.columns.str.strip()
        
        # Debug information
        st.write("**BeamBalance Debug Info:**")
        st.write(f"BeamBalance columns: {list(beam_df.columns)}")
        st.write(f"POST has Project column: {'Project' in post_df.columns}")
        if 'Project' in post_df.columns:
            st.write(f"Unique Projects in POST: {post_df['Project'].dropna().unique()[:5]}...")  # Show first 5
        if 'Project' in beam_df.columns:
            st.write(f"Unique Projects in BeamBalance: {beam_df['Project'].dropna().unique()[:5]}...")  # Show first 5

        if "Project" in post_df.columns and "Project" in beam_df.columns and "IT" in beam_df.columns:
            # Group by Project and concatenate IT values with comma
            beam_grouped = (
                beam_df.groupby("Project", dropna=True)["IT"]
                .apply(lambda s: ",".join([str(val) for val in s if pd.notna(val)]))
                .reset_index()
            )
            
            st.write(f"BeamBalance grouped data: {len(beam_grouped)} unique projects")
            
            post_df = post_df.merge(beam_grouped, how="left", on="Project")
            post_df["IT"] = post_df["IT"].fillna("Not Found")

            # move IT to second to last column (before TT_CODE if it exists)
            if "IT" in post_df.columns:
                cols = list(post_df.columns)
                it_col = cols.pop(cols.index("IT"))
                if "TT_CODE" in cols:
                    # Insert before TT_CODE
                    tt_idx = cols.index("TT_CODE")
                    cols.insert(tt_idx, it_col)
                else:
                    # Insert at end
                    cols.append(it_col)
                post_df = post_df[cols]
                st.success("✅ IT column added successfully!")
        else:
            missing = []
            if "Project" not in post_df.columns:
                missing.append("Project column missing in POST data (upload Demand file first)")
            if "Project" not in beam_df.columns:
                missing.append("Project column missing in BeamBalance file")
            if "IT" not in beam_df.columns:
                missing.append("IT column missing in BeamBalance file")
            st.error(f"❌ Cannot merge BeamBalance: {', '.join(missing)}")
    else:
        if beam_balance_file is None:
            st.info("💡 Upload BeamBalance file to add IT column")

    # --- Create Filtered IT Dataset (Remove rows with BCY/BYN) ---
    filtered_df = None
    if "IT" in post_df.columns:
        def has_bcy_byn_keywords(it_string):
            """Check if IT string contains BCY or BYN keywords"""
            if pd.isna(it_string) or it_string == "Not Found":
                return False
            
            # Split by comma and check each value for keywords
            it_values = [val.strip() for val in str(it_string).split(",")]
            return any(any(keyword in val.upper() for keyword in ["BCY", "BYN"]) for val in it_values)
        
        # Create filtered dataset by removing rows with BCY/BYN
        filtered_df = post_df.copy()
        
        # Remove rows where IT column contains BCY or BYN
        mask = ~filtered_df["IT"].apply(has_bcy_byn_keywords)
        filtered_df = filtered_df[mask].reset_index(drop=True)

    # --- Preview ---
    def highlight_action(val):
        return "color: red;" if pd.notna(val) else ""

    st.subheader("Modified POST Preview")
    if "Action Qty Befor Post" in post_df.columns:
        st.dataframe(post_df.style.applymap(highlight_action, subset=["Action Qty Befor Post"]),
                     use_container_width=True)
    else:
        st.dataframe(post_df, use_container_width=True)
    
    # Show filtered dataset preview if IT column exists
    if filtered_df is not None:
        st.subheader("Filtered IT Dataset Preview (Rows with BCY/BYN Removed)")
        st.write(f"**Original rows:** {len(post_df)} | **Filtered rows:** {len(filtered_df)} | **Removed:** {len(post_df) - len(filtered_df)}")
        if "Action Qty Befor Post" in filtered_df.columns:
            st.dataframe(filtered_df.style.applymap(highlight_action, subset=["Action Qty Befor Post"]),
                         use_container_width=True)
        else:
            st.dataframe(filtered_df, use_container_width=True)

    # --- Export Excel ---
    output = BytesIO()
    post_df.to_excel(output, index=False, sheet_name="ModifiedPost")
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                            min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin

    def col_idx(col_name: str):
        try:
            return list(post_df.columns).index(col_name) + 1
        except ValueError:
            return None

    beam_idx = col_idx("Beam Issue To PO")
    weft_idx = col_idx("Weft Issue To PO")
    action_idx = col_idx("Action Qty Befor Post")

    if beam_idx:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=beam_idx).number_format = "0.00"
    if weft_idx:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=weft_idx).number_format = "0.00"
    if action_idx:
        for r in range(2, ws.max_row + 1):
            c = ws.cell(row=r, column=action_idx)
            c.number_format = "0.000"
            c.font = Font(color="FF0000")

    final_buf = BytesIO()
    wb.save(final_buf)
    final_buf.seek(0)

    # --- Export Filtered Dataset if IT column exists ---
    filtered_buf = None
    if filtered_df is not None:
        filtered_output = BytesIO()
        filtered_df.to_excel(filtered_output, index=False, sheet_name="FilteredPost")
        filtered_output.seek(0)
        
        # Apply same formatting to filtered dataset
        wb_filtered = load_workbook(filtered_output)
        ws_filtered = wb_filtered.active
        
        for row in ws_filtered.iter_rows(min_row=1, max_row=ws_filtered.max_row,
                                min_col=1, max_col=ws_filtered.max_column):
            for cell in row:
                cell.border = thin
        
        # Apply number formatting to filtered dataset
        def col_idx_filtered(col_name: str):
            try:
                return list(filtered_df.columns).index(col_name) + 1
            except ValueError:
                return None

        beam_idx_f = col_idx_filtered("Beam Issue To PO")
        weft_idx_f = col_idx_filtered("Weft Issue To PO")
        action_idx_f = col_idx_filtered("Action Qty Befor Post")

        if beam_idx_f:
            for r in range(2, ws_filtered.max_row + 1):
                ws_filtered.cell(row=r, column=beam_idx_f).number_format = "0.00"
        if weft_idx_f:
            for r in range(2, ws_filtered.max_row + 1):
                ws_filtered.cell(row=r, column=weft_idx_f).number_format = "0.00"
        if action_idx_f:
            for r in range(2, ws_filtered.max_row + 1):
                c = ws_filtered.cell(row=r, column=action_idx_f)
                c.number_format = "0.000"
                c.font = Font(color="FF0000")
        
        filtered_buf = BytesIO()
        wb_filtered.save(filtered_buf)
        filtered_buf.seek(0)

    # Download buttons
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 Download Modified POST",
            data=final_buf,
            file_name="modified_post.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        if filtered_buf is not None:
            st.download_button(
                label="📥 Download Filtered IT Dataset",
                data=filtered_buf,
                file_name="filtered_post.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

# -----------------------------------------------------
# 🧩 Streamlit Page Setup
# -----------------------------------------------------
st.set_page_config(page_title="Supplier Summary App", layout="wide")
st.title("📊 Supplier Summary App")

# -----------------------------------------------------
# ⚙️ Cached Excel Loader
# -----------------------------------------------------
@st.cache_data
def load_excel(file):
    return pd.read_excel(file, header=None)

# -----------------------------------------------------
# 📂 File Upload
# -----------------------------------------------------
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = load_excel(uploaded_file)

        # --- Automatic header row detection ---
        header_row_index = None
        for i in range(len(df)):
            row_values = [str(x).strip().upper() for x in df.iloc[i].tolist()]
            if "NAME OF SUPPLIERS" in row_values and "TOTAL AMOUNT PAID" in row_values:
                header_row_index = i
                break

        if header_row_index is None:
            st.error("Could not find the header row.")
        else:
            df.columns = df.iloc[header_row_index]
            df = df[header_row_index + 1:].reset_index(drop=True)
            df.columns = [str(c).strip().upper() for c in df.columns]

            # --- Auto column detection ---
            supplier_col = next((c for c in df.columns if c == "NAME OF SUPPLIERS"), None)
            tin_col = next((c for c in df.columns if c == "TIN"), None)
            total_col = next((c for c in df.columns if c == "TOTAL AMOUNT PAID"), None)

            # --- Manual fallback ---
            if not all([supplier_col, total_col]):
                st.warning("Automatic detection failed. Please select columns manually:")
                supplier_col = st.selectbox("Supplier Name column:", df.columns)
                tin_col = st.selectbox("TIN column:", df.columns)
                total_col = st.selectbox("Total Amount Paid column:", df.columns)

            # --- Data Cleaning ---
            df = df[[supplier_col, tin_col, total_col]].copy()
            df.columns = ["Supplier Name", "TIN", "Total Amount Paid"]

            # --- Remove rows where all three columns are empty or None ---
            df = df.dropna(how='all', subset=["Supplier Name", "TIN", "Total Amount Paid"])
            df = df[~((df["Supplier Name"].astype(str).str.strip() == "") &
                      (df["TIN"].astype(str).str.strip() == "") &
                      (df["Total Amount Paid"].astype(str).str.strip() == ""))]

            # Keep Total Amount Paid as numeric where possible; empty stays empty
            df["Total Amount Paid"] = pd.to_numeric(df["Total Amount Paid"], errors="coerce").round(2)

            # -----------------------------------------------------
            # 📊 Display Results
            # -----------------------------------------------------
            if df.empty:
                st.warning("No valid supplier records found.")
            else:
                total_sum = df["Total Amount Paid"].sum(skipna=True)
                total_row = pd.DataFrame({
                    "Supplier Name": ["TOTAL"],
                    "TIN": [""],
                    "Total Amount Paid": [total_sum]
                })
                display_df = pd.concat([df, total_row], ignore_index=True)

                def highlight_total_row(row):
                    if str(row["Supplier Name"]).strip().upper() == "TOTAL":
                        return ["font-weight: bold; font-size: 1.1em; background-color: #f2f2f2;" for _ in row]
                    return [""] * len(row)

                styled_df = display_df.style.format({"Total Amount Paid": "{:,.2f}"}).apply(highlight_total_row, axis=1)

                st.subheader("Supplier Summary")
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                st.markdown("---")
                st.write(f"**Number of Entries (excluding total):** {len(df)}")

                # -----------------------------------------------------
                # 📤 Excel Export
                # -----------------------------------------------------
                def export_to_excel(df):
                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Supplier Summary"

                    for r in dataframe_to_rows(df, index=False, header=True):
                        ws.append(r)

                    for cell in ws[1]:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        if str(row[0].value).strip().upper() == "TOTAL":
                            for cell in row:
                                cell.font = Font(bold=True)
                                cell.fill = gray_fill

                    for i, col in enumerate(ws.columns, start=1):
                        max_length = max(len(str(cell.value)) for cell in col if cell.value is not None)
                        ws.column_dimensions[col[0].column_letter].width = max_length + (4 if i == 1 else 2)

                    wb.save(output)
                    output.seek(0)
                    return output

                excel_file = export_to_excel(display_df)

                st.download_button(
                    label="📥 Export to Excel",
                    data=excel_file,
                    file_name="Supplier_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Please upload an Excel file to start.")

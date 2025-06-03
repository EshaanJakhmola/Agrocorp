import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def map_ref(x):
    s = str(x).strip().upper()
    if ("PULSES" in s) and ("PAYMENT" not in s):
        return "pulses"
    if "SGD_GENERAL" in s:
        return "SGD General"
    if "CAD" in s:
        return "CAD"
    if s.isdigit() and len(s) > 0:
        return "payments"
    if "RB" in s:
        return "Rahul"
    if s.startswith("B"):
        return "Nitin"
    if "PAYMENT" in s:
        return "payments"
    if "NJ" in s:
        return "Nitin"
    if "OILSEEDS" in s:
        return "oilseeds"
    if "WHEAT" in s:
        return "Eur Wheat"
    return ""

def to_excel_bytes_with_pivot(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    if ("Currency" in df.columns) and ("Amount" in df.columns):
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
        pivot = (
            df.pivot_table(
                index=["ref", "Currency"],
                values="Amount",
                aggfunc="sum",
                margins=True,
                margins_name="Grand Total",
                fill_value=0
            )
        )
    else:
        pivot = pd.DataFrame(columns=["ref", "Currency", "Amount"]).set_index(["ref", "Currency"])

    # Try openpyxl first, fallback to xlsxwriter
    for engine in ("openpyxl", "xlsxwriter"):
        try:
            with pd.ExcelWriter(output, engine=engine) as writer:
                df.to_excel(writer, index=False, sheet_name="Processed")
                pivot.to_excel(writer, sheet_name="Pivot")
            break
        except ModuleNotFoundError:
            continue

    return output.getvalue()

def main():
    st.title("Excel Preprocessing & Pivot Generator")
    st.write("Upload an Excel, skip 10 rows, filter to current‐month, add `ref`, show pivot, download both sheets.")

    uploaded_file = st.file_uploader("Choose an Excel file (XLSX or XLS)", type=["xlsx","xls"])
    if uploaded_file is not None:
        try:
            # 1) Skip first 10 rows, use row 11 as header
            df = pd.read_excel(uploaded_file, header=10)

            # 2) Ensure “Value Date” exists, convert to datetime
            if "Value Date" not in df.columns:
                st.error("🚫 Could not find column 'Value Date'.")
                return
            df["Value Date"] = pd.to_datetime(df["Value Date"], errors="coerce")

            # 3) Filter to current month
            now = datetime.now()
            mask = (
                (df["Value Date"].dt.year == now.year) &
                (df["Value Date"].dt.month == now.month)
            )
            df = df.loc[mask].copy()

            # 4) Add “ref” column from “Reference”
            df["ref"] = df.get("Reference", "").apply(map_ref)

            # 5) Display processed data
            st.subheader("Processed Data")
            st.dataframe(df, use_container_width=True)

            # 6) Build and display pivot
            if ("Currency" in df.columns) and ("Amount" in df.columns):
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
                pivot = (
                    df.pivot_table(
                        index=["ref", "Currency"],
                        values="Amount",
                        aggfunc="sum",
                        margins=True,
                        margins_name="Grand Total",
                        fill_value=0
                    )
                )
                st.subheader("Pivot Table (Sum of Amount by ref × Currency)")
                st.dataframe(pivot, use_container_width=True)
            else:
                st.warning("⚠️ Missing 'Currency' or 'Amount' → cannot build pivot.")

            # 7) Download button for combined Excel
            excel_data = to_excel_bytes_with_pivot(df)
            st.download_button(
                label="📥 Download Processed + Pivot Excel",
                data=excel_data,
                file_name="processed_with_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

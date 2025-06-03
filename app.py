import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def map_ref(x):
    """
    Maps the input string to the appropriate 'ref' value, in this priority order:
      1. If "PULSES" (and NOT "PAYMENT") → "pulses"
      2. Elif "SGD_GENERAL" → "SGD General"
      3. Elif contains "CAD" anywhere → "CAD"
      4. Elif the entire string is numeric → "payments"
      5. Elif contains "RB" anywhere → "Rahul"
      6. Elif starts with "B" → "Nitin"
      7. Elif contains "PAYMENT" anywhere → "payments"
      8. Elif contains "NJ" anywhere → "Nitin"
      9. Elif contains "OILSEEDS" anywhere → "oilseeds"
     10. Elif contains "WHEAT" anywhere → "Eur Wheat"
     11. Else → ""
    """
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
    """
    1. Writes `df` to an in-memory Excel file under sheet "Processed".
    2. Builds a pivot (multi-index on ["ref","Currency"], sum of "Amount") under sheet "Pivot".
    3. Returns the raw bytes of that .xlsx file.
    """
    output = BytesIO()

    # Build the pivot table if possible
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

    # First try openpyxl; if missing, fall back to xlsxwriter
    for engine in ("openpyxl", "xlsxwriter"):
        try:
            with pd.ExcelWriter(output, engine=engine) as writer:
                df.to_excel(writer, index=False, sheet_name="Processed")
                pivot.to_excel(writer, sheet_name="Pivot")
            # If we reach here, writing succeeded. Break out of the loop.
            break
        except ModuleNotFoundError:
            continue

    return output.getvalue()

def main():
    st.title("Excel Preprocessing & Pivot Generator")

    st.write("""
    1. Upload an Excel file.  
    2. The app will automatically:
       • Remove the first 10 rows.  
       • Treat row 11 as the header row (so columns like **Value Date**, **Reference**, **Currency**, **Amount** line up correctly).  
       • Filter so that only rows whose **Value Date** (the “Value Date” column) fall in the current month remain.  
       • Add a new column called **ref** at the end, based on **Reference**, using these rules:
         1. “PULSES” (without “PAYMENT”) → “pulses”  
         2. “SGD_GENERAL” → “SGD General”  
         3. Contains “CAD” → “CAD”  
         4. Entirely numeric → “payments”  
         5. Contains “RB” → “Rahul”  
         6. Starts with “B” → “Nitin”  
         7. Contains “PAYMENT” → “payments”  
         8. Contains “NJ” → “Nitin”  
         9. Contains “OILSEEDS” → “oilseeds”  
        10. Contains “WHEAT” → “Eur Wheat”  
        11. Else → (blank)  
    3. You will see:
       • A “Processed” table in-app.  
       • A “Pivot” table (sum of **Amount** by `ref` × `Currency`, with subtotals).  
    4. Click **Download Processed + Pivot Excel** to grab a single `.xlsx` containing exactly two sheets.
    """)

    uploaded_file = st.file_uploader("Choose an Excel file (XLSX or XLS)", type=["xlsx","xls"])
    if uploaded_file is not None:
        try:
            # 1) Skip first 10 rows, use row 11 as header
            df = pd.read_excel(uploaded_file, header=10)

            # 2) Ensure “Value Date” exists, convert to datetime
            if "Value Date" not in df.columns:
                st.error("❌ Could not find a column named 'Value Date'.")
                return
            df["Value Date"] = pd.to_datetime(df["Value Date"], errors="coerce")

            # 3) Filter to current month
            now = datetime.now()
            mask = (
                (df["Value Date"].dt.year == now.year) &
                (df["Value Date"].dt.month == now.month)
            )
            df = df.loc[mask].copy()

            # 4) Add “ref” column
            df["ref"] = df.get("Reference", "").apply(map_ref)

            # 5) Show processed data
            st.subheader("Processed Data (Current Month)")
            st.dataframe(df, use_container_width=True)

            # 6) Build & show pivot
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
                st.warning("⚠️ Missing 'Currency' or 'Amount' column; pivot cannot be built.")

            # 7) Download button
            excel_output = to_excel_bytes_with_pivot(df)
            st.download_button(
                label="📥 Download Processed + Pivot Excel",
                data=excel_output,
                file_name="processed_with_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

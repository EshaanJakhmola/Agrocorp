import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def map_ref(x):
    """
    Priority‐ordered mapping rules (new + old). Given an input string `x`, returns the correct ref:
    
    1. If it contains "AUD_PULSES" → "AUD Pulses"
    2. Elif it contains "INR_PULSES" → "INR Pulses"
    3. Elif it contains "CAD CANOLA" → "OILSEEDS"
    4. Elif it contains "CAD-CANADA" → "CAD CANADA"
    5. Elif it contains "EUR COTTON" → "COTTON"
    6. Elif it contains "EUR WHEAT" → "WHEAT"
    7. Elif it contains "PULSES" (and NOT "PAYMENT") → "pulses"               # old
    8. Elif it contains "SGD" anywhere → "GENERAL"                            # updated from old "SGD_GENERAL"
    9. Elif it contains "CAD" anywhere → "CAD"                                # old generic
   10. Elif the entire string is numeric (only digits) → "payments"          # old
   11. Elif it contains "RB" anywhere → "RAHUL"                               # updated casing
   12. Elif it starts with "B" → "NITIN"                                      # old
   13. Elif it contains "PAYMENT" anywhere → "payments"                      # old
   14. Elif it contains "NJ" anywhere → "NITIN"                               # old
   15. Elif it contains "OILSEEDS" anywhere → "oilseeds"                     # old
   16. Elif it contains "WHEAT" anywhere → "Eur Wheat"                        # old generic
   17. Else → "" (blank string)
    """
    s = str(x).strip().upper()

    # 1. AUD_PULSES → "AUD Pulses"
    if "AUD_PULSES" in s:
        return "AUD Pulses"

    # 2. INR_PULSES → "INR Pulses"
    if "INR_PULSES" in s:
        return "INR Pulses"

    # 3. CAD CANOLA → "OILSEEDS"
    if "CAD CANOLA" in s:
        return "OILSEEDS"

    # 4. CAD-CANADA → "CAD CANADA"
    if "CAD-CANADA" in s:
        return "CAD CANADA"

    # 5. EUR COTTON → "COTTON"
    if "EUR COTTON" in s:
        return "COTTON"

    # 6. EUR WHEAT → "WHEAT"
    if "EUR WHEAT" in s:
        return "WHEAT"

    # 7. PULSES (but not PAYMENT) → "pulses"
    if ("PULSES" in s) and ("PAYMENT" not in s):
        return "pulses"

    # 8. SGD anywhere → "GENERAL"
    if "SGD" in s:
        return "GENERAL"

    # 9. CAD anywhere → "CAD"
    if "CAD-CANADA" in s:
        return "CAD-CANADA"
        
    if "CAD-CANOLA" in s:
        return "OILSEEDS"

    # 10. Entirely numeric → "payments"
    if s.isdigit() and len(s) > 0:
        return "payments"

    # 11. RB anywhere → "RAHUL"
    if s.startswith("RB"):
        return "RAHUL"

    # 12. Starts with B → "NITIN"
    if s.startswith("B"):
        return "NITIN"

    # 13. PAYMENT anywhere → "payments"
    if "PAYMENT" in s:
        return "payments"

    if s.startswith("NJ") or s.startswith("NJ-") or s.startswith("NJ ":
        return "NITIN"

    # 15. OILSEEDS anywhere → "oilseeds"
    if "OILSEEDS" in s:
        return "oilseeds"

    # 16. WHEAT anywhere → "Eur Wheat"
    if "WHEAT" in s:
        return "Eur Wheat"

    # 17. Default: no match → blank
    return ""


def to_excel_bytes_with_pivot(df: pd.DataFrame) -> bytes:
    """
    1. Write `df` to an in‐memory Excel file under sheet "Processed".
    2. Build a pivot (multi‐index on ["ref","Currency"], sum of "Amount") on sheet "Pivot".
    3. Return the raw bytes of that .xlsx file.
    """
    output = BytesIO()

    # Build pivot table if possible
    if ("Currency" in df.columns) and ("Amount" in df.columns):
        # Ensure Amount is numeric
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
        # Empty placeholder pivot
        pivot = pd.DataFrame(columns=["ref", "Currency", "Amount"]).set_index(["ref", "Currency"])

    # Write both DataFrames into one Excel file, two sheets
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
    st.title("Excel Preprocessing & Pivot Generator (Updated Mapping)")

    st.write("""
    1. Upload an Excel file.  
    2. The app will Process it.
    3. You will see:
       • A “Processed” table in‐app.  
       • A “Pivot” table (sum of **Amount** by `ref` × `Currency`, with subtotals).  
    4. Click **Download Processed + Pivot Excel** to grab a single .xlsx with both sheets.
    """)

    uploaded_file = st.file_uploader("Choose an Excel file (XLSX or XLS)", type=["xlsx","xls"])
    if uploaded_file is not None:
        try:
            # ───────────────────────────────────────────────────────────────
            # STEP 1: Read the sheet, using header=10 so that Excel row 11 becomes pandas' header.
            #         This discards rows 1–10 automatically.
            # ───────────────────────────────────────────────────────────────
            df = pd.read_excel(uploaded_file, header=10)

            # ───────────────────────────────────────────────────────────────
            # STEP 2: Ensure “Value Date” exists; convert it to datetime.
            # ───────────────────────────────────────────────────────────────
            if "Value Date" not in df.columns:
                st.error("❌ Could not find a column named 'Value Date'. Please check your file.")
                return

            df["Value Date"] = pd.to_datetime(df["Value Date"], errors="coerce")

            # ───────────────────────────────────────────────────────────────
            # STEP 3: Filter rows so only those with Value Date in the current month/year remain.
            # ───────────────────────────────────────────────────────────────
            

            # Find the previous calendar month (handling year boundary)
            today = datetime.now()
            if today.month == 1:
                prev_month = 12
                prev_year = today.year - 1
            else:
                prev_month = today.month - 1
                prev_year = today.year
            
            # Apply mask for “Value Date” in (prev_year, prev_month)
            mask = (
                (df["Value Date"].dt.year  == prev_year) &
                (df["Value Date"].dt.month == prev_month)
            )
            df = df.loc[mask].copy()
            

            # ───────────────────────────────────────────────────────────────
            # STEP 4: Create the new 'ref' column based on 'Reference' using map_ref(...)
            # ───────────────────────────────────────────────────────────────
            if "Reference" in df.columns:
                df["ref"] = df["Reference"].apply(map_ref)
            else:
                df["ref"] = ""

            # ───────────────────────────────────────────────────────────────
            # STEP 5: Show the “Processed” DataFrame
            # ───────────────────────────────────────────────────────────────
            st.subheader("Processed Data (filtered to current month)")
            st.dataframe(df, use_container_width=True)

            # ───────────────────────────────────────────────────────────────
            # STEP 6: Build & display the pivot table: sum of Amount by ref × Currency
            # ───────────────────────────────────────────────────────────────
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

                st.subheader("Pivot Table: Sum of Amount by ref × Currency")
                st.dataframe(pivot, use_container_width=True)
            else:
                st.warning("⚠️ Cannot build pivot table because 'Currency' or 'Amount' column is missing.")

            # ───────────────────────────────────────────────────────────────
            # STEP 7: Offer a single Excel download with BOTH sheets (“Processed” + “Pivot”)
            # ───────────────────────────────────────────────────────────────
            excel_bytes = to_excel_bytes_with_pivot(df)
            st.download_button(
                label="📥 Download Processed + Pivot Excel",
                data=excel_bytes,
                file_name="processed_with_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred while processing: {e}")

if __name__ == "__main__":
    main()

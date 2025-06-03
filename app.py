import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def map_ref(x):
    """
    Maps the input string to the appropriate 'ref' value, in this priority order:
      1. If it contains "PULSES" AND does NOT contain "PAYMENT" → "pulses"
      2. Elif it contains "SGD_GENERAL" → "SGD General"
      3. Elif it contains "CAD" anywhere → "CAD"
      4. Elif the entire string is numeric (only digits) → "payments"
      5. Elif it contains "RB" anywhere → "Rahul"
      6. Elif it starts with "B" → "Nitin"
      7. Elif it contains "PAYMENT" anywhere → "payments"
      8. Elif it contains "NJ" anywhere → "Nitin"
      9. Elif it contains "OILSEEDS" anywhere → "oilseeds"
     10. Elif it contains "WHEAT" anywhere → "Eur Wheat"
     11. Else → ""
    """
    s = str(x).strip().upper()

    # 1. "PULSES" without "PAYMENT"
    if ("PULSES" in s) and ("PAYMENT" not in s):
        return "pulses"
    # 2. "SGD_GENERAL"
    if "SGD_GENERAL" or "SGD GENERAL" in s:
        return "SGD General"
    # 3. "CAD"
    if "CAD" in s:
        return "CAD"
    # 4. Entirely numeric → "payments"
    if s.isdigit() and len(s) > 0:
        return "payments"
    # 5. "RB"
    if "RB" in s:
        return "Rahul"
    # 6. Starts with "B"
    if s.startswith("B"):
        return "Nitin"
    # 7. "PAYMENT"
    if "PAYMENT" in s:
        return "payments"
    # 8. "NJ"
    if "NJ" in s:
        return "Nitin"
    # 9. "OILSEEDS"
    if "OILSEEDS" in s:
        return "oilseeds"
    # 10. "WHEAT"
    if "WHEAT" in s:
        return "Eur Wheat"
    # 11. Default
    return ""

def to_excel_bytes_with_pivot(df: pd.DataFrame) -> bytes:
    """
    Given the processed DataFrame `df`, this function:
      1. Writes `df` to an in-memory Excel file under the sheet name "Processed".
      2. Builds a pivot table with a multi‐index on ["ref", "Currency"] and a single "Amount" column,
         then writes that to a second sheet named "Pivot".
    Returns the raw bytes of the Excel file for downloading.
    """
    output = BytesIO()

    # If "Currency" or "Amount" are missing, pivot will be empty
    if ("Currency" in df.columns) and ("Amount" in df.columns):
        # Ensure Amount is numeric
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
        # Create a multi‐index pivot table: index = ["ref", "Currency"], values = sum of "Amount"
        pivot = (
            df
            .pivot_table(
                index=["ref", "Currency"],
                values="Amount",
                aggfunc="sum",
                margins=True,
                margins_name="Grand Total",
                fill_value=0
            )
        )
    else:
        # Create an empty placeholder pivot
        pivot = pd.DataFrame(columns=["ref", "Currency", "Amount"]).set_index(["ref", "Currency"])

    # Write both DataFrames to a single Excel file, two sheets
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: Processed data
        df.to_excel(writer, index=False, sheet_name="Processed")
        # Sheet 2: Pivot table
        pivot.to_excel(writer, sheet_name="Pivot")
        # No explicit writer.save() needed—context manager handles it

    return output.getvalue()

def main():
    st.title("Excel Preprocessing & Pivot Generator")

    st.write("""
    ### Instructions
    1. Upload an Excel file.
    2. The app will automatically:
       - Remove the first 10 rows.
       - Treat row 11 as the header row (so columns like **Value Date**, **Reference**, **Currency**, **Amount** appear correctly).
       - Filter so that only rows whose **Value Date** (the column named "Value Date") fall in the current month remain.
       - Add a new column called **ref** at the end, based on the text in the **Reference** column (using the rules below).
    3. Once processing is complete, you can:
       - Preview the “Processed” sheet.
       - Preview the “Pivot” sheet (sum of **Amount** for each **ref** × **Currency**, with subtotals).
       - Download a single Excel file containing both sheets.
    ---
    **Mapping rules for `ref`:**  
      1. If **Reference** contains "PULSES" **and not** "PAYMENT" → `"pulses"`  
      2. Elif **Reference** contains "SGD_GENERAL" → `"SGD General"`  
      3. Elif **Reference** contains "CAD" anywhere → `"CAD"`  
      4. Elif **Reference** is entirely numeric → `"payments"`  
      5. Elif **Reference** contains "RB" anywhere → `"Rahul"`  
      6. Elif **Reference** starts with "B" → `"Nitin"`  
      7. Elif **Reference** contains "PAYMENT" anywhere → `"payments"`  
      8. Elif **Reference** contains "NJ" anywhere → `"Nitin"`  
      9. Elif **Reference** contains "OILSEEDS" anywhere → `"oilseeds"`  
     10. Elif **Reference** contains "WHEAT" anywhere → `"Eur Wheat"`  
     11. Else → `""` (empty string)
    """)

    uploaded_file = st.file_uploader("Choose an Excel file (XLSX or XLS)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            # ─────────────────────────────────────────────────────────────────
            # STEP 1: Read the sheet, using header=10 so that Excel row 11 becomes pandas' header.
            #         This automatically discards rows 1–10.
            # ─────────────────────────────────────────────────────────────────
            df = pd.read_excel(uploaded_file, header=10)

            # ─────────────────────────────────────────────────────────────────
            # STEP 2: Verify the “Value Date” column exists; convert it to datetime.
            # ─────────────────────────────────────────────────────────────────
            if "Value Date" not in df.columns:
                st.error("❌ Could not find a column named 'Value Date'. Please check your file.")
                return

            df["Value Date"] = pd.to_datetime(df["Value Date"], errors="coerce")

            # ─────────────────────────────────────────────────────────────────
            # STEP 3: Filter rows so only those with Value Date in the current month/year remain.
            # ─────────────────────────────────────────────────────────────────
            now = datetime.now()
            mask = (
                (df["Value Date"].dt.year  == now.year) &
                (df["Value Date"].dt.month == now.month)
            )
            df = df.loc[mask].copy()

            # ─────────────────────────────────────────────────────────────────
            # STEP 4: Create the new 'ref' column based on 'Reference', using map_ref(...)
            # ─────────────────────────────────────────────────────────────────
            if "Reference" in df.columns:
                df["ref"] = df["Reference"].apply(map_ref)
            else:
                df["ref"] = ""

            # ─────────────────────────────────────────────────────────────────
            # STEP 5: Show the “Processed” DataFrame
            # ─────────────────────────────────────────────────────────────────
            st.subheader("Processed Data (filtered to current month)")
            st.dataframe(df, use_container_width=True)

            # ─────────────────────────────────────────────────────────────────
            # STEP 6: Build & display the pivot table: sum of Amount by ref × Currency
            # ─────────────────────────────────────────────────────────────────
            if ("Currency" in df.columns) and ("Amount" in df.columns):
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)

                # Create pivot with subtotals (margins=True)
                pivot = (
                    df
                    .pivot_table(
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

            # ─────────────────────────────────────────────────────────────────
            # STEP 7: Offer a single Excel download with BOTH sheets ("Processed" + "Pivot")
            # ─────────────────────────────────────────────────────────────────
            excel_bytes = to_excel_bytes_with_pivot(df)
            st.download_button(
                label="📥 Download Excel (Processed + Pivot)",
                data=excel_bytes,
                file_name="processed_with_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred while processing: {e}")

if __name__ == "__main__":
    main()

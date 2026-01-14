import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from copy import copy

def copy_sheet(source_sheet, target_workbook, new_title):
    """
    Helper to copy a sheet within the same workbook or to a new one.
    For openpyxl, copying a sheet is simplest by using the built-in copy_worksheet.
    """
    target_sheet = target_workbook.copy_worksheet(source_sheet)
    target_sheet.title = new_title
    return target_sheet

def process_excel(input_file, template_file):
    # Load the template workbook
    # We use openpyxl to preserve formatting
    template_wb = openpyxl.load_workbook(template_file)
    template_sheet = template_wb.active
    
    # Load the input workbook
    # data_only=True to get values, not formulas
    input_wb = openpyxl.load_workbook(input_file, data_only=True)
    
    # We need a base template to copy from. 
    # We will keep the original template sheet as a "master" and copy it for each input sheet.
    # Or, we can fill the first one and copy for the rest.
    # Strategy: Keep the original 'Sheet1' (or active) as the master template.
    # For each input sheet, create a new copy of the master.
    # Finally, delete the master if it wasn't used (or keep it hidden).
    
    # Actually, the user said "Sheet 1 is first sheet, 2 is second sheet".
    # So we should probably map Input Sheet 1 -> Output Sheet 1, etc.
    
    output_sheets = []
    
    # Get all sheet names from input
    input_sheet_names = input_wb.sheetnames
    
    # We'll use the first sheet of the template as the base
    base_template_sheet = template_wb.active
    
    for i, sheet_name in enumerate(input_sheet_names):
        input_sheet = input_wb[sheet_name]
        
        # Prepare the target sheet
        if i == 0:
            # For the first sheet, use the existing active sheet
            target_sheet = base_template_sheet
            target_sheet.title = sheet_name # Rename to match input? Or keep "1", "2"? User said "Sheet 1=1". Let's use input name.
        else:
            # For subsequent sheets, copy the base template
            target_sheet = template_wb.copy_worksheet(base_template_sheet)
            target_sheet.title = sheet_name
        
        # Transfer data
        # Input: C2 to C48 (Rows 2 to 48, Column 3)
        # Output: C3 to C49 (Rows 3 to 49, Column 3)
        # 1-based index: C is column 3.
        
        # Loop 47 times (2 to 48 is 47 items)
        for row_offset in range(47):
            input_row = 2 + row_offset
            output_row = 3 + row_offset
            
            # Read value
            value = input_sheet.cell(row=input_row, column=3).value
            
            # Write value
            target_sheet.cell(row=output_row, column=3).value = value
            
    # Save result to memory
    output = BytesIO()
    template_wb.save(output)
    output.seek(0)
    return output

st.set_page_config(page_title="Excel Auto-Filler", layout="wide")

st.title("Excel Data Automation App")
st.markdown("""
**Instructions:**
1. Upload the **Input Data File** (contains the data in C2:C48).
2. Upload the **Template File** (the empty form).
3. The app will create a new file with a sheet for each sheet in your Input File, filling the data into the Template.
""")

col1, col2 = st.columns(2)

with col1:
    input_file = st.file_uploader("1. Upload Input File (Data)", type=['xlsx'])

with col2:
    template_file = st.file_uploader("2. Upload Template File (Form)", type=['xlsx'])

if input_file and template_file:
    if st.button("Process Excel Files"):
        try:
            with st.spinner("Processing..."):
                result_file = process_excel(input_file, template_file)
                
            st.success("Processing Complete!")
            
            st.download_button(
                label="Download Result Excel",
                data=result_file,
                file_name="processed_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")

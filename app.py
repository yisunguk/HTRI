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
    # Mapping Dictionary (Label -> Cell Address)
    # Based on the user's provided image
    MAPPING = {
        "Service of Unit": "I8",
        "Item No.": "AV8",
        "Size": "E9",
        "Type": "Y9",
        "Surf/Unit (Gross/Eff)": "K10",
        "Fluid Name": "T13",
        "Fluid Quantity, Total": "T14",
        "Temperature (In/Out)": "T20",
        "Inlet Pressure": "AB28",
        "Velocity": "AB29",
        "Pressure Drop, Allow/Calc": "T30",
        "Heat Exchanged": "M32",
        "MTD (Corrected)": "BB32",
        "Transfer Rate, Service": "M33",
        "Clean": "AH33",
        "Actual": "BB33",
        "Design/Test Pressure": "T36",
        "Design Temperature": "T37",
        "No Passes per Shell": "T38",
        "Tube No.": "F43",
        "OD": "N43",
        "Thk(Avg)": "AC43",
        "Length": "AR43",
        "Pitch": "BG43",
        "Tube Type": "F44",
        "Material": "AH44",
        "Tube pattern": "BM44",
        "Shell": "E45",
        "ID": "U45",
        # "OD": "AC45", # Duplicate Key "OD". Need to handle.
        # "Shell Cover": "AU45",
        "Channel or Bonnet": "K46",
        "Channel Cover": "AU46",
        "Tubesheet-Stationary": "K47",
        "Tubesheet-Floating": "AW47",
        "Floating Head Cover": "K48",
        "Impingement Plate": "AW48",
        "Baffles-Cross": "H49",
        # "Type": "V49", # Duplicate Key "Type".
        "%Cut (Diam)": "AM49",
        "Spacing(c/c)": "AX49",
        "Inlet": "BG49",
        "TEMA Class": "BA57"
    }
    
    # Handling Duplicate Keys manually by checking context or using specific labels if possible.
    # Since we iterate Template Rows, we can check the Label.
    # But "OD" appears for Tube (N43) and Shell (AC45).
    # "Type" appears for Unit (Y9) and Baffles (V49).
    # We need a smarter lookup or a list of values.
    
    # Enhanced Mapping with Context or List
    # If a label appears multiple times, we can use a list of addresses to pop from.
    # Or we can just map specific unique strings if the template has them.
    # Assuming the Template has exact strings.
    # Let's use a list for duplicates.
    
    MAPPING_LIST = {
        "OD": ["N43", "AC45"], # 1st: Tube, 2nd: Shell
        "Type": ["Y9", "V49"], # 1st: Unit, 2nd: Baffles
    }

    template_wb = openpyxl.load_workbook(template_file)
    template_sheet = template_wb.active
    
    input_wb = openpyxl.load_workbook(input_file, data_only=True)
    input_sheet_names = input_wb.sheetnames
    
    # Iterate through each input sheet
    for i, sheet_name in enumerate(input_sheet_names):
        input_sheet = input_wb[sheet_name]
        
        # Determine Target Column in Template
        # Sheet 1 -> Column C (3)
        # Sheet 2 -> Column D (4)
        target_col_idx = 3 + i
        
        # Write Sheet Name/Index at the top (Row 1)
        template_sheet.cell(row=1, column=target_col_idx).value = i + 1
        
        # Track usage of duplicates
        duplicate_counters = {key: 0 for key in MAPPING_LIST}
        
        # Iterate through Template Rows (e.g., 2 to 100)
        # We look at Column A for Labels
        for row_idx in range(2, 100):
            label_cell = template_sheet.cell(row=row_idx, column=1)
            label = label_cell.value
            
            if label:
                label = str(label).strip()
                
                cell_address = None
                
                # Check Duplicates first
                if label in MAPPING_LIST:
                    counter = duplicate_counters[label]
                    if counter < len(MAPPING_LIST[label]):
                        cell_address = MAPPING_LIST[label][counter]
                        duplicate_counters[label] += 1
                
                # Check Normal Mapping
                elif label in MAPPING:
                    cell_address = MAPPING[label]
                
                if cell_address:
                    # Read from Input
                    try:
                        value = input_sheet[cell_address].value
                        # Write to Template
                        template_sheet.cell(row=row_idx, column=target_col_idx).value = value
                    except Exception as e:
                        print(f"Error reading {cell_address}: {e}")

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

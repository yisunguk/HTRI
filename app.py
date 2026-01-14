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
    # Advanced Mapping Rules
    # Key: Label in Template
    # Value: List of rules for each occurrence of the label.
    #   - Single String: Cell Address (e.g., "I8")
    #   - List of Strings: Multiple Cells to merge (e.g., ["E9", "M9", "N9"])
    
    MAPPING_RULES = {
        "Service of Unit": ["I8"],
        "Item No.": ["AV8"],
        "Size": [ ["E9", "M9", "N9"] ], # Merge 3 cells
        "Type": ["Y9", "V49"], # 1st occurrence (Unit), 2nd (Baffles)
        "Surf/Unit (Gross/Eff)": [ ["K10", "O10", "P10"] ],
        
        # Shell Side (1) and Tube Side (2)
        "Fluid Name": ["T13", "AR13"], 
        "Fluid Quantity, Total": ["T14", "AR14"],
        "Temperature (In/Out)": [ ["T20", "AF20"], ["AR20", "BD20"] ], # Merge In/Out
        "Inlet Pressure": ["AB28", "AZ28"],
        "Velocity": ["AB29", "AZ29"],
        "Pressure Drop, Allow/Calc": [ ["T30", "AF30"], ["AR30", "BD30"] ],
        
        "Heat Exchanged": ["M32"],
        "MTD (Corrected)": ["BB32"],
        "Transfer Rate, Service": ["M33"],
        "Clean": ["AH33"],
        "Actual": ["BB33"],
        
        "Design/Test Pressure": ["T36"],
        "Design Temperature": ["T37"],
        "No Passes per Shell": ["T38"],
        
        "Tube No.": ["F43"],
        "OD": ["N43", "AC45"], # 1st (Tube), 2nd (Shell)
        "Thk(Avg)": ["AC43"],
        "Length": ["AR43"],
        "Pitch": ["BG43"],
        "Tube Type": ["F44"],
        "Material": ["AH44"],
        "Tube pattern": ["BM44"],
        
        "Shell": ["E45"],
        "ID": ["U45"],
        "Shell Cover": ["AU45"],
        "Channel or Bonnet": ["K46"],
        "Channel Cover": ["AU46"],
        "Tubesheet-Stationary": ["K47"],
        "Tubesheet-Floating": ["AW47"],
        "Floating Head Cover": ["K48"],
        "Impingement Plate": ["AW48"],
        "Baffles-Cross": ["H49"],
        
        "%Cut (Diam)": ["AM49"],
        "Spacing(c/c)": ["AX49"],
        "Inlet": ["BG49"],
        "TEMA Class": ["BA57"]
    }

    def get_cell_value(sheet, address_rule):
        """
        Retrieves value from sheet based on rule.
        rule can be "A1" or ["A1", "B1"].
        """
        if isinstance(address_rule, list):
            # Merge values
            values = []
            for addr in address_rule:
                val = sheet[addr].value
                if val is not None:
                    values.append(str(val))
            return " ".join(values) # Join with space
        else:
            # Single cell
            return sheet[address_rule].value

    template_wb = openpyxl.load_workbook(template_file)
    template_sheet = template_wb.active
    
    input_wb = openpyxl.load_workbook(input_file, data_only=True)
    input_sheet_names = input_wb.sheetnames
    
    # Iterate through each input sheet
    for i, sheet_name in enumerate(input_sheet_names):
        input_sheet = input_wb[sheet_name]
        
        # Determine Target Column in Template
        # Sheet 1 -> Column C (3)
        target_col_idx = 3 + i
        
        # Write Sheet Name/Index at the top (Row 1)
        template_sheet.cell(row=1, column=target_col_idx).value = i + 1
        
        # Track usage of duplicates for this column
        # We reset this for each input sheet (column)
        duplicate_counters = {key: 0 for key in MAPPING_RULES}
        
        # Iterate through Template Rows
        for row_idx in range(2, 150): # Increased range just in case
            label_cell = template_sheet.cell(row=row_idx, column=1)
            label = label_cell.value
            
            if label:
                label = str(label).strip()
                
                if label in MAPPING_RULES:
                    rules = MAPPING_RULES[label]
                    counter = duplicate_counters[label]
                    
                    if counter < len(rules):
                        rule = rules[counter]
                        
                        try:
                            value = get_cell_value(input_sheet, rule)
                            template_sheet.cell(row=row_idx, column=target_col_idx).value = value
                        except Exception as e:
                            print(f"Error processing {label} at row {row_idx}: {e}")
                        
                        duplicate_counters[label] += 1

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

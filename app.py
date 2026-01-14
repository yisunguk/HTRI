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
    #   - String: Single Cell Address (e.g., "I8")
    #   - List: Multiple Cells to MERGE into one (e.g., ["E9", "M9", "N9"])
    #   - Dict: Special Action (e.g., {"action": "vertical", "cells": ["T20", "AF20"]})
    
    MAPPING_RULES = {
        "Service of Unit": ["I8"],
        "Item No.": ["AV8"],
        "Size": [ ["E9", "M9", "N9"] ], # Merge 3 cells
        "Type": ["Y9", "V49"], # 1st occurrence (Unit), 2nd (Baffles)
        "Surf/Unit (Gross/Eff)": [ ["K10", "O10", "P10"] ],
        
        # Shell Side (1) and Tube Side (2)
        "Fluid Name": ["T13", "AR13"], 
        "Fluid Quantity, Total": ["T14", "AR14"],
        
        # Temperature: Vertical Split (In / Out)
        "Temperature (In/Out)": [ 
            {"action": "vertical", "cells": ["T20", "AF20"]}, # Shell Side
            {"action": "vertical", "cells": ["AR20", "BD20"]} # Tube Side
        ],
        
        "Inlet Pressure": ["AB28", "AZ28"],
        "Velocity": ["AB29", "AZ29"],
        
        # Pressure Drop: Vertical Split (Allow / Calc)
        "Pressure Drop, Allow/Calc": [ 
            {"action": "vertical", "cells": ["T30", "AF30"]}, # Shell Side
            {"action": "vertical", "cells": ["AR30", "BD30"]} # Tube Side
        ],
        
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

    def get_cell_value(sheet, addr):
        val = sheet[addr].value
        return str(val) if val is not None else ""

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
        duplicate_counters = {key: 0 for key in MAPPING_RULES}
        
        # Iterate through Template Rows
        for row_idx in range(2, 150):
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
                            if isinstance(rule, dict) and rule.get("action") == "vertical":
                                # Vertical Split
                                cells = rule["cells"]
                                # Write first value to current row
                                val1 = get_cell_value(input_sheet, cells[0])
                                template_sheet.cell(row=row_idx, column=target_col_idx).value = val1
                                
                                # Write second value to next row (row_idx + 1)
                                if len(cells) > 1:
                                    val2 = get_cell_value(input_sheet, cells[1])
                                    template_sheet.cell(row=row_idx + 1, column=target_col_idx).value = val2
                                    
                            elif isinstance(rule, list):
                                # Merge
                                values = [get_cell_value(input_sheet, addr) for addr in rule]
                                merged_value = " ".join([v for v in values if v])
                                template_sheet.cell(row=row_idx, column=target_col_idx).value = merged_value
                                
                            else:
                                # Single String
                                val = get_cell_value(input_sheet, rule)
                                template_sheet.cell(row=row_idx, column=target_col_idx).value = val
                                
                        except Exception as e:
                            print(f"Error processing {label} at row {row_idx}: {e}")
                        
                        duplicate_counters[label] += 1

    output = BytesIO()
    template_wb.save(output)
    output.seek(0)
    return output

import os

st.set_page_config(page_title="Excel Auto-Filler", layout="wide")

st.title("Excel Data Automation App")
st.markdown("""
**Instructions:**
1. Upload the **Input Data File** (contains the data).
2. The app will use the default **Template File** (`template.xlsx`) if available. Otherwise, please upload one.
""")

col1, col2 = st.columns(2)

with col1:
    input_file = st.file_uploader("1. Upload Input File (Data)", type=['xlsx'])

with col2:
    if os.path.exists("template.xlsx"):
        template_file = "template.xlsx"
        st.success("Using default 'template.xlsx' from repository.")
    else:
        template_file = st.file_uploader("2. Upload Template File (Form)", type=['xlsx'])
        st.warning("Default 'template.xlsx' not found in repository. Please upload it to GitHub if you want to fix it.")

# Debugging: Show available files
with st.sidebar.expander("Debug: File System"):
    st.write("Current Directory:", os.getcwd())
    st.write("Files:", os.listdir())

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

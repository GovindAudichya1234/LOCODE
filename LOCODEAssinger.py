import streamlit as st
import pandas as pd
import numpy as np
import re
from fuzzywuzzy import process

def load_excel(file):
    """Load the uploaded Excel file."""
    return pd.ExcelFile(file)  

def preprocess_text(text):
    """Normalize text by removing extra spaces and making it lowercase."""
    return re.sub(r'\s+', ' ', str(text).strip()).lower()

def match_lo_code(learning_objective, lo_mapping):
    """Find the closest match for the learning objective in the TTCA_FDT_LO mapping."""
    match, score = process.extractOne(preprocess_text(learning_objective), lo_mapping.keys())
    if score > 80:  # Threshold to ensure relevant matching
        return lo_mapping[match]
    return np.nan

def process_file(uploaded_file, learning_objectives, lo_mapping, file_name, question_range):
    """Process the uploaded FDT_AMT file."""
    df = pd.read_excel(uploaded_file,header=1)
    
    # Identify LO columns dynamically
    lo_columns = [col for col in df.columns if col.startswith("LO")]
    
    # Add new LO Code columns after each LO column
    for col in lo_columns:
        df.insert(df.columns.get_loc(col) + 1, col + "Code", np.nan)
    
    # Process the relevant question range
    start, end = map(int, question_range.split('-'))
    df_subset = df.iloc[start-3:end].copy()
    
    # Process each row in the relevant range
    for index, row in df_subset.iterrows():
        for col in lo_columns:
            lo_value = row[col]
            if pd.notna(lo_value):
                lo_value = preprocess_text(lo_value)
                matched_code = match_lo_code(lo_value, lo_mapping)
                df.at[index, col + "Code"] = matched_code
    
    return df

# Streamlit UI
st.title("Learning Objective Matching Tool")

st.write("Instructions")
st.write("1. The Main AMT Sheet should be first in excel")
st.write("2. The columns should have the exact name mentioned here :  Level,Skill,Topic,LO1,LO2,LO3,LO4,Question Type,Question Statement,Complexity Level,Difficulty Level Tag (Auto Populated- Do not Edit),Correct Answer (option keyin capital e.g. A ),Answer Explanation,Bloom's Taxonomy,optionKey1 ,optionKey2 ,optionKey3 ,optionKey4 ,optionValue1 ,optionValue2 ,optionValue3 ,optionValue4")
st.write("3. Make sure there is no space in the column name")
st.write("4. Select the LO Bank (Foundational or Preparatory)")
st.write("5. Upload the AMT Excel")
st.write("6. Write the file name as instructed Level_CourseNo_ModuleNo_SubUnitNo (E.g: FDT_7_2_10)")
st.write("7. Select the row range or data range basically the row number from which your question starts and ends (for E.g: 3-38)")
st.write("8. Let it process and check the processed data table appearing on the screen")
st.write("9. Download the file")

# Dropdown for LO selection
lo_option = st.selectbox("Select LO Type", ["Foundational LO", "Preparatory LO"])

# Set file name based on selection
ttca_file = "TTCA_FDT_LO_Final.xlsx" if lo_option == "Foundational LO" else "TTCA_Prep_LO_Final.xlsx"
template_file = "QTemplates.xlsx"

# Upload AMT or question file
uploaded_file = st.file_uploader("Upload the FDT_AMT Excel file", type=["xlsx"])

if uploaded_file:
    # Load TTCA_FDT_LO data
    xls_ttca = load_excel(ttca_file)
    sheet1 = pd.read_excel(xls_ttca, sheet_name="PRINFO")
    lo_code_sheet = pd.read_excel(xls_ttca, sheet_name="LOCODE")
    template_df = pd.read_excel(template_file)
    
    # User Inputs
    file_name = st.text_input("Enter File Name (e.g., FDT_1_2_3)")
    question_range = st.text_input("Enter Question Range (e.g., 3-38)")
    
    if file_name and question_range:
        # Filter relevant LOs for the provided file name
        relevant_los = sheet1[sheet1.iloc[:, 0] == file_name].iloc[:, 1].dropna().tolist()
        
        # Create a mapping of LO descriptions to codes
        lo_mapping = {preprocess_text(desc).rstrip('.'): code for code, desc in zip(lo_code_sheet.iloc[:, 0], lo_code_sheet.iloc[:, 1])}
        
        # Process the FDT_AMT file 
        processed_df = process_file(uploaded_file, relevant_los, lo_mapping, file_name, question_range)
        
        # Display and download processed file
        st.write("### Processed Data Preview")
        st.dataframe(processed_df)
        
        column_mapping = {
            "Level": "level",
            "Skill": "skill",
            "Topic": "topic",
            "LO1Code": "lo1",
            "LO2Code": "lo2",
            "LO3Code": "lo3",
            "LO4Code": "lo4",
            "Question Type": "questionType",
            "Question Statement": "questionStatement",
            "Complexity Level": "complexityLevel",
            "Difficulty Level Tag (Auto Populated- Do not Edit)": "difficultyLevel",
            "Correct Answer (option keyin capital e.g. A )": "correctAnswer",
            "Answer Explanation": "answerExplanation",
            "Bloom's Taxonomy": "bloomsTaxonomy",
            "optionKey1": "optionKey1",
            "optionKey2": "optionKey2",
            "optionKey3": "optionKey3",
            "optionKey4": "optionKey4",
            "optionValue1": "optionValue1",
            "optionValue2": "optionValue2",
            "optionValue3": "optionValue3",
            "optionValue4": "optionValue4"
        }
        
        # Insert processed data into the template with mapped column names
        for processed_col, template_col in column_mapping.items():
            if processed_col in processed_df.columns and template_col in template_df.columns:
                template_df[template_col] = processed_df[processed_col]
        
        # Fill remaining empty columns with "NA"
        template_df.fillna("NA", inplace=True)
        
        # Save and allow download of processed template file
        output_file = file_name + "_Processed.xlsx"
        template_df.to_excel(output_file, index=False)
        st.download_button("Download Updated Template File", data=open(output_file, "rb"), file_name=output_file)

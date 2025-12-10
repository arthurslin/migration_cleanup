import pandas as pd
import numpy as np
import os

folder = "reports"
masrev = "1ERP MAS Rev CLEAN"
mlprev = "1ERP MLP Rev CLEAN"
nusrev = "Legacy NUS CLEAN"
all_revisions = (masrev,mlprev)

report_path = os.path.join(os.path.dirname(__file__), folder)

def read_all_excel_sheets(folder_path):
    sheets = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                print("Reading sheet:", sheet)
                reportdf = pd.read_excel(xls, sheet_name=sheet)
                if not reportdf.empty:
                    sheets.append(reportdf)
    return sheets

all_reports = read_all_excel_sheets(report_path)

def clean_revision_data(reports):
    res = []

    for df in reports:

        if df.empty or len(df.columns) == 0:
            break

        if masrev in df.columns and mlprev in df.columns and nusrev in df.columns:
            for index, row in df.iterrows():
                # Process the values from both columns
                for rev in tuple(all_revisions):
                    new_col_name = f"{rev}_NEW"
                    if new_col_name not in df.columns:
                        df[new_col_name] = ""
                        
                    rev_value = row[rev]
                    nus_value = row[nusrev]

                    print(f"Processing row {index}: {rev}={rev_value}, {nusrev}={nus_value}")
                    if pd.isna(rev_value) or rev_value == "--":
                        value = nus_value
                        print(f"Using {nusrev} value: {value}")
                    else:
                        value = rev_value
                    
                    if pd.notna(value) and len(str(value)) == 1:
                        value = "-" + str(value)
                    
                    df.at[index, new_col_name] = value
                
            res.append(df)
    return res

final_dfs = clean_revision_data(all_reports)  
for final_df in final_dfs:
    final_df.to_excel("cleaned_revisions.xlsx", index=False)



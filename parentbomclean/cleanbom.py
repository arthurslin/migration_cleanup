import pandas as pd
import numpy as np
import os

folder = "reports"
exportstr = "Export"
sfile = "SOURCE_FILE"

report_path = os.path.join(os.path.dirname(__file__), folder)

def read_all_excel_sheets(folder_path):
    sheets = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                if exportstr in sheet:
                    print("Reading sheet:", sheet)
                    reportdf = pd.read_excel(xls, sheet_name=sheet)
                    if not reportdf.empty:
                        reportdf.insert(0, sfile, filename[:3])
                        sheets.append(reportdf)
    return sheets

all_reports = read_all_excel_sheets(report_path)

parent = "PARENT_PART"
countdict = []

for report in all_reports:
    parent_counts = report[parent].value_counts().to_dict()
    source_file = report[sfile].iloc[0]
    # print(f"{source_file} - Parent Part Counts:", len(parent_counts))
    countdict += [(parent_counts, source_file)]

print(len(countdict))
if len(countdict) == 2:
    dict1, source1 = countdict[0]
    dict2, source2 = countdict[1]

    common_keys = set(list(dict1.keys()) + list(dict2.keys()))

    dict1_filtered = {k: dict1.get(k, 0) for k in common_keys}
    dict2_filtered = {k: dict2.get(k, 0) for k in common_keys}
    
    # print(f"{source1} filtered:", dict1_filtered)
    # print(f"{source2} filtered:", dict2_filtered)
    comparison_df = pd.DataFrame({
        source1: dict1_filtered,
        source2: dict2_filtered
    }).fillna(0).astype(int)

    print(comparison_df)
    comparison_df.to_excel("parent_part_comparison.xlsx")
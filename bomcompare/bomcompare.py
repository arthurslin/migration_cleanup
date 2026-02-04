import pandas as pd
import numpy as np
import os

folder = "reports"
items = "items"
prefix = "ORIGIN_FILE"

report_path = os.path.join(os.path.dirname(__file__), folder)
items_path = os.path.join(os.path.dirname(__file__), items)

def read_all_excel_sheets(folder_path):
    sheets = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            xls = pd.ExcelFile(file_path)
            file_prefix = filename[:3]
            for sheet in xls.sheet_names:
                    reportdf = pd.read_excel(xls, sheet_name=sheet)
                    reportdf[prefix] = file_prefix
                    sheets.append(reportdf)
    return sheets

child_part = "CHILD_PART_NUMBER"
parent_part = "PARENT_PART"
pn = "PARTNUMBER"
all_reports = read_all_excel_sheets(report_path)
all_items =  read_all_excel_sheets(items_path)
all_items = set(all_items[0][pn]) # Assuming only one sheet with items

filtered_reports = []
print(len(all_reports))
for report in all_reports:
    print(report.columns)
    if parent_part in report.columns:
        filtered = report[report[parent_part].isin(all_items)][[child_part, parent_part, prefix]]
        filtered_reports.append(filtered)

combined_reports = pd.concat(filtered_reports, ignore_index=True)

combined_reports = combined_reports[~combined_reports.duplicated(subset=[child_part, parent_part], keep=False)]
combined_reports.to_excel("parent_part_comparison.xlsx", index=False)

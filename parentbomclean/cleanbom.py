import pandas as pd
import numpy as np
import os

folder = "reports"

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
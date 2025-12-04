import pandas as pd
import os

folder = "reports"
revstr = "Revision"
orgstr = "Org"
nanstr = "Unnamed"

report_path = os.path.join(os.path.dirname(__file__), folder)

def read_all_excel_sheets(folder_path):
    sheets = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                print("reading sheet:", sheet)
                reportdf = pd.read_excel(xls, sheet_name=sheet)
                sheets.append(reportdf)
    return sheets

all_reports = read_all_excel_sheets(report_path)

def clean_revision_data(reports):
    for df in reports:
        print(df.columns.tolist())
        


        # for col in df.columns.tolist():
        #     if orgstr in col:
        #         df.rename(columns={col: orgstr}, inplace=True)
        # while orgstr not in any(df.columns.tolist()):     
        #     df.drop(df.index[0], inplace=True)
        #     df.columns = df.iloc[0]
        #     df = df[1:].reset_index(drop=True)
        #     print("Adjusted header row", df.columns.tolist())

clean_revision_data(all_reports)
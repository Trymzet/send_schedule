import pandas as pd
import openpyxl as oxl
from datetime import date
from os import getcwd
from settings import TODAY, WDEM_FILE_FORMAT, CRC_FILE_FORMAT

input_WDEM_file_path = getcwd().replace("\\", "/").replace("Script", "") + f"Input/Audit_queue_{TODAY}.{WDEM_FILE_FORMAT}"
output_WDEM_file_path = getcwd().replace("\\", "/").replace("Script", "") + f"Output/Audit_queue_{TODAY}.{WDEM_FILE_FORMAT}"

# TODO: change to the CRC path
input_CRC_file_path = getcwd().replace("\\", "/").replace("Script", "") + f"Input/Audit_queue_{TODAY}.{CRC_FILE_FORMAT}"
output_CRC_file_path = getcwd().replace("\\", "/").replace("Script", "") + f"Output/Audit_queue_{TODAY}.{CRC_FILE_FORMAT}"


def generate_attachment(which="WDEM"):
    if which == "WDEM":
        our_business_processes = ["Approval by Receipt Processor", "Approval by Expense Partner"]
        WDEM_report = load_file(input_WDEM_file_path, which="WDEM")
        # deleting useless itemization
        WDEM_report.drop_duplicates(subset=["Expense Number"], inplace=True)
        WDEM_report = WDEM_report[WDEM_report["Awaiting BP Step"].isin(our_business_processes)]
        WDEM_report = WDEM_report[WDEM_report["Approver(s)"].str.contains("Stachura")]

        # below lines delete the useless "Expense Report: " and "Approval by " prefixes
        WDEM_report["Expense Number"] = WDEM_report["Expense Number"].apply(lambda x: x[16:])
        WDEM_report["Awaiting BP Step"] = WDEM_report["Awaiting BP Step"].apply(lambda x: x[12:])

        attachment_path = output_WDEM_file_path
        report = WDEM_report

    elif which == "CRC":
        CRC_schedule = load_file(input_CRC_file_path, which="CRC")

        attachment_path = output_CRC_file_path
        report = CRC_schedule

    with pd.ExcelWriter(attachment_path, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="Schedule", index=False)

    prettify(attachment_path)


def load_file(path, which=None):
    # loads the Excel report to memory
    if which == "WDEM":
        try:
            report = pd.read_excel(path, skiprows=5, usecols=[0, 1, 2, 4, 5, 6, 7, 8, 9])
        except Exception as e:
            print(e)
    elif which == "CRC":
        try:
            report = pd.read_csv(path)
        except Exception as e:
            print(e)
    return report


def prettify(file_path):
    # adjusts column width and adds a filter dropdown to each column
    wb = oxl.load_workbook(file_path)
    ws = wb.active

    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column].width = length

    af = oxl.worksheet.filters.AutoFilter('A1:I1')
    ws.auto_filter = af

    wb.save(file_path)
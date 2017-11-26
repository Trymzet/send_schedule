import win32com.client as win32
from datetime import date
import pandas as pd
import openpyxl as oxl

today = date.today().strftime("%d.%m") # e.g. 25.11
report_path = "../Rozpiska/Audit_queue_{}{}".format(today, ".xlsx")


def prepare_schedule():
    our_business_processes = ["Approval by Receipt Processor", "Approval by Expense Partner"]
    queue_report = load_report(report_path)
    # deleting useless itemization
    queue_report.drop_duplicates(subset=["Expense Number"], inplace=True)
    queue_report = queue_report[queue_report["Awaiting BP Step"].isin(our_business_processes)]
    queue_report = queue_report[queue_report["Approver(s)"].str.contains("Stachura")]
    # below lines delete the useless "Expense Report: " and "Approval by " prefixes
    queue_report["Expense Number"] = queue_report["Expense Number"].apply(lambda x: x[16:])
    queue_report["Awaiting BP Step"] = queue_report["Awaiting BP Step"].apply(lambda x: x[12:])

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        queue_report.to_excel(writer, sheet_name="Schedule", index=False)

    style_report()


def load_report(path):
    # loads the Excel report to memory
    try:
        report = pd.read_excel(path, skiprows=1, usecols=[0, 1, 2, 4, 5, 6, 7, 8, 9])
    except Exception as e:
        print(e)
    return report


def style_report():
    # adjusts column width and adds a filter dropdown to each column
    wb = oxl.load_workbook(report_path)
    ws = wb.active

    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column].width = length

    af = oxl.worksheet.filters.AutoFilter('A1:I1')
    ws.auto_filter = af

    wb.save("Audit_queue_{}{}".format(today, ".xlsx"))


def send_email(subject, body, attachment, to):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()

# Prepare the file

prepare_schedule()

# Send the email

today = date.today().strftime("%d.%m")

to = "x@y.com"
WDEM_schedule_subject = "Rozpiska WDEM {}".format(today)
WDEM_schedule_body = "Siema, <br><br> Rozpiska w zalaczniku - milego!<br><br>Michal"
WDEM_attachment = r"C:\Users\zawadzmi\Desktop\Rozpiska\Audit_queue_{}{}".format(today, ".xlsx")

CRC_schedule_subject = "Rozpiska CRC {}".format(today)
CRC_schedule_body = "Siema, <br><br> Rozpiska w zalaczniku - milego!<br><br>Michal"
CRC_attachment = r"C:\Users\zawadzmi\Desktop\Rozpiska\Audit_queue_{}{}".format(today, ".xlsx")

send_email(WDEM_schedule_subject, WDEM_schedule_body, WDEM_attachment, to)
send_email(CRC_schedule_subject, CRC_schedule_body, CRC_attachment, to)

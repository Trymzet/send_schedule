import win32com.client as win32
from os import chdir, path
from sys import argv
from settings import TO, TODAY
from prepare import generate_attachment, output_WDEM_file_path, output_CRC_file_path

# change working directory to wherever the script is located
# the try is for running the script via python, except is for running from a .bat. I will freeze it later anyway.
try:
    chdir(path.dirname(argv[0]))
except:
    chdir(path.dirname(argv[1]))


def send_email(subject, body, attachment, to):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()


def prepare_email(which):
    if which == "WDEM":
        subject = f"Rozpiska WDEM {TODAY}"
        body = "Siema, <br><br> Rozpiska w zalaczniku. Milego!<br><br>Michal"
        attachment = output_WDEM_file_path
    elif which == "CRC":
        subject = f"Rozpiska CRC {TODAY}"
        body = "Siema, <br><br> Rozpiska w zalaczniku. Milego!<br><br>Michal"
        attachment = output_CRC_file_path
    data = (subject, body, attachment)
    return data


generate_attachment(which="WDEM")

WDEM_email = prepare_email("WDEM")
CRC_email = prepare_email("CRC")

send_email(*WDEM_email, TO)
send_email(*CRC_email, TO)

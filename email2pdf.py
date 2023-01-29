
import os
import shutil
from PyPDF2 import PdfFileMerger
from pypdf import PdfWriter, PdfReader
from chilkat import CkEmail
from imap_tools import MailBox
import pdfkit
import time

options = {
    'page-size': 'A4',
    'margin-top': '15mm',
    'margin-right': '15mm',
    'margin-bottom': '15mm',
    'margin-left': '15mm',
    'encoding': "UTF-8",
    'no-outline': None
}

IMAP_SERVER = os.environ['IMAP_SERVER']
IMAP_USERNAME = os.environ['IMAP_USERNAME']
IMAP_PASSWORD = os.environ['IMAP_PASSWORD']
IMAP_INPUT_FOLDER = os.environ['IMAP_INPUT_FOLDER']
IMAP_SCAN_INTERVAL = int(os.environ['IMAP_SCAN_INTERVAL'])

OUTPUT_DIR = "/output/"
TMP_DIR = "/tmp/email2pdf/"

ALLOWED_TYPES = [
    "application/pdf",
    "application/octet-stream"
]

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
while True:
    try:
        print("Starting!")
        with MailBox(IMAP_SERVER).login(IMAP_USERNAME, IMAP_PASSWORD, initial_folder=IMAP_INPUT_FOLDER) as mailbox:
            for mail in mailbox.fetch():
                mailsubject = f'{mail.subject.replace(".", "_").replace(" ", "-")[:50]}'
                for bad_char in ["/", "*", ":", "<", ">", "|", '"', "’", "–"]:
                    mailsubject = mailsubject.replace(bad_char, "_")
                print(f"\nPDF#######################: {mailsubject}")
                print("Processing Message: " + mailsubject)
                if not os.path.exists(TMP_DIR + mailsubject + "/attachments/"):
                    os.makedirs(TMP_DIR + mailsubject + "/attachments/")

                for attachment in mail.attachments:
                    print(attachment.content_type)
                    if attachment.content_type in ALLOWED_TYPES:
                        attachmentfilename = attachment.filename.replace(" ", "_")
                        attachmentfilename = ''.join(filter(str.isalnum, attachmentfilename))
                        attachmentfilename = attachmentfilename + ".pdf"
                        print("Processing attachment: ", attachmentfilename, " | Type: ", attachment.content_type)
                        with open(TMP_DIR + mailsubject + "/attachments/" + attachmentfilename, "wb") as attachment_file:
                            attachment_file.write(attachment.payload)

                HEADER= "<hr><br><b>Von:</b> " + mail.from_ + "<br><b>An:</b> " + mail.to[0] + "<br><b>Datum:</b> " + mail.date_str + "<br><br><b>Betreff: " + mail.subject + "</b>"
#                HEADER=HEADER + "<br><br>------------------------------------------------------------<br><br>"
                HEADER=HEADER + "<br><br><hr><br><br>"
                if not mail.html.strip() == "":  # handle text only emails
                    pdftext = ('<meta http-equiv="Content-type" content="text/html; charset=utf-8"/>' + HEADER + mail.html)
                else:
                    pdftext = HEADER + mail.text

                print("++++++++++++++++++++++++++")
                print(TMP_DIR + mailsubject)
                print("++++++++++++++++++++++++++")
                pdfkit.from_string(pdftext, TMP_DIR + mailsubject + "/" + mailsubject + ".pdf", options=options)

                merger = PdfWriter()
                merger.append(TMP_DIR + mailsubject + "/" + mailsubject + ".pdf")
                for pdf in os.listdir(TMP_DIR + mailsubject + "/attachments/"):
                    if pdf.endswith(".pdf"):
                        merger.append(TMP_DIR + mailsubject + "/attachments/" + pdf)

                NEWFILE=OUTPUT_DIR + mailsubject + ".pdf"
                merger.write(NEWFILE)

                shutil.rmtree(TMP_DIR)

#            print("Moving all mails to IMAP_OUTPUT folder")
#            mailbox.move(mailbox.fetch(), IMAP_OUTPUT_FOLDER)

        print("Finished!")

    except TimeoutError:
        print("Connection timeout!")

    time.sleep(IMAP_SCAN_INTERVAL)

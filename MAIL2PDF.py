import os
import email
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime

def convert_html_to_text(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup.get_text()

def save_email_to_pdf(subject, content, pdf_path):
    pdfmetrics.registerFont(TTFont('Japanese', '[フォントのPass]'))  # 日本語フォントのパスを指定
    c = canvas.Canvas(pdf_path, pagesize=A4)
    c.setFont('Japanese', 10)  # 日本語フォントを使用
    c.drawString(30, 800, f"Subject: {subject}")
    textobject = c.beginText(30, 780)
    
    page_height = A4[1]
    line_height = 12
    bottom_margin = 20

    for line in content.split('\n'):
        textobject.textLine(line)
        if textobject.getY() < bottom_margin:  # Check if the bottom of the page is reached
            c.drawText(textobject)
            c.showPage()  # Start a new page
            c.setFont('Japanese', 10)
            textobject = c.beginText(30, page_height - line_height - bottom_margin)

    c.drawText(textobject)
    c.save()

def read_eml_files_and_save_as_pdf(folder_path):
    if not os.path.isdir(folder_path):
        print("Folder not found.")
        return

    for file in os.listdir(folder_path):
        if file.endswith('.eml'):
            eml_path = os.path.join(folder_path, file)
            with open(eml_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)

            # 受信日時を取得し、ファイル名に使用
            received_time = msg['Date']
            if received_time:
                received_datetime = email.utils.parsedate_to_datetime(received_time)
                timestamp = received_datetime.strftime("%Y%m%d%H%M")
            else:
                timestamp = "unknown_date"

            pdf_file_name = f"{timestamp}_Starbucks.pdf"
            pdf_path = os.path.join(folder_path, pdf_file_name)
            
            subject = msg['subject']
            body = msg.get_body(preferencelist=('plain', 'html'))
            content = body.get_content() if body else "No content available."
            if body.get_content_type() == 'text/html':
                content = convert_html_to_text(content)
            save_email_to_pdf(subject, content, pdf_path)

    print("Completed converting .eml files to PDF.")

# Set the folder path to the current directory
current_folder_path = os.getcwd()

# Convert .eml files in the current folder to PDF
read_eml_files_and_save_as_pdf(current_folder_path)

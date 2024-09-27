import os
import imaplib
import email
import re
import time
import pandas as pd
import smtplib
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from bs4 import BeautifulSoup
from email.mime.application import MIMEApplication


class EmailForwarder:
    def __init__(self, imap_server, smtp_server, mail_date, staff_number, sender_email,
                 password, excel_file, sender_filter, cc_email, pdf_filename,
                 insurance_message_file):
        self.imap_server = imap_server
        self.smtp_server = smtp_server
        self.staff_number = staff_number
        self.sender_email = sender_email
        self.password = password
        self.excel_file = excel_file
        self.sender_filter = sender_filter
        self.mail = None
        self.date = datetime.strptime(mail_date, "%d/%m/%Y")
        self.forwarded_file = "forwarded_emails_" + mail_date.replace("/", "_") + ".csv"
        self.log_file = "not_forwarded_log_" + mail_date.replace("/", "_") + ".txt"
        self.cc_email = cc_email
        self.pdf_filename = pdf_filename
        with open(insurance_message_file, 'r', encoding='utf-8') as file:
            self.insurance_message = file.read().replace('\n', '<br>')
        # Ensure the forwarded email tracking CSV file exists
        if not os.path.exists(self.forwarded_file):
            with open(self.forwarded_file, mode="w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Registration Number"])  # CSV Header

        # Ensure the log file exists
        if not os.path.exists(self.log_file):
            with open(self.log_file, "w") as f:
                pass

    def connect_imap(self):
        self.mail = imaplib.IMAP4_SSL(self.imap_server)
        self.mail.login(self.staff_number, self.password)
        self.mail.select("inbox")

    def fetch_emails_from_today(self):
        day = self.date.strftime('%d-%b-%Y')
        search_criteria = f'(SINCE "{day}" FROM "{self.sender_filter}")'
        result, data = self.mail.search(None, search_criteria)
        email_ids = data[0].split()
        return email_ids

    @staticmethod
    def extract_registration_number(body):
        soup = BeautifulSoup(body, "html.parser")
        plain_text = soup.get_text(separator=" ")

        match = re.search(r"Vehicle Registration\s*#\s*:\s*([A-Za-z0-9]+)", plain_text)

        if match:
            print(f"Registration Number Found: {match.group(1)}")
            return match.group(1)
        else:
            print("No Registration Number Found")
            return None

    @staticmethod
    def extract_other_fields(body):
        certificate_match = re.search(r"Certificate #\s*:\s*([A-Z0-9]+)", body)
        certificate_number = certificate_match.group(1) if certificate_match else None

        policy_match = re.search(r"Policy #\s*:\s*([A-Z0-9/]+)", body)
        policy_number = policy_match.group(1) if policy_match else None

        chassis_match = re.search(r"Chassis #\s*:\s*([A-Z0-9]+)", body)
        chassis_number = chassis_match.group(1) if chassis_match else None

        return {
            "certificate_number": certificate_number,
            "policy_number": policy_number,
            "chassis_number": chassis_number
        }

    def get_email_from_excel(self, registration_number):
        df = pd.read_excel(self.excel_file)
        email_match = df[df['REG NUMBER'] == registration_number]['EMAIL ADDRESS']
        if not email_match.empty:
            return email_match.values[0]
        return None

    def forward_email(self, msg, recipient_email, reg_number):
        forward_msg = MIMEMultipart()
        forward_msg["From"] = self.sender_email
        forward_msg["To"] = recipient_email
        forward_msg["Cc"] = self.cc_email
        forward_msg["Subject"] = "FWD: " + msg["Subject"]

        # Collect the recipients for sending the email
        recipients = [recipient_email, self.cc_email]

        body_content = ""

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is not None:
                forward_msg.attach(part)
            else:
                content_type = part.get_content_type()

                # Handle plain text content
                if content_type.startswith("text/plain"):
                    body_content = part.get_payload(decode=True).decode("utf-8")

                    # Prepend the insurance message to the plain text email
                    full_body_text = self.insurance_message + "\n\n" + body_content
                    forward_msg.attach(MIMEText(full_body_text, "plain"))

                # Handle HTML content
                elif content_type.startswith("text/html"):
                    body_content = part.get_payload(decode=True).decode("utf-8")

                    # Convert the plain text insurance message to HTML with <br> tags
                    insurance_message_html = self.insurance_message.replace("\n", "<br>")

                    # Prepend the insurance message to the HTML email
                    full_body_html = f"<p>{insurance_message_html}</p>{body_content}"
                    forward_msg.attach(MIMEText(full_body_html, "html"))

        # Attach the test.pdf file from the same directory
        try:
            with open(self.pdf_filename, "rb") as pdf_file:
                attach = MIMEApplication(pdf_file.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment', filename=self.pdf_filename)
                forward_msg.attach(attach)
        except FileNotFoundError:
            print(f"Attachment {self.pdf_filename} not found.")

        # Try to send the email via SMTP
        try:
            smtp_server = smtplib.SMTP(self.smtp_server, 587)
            smtp_server.starttls()
            smtp_server.login(self.staff_number, self.password)
            smtp_server.sendmail(self.sender_email, recipients, forward_msg.as_string())
            smtp_server.quit()

            print(
                f"Email sent successfully to {recipient_email} (CC: {self.cc_email}) with attachment {self.pdf_filename}")
            self.mark_as_forwarded(reg_number)
            time.sleep(15)

        except smtplib.SMTPException as e:
            print(f"Failed to send email: {e}")
            self.log_not_forwarded(reg_number, reason="Failed to send email")
            time.sleep(120)

    def check_if_already_forwarded(self, registration_number):
        """Check if the registration number has already been forwarded using the CSV file."""
        with open(self.forwarded_file, mode="r", newline="") as f:
            reader = csv.reader(f)
            next(reader)  # Skip the header
            for row in reader:
                if row[0] == registration_number:
                    return True
        return False

    def mark_as_forwarded(self, registration_number):
        """Mark a registration number as forwarded by adding it to the CSV file."""
        with open(self.forwarded_file, mode="a", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([registration_number])

    def log_not_forwarded(self, registration_number, reason="Unknown reason"):
        """Log the registration number and reason when forwarding fails."""
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"Registration Number: {registration_number}, Reason: {reason}\n")

    def process_emails(self):
        self.connect_imap()
        email_ids = self.fetch_emails_from_today()

        for email_id in email_ids:
            result, data = self.mail.fetch(email_id, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            if msg.is_multipart():
                body = None
                for part in msg.walk():
                    if part.get_content_type().startswith("text/plain"):
                        body = part.get_payload(decode=True).decode("utf-8")
                        break
                    elif part.get_content_type().startswith("text/html"):
                        body = part.get_payload(decode=True).decode("utf-8")

                if body:
                    registration_number = self.extract_registration_number(body)
                    fields = self.extract_other_fields(body)
                    if registration_number:
                        if self.check_if_already_forwarded(registration_number):
                            print(f"Registration number {registration_number} has already been forwarded.")
                            continue
                        recipient_email = self.get_email_from_excel(registration_number)
                        if recipient_email:
                            self.forward_email(msg, recipient_email, registration_number)
                        else:
                            self.log_not_forwarded(registration_number, reason="Email not found in spreadsheet")
            else:
                body = msg.get_payload(decode=True).decode("utf-8")
                registration_number = self.extract_registration_number(body)
                fields = self.extract_other_fields(body)
                if registration_number:
                    if self.check_if_already_forwarded(registration_number):
                        print(f"Registration number {registration_number} has already been forwarded.")
                        continue
                    recipient_email = self.get_email_from_excel(registration_number)
                    if recipient_email:
                        self.forward_email(msg, recipient_email, registration_number)
                    else:
                        self.log_not_forwarded(registration_number, reason="Email not found in spreadsheet")

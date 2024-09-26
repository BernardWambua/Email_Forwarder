import os
import imaplib
import email
import re
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from bs4 import BeautifulSoup


class EmailForwarder:
    def __init__(self, imap_server, smtp_server, mail_date, staff_number, sender_email, password, excel_file,
                 sender_filter, forwarded_file=""):
        self.imap_server = imap_server
        self.smtp_server = smtp_server
        self.staff_number = staff_number
        self.sender_email = sender_email
        self.password = password
        self.excel_file = excel_file
        self.sender_filter = sender_filter
        self.mail = None
        self.date = datetime.strptime(mail_date, "%d/%m/%Y")
        self.forwarded_file = "forwarded_emails_" + mail_date.replace("/", "_") + ".txt"
        self.log_file = "not_forwarded_log_" + mail_date.replace("/", "_") + ".txt"  # Log file for emails not forwarded

        # Ensure the forwarded email tracking file exists
        if not os.path.exists(self.forwarded_file):
            with open(self.forwarded_file, "w") as f:
                pass

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
        # Parse the HTML to extract the plain text
        soup = BeautifulSoup(body, "html.parser")
        plain_text = soup.get_text(separator=" ")  # Extract plain text from the HTML

        print(f"Extracted Plain Text: {plain_text}")  # Debugging: print the extracted plain text

        # Now apply the regular expression to extract the registration number
        match = re.search(r"Vehicle Registration\s*#\s*:\s*([A-Za-z0-9]+)", plain_text)

        if match:
            print(f"Registration Number Found: {match.group(1)}")  # Debugging output
            return match.group(1)
        else:
            print("No Registration Number Found")  # Debugging output
            return None

    @staticmethod
    def extract_other_fields(body):
        """Extract other relevant fields like certificate number, policy number."""
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

    def forward_email(self, msg, recipient_email):
        forward_msg = MIMEMultipart()
        forward_msg["From"] = self.sender_email
        forward_msg["To"] = recipient_email
        forward_msg["Subject"] = "FWD: " + msg["Subject"]

        for part in msg.walk():
            # Skip multipart container parts
            if part.get_content_maintype() == 'multipart':
                continue

            # For attachments
            if part.get('Content-Disposition') is not None:
                forward_msg.attach(part)
            # For plain text or HTML parts, relaxed conditions
            else:
                content_type = part.get_content_type()

                if content_type.startswith("text/plain"):
                    forward_msg.attach(MIMEText(part.get_payload(decode=True).decode("utf-8"), "plain"))
                elif content_type.startswith("text/html"):
                    forward_msg.attach(MIMEText(part.get_payload(decode=True).decode("utf-8"), "html"))

        smtp_server = smtplib.SMTP(self.smtp_server, 25)
        smtp_server.starttls()
        smtp_server.login(self.staff_number, self.password)
        smtp_server.sendmail(self.sender_email, recipient_email, forward_msg.as_string())
        smtp_server.quit()

    def check_if_already_forwarded(self, email_id):
        """Check if the email ID has already been forwarded."""
        with open(self.forwarded_file, "r") as f:
            forwarded_ids = f.read().splitlines()
        return email_id in forwarded_ids

    def mark_as_forwarded(self, email_id):
        """Mark an email ID as forwarded by adding it to the tracking file."""
        with open(self.forwarded_file, "a") as f:
            f.write(email_id + "\n")

    def log_not_forwarded(self, registration_number, email_body, fields):
        """Log the registration number, email body, and other fields when forwarding fails."""
        with open(self.log_file, "a", encoding="utf-8") as f:  # Specify UTF-8 encoding
            f.write(f"Registration Number: {registration_number}, Certificate: {fields['certificate_number']}, "
                    f"Policy: {fields['policy_number']}, Chassis: {fields['chassis_number']}, Email Body: {email_body}\n")

    def process_emails(self):
        self.connect_imap()
        email_ids = self.fetch_emails_from_today()

        for email_id in email_ids:
            if self.check_if_already_forwarded(email_id.decode()):  # Check if already forwarded
                print(f"Email {email_id.decode()} has already been forwarded.")
                continue

            result, data = self.mail.fetch(email_id, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            if msg.is_multipart():
                body = None  # Initialize to None to handle cases where no text part is found
                for part in msg.walk():
                    print(part.get_content_type())  # Debugging to see the actual content types

                    # Relaxed check for plain text
                    if part.get_content_type().startswith("text/plain"):
                        body = part.get_payload(decode=True).decode("utf-8")
                        # print("Plain Text Body Extracted: ", body)  # Debugging: print the extracted body
                        break  # If plain text is found, no need to check further parts

                    # Optionally, check for HTML content if no plain text is found
                    elif part.get_content_type().startswith("text/html"):
                        body = part.get_payload(decode=True).decode("utf-8")
                        # print("HTML Body Extracted: ", body)  # Debugging: print the extracted body

                if body:
                    # Process the email body using BeautifulSoup to extract text from HTML
                    registration_number = self.extract_registration_number(body)
                    fields = self.extract_other_fields(body)  # Extract other fields
                    if registration_number:
                        recipient_email = self.get_email_from_excel(registration_number)
                        if recipient_email:
                            self.forward_email(msg, recipient_email)
                            self.mark_as_forwarded(email_id.decode())  # Mark email as forwarded
                        else:
                            self.log_not_forwarded(registration_number, body, fields)  # Log not forwarded email
            else:
                # Handle non-multipart emails
                body = msg.get_payload(decode=True).decode("utf-8")
                # print("Non-Multipart Body Extracted: ", body)  # Debugging: print the extracted body
                registration_number = self.extract_registration_number(body)
                fields = self.extract_other_fields(body)  # Extract other fields
                if registration_number:
                    recipient_email = self.get_email_from_excel(registration_number)
                    if recipient_email:
                        self.forward_email(msg, recipient_email)
                        self.mark_as_forwarded(email_id.decode())  # Mark email as forwarded
                    else:
                        self.log_not_forwarded(registration_number, body, fields)  # Log not forwarded email

from email_forwarder import EmailForwarder


if __name__ == "__main__":
    print("Starting program")
    email_forwarder = EmailForwarder(
        imap_server="",
        smtp_server="",
        staff_number="",
        sender_email="",
        password="",
        excel_file="",
        sender_filter="",
        mail_date="26/09/2025"
    )
    email_forwarder.process_emails()

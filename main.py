from email_forwarder import EmailForwarder


if __name__ == "__main__":
    print("Starting program")
    email_forwarder = EmailForwarder(
        imap_server="mail.kengen.co.ke",
        smtp_server="mail.kengen.co.ke",
        staff_number="ISDesk",
        sender_email="insurance@kengen.co.ke",
        password="Password1234",
        excel_file="C:\\Users\\yxd\\Desktop\\python\\2026 Car Insurance Renewal DB.xlsx",
        sender_filter="aki@dmvic.com",
        mail_date="26/09/2025"
    )
    email_forwarder.process_emails()

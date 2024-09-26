from flask import Flask, render_template, request, redirect, url_for, flash
from email_forwarder import EmailForwarder  # Assuming the above class is saved in email_forwarder.py
import os
import secrets

app = Flask(__name__)
secret_key = secrets.token_hex(16)
app.secret_key = secret_key


# Define the route for the homepage
@app.route("/", methods=["GET", "POST"])
def index():
    default_values = {
        "imap_server": "mail.kengen.co.ke",
        "smtp_server": "172.16.4.7",
        "mail_date": "26/09/2024",
        "staff_number": "ISDesk",
        "sender_email": "insurance@kengen.co.ke",
        "password": "",
        "sender_filter": "aki@dmvic.com"
    }
    if request.method == "POST":
        # Retrieve form values
        imap_server = request.form.get("imap_server")
        smtp_server = request.form.get("smtp_server")
        mail_date = request.form.get("mail_date")
        staff_number = request.form.get("staff_number")
        sender_email = request.form.get("sender_email")
        password = request.form.get("password")
        excel_file = request.files["excel_file"]
        sender_filter = request.form.get("sender_filter")

        # Save the uploaded Excel file temporarily
        excel_file_path = os.path.join("uploads", excel_file.filename)
        excel_file.save(excel_file_path)

        try:
            # Initialize EmailForwarder and process emails (pseudo-code)
            forwarder = EmailForwarder(
                imap_server=imap_server,
                smtp_server=smtp_server,
                mail_date=mail_date,
                staff_number=staff_number,
                sender_email=sender_email,
                password=password,
                excel_file=excel_file_path,
                sender_filter=sender_filter
            )
            forwarder.process_emails()  # Process emails logic
            flash("Emails processed and forwarded successfully!", "success")
        except Exception as e:
            flash(f"An error occurred: {e}", "danger")

        return redirect(url_for("index"))

        # Render the template with the default values
    return render_template("index.html", **default_values)


# Run the app
if __name__ == "__main__":
    if not os.path.exists("uploads"):
        os.makedirs("uploads")
    app.run(debug=True)

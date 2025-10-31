# Email Forwarding Automation for Insurance Team

## 📘 Overview

This project is a **Python-based automation tool** designed to streamline the process of forwarding emails from an underwriting team to the appropriate insurance team members.

The application automatically:

1. Reads incoming emails from a shared mailbox.
2. Extracts the **vehicle registration number** from the email body.
3. Looks up the corresponding **employee email address** from an Excel mapping file.
4. Forwards the original message to the correct recipient — **without any manual intervention**.

This automation saves significant time, reduces human error, and ensures prompt handling of insurance-related communications.

---

## 🎯 Objectives

* Eliminate manual lookup and forwarding of emails by underwriters.
* Improve turnaround time in insurance case handling.
* Ensure accuracy and consistency in email routing using a single verified source of data.
* Provide configurable and maintainable Python automation for future scaling.

---

## ⚙️ How It Works

1. **Email Retrieval:**
   The app connects to a **shared mailbox** (e.g., via Microsoft Graph API, Outlook REST API, or IMAP).

2. **Text Extraction:**
   It scans the **email body** for a **vehicle registration number** using regular expressions.

3. **Excel Lookup:**
   The registration number is matched against an **Excel sheet** (e.g., `registration_mapping.xlsx`) containing mappings like:

   ```
   | Registration No | Employee Email          |
   |-----------------|-------------------------|
   | ABC1234         | john.doe@company.com    |
   | XYZ5678         | jane.smith@company.com  |
   ```

4. **Email Forwarding:**
   The email is forwarded automatically to the matched employee’s email address.
   If no match is found, it can:

   * Send a notification to an admin mailbox, or
   * Store the unmatched entry in a log file for later review.

---

## 🧩 Core Features

* **Automated Email Parsing** — Extracts key identifiers (registration numbers) using regex.
* **Excel Integration** — Reads employee mappings via `pandas`.
* **Smart Forwarding** — Uses the `smtplib` and `imaplib` (or Microsoft Graph API) to send and forward emails.
* **Error Handling & Logging** — Tracks unmatched cases, delivery issues, and system errors.
* **Configurable Settings** — Mailbox credentials, Excel path, and logging details can be customized via a config file.

---

## 🧠 Tech Stack

| Component            | Technology                                            |
| -------------------- | ----------------------------------------------------- |
| Programming Language | Python 3.x                                            |
| Email Handling       | `imaplib`, `smtplib`, or `O365` / Microsoft Graph API |
| Data Handling        | `pandas`, `openpyxl`                                  |
| Automation           | `schedule` or `cron` for periodic checks              |
| Logging              | `logging` module                                      |

---

## 🧾 Folder Structure

```
email_forwarding_automation/
│
├── config/
│   ├── config.json                # Mailbox credentials and file paths
│
├── data/
│   ├── registration_mapping.xlsx  # Mapping of vehicle reg no to employee email
│
├── src/
│   ├── email_forwarder.py         # Sends emails to matched recipients
│   ├── main.py                    # Entry point for running automation
│
├── logs/
│   ├── process_log.txt            # Logs of processed and failed forwards
│
├── requirements.txt
└── README.md
```

---

## 🚀 How to Run

1. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

2. **Update configuration**
   Edit `config/config.json` with:

   ```json
   {
     "mailbox_user": "shared_mailbox@company.com",
     "mailbox_password": "yourpassword",
     "excel_path": "data/registration_mapping.xlsx",
     "smtp_server": "smtp.office365.com",
     "imap_server": "outlook.office365.com"
   }
   ```

3. **Run the script**

   ```bash
   python src/main.py
   ```

4. **(Optional)** Schedule it to run periodically:

   * **Windows:** Use Task Scheduler
   * **Linux/macOS:** Add a cron job

---

## 📊 Example Workflow

**Email body example:**

```
Dear Insurance Team,

Please find attached documents for registration number XYZ5678.

Regards,
Underwriting Team
```

**Automation steps:**

1. Extracts `XYZ5678`
2. Looks up corresponding email (`jane.smith@company.com`)
3. Forwards email automatically with the same subject and attachments

✅ **Result:** Email instantly reaches the right person, no manual lookup needed.

---

## 🧩 Error Handling

* Logs all actions in `logs/process_log.txt`
* Unmatched registration numbers are stored with timestamps
* Email errors (e.g., failed send) trigger an alert to a configured admin address

---

## 📈 Future Enhancements

* Add a **web dashboard** for monitoring forwarded and pending emails
* Integrate with **Active Directory** or internal HR database instead of Excel
* Support **multi-level approval routing**
* Include **email content analysis** for prioritization (urgent, claim, renewal, etc.)

---

## 🧑‍💻 Author

**[Your Name]**
Automation Developer / Data Engineer
📧 [[your.email@example.com](mailto:your.email@example.com)]
📅 Last updated: October 2025

# Email-Bot

This is a python email bot that sends mails to a list of recipients listed in a excel workbook. It uses the python libraries smtplib to send the mail and openpyxl to extract the emails from the excel workbook. 

## Features

- Reads email addresses from an Excel workbook.
- Constructs and sends emails to each recipient.
- Supports plain text email bodies.
- Optional attachment functionality.
- Error handling and logging for unsuccessful email deliveries.

## Prerequisites

- Python 3.x
- `openpyxl` library
- app password for your sender email: refer to this article to [this article](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237) to know how to set up app passwords for your email.
- An SMTP server (e.g., Gmail SMTP)(as long as you have an email, you'll have an SMTP server) 
- Note: - `smtplib` and `email` libraries are python standard libeararies need not be installed seperatly

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/email-bot.git
   cd email-bot
   ```

2. Install the required Python libraries:

   ```bash
   pip install openpyxl
   ```

## Setup

1. Update the script with your email details:
   
   - Update the `fromAddress` with your email address.
   - Update the `passw` with your email app password (for Gmail, you might need to generate an app-specific password).
   - Update the `toAddress` with the recipient's address (though this will be overridden by the addresses in the Excel file).

2. Prepare your Excel workbook:

   - Ensure your Excel workbook is named `workbook.xlsx`.
   - The sheet containing the email addresses should be named `Sheet1`.
   - **Email addresses should be in the first column (A) of the sheet.**

3. (Optional) If you have an attachment, uncomment the relevant lines in the script and set the `filename` variable to the relative path of your attachment.
   
   line 14 and lines 62-75 have to be uncommented

## Usage

Run the script:

```bash
python email_bot.py
```

The script will:

1. Log in to the SMTP server.
2. Read email addresses from the Excel file.
3. Construct and send emails to each address.
4. Log the status of each email sent.

## Script Breakdown

### Import Libraries

```python
import smtplib
from email import encoders 
from email.mime.base import MIMEBase 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import openpyxl
```

### Load Excel Workbook

```python
database = openpyxl.load_workbook("workbook.xlsx", data_only=True)
database_sheet = database["Sheet1"]
```

### SMTP Server Configuration

```python
fromAddress = ""  # your email address
passw = ""  # your email app password
#set smtp server and port accoridng to your email provider
smtpServer = "smtp.gmail.com"
port = 587
```

### Login to SMTP Server

```python
server = smtplib.SMTP(smtpServer, port)
server.ehlo()
server.starttls()
server.login(fromAddress, passw)
print("Login successful")
```

### Send Emails

```python
for row in range(0, database_sheet.max_row + 1):
    try:
        toAddress = database_sheet.cell(row, 1).value.strip()
        
        # Create MIME object
        msg = MIMEMultipart()
        msg["From"] = fromAddress
        msg["To"] = toAddress
        msg["Subject"] = "Subject of the mail"

        body = "Hey there,\n\nthis is a sample mail\n\nThanks and regards"
        msg.attach(MIMEText(body, 'plain'))

        # Optional: Add attachment
        # with open(filename, "rb") as attachment:
        #     part = MIMEBase("application", "octet-stream")
        #     part.set_payload(attachment.read())
        #     encoders.encode_base64(part)
        #     part.add_header("Content-Disposition", f"attachment; filename= {filename}")
        #     msg.attach(part)

        txt = msg.as_string()
        server.sendmail(fromAddress, toAddress, txt)
        print(f"Mail sent to {toAddress}")
    except Exception as e:
        print(f"Failed to send email to {toAddress}: {e}")
```

### Logout from SMTP Server

```python
server.quit()
```

## Troubleshooting

- Ensure you have enabled "less secure apps" on your Gmail account or use an app-specific password if you have 2-factor authentication enabled.
- Verify that your Excel file is correctly formatted and the email addresses are in the first column.
- Check your internet connection and SMTP server details.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

For any issues or contributions, please open an issue or submit a pull request on [GitHub](https://github.com/yourusername/email-bot).

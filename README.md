# ACBM Auto Mailer

An Auto Emailer for ACBM designed to automate the process of sending emails with unique attachments (e.g., certificates) to a list of recipients.

## Purpose

The main purpose of this tool is to send bulk emails where each email is customized with a specific attachment (usually an image or document) intended for a specific recipient. It matches files from a local directory to a list of students/recipients provided in an Excel sheet.

## Getting Started

### Prerequisites

Ensure you have Python 3.9 or higher installed.

You will also need a valid Gmail account to use as the sender, with "Less Secure Apps" access enabled or an App Password generated if 2-Factor Authentication is on.

### Installation

1. Clone the repository.
2. Install the required Python packages:

```bash
pip install pandas openpyxl
```

*Note: `smtplib` and `email` are part of the Python standard library.*

## Configuration

The application is controlled via the `config.json` file. You must configure this file before running the script.

### `config.json` Structure

```json
{
    "email": "[your email]",
    "password": "[your password]",
    "subject": "Email Subject Line",
    "body": "body.txt",
    "data": "students.xlsx",
    "folder": "certificates/",
    "files": [
        {
            "type": "images",
            "extension": ".jpg"
        }
    ]
}
```

*   **email**: Your Gmail address.
*   **password**: Your Gmail password (or App Password).
*   **subject**: The subject line for the emails.
*   **body**: The filename of the text file containing the email body content (e.g., `body.txt`).
*   **data**: The filename of the Excel file containing recipient data (e.g., `students.xlsx`).
*   **folder**: The directory containing the files to be attached (e.g., `certificates/`).
*   **files**: A list of file types to scan for.
    *   **type**: An arbitrary name for the file type (used internally for column mapping).
    *   **extension**: The file extension to look for (e.g., `.jpg`, `.pdf`).

### Data Preparation

1.  **Recipients Excel (`students.xlsx`)**:
    *   Must contain a sheet with at least two columns: `Name` and `Email`.
    *   The script reads these columns to personalize the email and send it.

2.  **Attachments (`folder`)**:
    *   Place all files to be attached in the directory specified by `folder` in `config.json`.
    *   **Important**: The script assigns files to students based on the **numerical order** of the filenames.
    *   Example:
        *   `1.jpg` will be sent to the 1st student in the Excel sheet.
        *   `2.jpg` will be sent to the 2nd student, and so on.
    *   Ensure your filenames are named sequentially (e.g., `1.jpg`, `2.jpg`, `3.jpg`) and correspond to the row order in your Excel file.

3.  **Email Body (`body.txt`)**:
    *   Edit `body.txt` to contain the plain text message you want to send.

## Usage

To run the auto-mailer:

```bash
python auto-mailer.py
```

The script will:
1.  Load the configuration.
2.  Read the recipient list.
3.  Scan the attachment folder and map files to recipients.
4.  Log in to the SMTP server.
5.  Iterate through the list and send each email with its corresponding attachment.

## Troubleshooting

*   **SMTP Authentication Error**: If you see an authentication error, check if your email and password are correct. If using Gmail, you likely need to generate an [App Password](https://support.google.com/accounts/answer/185833).
*   **FileNotFoundError**: Ensure all paths in `config.json` are correct and relative to the script's location.
*   **Attachment Mismatch**: Verify that the number of files in the folder matches the number of students and that the filenames are numbered correctly to match the Excel row order.

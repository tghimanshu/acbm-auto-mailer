# Future Plan (Phase 2)

This document outlines the proposed enhancements and features for the next phase of the ACBM Auto Mailer project. The goal is to make the application more robust, user-friendly, and versatile.

## 1. Improved Error Handling and Logging

*   **Robust Logging**: Replace `print` statements with a proper logging mechanism (using Python's `logging` module) to record success, failures, and debug information to a file (`app.log`).
*   **Granular Error Handling**: Implement specific try-except blocks for individual emails. If one email fails, the script should continue to the next one instead of crashing or stopping.
*   **Retry Mechanism**: Automatically retry sending an email if it fails due to a temporary network issue.

## 2. Dynamic File Matching

*   **Name-Based Matching**: Instead of relying on strict numerical ordering (which is fragile), allow matching files based on the student's name or a unique ID column in the Excel sheet (e.g., matching `John_Doe.jpg` to the student "John Doe").
*   **Validation**: Add a pre-check step that verifies if every student in the Excel sheet has a corresponding file before attempting to send emails.

## 3. enhanced Configuration and Security

*   **Environment Variables**: Support loading sensitive credentials (email, password) from environment variables or a `.env` file instead of `config.json` to improve security.
*   **OAuth2 Support**: Implement OAuth2 for Gmail to replace the less secure password/app-password authentication method.

## 4. Templating and Content

*   **HTML Support**: Allow `body.txt` to contain HTML content for rich text emails.
*   **Dynamic Placeholders**: Support placeholders in the email body (e.g., `Hello {Name}`) that are replaced with data from the Excel sheet.

## 5. User Interface (CLI/GUI)

*   **Command Line Arguments**: Use `argparse` to allow overriding config paths via command line flags (e.g., `python auto-mailer.py --config my_config.json`).
*   **Progress Bar**: Add a progress bar (e.g., using `tqdm`) to visualize the sending progress.
*   **Simple GUI**: Create a basic graphical user interface (using `tkinter` or similar) for non-technical users to select files and monitor progress.

## 6. Testing

*   **Unit Tests**: Write unit tests for core logic (file matching, config loading) to ensure reliability.
*   **Mock SMTP**: Use a mock SMTP server for testing the sending process without actually sending emails.

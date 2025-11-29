"""
Auto Mailer Module

This module provides functionality to automatically send emails with attachments to a list of recipients
stored in an Excel file. It maps files from a directory to students based on the order of files (sorted numerically)
and the order of students in the Excel sheet.
"""

import pandas as pd
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class AutoMailer:
    """
    A class to handle the automatic sending of emails with attachments.

    Attributes:
        config (dict): The configuration dictionary loaded from the config file.
        students (pd.DataFrame): The DataFrame containing student data.
        smtp_session (smtplib.SMTP): The active SMTP session for sending emails.
    """

    def __init__(self, config_path='config.json'):
        """
        Initialize the AutoMailer with a configuration file.

        Args:
            config_path (str): The path to the JSON configuration file. Defaults to 'config.json'.
        """
        print('*' * 50)
        print('WELCOME TO AUTO EMAILER FOR ACBM')
        print('ver. 0.1.0')
        print('CREATED BY: TGHIMANSHU')
        print('*' * 50)
        print('\nPLEASE WAIT..........\n')

        self.config = self._load_config(config_path)
        self.students = None
        self.smtp_session = None

    def _load_config(self, path):
        """
        Load the configuration from a JSON file.

        Args:
            path (str): The path to the configuration file.

        Returns:
            dict: The loaded configuration dictionary.
        """
        with open(path, 'r') as content:
            return json.load(content)

    def load_data(self):
        """
        Load student data from the Excel file specified in the config.

        This method reads the Excel file and stores it in the `students` attribute.
        It also processes the file directory to map specific files to each student
        based on the sorted order of filenames.
        """
        self.students = pd.read_excel(self.config['data'])
        print(f"Loaded {len(self.students)} students from {self.config['data']}")

        self._map_files_to_students()

    def _map_files_to_students(self):
        """
        Map files from the configured folder to students.

        The method scans the directory for files matching the extensions defined in the config.
        It sorts these files numerically (assuming filenames are numbers like '1.jpg', '2.jpg')
        and assigns them to students in the order they appear in the DataFrame.
        """
        files_config = self.config['files']
        all_files = []

        # Prepare file categories
        for cfile in files_config:
            all_files.append({'type': cfile['type'], 'extension': cfile['extension'], 'files': []})

        # Scan directory
        try:
            dir_files = os.listdir(self.config['folder'])
        except FileNotFoundError:
            print(f"Error: The folder '{self.config['folder']}' was not found.")
            dir_files = []

        for curr_file in dir_files:
            for cfile in all_files:
                if os.path.splitext(curr_file)[1] == cfile['extension']:
                    cfile['files'].append(curr_file)

        # Assign files to students
        students_len = len(self.students)
        for cfiles in all_files:
            # Sort files by the integer value of their filename (e.g. '1.jpg' -> 1)
            # This assumes filenames are integers. If not, this might fail or sort alphabetically.
            try:
                sorted_files = sorted(cfiles['files'], key=lambda doc: int(os.path.splitext(doc)[0]))
            except ValueError:
                 # Fallback to string sorting if filenames are not integers
                sorted_files = sorted(cfiles['files'])

            # Pad the list if there are fewer files than students
            if len(sorted_files) < students_len:
                sorted_files.extend([None] * (students_len - len(sorted_files)))
            elif len(sorted_files) > students_len:
                sorted_files = sorted_files[:students_len]

            self.students[cfiles['type']] = sorted_files

        self.file_types = [cfile['type'] for cfile in files_config]

    def connect_smtp(self):
        """
        Initialize the SMTP connection using credentials from the config.
        """
        try:
            self.smtp_session = smtplib.SMTP('smtp.gmail.com', 587)
            self.smtp_session.starttls()
            self.smtp_session.login(self.config['email'], self.config['password'])
            print("SMTP connection established successfully.")
        except Exception as e:
            print(f"Failed to connect to SMTP server: {e}")
            raise

    def attach_file(self, msg, filename):
        """
        Attach a file to the email message.

        Args:
            msg (MIMEMultipart): The email message object.
            filename (str): The name of the file to attach (relative to the configured folder).
        """
        filepath = os.path.join(self.config['folder'], filename)
        try:
            with open(filepath, "rb") as attachment:
                p = MIMEBase('application', 'octet-stream')
                p.set_payload(attachment.read())
                encoders.encode_base64(p)
                p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                msg.attach(p)
        except FileNotFoundError:
            print(f"Warning: Attachment {filepath} not found.")

    def send_email(self, name, mail_to, filenames, body_text):
        """
        Send a single email to a student.

        Args:
            name (str): The name of the student.
            mail_to (str): The recipient's email address.
            filenames (list of str): List of filenames to attach.
            body_text (str): The plain text body of the email.
        """
        msg = MIMEMultipart()
        msg['From'] = self.config['email']
        msg['Subject'] = self.config['subject']
        msg.attach(MIMEText(body_text, 'plain'))
        msg['To'] = mail_to

        for filename in filenames:
            if isinstance(filename, str): # Verify filename is a string (not NaN from pandas)
                self.attach_file(msg, filename)

        text = msg.as_string()
        if self.smtp_session:
            self.smtp_session.sendmail(self.config['email'], mail_to, text)
            print(f'Mail Successfully Sent to {name}!')
        else:
            print("SMTP session is not active. Email not sent.")

    def run(self):
        """
        Execute the main mailer process: load data, connect to SMTP, and send emails.
        """
        self.load_data()

        # Load email body
        try:
            with open(self.config['body'], 'r') as content:
                body = content.read()
        except FileNotFoundError:
            print(f"Error: Body file '{self.config['body']}' not found.")
            return

        self.connect_smtp()

        for index, student in self.students.iterrows():
            student_email = student['Email']
            student_name = student['Name']
            student_files = []
            for file_type in self.file_types:
                # pandas might handle missing values, check if valid
                if pd.notna(student[file_type]):
                    student_files.append(student[file_type])

            self.send_email(student_name, student_email, student_files, body)

        if self.smtp_session:
            self.smtp_session.quit()

        input('\nTHANKS FOR USING IT, PRESS ANY KEY TO QUIT....')

if __name__ == "__main__":
    try:
        mailer = AutoMailer()
        mailer.run()
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press any key to exit...")

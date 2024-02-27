# Automated Certificate Generator and Emailer

This project automates the creation of personalized certificates using Microsoft Word templates and Excel spreadsheets for student data, then emails the certificates to each student using Microsoft Outlook.

## Features

- **Certificate Generation**: Automatically fills in student names, course details, and completion dates into Word document templates.
- **Data Handling**: Reads student information, including names and email addresses, from an Excel spreadsheet.
- **Email Automation**: Sends the generated certificates as email attachments to each student's email address using Outlook.
- **Customization**: Allows for easy customization of the certificate template and email content.

## Prerequisites

Before running this script, ensure you have the following installed:
- Python 3.x
- Microsoft Office (Word and Outlook specifically for this script)
- The following Python packages: `python-docx`, `openpyxl`, `pywin32` (for Windows users to automate Outlook emails)

## Installation

First, clone this repository to your local machine. Then, install the required Python packages using pip:

```bash
pip install python-docx openpyxl pywin32
```
## Setup
Template Preparation: Prepare a Word document (.docx) as your certificate template. Placeholders like @name should be used where dynamic content will be inserted.
Data Preparation: Fill an Excel (.xlsx) file with student information, including names, email addresses, and any other details to be included in the certificates.
Script Configuration: Modify the script paths to match the locations of your template and Excel files.
Usage
To generate and email the certificates, run the script from your terminal or command prompt:
```bas
python certificate_generator.py
```
Make sure to open Outlook on your computer before running the script to avoid any issues with the email automation.

## Customization
You can customize the Word template and Excel file format according to your needs. Just ensure the script is updated accordingly to correctly read data and fill the template.

## Contributing
Contributions are welcome! If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".

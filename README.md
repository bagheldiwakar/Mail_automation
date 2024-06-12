# Email Scheduler and Monitor

This project is a Python-based email scheduler and inbox monitor. It allows users to schedule emails to be sent at specific times, with an optional reminder email sent later. The application also monitors an email inbox for new emails and sends notifications for unseen emails. 

## Features

- **Email Scheduling**: Schedule emails to be sent at a specific time.
- **Reminder Emails**: Send a reminder email at a later specified time.
- **Inbox Monitoring**: Monitor an email inbox for new emails and send notifications for unseen emails.
- **GUI Interface**: User-friendly GUI built with Tkinter for easy interaction.
- **Excel File Upload**: Upload Excel files to populate email receivers, contract numbers, and order dates.

## Requirements

- Python 3.6+
- Required Python packages:
  - `tkinter`
  - `Pillow`
  - `pandas`
  - `openpyxl`

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/bagheldiwakar/Mail_automation.git
    cd Mail_automation
    ```

2. Install the required packages:
    ```bash
    pip install -r requirements.txt
    ```

3. Set up environment variables for email credentials (recommended to use a `.env` file or set environment variables directly):
    - `MONITOR_EMAIL`: Email address used for monitoring the inbox.
    - `MONITOR_PASSWORD`: Password for the monitoring email account.
    - `EMAIL_SENDER`: Email address used for sending emails.
    - `EMAIL_PASSWORD`: Password for the sending email account.
    - `NOTIFY_EMAIL`: Email address to receive notifications for new unseen emails.

## Usage

1. Run the application:
    ```bash
    python All_in_one.py
    ```

2. The GUI will open. Fill in the required fields:
    - **Receiver Emails**: Comma-separated list of receiver email addresses.
    - **Contract Numbers**: Comma-separated list of contract numbers corresponding to the receivers.
    - **Order Dates**: Comma-separated list of order dates corresponding to the receivers.
    - **Time to Send**: Time to send the initial email in HH:MM format.
    - **Reminder Time**: Time to send the reminder email in HH:MM format.

3. Alternatively, you can upload an Excel file with columns for receiver emails, contract numbers, and order dates.

4. Click "Process" to prepare the emails for sending. Click "Send" to start the email sending process. Use the "STOP MAIL" button to stop sending emails if needed. The "Done" button closes the application.

## Monitoring Inbox

The application will monitor the inbox of the specified email account for new unseen emails. When a new email is detected, a notification email is sent to the specified notification email address.

## Logging

The application logs events in the GUI, including:
- Emails processed and ready to send.
- Emails sent and their corresponding timestamps.
- Reminder emails sent and their corresponding timestamps.
- Errors encountered during the process.

## Security Considerations

- Make sure to use environment variables or a `.env` file to securely manage your email credentials.
- Avoid hardcoding sensitive information directly into the script.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request if you have any improvements or new features to add.


## Contact

For questions or issues, please open an issue on GitHub or contact (mail to:your-bagheldiwakar2000@gmail.com).

---

import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
import time
import threading
from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd
from PIL import ImageTk, Image
import imaplib
import email


# Global variables to control the email sending threads and button states
email_thread = None
reminder_thread = None
stop_sending = False
cooldown_time = 10  # 10 seconds cooldown

# IMAP configuration for monitoring inbox
MONITOR_EMAIL = 'ABCtest123@gmail.com'
MONITOR_PASSWORD = '**** **** **** ****'
MONITOR_SERVER = 'imap.gmail.com'

# SMTP configuration for sending notification email
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465  # SSL port

# Required column names in the Excel file
RECEIVER_COLUMN_NAME = "Receiver"
CONTRACT_COLUMN_NAME = "contract number"
DATE_COLUMN_NAME = "Order date"

def send_email(email_sender, email_password, email_receivers, subject, body):
    em = EmailMessage()
    em['From'] = f"XYZ International <{email_sender}>"
    em['To'] = ', '.join(email_receivers)
    em['Subject'] = subject
    em.set_content(body)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receivers, em.as_string())

def process_emails():
    global emails_to_send
    email_sender = 'ABCtest123@gmail.com'
    email_password = '**** **** **** ****'
    send_time_str = time_entry.get()
    reminder_time_str = reminder_time_entry.get()
    
    # Validate time format
    try:
        send_time = datetime.strptime(send_time_str, "%H:%M").time()
        reminder_time = datetime.strptime(reminder_time_str, "%H:%M").time()
    except ValueError:
        messagebox.showerror("Invalid time format", "Please enter time in HH:MM format")
        return
    
    email_receivers = [email.strip() for email in receiver_entry.get().split(',')]
    contract_numbers = contract_numbers_entry.get().split(',')
    order_dates = order_dates_entry.get().split(',')
    
    if len(email_receivers) != len(contract_numbers) or len(email_receivers) != len(order_dates):
        messagebox.showerror("Error", "The number of emails, contract numbers, and order dates must be the same")
        return
    
    send_datetime = datetime.combine(datetime.today(), send_time)
    reminder_datetime = datetime.combine(datetime.today(), reminder_time)
    current_time = datetime.now()
    
    if current_time > send_datetime:
        send_datetime += timedelta(days=1)
    if current_time > reminder_datetime:
        reminder_datetime += timedelta(days=1)
    
    send_delay = (send_datetime - current_time).total_seconds()
    reminder_delay = (reminder_datetime - current_time).total_seconds()
    
    emails_to_send = []
    for email, contract, order_date in zip(email_receivers, contract_numbers, order_dates):
        subject = f"First mail subject"
        body = f"""
First mail body
"""
        reminder_subject = f"Second mail subject"
        reminder_body = f""" 
Second mail body
"""
        emails_to_send.append((email_sender, email_password, [email], subject, body, send_delay, reminder_delay, reminder_subject, reminder_body))
    
    process_button.config(state=tk.DISABLED)
    
    # Displaying the format in the log
    example_contract = "[CONTRACT_NUMBER]"
    example_date = "[ORDER_DATE]"
    example_subject = f"Example mail subject"
    example_body = f"""
Example mail body
"""
    log_text.insert(tk.END, f"Emails processed and ready to send at {send_time_str} with a reminder at {reminder_time_str}.\n")
    log_text.insert(tk.END, f"\nExample Email Format:\nSubject: {example_subject}\nBody: {example_body}\n")
    
    # Enable the "Send" button after cooldown to prevent multiple times sent
    app.after(cooldown_time * 1000, send_button.config, {'state': tk.NORMAL})

def send_all_emails():
    global stop_sending, email_thread
    stop_sending = False

    def send_emails():
        global emails_to_send, stop_sending
        for email_sender, email_password, email_receivers, subject, body, send_delay, reminder_delay, reminder_subject, reminder_body in emails_to_send:
            if stop_sending:
                log_text.insert(tk.END, "Email sending process stopped by user.\n")
                break
            try:
                time.sleep(send_delay)
                send_email(email_sender, email_password, email_receivers, subject, body)
                sent_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_text.insert(tk.END, f"Email sent at {sent_time_str} to {email_receivers[0]}\n")
                
                # Wait for the reminder delay
                time.sleep(reminder_delay)
                
                # Send reminder email
                send_email(email_sender, email_password, email_receivers, reminder_subject, reminder_body)
                reminder_sent_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_text.insert(tk.END, f"Reminder email sent at {reminder_sent_time_str} to {email_receivers[0]}\n")
            except Exception as e:
                error_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_text.insert(tk.END, f"Error at {error_time_str} for {email_receivers[0]}: {str(e)}\n")
        
        process_button.config(state=tk.NORMAL)
        send_button.config(state=tk.NORMAL)

    email_thread = threading.Thread(target=send_emails)
    email_thread.start()
    send_button.config(state=tk.DISABLED)  # Disable the button after clicked

def stop_sending_emails():
    global stop_sending
    stop_sending = True

def close_app():
    app.quit()
    app.destroy()

def upload_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        try:
            df = pd.read_excel(filepath, engine='openpyxl')
            email_list = df[RECEIVER_COLUMN_NAME].dropna().tolist()
            contract_list = df[CONTRACT_COLUMN_NAME].dropna().astype(str).tolist()
            date_list = df[DATE_COLUMN_NAME].dropna().astype(str).tolist()
            
            receiver_entry.delete(0, tk.END)
            receiver_entry.insert(0, ', '.join(email_list))
            
            contract_numbers_entry.delete(0, tk.END)
            contract_numbers_entry.insert(0, ', '.join(contract_list))
            
            order_dates_entry.delete(0, tk.END)
            order_dates_entry.insert(0, ', '.join(date_list))
        except Exception as e:
            messagebox.showerror("Error", str(e))

def send_notification(sender):
    def send_email():
        # SMTP connection for sending notification email
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            # Create the notification email message
            msg = EmailMessage()
            msg['From'] = MONITOR_EMAIL  # Sender email address
            msg['To'] = 'XYZtest2@gmail.com'  # Recipient email address
            msg['Subject'] = 'New Email Received'
            msg.set_content(f'You have received a new email from {sender} in your inbox.')

            # Send the notification email
            smtp.login(MONITOR_EMAIL, MONITOR_PASSWORD)
            smtp.send_message(msg)

    try:
        # Start a new thread to send the email
        email_thread = threading.Thread(target=send_email)
        email_thread.start()
    except RuntimeError:
        # Catch RuntimeError when trying to create a new thread at interpreter shutdown
        pass

def get_unseen_emails():
    # Connect to the IMAP server for monitoring inbox
    mail = imaplib.IMAP4_SSL(MONITOR_SERVER)
    mail.login(MONITOR_EMAIL, MONITOR_PASSWORD)
    mail.select('inbox')

    # Search for unseen emails
    result, data = mail.search(None, 'UNSEEN')

    unseen_emails = []
    if result == 'OK':
        for num in data[0].split():
            # Fetch the email
            result, email_data = mail.fetch(num, '(RFC822)')
            if result == 'OK':
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                sender = msg['From']
                unseen_emails.append((sender, msg))
    
    # Close the connection
    mail.close()
    mail.logout()

    return unseen_emails

def process_emails_inbox():
    unseen_emails = get_unseen_emails()

    # Process the unseen emails
    for sender, msg in unseen_emails:
        # Send a notification for each unseen email
        send_notification(sender)

def monitor_inbox():
    while True:
        process_emails_inbox()
        time.sleep(5)

# Tkinter GUI code starts here
app = tk.Tk()
app.title("Email Scheduler")
app.geometry("900x600")  # Set window size

# Add background image
bg_image = Image.open("Designer.jpg")
bg_image = bg_image.resize((900, 600), Image.LANCZOS)
bg_photo = ImageTk.PhotoImage(bg_image)
bg_label = tk.Label(app, image=bg_photo)
bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# Main frame
main_frame = tk.Frame(app, bg="white")  # Set frame background color
main_frame = tk.Frame(app)
main_frame.pack(side=tk.LEFT, padx=10, pady=10)

# Add widgets to the main frame with respective background and foreground colors

tk.Label(main_frame, text="Receiver Emails (comma separated):", bg="white", fg="black").pack()
receiver_entry = tk.Entry(main_frame, width=50)
receiver_entry.pack()

tk.Label(main_frame, text="Contract Numbers (comma separated):", bg="white", fg="black").pack()
contract_numbers_entry = tk.Entry(main_frame, width=50)
contract_numbers_entry.pack()

tk.Label(main_frame, text="Order Dates (comma separated):", bg="white", fg="black").pack()
order_dates_entry = tk.Entry(main_frame, width=50)
order_dates_entry.pack()

upload_button = tk.Button(main_frame, text="Upload Excel File", command=upload_file, bg="sky blue", fg="black", relief=tk.GROOVE, bd=3)
upload_button.pack(pady=5)

tk.Label(main_frame, text="Time to Send (HH:MM):", bg="white", fg="black").pack()
time_entry = tk.Entry(main_frame, width=20)
time_entry.pack()

tk.Label(main_frame, text="Reminder Time (HH:MM):", bg="white", fg="black").pack()
reminder_time_entry = tk.Entry(main_frame, width=20)
reminder_time_entry.pack()

process_button = tk.Button(main_frame, text="Process", command=process_emails, bg="light green", fg="black", relief=tk.GROOVE, bd=3)
process_button.pack(pady=5)

send_button = tk.Button(main_frame, text="Send", command=send_all_emails, bg="light green", fg="black", relief=tk.GROOVE, bd=3)
send_button.pack(pady=5)

kill_button = tk.Button(main_frame, text="STOP MAIL", command=stop_sending_emails, bg="red", fg="white", relief=tk.GROOVE, bd=3)
kill_button.pack(pady=5)

done_button = tk.Button(main_frame, text="Done", command=close_app, bg="gray", fg="white", relief=tk.GROOVE, bd=3)
done_button.pack(pady=5)

# Create a text widget for logging
log_text = tk.Text(main_frame, height=10, width=70, bg="white", fg="black", wrap=tk.WORD)
log_text.pack(pady=10)

# Run the thread to monitor the inbox for new emails
inbox_thread = threading.Thread(target=monitor_inbox)
inbox_thread.start()

# Start the GUI main loop
app.mainloop()

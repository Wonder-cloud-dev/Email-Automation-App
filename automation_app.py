import pandas as pd  # For working with Excel files
import smtplib  # For sending emails
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, messagebox, Text, Scrollbar, RIGHT, Y, END  # For GUI
from email.mime.text import MIMEText  # For creating the email body
from email.mime.multipart import MIMEMultipart  # For creating multi-part emails
import threading  # To run email-sending operations in a separate thread


class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation Tool")
        self.root.geometry("500x400")

        # Variables
        self.file_path = StringVar()
        self.sender_email = StringVar()
        self.sender_password = StringVar()

        # UI Elements
        Label(root, text="Email Automation Tool", font=("Arial", 16)).pack(pady=10)

        # File selection
        Button(root, text="Select Excel File", command=self.select_file).pack(pady=5)
        Label(root, textvariable=self.file_path, wraplength=300).pack(pady=5)

        # Sender's email
        Label(root, text="Sender's Email:").pack(pady=5)
        Entry(root, textvariable=self.sender_email, width=30).pack(pady=5)

        # Sender's password
        Label(root, text="Email Password:").pack(pady=5)
        Entry(root, textvariable=self.sender_password, show="*", width=30).pack(pady=5)

        # Send emails button
        Button(root, text="Send Emails", command=self.start_sending_emails).pack(pady=10)

        # Log area with scrollbar
        self.log_text = Text(root, wrap="word", height=10, width=60)
        self.log_text.pack(pady=10)

        scrollbar = Scrollbar(root, command=self.log_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def log_message(self, message):
        """Logs a message to the text area."""
        self.log_text.insert(END, f"{message}\n")
        self.log_text.see(END)  # Auto-scroll to the bottom

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        self.file_path.set(file_path)

    def send_email(self, sender_email, sender_password, recipient_email, subject, body):
        try:
            # Set up the email server
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, sender_password)

            # Create the email
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            # Send the email
            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()
            self.log_message(f"[SUCCESS] Email sent to: {recipient_email}")
        except Exception as e:
            self.log_message(f"[ERROR] Could not send email to: {recipient_email}. Error: {e}")

    def send_emails(self):
        """Sends emails using data from the selected Excel file."""
        # Load file and credentials
        file_path = self.file_path.get()
        sender_email = self.sender_email.get()
        sender_password = self.sender_password.get()

        if not file_path or not sender_email or not sender_password:
            messagebox.showerror("Error", "All fields must be filled out.")
            return

        try:
            df = pd.read_excel(file_path)

            # Check required columns
            required_columns = ["RecipientEmail", "Name", "Message"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "Excel file must contain RecipientEmail, Name, and Message columns.")
                return

            # Initialize counters
            total_emails = len(df)
            emails_sent = 0
            errors = []

            self.log_message(f"Starting to send {total_emails} emails...\n")
            
            # Iterate through the data and send emails
            for index, row in df.iterrows():
                recipient_email = row["RecipientEmail"]
                name = row["Name"]
                message = row["Message"]

                subject = f"Hello, {name}!"
                body = f"Dear {name},\n\n{message}\n\nBest regards,\nYour Automation Tool"

                try:
                    self.send_email(sender_email, sender_password, recipient_email, subject, body)
                    emails_sent += 1
                except Exception as e:
                    errors.append((name, recipient_email, str(e)))

            # Summary
            self.log_message("\n--- Email Sending Summary ---")
            self.log_message(f"Total Emails: {total_emails}")
            self.log_message(f"Successfully Sent: {emails_sent}")
            self.log_message(f"Failed: {len(errors)}")

            if errors:
                self.log_message("\n--- Failed Emails ---")
                for error in errors:
                    self.log_message(f"Name: {error[0]}, Email: {error[1]}, Error: {error[2]}")

            # Show messagebox summary
            if emails_sent == total_emails:
                messagebox.showinfo("Success", "All emails were sent successfully!")
            else:
                messagebox.showwarning(
                    "Partial Success",
                    f"Emails sent: {emails_sent}/{total_emails}\nErrors encountered: {len(errors)}"
                )

        except Exception as e:
            self.log_message(f"[ERROR] An error occurred: {e}")

    def start_sending_emails(self):
        """Starts the email-sending process in a separate thread."""
        email_thread = threading.Thread(target=self.send_emails)
        email_thread.daemon = True  # Ensure thread exits when the main program ends
        email_thread.start()


# Run the application
if __name__ == "__main__":
    root = Tk()
    app = EmailAutomationApp(root)
    root.mainloop()

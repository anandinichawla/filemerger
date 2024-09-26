import os
import imaplib
import smtplib
import email
from email import policy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import decode_header
from PIL import Image
from flask import Flask
from dotenv import load_dotenv
from flask_apscheduler import APScheduler
import hashlib
import threading
import PyPDF2

# Load environment variables
load_dotenv()

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Create folders if they don't exist
RECEIVED_FOLDER = "./received"
SENT_FOLDER = "./sent"
PDF_DIR = 'pdf_attachments'
os.makedirs(RECEIVED_FOLDER, exist_ok=True)
os.makedirs(SENT_FOLDER, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

app = Flask(__name__)
scheduler = APScheduler()

# Global lock to ensure only one job runs at a time
lock = threading.Lock()

# Function to calculate hash of image content
def calculate_image_hash(image_path):
    hasher = hashlib.md5()
    with open(image_path, 'rb') as f:
        buf = f.read()
        hasher.update(buf)
    return hasher.hexdigest()

# Function to fetch unseen emails with image and PDF attachments
def fetch_unseen_emails():
    image_paths = []
    pdf_paths = []
    sender_email = None

    # Connect to the email account using IMAP
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        # Search for unseen emails
        status, messages = mail.search(None, "UNSEEN")
        email_ids = messages[0].split()

        if not email_ids:
            print("No new unseen emails.")
            mail.logout()
            return [], [], None

        # Fetch each unseen email
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1], policy=policy.default)

                    subject = msg['Subject']
                    sender_email = msg['From']
                    print(f"Email from: {sender_email} with subject: {subject}")

                    # Decode email subject and body
                    subject_decoded = decode_header(subject)[0][0]
                    subject_str = subject_decoded.decode() if isinstance(subject_decoded, bytes) else subject_decoded
                    body_str = get_email_body_text(msg)

                    # Check if the subject or body contains the keywords
                    if 'combine' in subject_str.lower() or 'merge' in subject_str.lower() or 'combine' in body_str.lower() or 'merge' in body_str.lower():
                        print("Found 'combine' or 'merge' in email. Processing images and PDFs...")

                        if msg.is_multipart():
                            for part in msg.iter_attachments():
                                content_type = part.get_content_type()
                                filename = part.get_filename()

                                if content_type.startswith("image"):
                                    # Save the image inside the "received" folder
                                    filepath = os.path.join(RECEIVED_FOLDER, filename)
                                    with open(filepath, "wb") as img_file:
                                        img_file.write(part.get_payload(decode=True))
                                        image_paths.append(filepath)
                                        print(f"Saved image: {filepath}")
                                elif content_type == "application/pdf":
                                    # Save the PDF inside the "pdf_attachments" folder
                                    filepath = os.path.join(PDF_DIR, filename)
                                    with open(filepath, "wb") as pdf_file:
                                        pdf_file.write(part.get_payload(decode=True))
                                        pdf_paths.append(filepath)
                                        print(f"Saved PDF: {filepath}")
                        # Mark the email as seen
                        mail.store(email_id, '+FLAGS', '\\Seen')
                    else:
                        print("Keywords 'combine' or 'merge' not found. Skipping email.")
                        # Mark the email as seen so it's not processed again
                        mail.store(email_id, '+FLAGS', '\\Seen')
    except Exception as e:
        print(f"Error connecting to email server: {e}")
        return [], [], None

    mail.logout()
    return image_paths, pdf_paths, sender_email

# Function to extract text from email body
def get_email_body_text(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode("utf-8")
    else:
        return msg.get_payload(decode=True).decode("utf-8")
    return ""

# Function to remove duplicate images by comparing their hashes
def remove_duplicate_images(image_paths):
    unique_images = []
    seen_hashes = set()

    for img_path in image_paths:
        img_hash = calculate_image_hash(img_path)
        if img_hash not in seen_hashes:
            unique_images.append(img_path)
            seen_hashes.add(img_hash)
        else:
            print(f"Duplicate image found and ignored: {img_path}")

    return unique_images

# Function to combine the unique images into a single PDF
def combine_images_to_pdf(image_paths, output_pdf="combined_images.pdf"):
    try:
        images = [Image.open(img).convert('RGB') for img in image_paths]
        output_pdf_path = os.path.join(SENT_FOLDER, output_pdf)  # Save PDF in the "sent" folder
        images[0].save(output_pdf_path, save_all=True, append_images=images[1:])
        print(f"Combined image PDF saved at: {output_pdf_path}")
        return output_pdf_path
    except Exception as e:
        print(f"Error creating image PDF: {e}")
        return None

# Function to combine PDFs
def combine_pdfs(pdf_paths, output_pdf):
    try:
        merger = PyPDF2.PdfMerger()
        for pdf_file in pdf_paths:
            merger.append(pdf_file)
        output_pdf_path = os.path.join(SENT_FOLDER, output_pdf)
        with open(output_pdf_path, 'wb') as merged_pdf:
            merger.write(merged_pdf)
        print(f"Combined PDF saved as {output_pdf_path}.")
        return output_pdf_path
    except Exception as e:
        print(f"Error combining PDFs: {e}")
        return None

# Function to send the combined PDF back via email
def send_email_with_attachment(receiver_email, subject, body, file_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with open(file_path, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype="pdf")
            attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
            msg.attach(attach)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL, PASSWORD)
            server.sendmail(EMAIL, receiver_email, msg.as_string())
            print(f"Sent email with combined PDF to {receiver_email}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Function to automate image and PDF processing
def process_email_images_and_pdfs():
    print("Checking for new emails...")

    with lock:
        image_paths, pdf_paths, sender_email = fetch_unseen_emails()

        if not image_paths and not pdf_paths:
            print("No new images or PDFs found.")
            return

        # Remove duplicates from the image paths
        unique_image_paths = remove_duplicate_images(image_paths)
        combined_image_pdf = None 
        combined_pdf_paths = []
        if unique_image_paths:
            # Combine unique images into one PDF
            combined_image_pdf = combine_images_to_pdf(unique_image_paths)
            if combined_image_pdf:
                combined_pdf_paths.append(combined_image_pdf)

        if pdf_paths:
            combined_pdf = combine_pdfs(pdf_paths, "combined_pdfs.pdf")
            if combined_pdf:
                combined_pdf_paths.append(combined_pdf)

         # New Logic: If both an image PDF and other PDFs exist, merge them together
        if combined_image_pdf and pdf_paths:
            print("Merging image PDF and other PDFs together...")
            all_pdfs_to_merge =  pdf_paths + [combined_image_pdf]
            final_combined_pdf = combine_pdfs(all_pdfs_to_merge, "final_combined_output.pdf")
            combined_pdf_paths = [final_combined_pdf]

        if combined_pdf_paths:
            # Prepare the email body
            body_message = "Here are your combined files."
            for pdf in combined_pdf_paths:
                send_email_with_attachment(
                    receiver_email=sender_email,
                    subject="Your Combined Files",
                    body=body_message,
                    file_path=pdf
                )
            print("Processing complete and emails sent!")
        else:
            print("No files to combine.")

# Scheduler job to periodically check for new emails
@scheduler.task('interval', id='check_email_task', seconds=60)
def scheduled_email_check():
    process_email_images_and_pdfs()

if __name__ == '__main__':
    scheduler.init_app(app)
    scheduler.start()
    app.run(port=5000)

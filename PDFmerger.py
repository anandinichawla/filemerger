

import imaplib
import time 
import email
from email.header import decode_header
import os
import PyPDF2
print(PyPDF2.__version__)
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO


#email account credentials 

sender_email="achawla6061@gmail.com"
my_password="obui hftq bloj rdwp"
receiver_email="anandini@vara.eco"
smtp_server="smtp.gmail.com" 
smtp_port=587   # Use port 465 for SSL


#body of email 
subject="returning your combined pdf" 
body="here is your combined pdf." 


# Keywords to trigger the script
TRIGGER_KEYWORD = "combine" 







# Directory to save the PDF
SAVE_DIR = 'pdf_attachments'



#Directory where the PDFs are stored
PDF_DIR = 'pdf_attachments'  # Replace with the directory where your PDFs are stored
OUTPUT_PDF = 'combined.pdf'  # The name of the final combined PDF

def combine_pdfs(pdf_dir,output_pdf):

    print("i am in combine pdf")

    # Create a PDF merger object 

    merger = PyPDF2.PdfMerger()

    # Get the list of PDF files in the directory
    pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]

    print("here are the pdf files")
    print(pdf_files); 

    for pdf_file in pdf_files: 
        file_path = os.path.join(pdf_dir, pdf_file)
        print(f"filepath={file_path}")
        print(f"Adding {pdf_file} to the combined PDF")
        with open(file_path, 'rb') as f:
            merger.append(file_path)

    print("wrote all the files ")
    print(f"output_pdf= {output_pdf}")
    with open(output_pdf, 'wb') as merged_pdf:
        merger.write(merged_pdf)
    # Write the combined PDF to an output file 
    # with open(output_pdf,'wb') as output_file:
    #     print(f"output_file= {output_file}")
    #     merger.write(output_file)

    # Close the merger object 
    merger.close()


    print(f"Combined PDF saved as {output_pdf}.")





def send_email_with_attachment(file_path):

    print("i am in email")
    print(file_path)
    try:
        # Create email message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))



        with open(file_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="pdf")
            part.add_header('Content-Disposition', f'attachment; filename="{file_path}"')
            msg.attach(part)
        
        # Establish a connection to the SMTP server
        with smtplib.SMTP(smtp_server, 587) as server:
            server.starttls()  # Upgrade the connection to TLS (for secure communication)
            server.login(sender_email,my_password)  # Replace 'your_password' with actual password
            server.sendmail(sender_email, receiver_email, msg.as_string())
            print("Email sent successfully!")
            

    except Exception as e:
        print(f"Failed to send email: {e}")







def fetch_pdfs():

    flag = 0 

    # Connect to the IMAP server
    mail = imaplib.IMAP4_SSL(smtp_server)
    mail.login(sender_email, my_password)
    mail.select('inbox')  # Select the inbox folder
    print("success")

    # Search for all emails
    status, data = mail.search(None, "ALL")
    email_ids = data[0].split()

    if not email_ids:
        print("No emails found.")
        mail.logout()


    # Fetch the most recent email (for simplicity)
    latest_email_id = email_ids[-1]
    status, data = mail.fetch(latest_email_id, '(RFC822)')
    raw_email = data[0][1]
    
    # Parse the email content
    msg = email.message_from_bytes(raw_email)

        # Decode email subject
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding if encoding else 'utf-8')
            
    print(f"Subject: {subject}")

    # Decode email body
    if msg.is_multipart():
        print("in message")

        for part in msg.walk():
            print("in for again")
                
            content_type = part.get_content_type()
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode()
                if TRIGGER_KEYWORD.lower() in body.lower():  
                    flag = 1 
                    print("trigger word is there and flaf is now one")
                    break 


        if(flag == 1):
            print("in flag 1 function")

            for part in msg.walk(): 
                content_disposition = str(part.get("Content-Disposition"))
                # print(f"Content-Type: {content_type}")
                print(f"Content-Disposition: {content_disposition}")
                if content_disposition and "attachment" in content_disposition: 
                #Get the attachment filename 
                    filename = part.get_filename() 
                    print(filename)
                    if filename and filename.endswith(".pdf"):
                        print("saving the pdf")
                        #Save the PDF to the specified directory 
                        filepath = os.path.join(SAVE_DIR, filename)
                        with open(filepath, "wb") as f: 
                            f.write(part.get_payload(decode=True))
                            print(f"Saved PDF: {filename}")

            return 1

        else:
            print("no trigger")
    

      

   


    else:
        payload = msg.get_payload(decode=True) 
        if payload is not None: 
            body = payload.decode()
            print("no attachments found in this email") 
            print(f"Body: {body}")
        else: 
            print("none")

    return 0

        # Logout from the email server
    mail.logout()

# function fetch_pdf ends here 

# Run all the functions 
if __name__ == "__main__":

    while True: 
        val = fetch_pdfs()

        if(val == 1):
            combine_pdfs(PDF_DIR, OUTPUT_PDF)
            send_email_with_attachment(
                file_path= OUTPUT_PDF  
            )
        time.sleep(5)







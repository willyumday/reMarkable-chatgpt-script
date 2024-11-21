import imaplib
import smtplib
import openai
from google.cloud import storage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from fpdf import FPDF
import os

# Set up email and OpenAI API credentials
EMAIL = "@gmail.com"
PASSWORD = ""
OPENAI_API_KEY = ""
BUCKET_NAME = ""  # Replace with your bucket name

# Set up OpenAI API key
openai.api_key = OPENAI_API_KEY

# Initialize Google Cloud Storage client
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\willy\"  # Update with the path to your JSON key file
storage_client = storage.Client()

print("Starting script...")
print(f"Email account: {EMAIL}")

# Connect to IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")
imap.login(EMAIL, PASSWORD)

# Fetch the latest email
imap.select("inbox")
status, messages = imap.search(None, "UNSEEN")
email_ids = messages[0].split()

if email_ids:
    print(f"Found {len(email_ids)} unseen emails.")
    for email_id in email_ids:
        res, msg = imap.fetch(email_id, "(RFC822)")
        for response_part in msg:
            if isinstance(response_part, tuple):
                # Parse email content
                email_content = response_part[1].decode("utf-8")
                
                # Send the content to ChatGPT
                response = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo",  # You can also use "gpt-4" if needed
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant."},
                        {"role": "user", "content": email_content}
                    ],
                    max_tokens=500
                )

                # Save the response as a PDF file
                response_text = response["choices"][0]["message"]["content"]
                pdf_file_name = "response.pdf"
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.multi_cell(0, 10, response_text)
                pdf.output(pdf_file_name)
                print(f"Generated PDF: {pdf_file_name}")

                # Upload the PDF file to Google Cloud Storage
                bucket = storage_client.bucket(BUCKET_NAME)
                blob = bucket.blob(pdf_file_name)
                blob.upload_from_filename(pdf_file_name)
                print(f"Uploaded {pdf_file_name} to {BUCKET_NAME}")

                # Optionally, email the response back with the PDF attachment
                smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
                smtp.login(EMAIL, PASSWORD)

                # Create a multipart message
                msg = MIMEMultipart()
                msg["From"] = EMAIL
                msg["To"] = EMAIL
                msg["Subject"] = "Your ChatGPT Response"

                # Attach the email body
                msg.attach(MIMEText("Please find the AI-generated response attached as a PDF."))

                # Attach the PDF file
                with open(pdf_file_name, "rb") as pdf_file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(pdf_file.read())

                # Encode the payload in Base64
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={pdf_file_name}",
                )
                msg.attach(part)

                # Send the email
                smtp.sendmail(EMAIL, EMAIL, msg.as_string())
                smtp.quit()
                print("Response email sent successfully.")

else:
    print("No unseen emails found.")

# Close connections
imap.logout()

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


def Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path):
    sender_email = "ybot@i-t-g.net"
    smtp_server = "email-smtp.us-east-2.amazonaws.com"
    smtp_port = 587  
    smtp_username = "AKIAXGIXUMEOTGEGHWYW"
    smtp_password = "BIC8sGYiSPduHzFArWtGw/j27gKWPNYp1xz0kDHnZVr5"

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_emails
    message['CC'] = cc_emails
    message['Subject'] = subject

    body_text = f"""Hi Team,
    
{body_message}
{body_message1}
                       
Regards,
Onboarding Team."""
    message.attach(MIMEText(body_text, 'plain'))
    
    if attachment_path:
        try:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            attachment_name = os.path.basename(attachment_path)
            part.add_header('Content-Disposition', f'attachment; filename= {attachment_name}')
            message.attach(part)
        except Exception as e:
            print("Error adding attachment:", str(e))
   
    
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, recipient_emails.split(",") + cc_emails.split(","), message.as_string())
    except Exception as e:
        print("An error occurred:", str(e))
    finally:
        server.quit()

if __name__ == "__main__":
    recipient_emails = "Recepients"
    cc_emails = "CC"
    subject = "Subject"
    body_message = "Body"
    body_message1 = "Body1"
    body_message2 = "Body2"
    body_message3 = "Body3"
    attachment_path = "Attachment"

    Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)




import os
import fetcher

from dotenv import load_dotenv

USERNAME = os.getenv("EMAIL_ADDRESS")
PASSWORD = os.getenv("EMAIL_PASSWORD")
OUTLOOK_IMAP = "outlook.office365.com"

def main():
    load_dotenv()
    
    imap = fetcher.login(USERNAME, PASSWORD, OUTLOOK_IMAP)
    
    sender, subject, body = fetcher.fetch_email(1, imap)
    print(f"Sender: {sender}")
    print(f"Subject: {subject}")
    print(f"Body: {body}")
    
    fetcher.logoutAndClose(imap)

if __name__ == "__main__":
    main()

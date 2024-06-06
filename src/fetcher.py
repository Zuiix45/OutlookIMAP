import imaplib
import email
from email.header import decode_header
import os

def login(username, password, imap_server):
    imap = imaplib.IMAP4_SSL(imap_server)
    
    imap.login(username, password)
    
    return imap

def decodeH(header):
    header_part, encoding = decode_header(header)[0]
    if isinstance(header_part, bytes):
        header_part = header_part.decode(encoding) # if it's a bytes, decode to str
    return header_part

def fetch_email(index, imap):
    imap.select("INBOX")
    
    res, msg = imap.fetch(str(index), "(RFC822)")
    
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            
            # decode the email subject and sender
            subject = decodeH(msg.get("Subject"))
            sender_address = decodeH(msg.get("Return-Path"))
                
            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    # download attachment
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        
                        if filename:
                            if not os.path.isdir("attachments"):
                                # make a folder for this email (named after the subject)
                                os.mkdir("attachments")
                            filepath = os.path.join("attachments", filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
                    
                    elif content_type == "text/plain":
                        body = part.get_payload(decode=True).decode()
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                
                # get the email body
                if content_type == "text/plain":
                    body = msg.get_payload(decode=True).decode()
                    
            if content_type == "text/html":
                if not os.path.isdir("attachments"):
                    # make a folder for this email (named after the subject)
                    os.mkdir("attachments")
                filepath = os.path.join("attachments", filename)
                # download attachment and save it
                open(filepath, "w").write(body)
                
            return sender_address, subject, body

def logoutAndClose(imap):
    imap.close()
    imap.logout()
    
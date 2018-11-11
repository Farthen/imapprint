from credentials import USERNAME, PASSWORD, HOSTNAME, PRINTERNAME

import email
import imaplib
import os
import pypandoc
import hashlib
import subprocess

class FetchEmail():

    connection = None
    error = None

    def __init__(self, mail_server, username, password):
        self.connection = imaplib.IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select(readonly=False) # so we can mark mails as read

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()
        self.connection.logout()

    def save_attachment(self, msg, download_folder="/tmp"):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)

        return: file path to attachment
        """
        att_paths = []
        for part in msg.walk():
            att_path = None
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                if part.get_content_type() != 'application/octet-stream':
                    continue

            filename = part.get_filename()
            if filename is None:
                continue
            base, ext = os.path.splitext(filename)
            ext = ext.lower()
            if ext not in ['.asc', '.sig', '.gpg']:
                att_data = part.get_payload(decode=True)
                shasum = hashlib.sha256(att_data).hexdigest()
                base = base + '-' + shasum[:10]
                att_path = os.path.join(download_folder, base + ext)

                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    fp.write(att_data)
                    fp.close()
                if ext not in ['.pdf', '.ps']:
                    try:
                        outputfilename = os.path.join(download_folder, base + '.pdf')
                        fmt = None
                        if ext == ".txt":
                            fmt = 'md'
                        pypandoc.convert_file(att_path, 'pdf', format=fmt, outputfile=outputfilename)
                        os.remove(att_path)
                        att_path = outputfilename
                    except RuntimeError:
                        os.path.rm(att_path)
                        att_path = None
            if att_path is not None:
                att_paths.append(att_path)
        return att_paths

    def fetch_unread_messages(self):
        """
        Retrieve unread messages
        """
        emails = []
        (result, data) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            for message in data[0].split():
                try:
                    ret, data = self.connection.fetch(message,'(RFC822)')
                except:
                    print("Can't fetch message {}".format(message))
                    continue

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                response, data = self.connection.store(message, '+FLAGS','\\Seen')

            return emails

        self.error = "Failed to retrieve emails."
        return emails

    def parse_email_address(self, email_address):
        """
        Helper function to parse out the email address from the message

        return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
        """
        return email.utils.parseaddr(email_address)

def print_file(filename):
    subprocess.call(["/usr/bin/lp", "-d", PRINTERNAME, filename])
    os.remove(filename)

if __name__ == "__main__":
    fetchmail = FetchEmail(HOSTNAME, USERNAME, PASSWORD)
    msgs = fetchmail.fetch_unread_messages()
    for msg in msgs:
        att_paths = fetchmail.save_attachment(msg, './attachments/')
        for att in att_paths:
            print_file(att)
    fetchmail.close_connection()


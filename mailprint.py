from credentials import USERNAME, PASSWORD, HOSTNAME, PRINTERNAME

import email
import imaplib
import os
import pypandoc
import hashlib
import subprocess
import time
import glob

os.environ.setdefault('PYPANDOC_PANDOC', '/usr/local/bin/pandoc')

LIBREOFFICE_EXT = [".ods", ".ots", ".fods", ".uos", ".xlsx", ".xml", ".xls", ".xlt", ".dif", ".dbf", ".html", ".slk", ".csv", ".xlsx", ".xlsm", ".odt", ".ott", ".fodt", ".uot", ".docx", ".xml", ".doc", ".dot", ".html", ".rtf", ".docx", ".odf", "mml", "odp", "otp", "odg", "fodp", "uop", "pptx", "ppsx", "potm", "ppt", "pps", "pot", "pptx", "odg", "otg", "fodg", "odm,", "oth,", "odb,", "oxt,", "otf"]
IMAGEMAGICK_EXT = [".jpg", ".jpeg", ".png", ".gif", ".tiff", ".bmp"]

def subprocess_execute(command, time_out=60):
    """executing the command with a watchdog"""

    # launching the command
    c = subprocess.Popen(command)

    # now waiting for the command to complete
    t = 0
    while t < time_out and c.poll() is None:
        time.sleep(1)  # (comment 1)
        t += 1

    # there are two possibilities for the while to have stopped:
    if c.poll() is None:
        # in the case the process did not complete, we kill it
        c.terminate()
        # and fill the return code with some error value
        returncode = -1  # (comment 2)

    else:                 
        # in the case the process completed normally
        returncode = c.poll()

    return returncode   

class ConnectionError(Exception):
    pass

class FetchEmail():

    connection = None
    error = None

    def __init__(self, mail_server, username, password):
        try:
            self.connection = imaplib.IMAP4_SSL(mail_server)
            self.connection.login(username, password)
            self.connection.select(readonly=False) # so we can mark mails as read
        except OSError:
            raise ConnectionError()

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
                print("Downloading attachment '{}'".format(filename))
                att_data = part.get_payload(decode=True)
                shasum = hashlib.sha256(att_data).hexdigest()
                base = base + '-' + shasum[:10]
                att_path = os.path.join(download_folder, base + ext)

                if not os.path.isfile(att_path):
                    fp = open(att_path, 'wb')
                    fp.write(att_data)
                    fp.close()
                if ext not in ['.pdf', '.ps']:
                    outputfilename = os.path.join(download_folder, base + '.pdf')
                    try:
                        print("Converting file '{}{}' to PDF".format(base, ext))
                        if ext in LIBREOFFICE_EXT:
                            print("Converting with LibreOffice")
                            timeout = 25
                            ret = subprocess_execute(["/usr/bin/libreoffice", "--headless", "--convert-to", "pdf", "--outdir", download_folder, att_path], time_out=timeout)
                            if ret == -1:
                                raise RuntimeError("Failed to convert with libreoffice. Timeout after {} seonds".format(timeout))
                            if ret != 0:
                                raise RuntimeError("Failed to convert with libreoffice. Return code: {}".format(ret))
                            if not os.path.exists(outputfilename):
                                raise RuntimeError("Failed to convert with libreoffice. Unknown Error. PDF file was not created.")
                        elif ext in IMAGEMAGICK_EXT:
                            print("Converting with imagemagick")
                            timeout = 25
                            ret = subprocess_execute(["/usr/bin/convert", att_path, outputfilename], time_out=timeout)
                            if ret == -1:
                                raise RuntimeError("Failed to convert with imagemagick. Timeout after {} seonds".format(timeout))
                            if ret != 0:
                                raise RuntimeError("Failed to convert with imagemagick. Return code: {}".format(ret))
                            if not os.path.exists(outputfilename):
                                raise RuntimeError("Failed to convert with imagemagick. Unknown Error. PDF file was not created.")
                        else:
                            print("Converting with pandoc")
                            fmt = None
                            if ext == ".txt":
                                fmt = 'md'
                            pypandoc.convert_file(att_path, 'pdf', format=fmt, outputfile=outputfilename)
                        print("Successfully created '{}'. Removing source file.".format(outputfilename))
                        os.remove(att_path)
                        att_path = outputfilename
                    except RuntimeError as e:
                        print("Error converting to PDF. Removing file. Error: {}".format(e))
                        try:
                            os.remove(att_path)
                        except OSError:
                            pass
                        try:
                            os.remove(outputfilename)
                        except OSError:
                            pass
                        for f in glob.glob(download_folder + "*.tmp"):
                            try:
                                os.remove(f)
                            except OSError:
                                pass
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
    print("Printing file: '{}'".format(filename))
    subprocess.call(["/usr/bin/lp", "-o", "fit-to-page", "-o", "media=A4", "-d", PRINTERNAME, filename])
    os.remove(filename)

if __name__ == "__main__":
    fetchmail = FetchEmail(HOSTNAME, USERNAME, PASSWORD)
    msgs = fetchmail.fetch_unread_messages()
    for msg in msgs:
        att_paths = fetchmail.save_attachment(msg, './attachments/')
        for att in att_paths:
            print_file(att)
    fetchmail.close_connection()


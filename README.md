# IMAP Mail Printer

This package looks for new messages in an email account and prints all attachments.

You can forward any mails with attachments to this account to have them printed out. You just need to run this script regularly (for instance every minute)

## Installation

This package requires pandoc and pypandoc. The more recent the version of pandoc the better.

Copy credentials.py.template to credentials.py and modify to your needs. Then run mailprint.py in a cron job.

## Behavior

All unread mail will be downloaded and printed once. After downloading the message will be marked as read so it will not be printed again.


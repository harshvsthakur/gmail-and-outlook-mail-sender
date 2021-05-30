# Send emails to all accounts in excel sheet column

Pre-requisite : Make sure third party log in is enabled on your gmail account.

Problem : 
I wanted to send marketing mails to a list of companies I had along with their email information on an excel sheet.

Solution:

- This code enables to batch send emails.

- The list of emails is read from an excel sheet column.

- The body of the mail can be edited by editing the included mail1.txt file. The code writes what ever is in the text file as the body of the mail making entering line spaces easier.

- Attachment can be added by editing the path. (PDF file in this example)

- Can easily edit Subject and CC fields as per requirement. 


Note: 

1. The decision to keep this as a text email is delibrate. By default outlook disables html mails and I wanted to make sure this email reaches the recipients.

2. I believe there is a limit to the number of mails gmail can send - 500 mails per day  as of 30/05/20. I only needed to send approximately 400 mails so this worked perfectly for me.  There are work arounds if you have more than 500 recipients to send mails to. 

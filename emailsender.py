'''
Email:

Hi,

There are currently {#no_of_invoices} in your Stampli cue. Could you please review and approve so we can process them for payment?

Thank you, 


***Need script to send email based on excel spreadsheet provided by user

'''
import smtplib
import csv
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def get_contact(filename):
	emails = []
	counts = []

	with open(filename) as file:
		contactlist = csv.reader(file)
		next(contactlist)
		for email,count in contactlist:
			emails.append(email)
			counts.append(count)

	return emails, counts

def read_template(filename):
	with open(filename, 'r', encoding='utf-8') as template_file:
		template_file_content = template_file.read()
	return Template(template_file_content)

	
def main():
    emails, counts = get_contact('test_contact.csv') # read contacts
    message_template = read_template('StampliEmail.txt')

    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()

    USERNAME = input('Enter email address: ')
    PASSWORD = input('What is password to ' + USERNAME + ": ")

    s.login(USERNAME, PASSWORD)

    # For each contact, send the email:
    for email, count in zip(emails, counts):
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(no_of_invoice=count.title())

        # Prints out the message body for our sake
        print(message)

        # setup the parameters of the message
        msg['From']=USERNAME
        msg['To']=email
        msg['Subject']="URGENT: Please approve your invoices"
        
        # add in the message body
        msg.attach(MIMEText(message, 'plain'))
        
        # send the message via the server set up earlier.
        s.send_message(msg)
        del msg
        
    # Terminate the SMTP session and close the connection
    s.quit()


if __name__ == '__main__':
    main()

print('done')
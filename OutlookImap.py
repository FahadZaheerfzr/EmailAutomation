from http import server
import imaplib
import smtplib
import email
from email.header import decode_header
import getpass


username = input("Enter Outlook Email: ")
password = getpass.getpass("Enter password for this email: ")
# account credentials
# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
# for office 365, it's this:
imap_server = "outlook.office365.com"
smtp_server = "smtp-mail.outlook.com"


# create an IMAP4 class with SSL 
imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
imap.login(username, password)


smtp = smtplib.SMTP(smtp_server, 587)
smtp.starttls()
smtp.login(username, password)




data_dic = {}
with open("config.ini", "r") as f:
        lines = f.readlines()
        for data in lines:
            data = data.strip()
            key_vals = data.split(":")
            data_dic[key_vals[0]] = key_vals[1]



reference_numbers = []
references_exist = False
try:
    with open("reference_numbers.txt", "r") as f:
        reference_numbers = f.readlines()
        reference_numbers = reference_numbers[-1]
        print(reference_numbers)
        references_exist = True
except FileNotFoundError:
    pass


prefixes = []
try:
    with open("prefixes.txt", "r") as f:
        prefixes = f.readlines()
except FileNotFoundError:
    pass



first = True
difference = 0
top_msg = 0
is_reply = False
while True:
    status, messages = imap.select("INBOX")
    top_msg+= difference
    difference = 0
    # number of top emails to fetch
    messages = int(messages[0])
    for i in range(messages, top_msg, -1):
        print("Currently id of the last message is: " + str(i))
        if first:
            top_msg = i
            first = False
            break
        
        if i == top_msg:
            break

        if references_exist:
            n= int(data_dic["Number_of_Digits"])-len(str(int(reference_numbers) + 1))
            reference_numbers = "0" * n +str(int(reference_numbers) + 1)
            refernece_string = "[REF-" + data_dic["PREFIX"] + "-" + reference_numbers + "]"
        else:
            n= int(data_dic["Number_of_Digits"])-len(data_dic["Start Number"])
            reference_numbers = "0" * n + data_dic["Start Number"]
            refernece_string = "[REF-" + data_dic["PREFIX"] + "-" + reference_numbers + "]"

        to_mail = data_dic["To Mail"].split(",")
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                    
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if "REF" in subject:
                    is_reply = True
                    break
                msg.replace_header("From", username)
                msg.replace_header("Subject", refernece_string+subject)
                
                for mail_address in to_mail:
                    msg.replace_header("To", mail_address.strip())
                    smtp.sendmail(username, mail_address, msg.as_string())
        

        if is_reply:
            is_reply = False
            break
        difference+=1
        with open("reference_numbers.txt", "a") as f:
            f.write(str(reference_numbers) + "\n")
        with open("prefixes.txt", "a") as f:
            f.write(data_dic["PREFIX"] + ", ")

smtp.quit()

# close the connection and logout
imap.close()
imap.logout()
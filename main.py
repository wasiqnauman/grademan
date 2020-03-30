import smtplib
import pandas as pd
from email.message import EmailMessage
from getpass import getpass



def getEmail(filename):
    """
    reads emails from excel file and returns them as a list
    """

    data = pd.read_excel(filename)
    emails = data['Email'].to_numpy()
    return emails

def createMsg(to, frm, content, *args):
    """
    fills in the headers for a message and returns it
    """

    msg = EmailMessage()
    msg['To'] = to
    msg['From'] = frm
    msg['Subject'] = 'Grades'
    msg.set_content(content)
    if args:   # check if CC list is provided as a paramater
        msg['CC'] = args[0]
    return msg

def initSMTP():
    """
    establishes connection to the SMTP server and returns the object
    ! not sure if this a good practice, let me know what you think !
    """

    mail = smtplib.SMTP('smtp.gmail.com', 587)   # connect to gmail SMTP server
    code = mail.ehlo()[0]   # get the return code
    if code == 250:
        print("Connection is successful!")
    else:
        print(f"Something went wrong! Error: {code}")
    mail.starttls()   # establish secure TLS encryption
    email = input('Email: ')
    passwd = getpass()   # Does not work on IDLE, use the terminal!
    mail.login(email, passwd)
    return mail


def main():
    filename = 'emails.xlxs'
    # to = getEmail(filename)
    frm = 'test@example.com'
    name = 'Person'
    g = [100,100,100]
    msgContent = \
    f"""Hi {name},
    Your grades for this semester uptil now are as follows
    {g[0]} {g[1]} {g[2]}
Love,
Grademan"""
    print(msgContent)
    mail = initSMTP()

if __name__ == "__main__":
    main()

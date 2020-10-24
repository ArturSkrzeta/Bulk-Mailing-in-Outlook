import csv
import time
import win32com.client as client

def get_recipients():
    with open('recipients.csv', newline='') as f:
        reader = csv.reader(f)
        recipients = [row for row in reader]
    return recipients

def get_template():
    with open('./templates/template.html', 'r') as myfile:
        template=myfile.read()
    return template

def save_drafts(outlook, recipients, template):
    for name, address in recipients:
        mail = outlook.CreateItem(0)
        mail.To = address
        mail.Subject = "Training is coming!"
        mail.HTMLBody = template.format(name.split(" ")[0])
        mail.Save()

def send_emails_in_chunks(drafts):
    chunks = [drafts[x:x+30] for x in range(0, len(drafts), 30)]
    for chunk in chunks:
        for message in chunk:
            message.Send()
        time.sleep(60)

def main():

    outlook = client.Dispatch('Outlook.Application')
    save_drafts(outlook, get_recipients(), get_template())

    drafts = outlook.GetNameSpace('MAPI').GetDefaultFolder(16)
    send_emails_in_chunks(list(drafts.Items))


if __name__ == '__main__':
    main()

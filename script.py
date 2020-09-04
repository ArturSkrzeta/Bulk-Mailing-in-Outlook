import csv
import time
import win32com.client as client

template ="""Hi {},

please schedule your spot for training.

Thanks"""

def get_recipients():
    with open('recipients.csv', newline='') as f:
        reader = csv.reader(f)
        recipients = [row for row in reader]
    return recipients

def save_drafts(recipients, outlook, template):
    for name, address in recipients:
        message = outlook.CreateItem(0)
        message.To = address
        message.Subject = "Training is coming!"
        message.Body = template.format(name)
        message.Save()

    namespace = outlook.GetNameSpace('MAPI')
    drafts = namespace.GetDefaultFolder(16)
    return list(drafts.Items)

def send_emails_in_chunks(drafts):
    chunks = [drafts[x:x+30] for x in range(0, len(drafts), 30)]
    for chunk in chunks:
        for message in chunk:
            ,message.Send()
        time.sleep(60)

def main():

    outlook = client.Dispatch('Outlook.Application')
    recipients = get_recipients()
    drafts = save_drafts(recipients, outlook, template)
    send_emails_in_chunks(drafts)


if __name__ == '__main__':
    main()

import win32com.client
import scraper




def send_mail(creds, to, subject, body, html_body=""):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNameSpace("MAPI")
    namespace.Logon(creds["long_username"], creds["password"])

    msg = outlook.CreateItem(0)
    print(to)
    msg.To = 'tony.bridges@wi.gov'
    msg.Subject = subject
    msg.Body = body
    if html_body:
        msg.HTMLBody = html_body
    msg.ReplyRecipients.Add("tbridges42@gmail.com")
    msg.SentOnBehalfOfName = "GABMove@wisconsin.gov"

    msg.Send()


def main():
    send_mail(scraper.get_creds(), '', 'test', 'test')


if __name__ == "__main__":
    main()

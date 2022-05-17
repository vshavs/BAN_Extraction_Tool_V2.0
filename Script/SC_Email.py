import win32com.client as wincl


def email_file(email, sub, abzipfilepath):
    outlook = wincl.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0x0)
    mail.To = email
    mail.Subject = 'Extract_V2_{}'.format(sub)
    mail.HTMLBody = '''<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Hi,</p>
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Please find attached extract for BAN as requested</p>
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Regards,</p>
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Vinay Sharma</p>'''
    # mail.Body = "This is the normal body"
    mail.Attachments.Add(abzipfilepath)
    # mail.Attachments.Add('c:\\sample2.xlsx')
    mail.BCC = 'FACTAutomation@birlasoft.com'
    mail.Send()
    print("email send")

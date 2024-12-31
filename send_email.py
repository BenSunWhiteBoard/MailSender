import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'TianlongLiangSDE@hotmail.com'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

# To attach a file to the email (optional):
attachment = r'G:\Playground\test.txt'
mail.Attachments.Add(attachment)

mail.Send()
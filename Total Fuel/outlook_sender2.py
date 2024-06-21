# import win32com.client as win32

# olAPP = win32.Dispatch("Outlook.Application")
# olNS = olAPP.GetNameSpace("MAPI")

# mailItem = olAPP.CreateItem(0)
# mailItem.Subject = "Hello 123"

# mailItem.BodyFormat =1
# mailItem.Body ="Hello There"

# mailItem.To = "pokuttu@yalelo.ug"
# mailItem._oleobj_.Invoke(*(64209,0,8,0,olNS.Accounts.Item("pokuttu@yalelo.ug")))

# mailItem.Display()
# mailItem.Save()
# mailItem.Send()
# # mailItem.BodyFormat = 2
# # mailItem.HTMLBody= "<HTML Markup>"
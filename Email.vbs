'  3 = "Deleted Items"
'  4 = "Outbox"
'  5 = "Sent Items"
'  6 = "Inbox"
'  9 = "Calendar"
' 10 = "Contacts"
' 11 = "Journal"
' 12 = "Notes"
' 13 = "Tasks"
' 15 = "Reminders"
' 16 = "Drafts"

' the connection to Outlook application
Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(6) 'Inbox

Dim subject
subject = "Weekly Learning Digest"

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[Unread]=true") ' reading of unread mails
Set colFilteredItems = colFilteredItems.Restrict("[Subject] = " & subject)

MsgBox colFilteredItems.Count

Dim countItem
For countItem = colFilteredItems.Count To 1 Step -1
  Set itemEmail  = colFilteredItems.Item(countItem)
  MsgBox itemEmail.body
Next 
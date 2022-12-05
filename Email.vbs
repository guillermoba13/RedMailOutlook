' `````` the property's for obtain information at depending the folder in email
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

Dim countItem, emailSubject, emailFrom, emailTo, emailCc, item, count
Dim emailBcc, emailMessage, emailReceivedTim, emailReceivedDat, emailSentTime, emailSentDate
Dim arrayExcel()

ReDim arrayExcel(colFilteredItems.Count, 9)
count = 0
countAllEmailItem = colFilteredItems.Count + 1

For countItem = colFilteredItems.Count To 1 Step -1
    ' asigna value the geting items in outlook application
    Set itemEmail  = colFilteredItems.Item(countItem)

    arrayExcel(count,0) = itemEmail.Subject
    arrayExcel(count,1) = itemEmail.SenderEmailAddress
    arrayExcel(count,2) = itemEmail.To
    arrayExcel(count,3) = itemEmail.Cc
    arrayExcel(count,4) = itemEmail.Bcc
    arrayExcel(count,5) = itemEmail.Body
    arrayExcel(count,6) = itemEmail.ReceivedTime
    arrayExcel(count,7) = itemEmail.ReceivedTime
    arrayExcel(count,8) = itemEmail.ReceivedTime
    arrayExcel(count,9) = itemEmail.ReceivedTime
    count = count + 1
Next 

' create file
Dim pathFileSave

pathFileSave = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\"
nameFileSave = "Test.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Add()
Set objSheet = objWorkBook.WorkSheets(1)

' header
objSheet.Range("A1") = "Subject"
objSheet.Range("B1") = "SenderEmailAddress"
objSheet.Range("C1") = "To"
objSheet.Range("D1") = "Cc"
objSheet.Range("E1") = "Bcc"
objSheet.Range("F1") = "Body"
objSheet.Range("G1") = "ReceivedTime"
objSheet.Range("H1") = "ReceivedDate"
objSheet.Range("I1") = "SendTime"
objSheet.Range("J1") = "SendDate"

objSheet.Range("A2:J"&countAllEmailItem) = arrayExcel

objWorkbook.SaveAs pathFileSave & nameFileSave

objWorkbook.Close
objExcel.Workbooks.Close
objExcel.Quit
Set objExcel = Nothing
Function formatDTYyyyymmdd(dateTime)
    formatDTYyyyymmdd = "(" & FormatDateTime(dateTime) & ")"
End Function ' formatDTYyyyymmdd

Function WriteLog(logLocation, lineType, message)
    If logLocation <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(logLocation) = False Then
            Set logFile = fso.CreateTextFile(logLocation)
        Else
            Set logFile = fso.OpenTextFile(logLocation, 8, False, 0)
        End If
        logFile.WriteLine(formatDTYyyyymmdd(Now) & " - " & lineType & " - " & message)
        logFile.Close
    End If
End Function ' WriteLog

Function CloseExcelInstance(infoLogFile, errorLogFile, nameScript)
    Const strComputer = "."
    Const findProc = "EXCEL.EXE"

    Set objWMIService = GetObject("Winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set processList = objWMIService.ExecQuery("Select Name from Win32_Process WHERE Name='" & findProc & "'")

    If processList.Count > 0 Then
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "taskkill /f /im excel.exe"
        If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Killing any lining Excel instance"
    Else
        If errorLogFile <> "" Then WriteLog errorLogFile, "INFO", nameScript & "It was not possible to close the Excel instance"
    End If
    Set objShell = Nothing
    Set objWMIService = Nothing
    Set processList = Nothing
End Function ' CloseExcelInstance

Function CreateExcelInstance(infoLogFile, errorLogFile, nameScript)
    Dim excelObj, counter
    counter = 0
    Do While (counter < 5)
        If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Attempt number " & counter & "CreateExcelInstance"
        Err.Clear
        On Error Resume Next
        Set excelObj = CreateObject("Excel.Application")
        If Err.Number = 0 Then
            If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Success Excel instance creation"
            Exit Do
        Else
            If infoLogFile <> "" Then WriteLog infoLogFile, "WARNING", nameScript & "Excel instance creation fail"
            Dim oShell : Set oShell = CreateObject("WScript.Shell")
            If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Killing any lining Excel instance"
            oShell.Run "taskkill /f /im excel.exe"
            counter = counter + 1
        End If
    Loop
    Set CreateExcelInstance = excelObj
End Function ' CreateExcelInstance

Function RunTimeMail(nameScript, infoLogFile, errorLogFile, pathFileSave, nameFileSave, subject, _
    fristRowHeader, letterSubject, letterColumnSenderEmailAddress, letterColumnTo, letterColumnCc, letterColumnBcc, _
    letterColumnBody, letterColumnReceivedTime, letterColumnReceivedDate, letterColumnSendTime, letterColumnSendDate, valueSubject, _
    valueSenderEmailAddress, valueTo, valueCc, valueBcc, valueBody, valueReceivedTime, valueReceivedDate, valueSendTime, valueSendDate, _
    letterColumnFristRangeArray, letterColumnLastRangeArray, fristRowRangeArray)
    
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
    Set objExcel = CreateExcelInstance(infoLogFile, errorLogFile, nameScript)
    objExcel.Application.DisplayAlerts = False
    Set objWorkbook = objExcel.Workbooks.Add()
    Set objSheet = objWorkBook.WorkSheets(1)

    ' add header
    objSheet.Range(letterSubject & fristRowHeader) = valueSubject
    objSheet.Range(letterColumnSenderEmailAddress & fristRowHeader) = valueSenderEmailAddress
    objSheet.Range(letterColumnTo & fristRowHeader) = valueTo
    objSheet.Range(letterColumnCc & fristRowHeader) = valueCc
    objSheet.Range(letterColumnBcc & fristRowHeader) =  valueBcc
    objSheet.Range(letterColumnBody & fristRowHeader) = valueBody
    objSheet.Range(letterColumnReceivedTime & fristRowHeader) = valueReceivedTime
    objSheet.Range(letterColumnReceivedDate & fristRowHeader) = valueReceivedDate
    objSheet.Range(letterColumnSendTime & fristRowHeader) = valueSendTime
    objSheet.Range(letterColumnSendDate & fristRowHeader) = valueSendDate
    ' insert array in excel file
    objSheet.Range(letterColumnFristRangeArray&fristRowRangeArray&":"&letterColumnLastRangeArray&countAllEmailItem) = arrayExcel

    objWorkbook.SaveAs pathFileSave & nameFileSave

    objWorkbook.Close
    objExcel.Workbooks.Close
    objExcel.Quit

    Set objExcel = Nothing
    Set objSheet = Nothing
End Function ' RunTimeMail

' -----------------------------------------------
' the function to manupulate excel applications
' -----------------------------------------------

On Error Resume Next

Dim nameScript, infoLogFile, errorLogFile, pathFileSave, subject
Dim fristRowHeader, letterSubject, letterColumnSenderEmailAddress, letterColumnTo, letterColumnCc, letterColumnBcc
Dim letterColumnBody, letterColumnReceivedTime, letterColumnReceivedDate, letterColumnSendTime, letterColumnSendDate, valueSubject
Dim valueSenderEmailAddress, valueTo, valueCc, valueBcc, valueBody, valueReceivedTime, valueReceivedDate, valueSendTime, valueSendDate
Dim letterColumnFristRangeArray, letterColumnLastRangeArray, fristRowRangeArray

infoLogFile = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\infoLog.txt"
errorLogFile = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\errorLog.txt"

' ---- ---- the excel information dinamics  ---------------
pathFileSave = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\"
nameFileSave = "Test.xlsx"
letterColumnFristRangeArray = "A"
letterColumnLastRangeArray = "J"
fristRowRangeArray = 2
fristRowHeader = 1
letterSubject = "A"
letterColumnSenderEmailAddress = "B" 
letterColumnTo = "C"
letterColumnCc = "D"
letterColumnBcc = "E"
letterColumnBody = "F"
letterColumnReceivedTime = "G"
letterColumnReceivedDate = "H"
letterColumnSendTime = "I"
letterColumnSendDate = "J"
valueSubject = "Subject"
valueSenderEmailAddress = "SenderEmailAddress"
valueTo = "To"
valueCc = "Cc"
valueBcc = "Bcc"
valueBody = "Body"
valueReceivedTime = "ReceivedTime"
valueReceivedDate = "ReceivedDate"
valueSendTime = "SendTime"
valueSendDate = "SendDate"
' -------------------------------------------

subject = "Weekly Learning Digest"

nameScript = " "&Wscript.ScriptName

Call RunTimeMail(nameScript, infoLogFile, errorLogFile, pathFileSave, nameFileSave, subject, _
    fristRowHeader, letterSubject, letterColumnSenderEmailAddress, letterColumnTo, letterColumnCc, letterColumnBcc, _
    letterColumnBody, letterColumnReceivedTime, letterColumnReceivedDate, letterColumnSendTime, letterColumnSendDate, valueSubject, _
    valueSenderEmailAddress, valueTo, valueCc, valueBcc, valueBody, valueReceivedTime, valueReceivedDate, valueSendTime, valueSendDate, _
    letterColumnFristRangeArray, letterColumnLastRangeArray, fristRowRangeArray)

If Err.Number <> 0 Then
    ' return
    MsgBox Err.Description
    ' WScript.StdOut.WriteLine "value retunr" 
    'WScript.StdOut.WriteLine Err.Description
    Call CloseExcelInstance(infoLogFile, errorLogFile, nameScript)
Else
    MsgBox "successfull"
    'WScript.StdOut.WriteLine "value retunr"
    Call CloseExcelInstance(infoLogFile, errorLogFile, nameScript)
End If
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
        If errorLogFile <> "" Then WriteLog errorLogFile, "INFO", nameScript & " It was not possible to close the Excel instance"
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

Function GetValueNodeXml(nameNode, pathConfigXmlFile, errorLogFile, errorMessage)
    Dim arrayNode, count
    Set objXML = CreateObject("Msxml2.DOMDocument")
    objXML.async = False
    objXML.load(pathConfigXmlFile)

    Set objError = objXML.parseError  
    With objError  
        If .errorCode = 0 Then  
            Set objNodes = objXML.selectNodes(nameNode)
            item = Split(objNodes.item(0).text, "##")
            If UBound(item) >= 2 Then
                ReDim arrayNode(UBound(item) - 2)
                count = 0
                For Each elem In item
                    If Trim(elem) <> "" Then
                        arrayNode(count) = Trim(elem)
                        count = count + 1
                    End If
                Next
                GetValueNodeXml = arrayNode
            Else
                'error
                errorMessage = "The node is null"
                If errorLogFile <> "" Then WriteLog errorLogFile, "ERROR", nameScript & errorMessage
            End If
        Else  
            errorMessage = "XML Document could not be parsed!!!" & vbCrLf &_  
                            "ErrorCode: " & .errorCode & vbCrLf &_  
                            "Line: " & .line & vbCrLf &_  
                            "Reason: " & .reason & vbCrLf &_  
                            "Path: " & .URL 
            If errorLogFile <> "" Then WriteLog errorLogFile, "ERROR", nameScript & errorMessage
        End If  
    End With  
    Set objXML = Nothing  
End Function ' GetValueNodeXml

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
    Const inbox = 6

    ' the connection to Outlook application
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(inbox)

    Set colItems = objFolder.Items
    Set colFilteredItems = colItems.Restrict("[Unread]=true" & " And [Subject] = " & subject) ' reading of unread mails
    MsgBox "colFilteredItems  "&colFilteredItems.count
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

' ----------------------------------------------- jose luis duarte
' the function to manupulate excel applications
' -----------------------------------------------

On Error Resume Next

Dim nameScript, infoLogFile, errorLogFile, pathFileSave, subject
Dim fristRowHeader, letterSubject, letterColumnSenderEmailAddress, letterColumnTo, letterColumnCc, letterColumnBcc
Dim letterColumnBody, letterColumnReceivedTime, letterColumnReceivedDate, letterColumnSendTime, letterColumnSendDate, valueSubject
Dim valueSenderEmailAddress, valueTo, valueCc, valueBcc, valueBody, valueReceivedTime, valueReceivedDate, valueSendTime, valueSendDate
Dim letterColumnFristRangeArray, letterColumnLastRangeArray, fristRowRangeArray


' geting values 
Dim objXML, GroupName, Games, pathConfigExcelXmlFile, pathConfigXmlFile
Dim plot, GameName, GameRating, errorMessage

pathConfigExcelXmlFile = "C:\Users\gbarajas\documents\Curses\Bots\BotEmail\Config\ConfigExcel.xml"
pathConfigXmlFile = "C:\Users\gbarajas\documents\Curses\Bots\BotEmail\Config\Config.xml"
infoLogFile = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\infoLog.txt"
errorLogFile = "C:\Users\gbarajas\Documents\Curses\Bots\BotEmail\errorLog.txt"

' get paths save files
errorMessage = ""
nameNode = "//PathsLocal"
listItemsXml = GetValueNodeXml(nameNode, pathConfigXmlFile, errorLogFile, errorMessage)
If errorMessage <> "" Then
    ' error
    Stop
Else
    pathFileSave = listItemsXml(0)
    nameFileSave = listItemsXml(1)
    Set listItemsXml = Nothing
    Set item = Nothing
End If

' data filter in mail
errorMessage = ""
nameNode = "//Filter"
listItemsXml = GetValueNodeXml(nameNode, pathConfigXmlFile, errorLogFile, errorMessage)
If errorMessage <> "" Then
    ' error
    Stop
Else
    subject = listItemsXml(0)
    Set listItemsXml = Nothing
    Set item = Nothing
End If

' letters excel 
errorMessage = ""
nameNode = "//Columns"
listItemsXml = GetValueNodeXml(nameNode, pathConfigExcelXmlFile, errorLogFile, errorMessage)
If errorMessage <> "" Then
    ' error
    Stop
Else
    letterColumnFristRangeArray = listItemsXml(0)
    letterColumnLastRangeArray = listItemsXml(1)
    letterSubject = listItemsXml(2)
    letterColumnSenderEmailAddress = listItemsXml(3)
    letterColumnTo = listItemsXml(4)
    letterColumnCc = listItemsXml(5)
    letterColumnBcc = listItemsXml(6)
    letterColumnBody = listItemsXml(7)
    letterColumnReceivedTime =listItemsXml(8)
    letterColumnReceivedDate = listItemsXml(9)
    letterColumnSendTime = listItemsXml(10)
    letterColumnSendDate = listItemsXml(11)
    Set listItemsXml = Nothing
    Set item = Nothing
End If

' rows excel
errorMessage = ""
nameNode = "//Rows"
listItemsXml = GetValueNodeXml(nameNode, pathConfigExcelXmlFile, errorLogFile, errorMessage)
If errorMessage <> "" Then
    ' error
    Stop
Else
    fristRowRangeArray = listItemsXml(0)
    fristRowHeader = listItemsXml(1)
    Set listItemsXml = Nothing
    Set item = Nothing
End If

' values excel
errorMessage = ""
nameNode = "//Value"
listItemsXml = GetValueNodeXml(nameNode, pathConfigExcelXmlFile, errorLogFile, errorMessage)
If errorMessage <> "" Then
    ' error
    Stop
Else
    valueSubject = listItemsXml(0)
    valueSenderEmailAddress = listItemsXml(1)
    valueTo = listItemsXml(2)
    valueCc = listItemsXml(3)
    valueBcc = listItemsXml(4)
    valueBody = listItemsXml(5)
    valueReceivedTime =listItemsXml(6)
    valueReceivedDate = listItemsXml(7)
    valueSendTime = listItemsXml(8)
    valueSendDate = listItemsXml(9)
    Set listItemsXml = Nothing
    Set item = Nothing
End If

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
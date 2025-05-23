
Function SendAuthEmail(Recipient, Subject, Body, AttachmentPath)

    Dim objMessage, objConfig, Fields, senderEmail, senderPass
    Dim OutlookApp, MailItem

'########################################################################################
'Input your email credentials:

    senderEmail = "your.user.email@triumph.com"
    senderPass  = "YourEmailPassword"

'########################################################################################
    Set OutlookApp = Nothing
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)
    Set objMessage = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")

    Set Fields = objConfig.Fields
    With Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587  ' port  for TLS
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  ' Enable authentication
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderEmail
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = senderPass
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With

     MailItem.To = Recipient
     MailItem.Subject = Subject
     MailItem.HTMLBody = "<html><body><pre>" & Body & "</pre></body></html>"
      If AttachmentPath <> "" Then
        MailItem.Attachments.Add AttachmentPath
      End If
    MailItem.Send

    Set MailItem = Nothing
    Set OutlookApp = Nothing
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set Fields = Nothing
    WScript.Sleep 350

End Function

Function SendEmail(Recipient, Subject, Body, AttachmentPath)

    Dim OutlookApp, MailItem
    Set OutlookApp = Nothing
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

     MailItem.To = Recipient
     MailItem.Subject = Subject
     MailItem.HTMLBody = "<html><body><pre>" & Body & "</pre></body></html>"
      If AttachmentPath <> "" Then
        MailItem.Attachments.Add AttachmentPath
      End If
    MailItem.Send

    Set MailItem = Nothing
    Set OutlookApp = Nothing
    WScript.Sleep 350

End Function

Sub KillProcess(processName)
    Dim objShell, command
    Set objShell = CreateObject("WScript.Shell")
    command = "taskkill /F /IM " & processName
    objShell.Run command, 0, True
    Set objShell = Nothing
End Sub

Function Cleaner()
    On Error Resume Next 

    KillProcess("excel.exe")
    KillProcess("wscript.exe")

    Dim objWMIService, objProcessList, objProcess
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objProcessList = objWMIService.ExecQuery("Select * from Win32_Process WHERE Name = 'excel.exe'")

    For Each objProcess In objProcessList
        objProcess.Terminate()
    Next

    Set objProcessList = Nothing
    Set objWMIService = Nothing
End Function

Dim inputFolder, outputFolder, folderExist, rootFolder,configFolder, reportFileExist, reportFile
Dim emails, file2
Dim emailSubject, emailBody, emailAttach, inputCustPoFile, objShell, scriptPath, targetPath
Dim xlsxFile, csvFile, sendTo
Dim custPO,order, custPoLastRow
Dim emailName, emailValue, emailLastRow, dictEmails 
Dim userChoice, promptMessage, selectedLanguage, headers2

promptMessage = "Please choose a language for the header:" & vbCrLf & vbCrLf & _
                "1: English" & vbCrLf & _
                "2: German" & vbCrLf & vbCrLf & _
                "Enter 1 or 2:"

userChoice = InputBox(promptMessage, "Language Selection")

Select Case userChoice
    Case "1"
        headers2 = "Name	Order	Customer PO	    Line	Status	Item	Description	Customer Item	Qty Ordered	U/M	Desired date	Unit Price	Net Price	Currency"  

    Case "2"
        headers2 = "Name	Auftr.	Kunden-BS	    Pos.	Status	Teil	Beschreibung	Kundenartikel	Bestellte Mge	ME	Wunschtermin	Preis/ME	Nettopreis	Wahrung"

    Case "" 
        MsgBox "Operation cancelled.", vbExclamation, "Cancelled"
        WScript.Quit
        
    Case Else 
        MsgBox "Invalid input. Please enter only 1 or 2.", vbCritical, "Input Error"
        WScript.Quit
        
End Select

Set shell = CreateObject("WScript.Shell")
Set dvs = CreateObject("Scripting.FileSystemObject")
Set custPO_dict = CreateObject("Scripting.Dictionary")
Set dictFiles = CreateObject("Scripting.Dictionary")
Set dictExcelFiles = CreateObject("Scripting.Dictionary")
rootFolder = shell.CurrentDirectory & "\"
custPO_dict.RemoveAll()
dictFiles.RemoveAll()
dictExcelFiles.RemoveAll()
emailSubject = "Open Orders Customer File"
inputFolder = rootFolder & "Input\"
outputFolder = rootFolder & "Output\"
configFolder = rootFolder & "Config\"
reportFile = rootFolder & "ReportFile_" & year(Now()) & month(Now()) & day(Now()) & ".txt"
inputCustPoFile = inputFolder & "CustomerOrdersExport1.csv"
inputCustOrderLineFile = inputFolder & "CustomerOrderLinesExport1.csv"

    folderExist = dvs.FolderExists(inputFolder)
        If Not folderExist then 
            dvs.CreateFolder(inputFolder)
        Else
            If Not dvs.FileExists(inputCustPoFile) then 
                WScript.Echo "CustomerOrdersExport1.csv not found in the Input folder. Please check."
                WScript.Quit
            End if
           If Not dvs.FileExists(inputFolder & "CustomerOrderLinesExport1.csv") then 
                WScript.Echo "CustomerOrderLinesExport1.csv not found in the Input folder. Please check."
                WScript.Quit
            End if
            
        End if

    folderExist = dvs.FolderExists(outputFolder)
        If Not folderExist then 
            dvs.CreateFolder(outputFolder)
        Else 
            dvs.DeleteFile(outputFolder & "\*.*")
        End if

    folderExist = dvs.FolderExists(configFolder)
        If Not folderExist then 
            dvs.CreateFolder(configFolder)
            WScript.Echo "Check if config files exist in folder Config"
            WScript.Quit
        Else
            If Not dvs.FileExists(configFolder & "SubjectText.txt") then 
                WScript.Echo "File 'SubjectText.txt' is missing from Config folder. Please check."
                WScript.Quit
            End if
           If Not dvs.FileExists(configFolder & "Emails.xlsx") then 
                WScript.Echo "File 'Emails.xlsx' with list of client emails is missing from Config folder. Please check."
                WScript.Quit
            End if
        End if

Set fileForEmailSubject = dvs.OpenTextFile(configFolder & "SubjectText.txt", 1)
emailBody = fileForEmailSubject.ReadAll

'############################################# Mapping table - Populate the dictionary ###############################################
Set dvsCustPoExcel = CreateObject("Excel.Application")
Set dvsCustPoWB = dvsCustPoExcel.Workbooks.Open(inputCustPoFile)
Set dvsCustPoWSheet = dvsCustPoWB.Sheets(1) 
dvsCustPoExcel.Visible = False
dvsCustPoExcel.DisplayAlerts=False
custPoLastRow = dvsCustPoWSheet.UsedRange.Rows.Count

For iOrder = 1 to custPoLastRow
    order = TRIM(dvsCustPoWSheet.Cells(iOrder, 1).Value)
    custPO = TRIM(dvsCustPoWSheet.Cells(iOrder, 2).Value)
      If Not custPO_dict.Exists(order) Then
        custPO_dict.Add order, custPO
      End If
Next

dvsCustPoExcel.Quit
Set dvsCustPoExcel = Nothing
Set dvsCustPoWB = Nothing
Set dvsCustPoWSheet = Nothing

'############################################# Email table - Populate the dictionary #############################################
Set dictEmails = CreateObject("Scripting.Dictionary")
dictEmails.RemoveAll()
Set dvsExcel = CreateObject("Excel.Application")
Set dvsWorkbook = dvsExcel.Workbooks.Open(configFolder & "Emails.xlsx")
Set dvsWorksheet = dvsWorkbook.Sheets(1)
dvsExcel.Visible = True
dvsExcel.DisplayAlerts=False
emailLastRow = dvsWorksheet.UsedRange.Rows.Count

For iEmail = 1 to emailLastRow
    emailName = dvsWorksheet.Cells(iEmail, 1).Value
    emailValue = dvsWorksheet.Cells(iEmail, 2).Value
      If Not dictEmails.Exists(emailName) Then
        dictEmails.Add emailName, emailValue
      End If
Next

dvsExcel.Quit
Set dvsExcel = Nothing
Set dvsWorkbook = Nothing
Set dvsWorksheet = Nothing

'############################################# Main table - CustomerOrderLinesExport1 ############################################
Dim rawDueDate, formattedDueDate

Set file2 = dvs.OpenTextFile(inputFolder & "CustomerOrderLinesExport1.csv", 1, False, -1)

Do While Not file2.AtEndOfStream
    rowLine = file2.ReadLine
    parts = Split(rowLine, vbTab)
     name = parts(15) 
     tcgOrder = TRIM(parts(0))
     line = parts(1)
     orderStatus = parts(2)
     orderItem = parts(3)
     descriptionItem = parts(4)
     customerItem = parts(17)
     qtyOrdered = parts(5)
     um = parts(6)
     unitPrice = parts(7)
     'dueDate = parts(29)
    rawDueDate = parts(29)
    If IsDate(rawDueDate) Then
        Dim dtDueDate
        dtDueDate = CDate(rawDueDate)
        dueDate = Right("0" & Day(dtDueDate), 2) & "." & Right("0" & Month(dtDueDate), 2) & "." & Year(dtDueDate)
    Else
        dueDate = rawDueDate
    End If
     netPrice = parts(19)
     currencyBQ = parts(68)
    If custPO_dict.Exists(tcgOrder) Then
        custPO = custPO_dict(tcgOrder)
    End If
    filteredName = Replace(name,"/"," ")

    If name = "Name" Then
        'no addition handle needed atm
    Else 
        If Not dictFiles.Exists(name) Then
            Set fileOut = dvs.CreateTextFile(outputFolder & "Open Orders Customer " & filteredName & ".tsv", 8)
            fileOut.WriteLine headers2
            dictFiles.Add name, fileOut
        End If
        dictFiles(name).WriteLine name & vbTab & tcgOrder & vbTab & custPo & vbTab & line & vbTab & orderStatus & vbTab & orderItem & vbTab & descriptionItem & vbTab & customerItem & vbTab & qtyOrdered & vbTab & um & vbTab & dueDate & vbTab & unitPrice & vbTab & netPrice & vbTab & currencyBQ
    End If

Loop
file2.Close

'############################################# Close all output files ###########################################################
For Each key In dictFiles.Keys
    dictFiles(key).Close
Next

'############################################# Convert CSV files to XLSX ########################################################
On error resume next
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.DisplayAlerts = False
Set writeToReport = dvs.OpenTextFile(reportFile,2,True)
writeToReport.WriteLine  Time & "- Sending email started!"

For Each key In dictFiles.Keys
    sendTo = ""
    filteredName = Replace(key,"/"," ")
    emailSubject = "Open Orders Customer File " & filteredName
    csvFile = outputFolder & "Open Orders Customer " & filteredName & ".tsv"
    xlsxFile = outputFolder & "Open Orders Customer " & filteredName & ".xlsx"
    'Set objWorkbook = objExcel.Workbooks.Open(csvFile,,,,,,,,vbTab)
    'objWorkbook.SaveAs xlsxFile, 51 
    'objWorkbook.Close False
'test part
Dim lastRow, i
Set objWorkbook = objExcel.Workbooks.Open(csvFile,,,,,,,,vbTab)
Set objSheet = objWorkbook.Sheets(1)
lastRow = objSheet.Cells(objSheet.Rows.Count, 11).End(-4162).Row 
 For i = 2 To lastRow 
    objSheet.Cells(i, 11).NumberFormat = "@"
 Next

objWorkbook.SaveAs xlsxFile, 51
objWorkbook.Close False
Set objSheet = Nothing
'end


    If dictEmails.Exists(key) Then
        sendTo =  dictEmails(key)
        If sendTo = "" Then
            writeToReport.WriteLine Time & "- Missing email from the table for " & key
        Else
            ' no auth
          SendEmail sendTo, emailSubject, emailBody, xlsxFile
            ' with auth 
         'SendAuthEmail sendTo, emailSubject, emailBody, xlsxFile
            writeToReport.WriteLine Time & "- Successful email sent to " & sendTo
        End If

    Else
        writeToReport.WriteLine Time & "- Email " & dictEmails(key) & " not found in email config table."
    End If
Next
On Error Goto 0
writeToReport.WriteLine Time & "- Sending email ended! "

'############################################# Cleanup #########################################################################
objExcel.Quit
Set dvs = Nothing
Set custPO_dict = Nothing
Set dictFiles = Nothing
Set dictExcelFiles = Nothing
Set objExcel = Nothing
Set file2 = Nothing
Set fileOut = Nothing
Set objWorkbook = Nothing

WScript.Echo "Files successfully split and sent."
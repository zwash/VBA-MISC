Attribute VB_Name = "Module5"
Sub MSAG_ECD_Weekly()

Dim answer As Integer
Dim fsoObj As Object, TheDate As String
Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim title As String
Dim titlerow As Integer
Dim j As Long

Cells.AutoFilter

answer = MsgBox("Did you do the Weekly CSP MSAG Change Request first?", vbYesNo + vbQuestion, "Empty Sheet")
If answer = vbYes Then
    'Create a folder named MMDDYY in the Weekly ECD Reports
    TheDate = Format(Date, "MMDDYY")
    enddir = ("\\sea-c-fs1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Weekly ECD Reports\" & TheDate & "\")
    Set fsoObj = CreateObject("Scripting.FileSystemObject")
    With fsoObj
    If Not .FolderExists(enddir) Then
    .CreateFolder (enddir)
    End If
    End With

    'Remove any records that is Status = Completed and "Waiting for TrueNorth to update database"
    For j = Cells(Rows.Count, 3).End(xlUp).Row To 2 Step -1
        If Cells(j, 3).Value = "Completed" Then
            Rows(j).Delete
        ElseIf Cells(j, 3).Value = "Canceled" Then
            Rows(j).Delete
        ElseIf Cells(j, 15).Value = "Waiting for TrueNorth to update database" Then
            Rows(j).Delete
        ElseIf Cells(j, 15).Value = "True North: We recently discovered that all of the OIRIDs were renumbered for this district. We are working with their vendor to correct this issue." Then
            Rows(j).Delete
        End If
    Next j

    'Delete Column O-S
    Range("O:S").Delete
    
    'Find how many Rows and Columns there are
    With ActiveSheet
        Lastrow = .Cells(.Rows.Count, "K").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    'Unhighlight the records except the header
    With Range(Cells(2, 1), Cells(Lastrow, LastCol))
        .Interior.Color = xlNone
        .WrapText = False
    End With
    
    
    vcol = 13
    Set ws = Sheets(1)
    lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
    title = "A1:N1"
    titlerow = ws.Range(title).Cells(1).Row
    icol = LastCol + 1
    
    'Create a new column titles "Unique" and list the unique ECD in the spreadsheet
    ws.Cells(1, icol) = "Unique"
    For i = 3 To lr
        On Error Resume Next
        If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
            ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
        End If
    Next
    
    'Copy the Unique ECD and clear the column
    myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    ws.Columns(icol).Clear
    
    'Create individual spreadsheet containing the ECD record and name the spreadsheet
    For i = 2 To UBound(myarr)
        ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
        If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
            Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
        Else
            Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
        End If
        ws.Range("A" & titlerow & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A1")
        Sheets(myarr(i) & "").Columns.AutoFit
    Next
    
    ws.AutoFilterMode = False
    ws.Activate
        
    'Disable excel alert
    XPath = Application.ActiveWorkbook.Path
	Dim signame As String
	signame = InputBox("Enter name for email signature")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim contacts As Worksheet, password as String, countyName as String
	Set contacts = Worksheets("Contacts")
	Dim sendTo As String, crFile As String
	Dim olApp As Outlook.Application
	Set olApp = CreateObject("Outlook.Application")
    'Save each individual sheet to Weekly ECD Reports folder directory with name ECD_MSAG_CR_Report_MMDDYY, except for initial "Sheet1"
    For Each xWs In ActiveWorkbook.Sheets
        If xWs.Name = "Master" Or xWs.Name = "Legend" Or xWs.Name = "Contacts" Then
            'Do nothing
        Else
			countyName = xWs.Name & ""
			password = Application.VLookup(xWs.Name, contacts.Range("A2:B69"), 2, False)
			sendTo = Application.VLookup(xWs.Name, contacts.Range("A2:C69"), 3, False)
			crFile = countyName & "_MSAG_CR_Report_" & Format(Date, "yyyymmdd")
			xWs.Copy
			Application.ActiveWorkbook.SaveAs FileName:=enddir & xWs.Name & "_MSAG_CR_Report_" & Format(Date, "yyyymmdd") & ".xlsx", password:=password
			Application.ActiveWorkbook.Close False
			Dim olMail As Outlook.MailItem, attachment As Outlook.Attachments
			Set olMail = olApp.CreateItem(olMailItem)
			Set attachment = olMail.Attachments
			olMail.To = sendTo
			olMail.CC = "ng-data-services@comtechtel.com; Anthony.Mobley@comtechtel.com"
			olMail.Subject = "Outstanding MSAG Change Request Report - " & Format(Date, "M/D/YY")
			olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Dear " & Application.WorksheetFunction.Proper(countyName) & " County ECD," & vbNewLine & vbNewLine & "Please see the attachment. This is a follow-up on outstanding MSAG Change Requests that we have not heard back about. " _ 
			& "Based upon our records, we have verified that the address(es) attached (" & crFile & ") do(es) not have a GIS Address Point record.  We ask that you please verify and submit a MSAG ledger, as well as create the new Main Address Point for your next GIS sync update. " _ 
			& "This way your data will be in alignment with the legacy data.  Once you have created the address point(s), please respond back with the new OIRID(s)." & vbNewLine & "If you do not feel this is a valid address for the CSP to use, please provide Comtech the correct OIRID to use and we will follow up with the CSP." _ 
			& vbNewLine & vbNewLine & "The spreadsheet is locked because it contains private customer information, you will receive the password in another email." & vbNewLine & vbNewLine & "If you have any questions, feel free to contact us." _ 
			& vbNewLine & vbNewLine & "Regards," & vbNewLine & vbNewLine & signame & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
			olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
			attachment.Add enddir & xWs.Name & "_MSAG_CR_Report_" & Format(Date, "yyyymmdd") & ".xlsx"
			olMail.Display
			Set olMail = olApp.CreateItem(olMailItem)
			olMail.To = sendTo
			olMail.Subject = "Important Information: Do not discard"
			olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
			olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Hello " & Application.WorksheetFunction.Proper(countyName) & " county," & vbNewLine & vbNewLine & "The code to open the spreadsheet from the previous email is: " & password & _
			vbNewLine & vbNewLine & "Please let us know if you experience any issues with unlocking the spreadsheet." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & _
			signame & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
			olMail.Display
        End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'Open folder directory containing newly saved ECD MSAG CR
    Call Shell("explorer.exe" & " " & enddir, vbNormalFocus)
Else
    'Do nothing
End If
End Sub



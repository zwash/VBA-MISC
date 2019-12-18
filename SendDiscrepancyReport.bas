Sub SendDiscrepancyReport()

Dim reports() As String, master As Workbook, lookup as Workbook, finalRow as Long, myRange As Range, size as Long, theDate As String, fsoObj as Object, today as Double, lookupfinalrow as Long
theDate = Format(Date, "MMDDYY")
enddir = ("\\sea-fs-1\Teams\AQPS_2\AQPS\Washington\Requests\Request Escalation Process\ANI ALI DR Spreadsheets\" & theDate & "\")
Set fsoObj = CreateObject("Scripting.FileSystemObject")
With fsoObj
    If Not .FolderExists(enddir) Then
    .CreateFolder (enddir)
    End If
End With
Application.ScreenUpdating = False
Set master = ActiveWorkbook
finalRow = Cells(Rows.Count, 1).End(xlUp).Row
Set myRange = Range("A2:AA" & finalRow)
size = 0
Set lookup = Workbooks.Open(fileName:= "\\sea-fs-1\Teams\AQPS_2\AQPS\Washington\Requests\Request Escalation Process\ANI ALI DR Spreadsheets\Contacts.xlsx") 'Add lookup file
master.Activate
today = Date
For Each rw in myRange.Rows
	'CHECK THE DATE
	Dim lastUpdated As Double, status As String
	lastUpdated = Cells(rw.Row, 25)
	status = Cells(rw.Row, 2)
	If status = "Referred" AND today > lastUpdated + 14 Then
		Dim contact As String, copyRange As String, password As Variant
		contact = Cells(rw.Row, 18).Value
		If Len(contact) = Null OR Len(contact) = 0 Then
			contact = Cells(rw.Row, 19).Value
		End If
		copyRange = "A" & rw.Row & ":AA" & rw.Row
		lookup.Activate
		lookupfinalrow = Cells(Rows.Count, 1).End(xlUp).Row
		master.Activate
		password = Application.VLookup(contact, lookup.Worksheets("Sheet1").Range("A2:B" & lookupfinalrow), 2, False) 'ADD VLOOKUP RANGE
		If IsError(password) Then 
			Dim newpass as String
			newpass = InputBox("Enter new password for contact: " & contact)
			lookup.Activate
			lookupfinalrow = Cells(Rows.Count, 1).End(xlUp).Offset(1).Row
			Cells(lookupfinalrow, "A").Value = contact
			Cells(lookupfinalrow, "B").Value = newpass
			password = newpass
			master.Activate
		End If
		fileName = "ANI_ALI_DR_Outstanding_" & contact & "_" & theDate & ".xlsx"
		Dim wb As Workbook
		If Len(Dir(enddir & fileName)) = 0 Then
			Set wb = Workbooks.Add("\\sea-fs-1\Teams\AQPS_2\AQPS\Washington\Requests\Request Escalation Process\ANI ALI DR Spreadsheets\ANI_ALI_DR_Outstanding_Template.xlsx" )
			wb.SaveAs fileName:=enddir & fileName, password:=password
			size = size + 1
			ReDim Preserve reports(size)
			reports(size) = contact
		Else
			Set wb = Workbooks.Open(fileName:=enddir & fileName, password:=password)
		End If
		finalRow = Cells(Rows.Count, 1).End(xlUp).Offset(1).Row
		Workbooks(fileName).Worksheets(1).Range("A" & FinalRow & ":AA" & FinalRow).Value = _
		master.Worksheets(1).Range(copyRange).Value
		wb.Close SaveChanges:=True
	End If
Next rw

Application.ScreenUpdating = True
Dim sendTo As String, crFileFull As String
Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim sig as String
sig = InputBox("Enter name for email signature")
For Each escalation In reports 
	If Not escalation = "" Then
		password = Application.VLookup(escalation, lookup.Worksheets("Sheet1").Range("A2:B" & lookupfinalrow), 2, False)
		crFileFull = enddir & "ANI_ALI_DR_Outstanding_" & escalation & "_" & theDate & ".xlsx"
		Dim olMail As Outlook.MailItem, attachment As Outlook.Attachments
		Set olMail = olApp.CreateItem(olMailItem)
		Set attachment = olMail.Attachments
		olMail.To = ""
		olMail.CC = "ng-data-services@comtechtel.com"
		olMail.Subject = "Outstanding ANI/ALI Discrepancy Reports"
		olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Hello " & escalation & "," & vbNewLine & vbNewLine & "You are receiving this email because you have outstanding ANI/ALI Discrepancy Reports referred to your company. " & vbNewLine & vbNewLine & _
		"You should login to the Comtech ALI DBMS, navigate to the ALI > Workflow > ANI/ALI DR page and under the ""ANI/ALI Referred for Action"" queue is where you will find the attached ANI ALI DRs. We have locked the spreadsheet because it contains customer information, we will provide the password in another email." _
		& vbNewLine & vbNewLine & "Once you have reviewed/taken action on the report(s) depending on the comments, please update their status accordingly in the ALI and hit save." & vbNewLine & vbNewLine & "Please let us know if you have any questions." & vbNewLine & vbNewLine & "Thank you, " _
		& vbNewLine & vbNewLine & sig & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
		olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
		attachment.Add crFileFull
		olMail.Display
		Set olMail = olApp.CreateItem(olMailItem)
		olMail.To = ""
		olMail.Subject = "Important Information: Do not discard"
		olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
		olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Hello " & escalation & "," & vbNewLine & vbNewLine & "The code to open the spreadsheet from the previous email is: " & password & vbNewLine & vbNewLine & _
		"Please let us know if you experience any issues with unlocking the spreadsheet." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & sig & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
		olMail.Display
	End if
Next escalation
lookup.Close SaveChanges:= True
End Sub
Sub Send_MSAG_CR_Escalations()

Dim reports() As String, master As Workbook, finalRow as Long, myRange As Range, size as Long, theDate As String, fsoObj as Object, today as Double
theDate = Format(Date, "MMDDYY")
enddir = ("\\sea-fs-1\Teams\AQPS_2\AQPS\Washington\Requests\Request Escalation Process\MSAG CR Spreadsheets\" & theDate & "\")
Set fsoObj = CreateObject("Scripting.FileSystemObject")
With fsoObj
    If Not .FolderExists(enddir) Then
    .CreateFolder (enddir)
    End If
End With
Application.ScreenUpdating = False
Set master = ActiveWorkbook
finalRow = Cells(Rows.Count, 1).End(xlUp).Row
Set myRange = Range("A2:V" & finalRow)
size = 0
today = Date
For Each rw in myRange.Rows
	'CHECK THE DATE
	Dim lastUpdated As Double, status As String
	lastUpdated = Cells(rw.Row, 2)
	status = Cells(rw.Row, 1)
	If status = "Referred" AND today > lastUpdated + 14 Then
		Dim contact As String, copyRange As String
		contact = Cells(rw.Row, 13).Value
		If Len(contact) = Null OR Len(contact) = 0 Then
			contact = Cells(rw.Row, 14).Value
		End If
		copyRange = "A" & rw.Row & ":V" & rw.Row
		fileName = "MSAG_CR_Outstanding_Report_" & contact & "_" & theDate & ".xlsx"
		Dim wb As Workbook
		If Len(Dir(enddir & fileName)) = 0 Then
			Set wb = Workbooks.Add("\\sea-fs-1\Teams\AQPS_2\AQPS\Washington\Requests\Request Escalation Process\MSAG CR Spreadsheets\MSAG_CR_Outstanding_Report_Template.xlsx" )
			wb.SaveAs fileName:=enddir & fileName
			size = size + 1
			ReDim Preserve reports(size)
			reports(size) = contact
		Else
			Set wb = Workbooks.Open(fileName:=enddir & fileName)
		End If
		finalRow = Cells(Rows.Count, 1).End(xlUp).Offset(1).Row
		Workbooks(fileName).Worksheets(1).Range("A" & FinalRow & ":V" & FinalRow).Value = _
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
		crFileFull = enddir & "MSAG_CR_Outstanding_Report_" & escalation & "_" & theDate & ".xlsx"
		Dim olMail As Outlook.MailItem, attachment As Outlook.Attachments
		Set olMail = olApp.CreateItem(olMailItem)
		Set attachment = olMail.Attachments
		olMail.To = ""
		olMail.CC = "ng-data-services@comtechtel.com"
		olMail.Subject = "Outstanding MSAG CR Report"
		olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Hello " & escalation & "," & vbNewLine & vbNewLine & "You are receiving this email because you have outstanding MSAG Change Requests referred to your jurisdiction that are awaiting your action. We have downloaded a report of them and attached them to this email." _
		& vbNewLine & vbNewLine & "You should login to the Comtech ALI DBMS, navigate to the MSAG > Workflow > MSAG CR page and under the ""MSAG CR Referred for Action"" queue is where you will find the MSAG Change Requests." & vbNewLine & vbNewLine & _
		"Once you have reviewed/taken action on the request(s) depending on the comments, please update their status accordingly within the ALI and hit save." & vbNewLine & vbNewLine & " Please let us know if you have any questions." & vbNewLine & vbNewLine & _
		"Thank you," & vbNewLine & vbNewLine & sig & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
		olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
		attachment.Add crFileFull
		olMail.Display
	End if
Next escalation

End Sub
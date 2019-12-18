Sub SendGISConfirmation()

Dim thedate As String, myRange As Range, copyRange As String, countyName As String, fileName As String, FinalRow As Long, wb As Workbook, x as Long, master as Workbook, lookup as Workbook, sendTo as String, lookupfinalrow as Long
Application.ScreenUpdating = False
Set master = ActiveWorkbook
thedate = Format(Date, "MMDDYY")
Set myRange = Selection
x = Selection.Rows(1).Row
countyName = Cells(x, 4)
Set wb = Workbooks.Add("\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\MSAG Related Docs\GIS Change Confirmations\Spreadsheet Template\COUNTY_MSAG_Mismatch_MMDDYY.xlsx")
fileName = countyName & "_MSAG_Mismatch_" & Format(Date, "MMDDYY") & ".xlsx"
wb.SaveAs fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\MSAG Related Docs\GIS Change Confirmations\Spreadsheets to Send\" & fileName
For Each rw In myRange.Rows
	wb.Activate
	FinalRow = Cells(Rows.Count, 1).End(xlUp).Offset(1).Row
	master.Activate
	copyRange = "D" & rw.Row & ":K" & rw.Row
	Workbooks(fileName).Worksheets(1).Range("A" & FinalRow & ":H" & FinalRow).Value = _
    master.Worksheets("GIS Change Confirmation").Range(copyRange).Value
Next rw
wb.Close SaveChanges:=True
Application.ScreenUpdating = True
Set lookup = Workbooks.Open(fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\ECD Related Docs\TN Contact Docs\ECD_Contacts_for_GIS_Change_Confirmations.xlsx")
lookupfinalrow = Cells(Rows.Count, 1).End(xlUp).Row
sendTo = Application.VLookup(countyName, lookup.Worksheets(1).Range("A2:B" & lookupfinalrow), 2, False)
lookup.Close
Dim fileFull As String
Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim sig as String
sig = InputBox("Enter name for email signature")
fileFull = "\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\MSAG Related Docs\GIS Change Confirmations\Spreadsheets to Send\" & fileName
Dim olMail As Outlook.MailItem, attachment As Outlook.Attachments
Set olMail = olApp.CreateItem(olMailItem)
Set attachment = olMail.Attachments
olMail.To = sendTo
olMail.CC = "ng-data-services@comtechtel.com"
olMail.Subject = "GIS Change Confirmation"
olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Dear " & Application.WorksheetFunction.Proper(countyName) & " County ECD," & vbNewLine & vbNewLine & "Comtech has been working with True North to implement a new GIS verification check similar to MSAG change requests. It has come to our attention that you have made changes to your GIS data that we need you to confirm." _ 
& vbNewLine & vbNewLine & "Please see the attached spreadsheet with each change that needs to be validated. Please use column N labelled: Valid? and mark whether the change is Valid or Invalid and send it back with your reply." _
& vbNewLine & vbNewLine & "Please respond to this email at your soonest convenience. If this is a valid change, please also update your legacy MSAG data. To expedite this verification process going forward, please feel free to email us at NG-Data-Services@comtechtel.com with similar address changes in the future, using the same layout as the attached spreadsheet. That information would be used to approve these types of changes, thus no longer requiring us to reach out for further approval." _
& vbNewLine & vbNewLine & "Please let us know if you have any questions or concerns." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & sig & " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
attachment.Add fileFull
olMail.Display
End Sub
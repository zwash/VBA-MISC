Attribute VB_Name = "Module3"
Sub sendCR()

Dim thedate As String, myRange As Range
thedate = Format(Date, "MMDDYY")
Set myRange = Selection

Application.ScreenUpdating = False
Dim contacts As Worksheet, master As Workbook
Set contacts = Worksheets("Contacts")
Set master = ActiveWorkbook
Dim counties() As String, size As Long, FinalRow As Long, copyRange As String
size = 0
For Each rw In myRange.Rows
    Dim countyName As String, fileName As String, password As Variant
    countyName = Cells(rw.Row, 13).Value
    copyRange = "B" & rw.Row & ":N" & rw.Row
    password = Application.VLookup(countyName, contacts.Range("A2:B69"), 2, False)
    fileName = countyName & "_MSAG_CR_" & thedate & ".xlsx"
    Dim wb As Workbook
    If Len(Dir("\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Change Requests\" & fileName)) = 0 Then
        Set wb = Workbooks.Add("\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Email Templates\COUNTY_MSAG_CR_MMDDYY.xlsx")
        wb.SaveAs fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Change Requests\" & fileName, password:=password
        size = size + 1
        ReDim Preserve counties(size)
        counties(size) = countyName
    Else
        Set wb = Workbooks.Open(fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Change Requests\" & fileName, password:=password)
        
    End If
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Offset(1).Row
    Workbooks(fileName).Worksheets(1).Range("A" & FinalRow & ":M" & FinalRow).Value = _
    master.Worksheets("Master").Range(copyRange).Value
    wb.Close SaveChanges:=True
Next rw

Application.ScreenUpdating = True

Dim sendTo As String, crFile As String, crFileFull As String
Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim name as String
name = InputBox("Enter name for email signature")
For Each county In counties
    If Not county = "" Then
    countyName = county
    password = Application.VLookup(countyName, contacts.Range("A2:B69"), 2, False)
    sendTo = Application.VLookup(countyName, contacts.Range("A2:C69"), 3, False)
    crFile = countyName & "_MSAG_CR_" & thedate & ".xlsx"
    crFileFull = "\\sea-fs-1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Change Requests\" & countyName & "_MSAG_CR_" & thedate & ".xlsx"
    Dim olMail As Outlook.MailItem, attachment As Outlook.Attachments
    Set olMail = olApp.CreateItem(olMailItem)
    Set attachment = olMail.Attachments
    olMail.To = sendTo
    olMail.CC = "ng-data-services@comtechtel.com"
    olMail.Subject = "New MSAG Change Request " & Format(Date, "M/D/YY")
    olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Dear " & Application.WorksheetFunction.Proper(countyName) & " County ECD," & vbNewLine & vbNewLine & "Comtech has been working with CSPs in the Jackson region. " _
    & "Because of this the CSPs have been submitting MSAG CR (Change Requests) containing new customer address information to Comtech.  Based upon " _
    & "our records, we have verified that the address(es) attached (" & crFile & ") does not have a GIS Address Point record.  We ask that you please " _
    & "verify and submit a MSAG ledger, as well as create the new Main Address Point for your next GIS sync update. This way your data will be in " _
    & "alignment with the legacy data.  Once you have created the address point, please respond back with the new OIRID." & vbNewLine & "If you do not feel this " _
    & "is a valid address for the CSP to use, please provide Comtech the correct OIRID to use and we will follow up with the CSP." & vbNewLine & vbNewLine & "Please note, " _
    & "the spreadsheet has been locked because it contains sensitive information. You will receive a password to unlock the spreadsheet in another email." & vbNewLine _
    & vbNewLine & "If you have any questions, feel free to contact us." & vbNewLine & "Regards," & vbNewLine & vbNewLine & name & _ 
    " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
    olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
    attachment.Add crFileFull
    olMail.Display
    Set olMail = olApp.CreateItem(olMailItem)
    olMail.To = sendTo
    olMail.Subject = "Important Information: Do not discard"
    olMail.SentOnBehalfOfName = "ng-data-services@comtechtel.com"
    olMail.Body = Format(Date, "M/D/YY") & vbNewLine & vbNewLine & "Hello " & Application.WorksheetFunction.Proper(countyName) & " county," & vbNewLine & vbNewLine & "The code to open the spreadsheet from the previous email is: " & password & _
    vbNewLine & vbNewLine & "Please let us know if you experience any issues with unlocking the spreadsheet." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & name & _
    " | NG911 Data Integrity Group |  Safety & Security Technologies | Comtech Telecommunications Corp. | 2401 Elliott Ave, 2nd floor, Seattle, WA 98121 | p. 206-792-2285 | f. 206-792-2001 | ng-data-services@comtechtel.com |  www.comtech911.com"
    olMail.Display
    End If
Next county

End Sub


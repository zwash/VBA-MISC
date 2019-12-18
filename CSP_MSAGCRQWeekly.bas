Attribute VB_Name = "CSP_MSAGCRQWeekly"
Sub CSP_MSAG_CR_Weekly()

Cells.AutoFilter

'Dim xPath As String
Dim fsoObj As Object, TheDate As String
TheDate = Format(Date, "MMDDYY")
enddir = ("\\sea-c-fs1\Teams\AQPS_2\AQPS\Tennessee\Requests\MSAG Requests\MSAG Change Request\Weekly CSP Reports\" & TheDate & "\")
Set fsoObj = CreateObject("Scripting.FileSystemObject")
With fsoObj
If Not .FolderExists(enddir) Then
.CreateFolder (enddir)
End If
End With

Dim j As Long
    For j = Cells(Rows.Count, 3).End(xlUp).Row To 2 Step -1
        If Cells(j, 3).Value = "Completed" Then
            Rows(j).Delete
        ElseIf Cells(j, 3).Value = "Canceled" Then
            Rows(j).Delete
        'Update CSP Recommendation to say Client Services has begun actively working on reaching out via email and phone if cell color is PURPLE
        ElseIf Cells(j, 3).Interior.Color = RGB(177, 160, 199) Then
            Cells(j, 15).Value = "Client Services has begun actively working on reaching out via email and phone"
            Range(Cells(j, 1), Cells(j, 15)).Interior.Color = RGB(255, 192, 0)
        'Update CSP Recommendation to say Client Services has escalated to the state if cell color is ORANGE
        ElseIf Cells(j, 3).Interior.Color = RGB(255, 192, 0) Then
            Cells(j, 15).Value = "Client Services has escalated to the state"
        'Update CSP Recommendation to say Client Services has escalated to the state if cell color is RED
        ElseIf Cells(j, 3).Interior.Color = RGB(255, 0, 0) Then
            Cells(j, 15).Value = "Client Services has escalated to the state"
            Range(Cells(j, 1), Cells(j, 15)).Interior.Color = RGB(255, 192, 0)
        End If
    Next j
    Range("A:A,N:N,P:S").Delete


Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim title As String
Dim titlerow As Integer
vcol = 3
Set ws = Sheets(1)
lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
title = "A1:M1"
titlerow = ws.Range(title).Cells(1).Row
icol = ws.Columns.Count
ws.Cells(1, icol) = "Unique"
For i = 3 To lr
On Error Resume Next
If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
End If
Next
myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
ws.Columns(icol).Clear
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
    
'xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim contacts As Worksheet, password as String
Set contacts = Worksheets("Contacts")
For Each xWs In ActiveWorkbook.Sheets
    If xWs.Name = "Master" Or xWs.Name = "Legend" Or xWs.Name = "Contacts" Then
            'Do nothing
    Else
    xWs.Copy
	password = Application.VLookup(xWs.Name, contacts.Range("A2:B69"), 2, False)
    Application.ActiveWorkbook.SaveAs FileName:=enddir & xWs.Name & " Weekly_MSAG_Change_Request_" & Format(Date, "yyyymmdd") & ".xlsx", password:=password
    Application.ActiveWorkbook.Close False
    End If
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True

'Open folder directory containing newly saved ECD MSAG CR
Call Shell("explorer.exe" & " " & enddir, vbNormalFocus)

End Sub



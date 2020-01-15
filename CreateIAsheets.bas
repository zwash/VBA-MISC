Sub CreateIAsheets()

Dim myRange as Range
Set myRange = Range("A2:B117")
Application.ScreenUpdating = False
For Each rw in myRange.Rows
	Dim fccid As String, psapName As String, fileName As String
	fccid = Cells(rw.Row, 1).Value
	psapName = Cells(rw.Row, 2).Value
	fileName = Replace("Iowa_ESN_ELT_Managment_" & psapName & ".xlsx", " ", "_")
	Dim wb As Workbook
	If Len(Dir("\\sea-fs-1\Teams\AQPS_2\AQPS\Iowa\Duplicate ESN research\ESN ELT Management forms2\" & fileName)) = 0 Then
		Set wb = Workbooks.Add("\\sea-fs-1\Teams\AQPS_2\AQPS\Iowa\Duplicate ESN research\Comtech_ESN_ELT_Change_Management_Form_Iowa_zw.xlsx")
        wb.SaveAs fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Iowa\Duplicate ESN research\ESN ELT Management forms2\" & fileName
		wb.Worksheets(2).Cells(3, 2).Value = fccid
		wb.Worksheets(2).Cells(4, 2).Value = psapName
		wb.Close SaveChanges:=True
	End If
Next rw
Application.ScreenUpdating = True

End Sub
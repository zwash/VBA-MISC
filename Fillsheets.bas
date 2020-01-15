Sub Fillsheets() 

Application.ScreenUpdating = False
Dim myRange as Range, fileName as String, currentName as String, currentRow as Integer, wb as Workbook, master as Workbook, finalRow
currentName = Cells(5, 5).Value
fileName = Replace("Iowa_ESN_ELT_Managment_" & currentName & ".xlsx", " ", "_")
currentRow = 7
finalRow = 2
Set master = ActiveWorkbook
Set wb = Workbooks.Open(fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Iowa\Duplicate ESN research\ESN ELT Management forms2\" & fileName)
Set myRange = master.Worksheets("All ESNs").Range("A5:M3648")
For Each rw in myRange.Rows
	If StrComp(master.Worksheets("All ESNs").Cells(rw.Row, 5).Value, "?") <> 0 Then
		Dim oldESN as Long, newESN as Long
		If StrComp(master.Worksheets("All ESNs").Cells(rw.Row, 5).Value, currentName) = 0 Then
			'Skip if
		Else
			wb.Close SaveChanges:=True
			currentName = master.Worksheets("All ESNs").Cells(rw.Row, 5).Value
			fileName = Replace("Iowa_ESN_ELT_Managment_" & currentName & ".xlsx", " ", "_")
			Set wb = Workbooks.Open(fileName:="\\sea-fs-1\Teams\AQPS_2\AQPS\Iowa\Duplicate ESN research\ESN ELT Management forms2\" & fileName)
			currentRow = 7
			finalRow = 2
		End If
		oldESN = master.Worksheets("All ESNs").Cells(rw.Row, 1).Value
		newESN = master.Worksheets("All ESNs").Cells(rw.Row, 2).Value
		wb.Worksheets(3).Cells(currentRow, 1).Value = master.Worksheets("All ESNs").Cells(rw.Row,2)
		wb.Worksheets(3).Cells(currentRow, 2).Value = currentName
		wb.Worksheets(3).Cells(currentRow, 3).Value = master.Worksheets("All ESNs").Cells(rw.Row,11)
		wb.Worksheets(3).Cells(currentRow, 4).Value = "=len(C" & currentRow & ")"
		wb.Worksheets(3).Cells(currentRow, 5).Value = master.Worksheets("All ESNs").Cells(rw.Row,12)
		wb.Worksheets(3).Cells(currentRow, 6).Value = "=len(E" & currentRow & ")"
		wb.Worksheets(3).Cells(currentRow, 7).Value = master.Worksheets("All ESNs").Cells(rw.Row,13)
		wb.Worksheets(3).Cells(currentRow, 8).Value = "=len(G" & currentRow & ")"
		wb.Worksheets(3).Cells(currentRow, 9).Value = "=sum(D" & currentRow & ",F" & currentRow & ",H" & currentRow & ")"
		wb.Worksheets(3).Cells(currentRow, 10).Value = master.Worksheets("All ESNs").Cells(rw.Row,3)
		currentRow = currentRow + 1
		If oldESN <> newESN Then
			wb.Worksheets(4).Cells(finalRow,1).Value = oldESN
			wb.Worksheets(4).Cells(finalRow,2).Value = newESN
			finalRow = finalRow + 1
		End IF
	End If
Next rw
Application.ScreenUpdating = True
End Sub
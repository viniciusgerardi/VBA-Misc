'Refresh all Pivot Tables in a Worksheet
Sub refreshWorksheetPivotTable()
	Dim pt As PivotTable
	For Each pt In ActiveSheet.PivotTables
		pt.RefreshTable
	Next pt
End Sub


'Refresh all Pivot Tables in a Workbook
Sub AllWorkbookPivots()
	Dim pt As PivotTable
	Dim ws As Worksheet
	For Each ws In ActiveWorkbook.Worksheets
		For Each pt In ws.PivotTables
			pt.RefreshTable
		Next pt
	Next ws
End Sub